"""
enrich_remarks.py
-----------------
Step 2 after clean_statements.py.

1. Reads statements_clean.xlsx and breaks every Transaction_Remarks string into
   parts (who, UPI ID, your note, bank, reference no.).
2. Matches those parts against your dictionary (config/remark_dictionary.csv).
3. Writes statements_enriched.xlsx (original columns untouched + new columns).
4. Writes remark_review.xlsx: ONE ROW PER UNIQUE PAYEE / NOTE (not per
   transaction), sorted by how many transactions each one covers.
   You fill only the yellow cells.
5. `--apply-review` copies your yellow-cell edits into the dictionary
   (a backup is taken first) and re-runs everything.

Loop:
    python enrich_remarks.py                 # first run
    (open remark_review.xlsx, fill yellow cells, save, close)
    python enrich_remarks.py --apply-review  # saves edits to dictionary + re-runs
    ...repeat until you are happy

Method = rules + your dictionary. No LLM. Data never leaves your machine.
"""

from __future__ import annotations

import argparse
import difflib
import logging
import re
import shutil
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

# ==========================================================================
# Defaults
# ==========================================================================
PROJECT = Path(r"C:\Users\napatil1\Music\project\personal-finance-agent")
DEFAULT_INPUT = PROJECT / "data" / "processed" / "statements_clean.xlsx"
DEFAULT_OUTPUT = PROJECT / "data" / "processed" / "statements_enriched.xlsx"
DEFAULT_REVIEW = PROJECT / "data" / "processed" / "remark_review.xlsx"
DEFAULT_DICT = PROJECT / "config" / "remark_dictionary.csv"

log = logging.getLogger("enrich_remarks")

# ==========================================================================
# Rules you may want to edit
# ==========================================================================

# Text the bank / UPI app puts in the "note" slot when YOU wrote nothing.
SYSTEM_NOTES = [
    "upi", "upi transaction", "payment from ph", "paid via", "pay to",
    "sent using", "collect", "na", "payment fr", "upi scan", "upiintent",
    "no remarks",
]

# Words ignored when a note is matched word by word (Marathi/Hindi/English joiners, cut-offs).
STOPWORDS = {"ani", "and", "aur", "ka", "ki", "ke", "che", "chi", "cha", "la",
             "for", "the", "of", "to", "n", "w", "&", "+"}

# Generic patterns for UPI IDs (used only if the dictionary has no hit).
VPA_PATTERNS = [
    (re.compile(r"^paytmqr", re.I), "Paytm QR"),
    (re.compile(r"^paytm", re.I), "Paytm"),
    (re.compile(r"^bharatpe", re.I), "BharatPe"),
    (re.compile(r"^(gpay|goog)", re.I), "Google Pay"),
    (re.compile(r"^q[0-9a-z]{6,}@ybl", re.I), "QR Code UPI"),
    (re.compile(r"^(redbus|rb\d{8,})", re.I), "RedBus"),
    (re.compile(r"^\d{10}(-\d+)?(@|$)"), "Person UPI"),
]

# Labels that describe the APP/QR used to pay, not the real person/shop.
AGGREGATOR_LABELS = {
    "paytm", "paytm qr", "bharatpe", "google pay", "phonepe merchant",
    "qr code upi", "person upi", "razorpay", "billdesk", "payment", "upi scan",
    "bajaj pay", "amazon pay", "merger pay", "ph",
}

IFSC_BANK = {"ICIC": "ICICI Bank", "HDFC": "HDFC Bank", "SBIN": "State Bank of India",
             "UTIB": "Axis Bank", "BARB": "Bank of Baroda", "KKBK": "Kotak Bank",
             "YESB": "Yes Bank", "CNRB": "Canara Bank", "PUNB": "PNB"}

# Suggestions shown as a drop-down in the review file (you can still type anything).
ROLE_SUGGESTIONS = [
    "Family", "Friend", "Self", "Milkman", "Vegetable Vendor", "Grocery Shop",
    "Restaurant / Food", "Street Food", "Fuel Station", "Online Service",
    "Shopping", "Travel", "Medical", "Employer", "Investment Platform",
    "Bank / Credit Card", "Utility / Bill", "Landlord", "Other",
]

MIN_PREFIX_KEY = 4
MIN_PREFIX_CAND = 6
FUZZY_CUTOFF = 0.85
DICT_COLUMNS = ["Field", "Extracted_Info", "Refined_Label", "Role", "Default_Purpose", "Source"]

# ==========================================================================
# Text helpers
# ==========================================================================
_ws = re.compile(r"\s+")
_nonalnum = re.compile(r"[^a-z0-9]")


def norm(s) -> str:
    return _ws.sub(" ", str(s).strip().lower())


def compact(s) -> str:
    return _nonalnum.sub("", str(s).lower())


def is_blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "None", "NaT")


def clean_cell(v) -> str:
    return "" if is_blank(v) else str(v).strip()


def is_ref(seg: str) -> bool:
    return bool(re.fullmatch(r"\d{6,}", seg.strip()))


def is_hash(seg: str) -> bool:
    s = seg.replace(" ", "").replace("-", "")
    return len(s) >= 14 and bool(re.search(r"\d", s)) and bool(re.search(r"[A-Za-z]", s))


def looks_like_vpa(seg: str) -> bool:
    s = seg.strip()
    if "@" in s:
        return True
    if " " in s or not s:
        return False
    return bool(re.search(r"[\d.\-_]", s)) or (s.islower() and len(s) > 6)


def is_system_note(note: str) -> bool:
    n = norm(note)
    return any(n == s or n.startswith(s + " ") or (n.startswith(s) and len(s) > 3) for s in SYSTEM_NOTES)


def title(s: str) -> str:
    """'SANDIP RAM' -> 'Sandip Ram'; mixed case keeps short caps: 'BP Petrol Pump' stays."""
    s = str(s).strip()
    if s.isupper():
        return " ".join(w.capitalize() for w in s.split())
    return " ".join(w if (w.isupper() and len(w) <= 3) else w.capitalize() for w in s.split())


# ==========================================================================
# Dictionary
# ==========================================================================
class Matcher:
    """Exact, prefix, word-level and fuzzy look-ups over (key, label) pairs. First key wins."""

    def __init__(self, pairs):
        self.by_norm: dict[str, str] = {}
        self.by_compact: dict[str, str] = {}
        for key, label in pairs:
            nk, ck = norm(key), compact(key)
            if nk and nk not in self.by_norm:
                self.by_norm[nk] = label
            if ck and ck not in self.by_compact:
                self.by_compact[ck] = label
        self.prefix_keys = sorted((k for k in self.by_compact if len(k) >= MIN_PREFIX_KEY),
                                  key=len, reverse=True)
        self.compact_keys = list(self.by_compact)

    def exact(self, text):
        if is_blank(text):
            return None
        return self.by_norm.get(norm(text)) or self.by_compact.get(compact(text))

    def prefix(self, text):
        c = compact(text)
        if len(c) < MIN_PREFIX_KEY:
            return None
        for k in self.prefix_keys:
            if c.startswith(k):
                return self.by_compact[k]
        if len(c) >= MIN_PREFIX_CAND:
            labels = [self.by_compact[k] for k in self.prefix_keys if k.startswith(c)]
            if labels:
                return max(set(labels), key=labels.count)
        return None

    def words(self, text):
        toks = [t for t in re.split(r"[\s,;.]+", norm(text)) if t]
        used = [False] * len(toks)
        found = []
        for n in (3, 2, 1):
            for i in range(len(toks) - n + 1):
                if any(used[i:i + n]):
                    continue
                chunk = toks[i:i + n]
                if n == 1 and (chunk[0] in STOPWORDS or len(chunk[0]) < 3):
                    continue
                lab = self.exact(" ".join(chunk))
                if lab:
                    found.append((i, lab))
                    for j in range(i, i + n):
                        used[j] = True
        labels = []
        for _, lab in sorted(found):
            if lab not in labels:
                labels.append(lab)
        return labels

    def fuzzy(self, text):
        c = compact(text)
        if len(c) < 5:
            return None
        hit = difflib.get_close_matches(c, self.compact_keys, n=1, cutoff=FUZZY_CUTOFF)
        return self.by_compact[hit[0]] if hit else None


class DictionaryStore:
    """
    remark_dictionary.csv with columns:
      Field          any | payee | note   (blank = any)
                     payee = only used to find WHO, note = only used for WHAT
      Extracted_Info text as it appears in the remark (UPI id, name, note)
      Refined_Label  clean name you want to see
      Role           optional, payee only (e.g. Milkman, Family)
      Default_Purpose optional, payee only; used when you wrote no note (e.g. Milk)
      Source         user | added_by_claude | review
    """

    def __init__(self, path: Path):
        self.path = Path(path)
        df = pd.read_csv(self.path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        df.columns = [c.strip() for c in df.columns]
        if not {"Extracted_Info", "Refined_Label"}.issubset(df.columns):
            raise SystemExit("Dictionary needs columns: Extracted_Info, Refined_Label")
        for c in DICT_COLUMNS:
            if c not in df.columns:
                df[c] = ""
        df["Field"] = df["Field"].str.strip().str.lower().replace("", "any")
        df["Source"] = df["Source"].replace("", "user")
        for c in DICT_COLUMNS:
            df[c] = df[c].astype(str).str.strip()
        self.df = df[DICT_COLUMNS + [c for c in df.columns if c not in DICT_COLUMNS]]
        self._build()
        self._warn_conflicts()

    # Priority: review edits first, then field-specific rows, then 'any' rows.
    def _ordered(self, fields):
        d = self.df[self.df["Field"].isin(fields) & self.df["Extracted_Info"].ne("") & self.df["Refined_Label"].ne("")]
        return d.assign(_p1=~d["Source"].isin(["review", "review_llm"]), _p2=d["Field"].eq("any")) \
                .sort_values(["_p1", "_p2"], kind="stable")

    def _build(self):
        pay = self._ordered(("any", "payee"))
        note = self._ordered(("any", "note"))
        self.payee = Matcher(zip(pay["Extracted_Info"], pay["Refined_Label"]))
        self.note = Matcher(zip(note["Extracted_Info"], note["Refined_Label"]))
        self.meta: dict[str, dict] = {}
        for lab, role, dp in zip(pay["Refined_Label"], pay["Role"], pay["Default_Purpose"]):
            m = self.meta.setdefault(lab.lower(), {"Role": "", "Default_Purpose": ""})
            m["Role"] = m["Role"] or role
            m["Default_Purpose"] = m["Default_Purpose"] or dp

    def _warn_conflicts(self):
        d = self.df.assign(_k=self.df["Extracted_Info"].map(norm))
        for (fld, key), g in d.groupby(["Field", "_k"]):
            labs = g["Refined_Label"].unique()
            if key and len(labs) > 1:
                log.warning("Dictionary conflict [%s] '%s': %s -> using '%s'", fld, key, list(labs), labs[0])

    def meta_for(self, label: str) -> dict:
        return self.meta.get(str(label).lower(), {"Role": "", "Default_Purpose": ""})

    # ---- edits (used by --apply-review) --------------------------------
    def upsert(self, fld, key, label, role="", dp="", source="review") -> str:
        key = key.strip()
        mask = self.df["Field"].eq(fld) & self.df["Extracted_Info"].map(norm).eq(norm(key))
        if mask.any():
            i = self.df.index[mask][0]
            before = tuple(self.df.loc[i, ["Refined_Label", "Role", "Default_Purpose"]])
            self.df.loc[i, "Refined_Label"] = label
            if role:
                self.df.loc[i, "Role"] = role
            if dp:
                self.df.loc[i, "Default_Purpose"] = dp
            self.df.loc[i, "Source"] = source
            after = tuple(self.df.loc[i, ["Refined_Label", "Role", "Default_Purpose"]])
            return "updated" if before != after else "same"
        row = {c: "" for c in self.df.columns}
        row.update({"Field": fld, "Extracted_Info": key, "Refined_Label": label,
                    "Role": role, "Default_Purpose": dp, "Source": source})
        self.df = pd.concat([self.df, pd.DataFrame([row])], ignore_index=True)
        return "added"

    def rename_label(self, old, new, fields) -> int:
        mask = self.df["Refined_Label"].eq(old) & self.df["Field"].isin(fields)
        self.df.loc[mask, "Refined_Label"] = new
        return int(mask.sum())

    def set_meta(self, label, role, dp) -> None:
        mask = self.df["Refined_Label"].eq(label) & self.df["Field"].isin(("any", "payee"))
        if role:
            self.df.loc[mask, "Role"] = role
        if dp:
            self.df.loc[mask, "Default_Purpose"] = dp

    def save(self) -> Path:
        backup_dir = self.path.parent / "backup"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = backup_dir / f"{self.path.stem}_{stamp}{self.path.suffix}"
        shutil.copy2(self.path, backup)
        self.df.to_csv(self.path, index=False, encoding="utf-8-sig")
        self._build()
        return backup


# ==========================================================================
# Parsing one remark
# ==========================================================================
@dataclass
class Parsed:
    Channel_Clean: str = ""
    Sub_Channel: str = ""
    Remark_Layout: str = ""
    Party_Name: str = ""
    Party_VPA: str = ""
    Party_Bank: str = ""
    Ref_No: str = ""
    System_Note: str = ""
    User_Note: str = ""
    Fixed_Purpose: str = ""
    extra: dict = field(default_factory=dict)


def detect_channel(r: str) -> str:
    if re.search(r":int\.pd:", r, re.I):
        return "INT"
    m = re.match(r"\s*([A-Za-z]+)", r)
    return m.group(1).upper() if m else ""


def split_note(p: Parsed, note: str) -> None:
    note = (note or "").strip()
    if not note:
        return
    if is_system_note(note):
        p.System_Note = note
    else:
        p.User_Note = note


def parse_upi(r: str, p: Parsed) -> None:
    segs = [s.strip() for s in r.split("/")]
    while segs and segs[-1] == "":
        segs.pop()
    get = lambda i: segs[i] if i < len(segs) else ""
    ref_idx = next((i for i, s in enumerate(segs) if i >= 1 and is_ref(s)), None)

    def set_party(seg):
        if looks_like_vpa(seg):
            p.Party_VPA = seg
        else:
            p.Party_Name = seg

    if ref_idx == 1:                 # 2022-24: UPI/<ref>/<note>/<vpa or name>/<bank>/<hash>
        p.Remark_Layout = "UPI-A"
        p.Ref_No = get(1)
        split_note(p, get(2))
        set_party(get(3))
        p.Party_Bank = get(4)
    elif ref_idx == 4:               # 2025: UPI/<vpa or name>/<note>/<bank>/<ref>/<hash>
        p.Remark_Layout = "UPI-B"
        set_party(get(1))
        split_note(p, get(2))
        p.Party_Bank = get(3)
        p.Ref_No = get(4)
    elif ref_idx == 5:               # 2026: UPI/<name>/<vpa>/<note>/<bank>/<ref>/<hash>
        p.Remark_Layout = "UPI-C"
        p.Party_Name = get(1)
        p.Party_VPA = get(2)
        split_note(p, get(3))
        p.Party_Bank = get(4)
        p.Ref_No = get(5)
    else:
        p.Remark_Layout = "UPI-unknown"
        texts = []
        for s in segs[1:]:
            if not s or is_hash(s):
                continue
            if is_ref(s):
                p.Ref_No = p.Ref_No or s
            elif "@" in s and not p.Party_VPA:
                p.Party_VPA = s
            else:
                texts.append(s)
        if texts:
            p.Party_Name = texts[0]
        if len(texts) > 1:
            split_note(p, texts[1])


def text_segments(segs):
    out = []
    for s in segs:
        s = s.strip()
        if not s or is_ref(s) or is_hash(s) or re.fullmatch(r"\d{1,2}-\d{1,2}-\d{2,4}", s):
            continue
        if re.fullmatch(r"[\dA-Z]{6,}", s) and re.search(r"\d", s):
            continue
        out.append(s)
    return out


def parse_other(r: str, ch: str, p: Parsed) -> None:
    segs = [s.strip() for s in r.split("/")]
    body = segs[1:]

    if ch == "INT":
        p.Remark_Layout = "INT"
        m = re.search(r"Int\.Pd:(.+)$", r, re.I)
        p.Fixed_Purpose = "Interest Credited"
        p.System_Note = m.group(1).strip() if m else ""
        return
    if ch == "NEFT":
        p.Remark_Layout = "NEFT"
        parts = [s.strip() for s in r.split("-")]
        p.Ref_No = parts[1] if len(parts) > 1 else ""
        p.Party_Name = parts[2] if len(parts) > 2 else ""
        ifsc = parts[-1] if parts else ""
        p.Party_Bank = IFSC_BANK.get(ifsc[:4].upper(), ifsc)
        return
    if ch == "RTGS":
        p.Remark_Layout = "RTGS"
        head, _, name = r.partition("/")
        p.Party_Name = name.strip()
        m = re.match(r"RTGS:\s*[A-Z]{4,5}\d+(.*)$", head.strip(), re.I)
        p.Party_Bank = m.group(1).strip() if m else ""
        return
    if ch in ("NFS", "CAM") or "CASH WDL" in r.upper():
        p.Remark_Layout = "ATM"
        p.Fixed_Purpose = "Cash Withdrawal"
        p.System_Note = "CASH WDL"
        return
    if ch == "MMT":
        p.Remark_Layout = "MMT"
        p.Sub_Channel = body[0] if body else ""
        p.Ref_No = body[1] if len(body) > 1 else ""
        split_note(p, body[2] if len(body) > 2 else "")
        p.Party_Name = body[3] if len(body) > 3 else ""
        p.Party_Bank = body[4] if len(body) > 4 else ""
        return
    if ch == "CLG":
        p.Remark_Layout = "CLG"
        p.Party_Name = body[0] if body else ""
        p.Party_Bank = body[1] if len(body) > 1 else ""
        return
    if ch in ("ICSP", "ATD"):
        p.Remark_Layout = ch
        split_note(p, body[0] if body else "")
        return
    if ch == "CMS":
        p.Remark_Layout = "CMS"
        split_note(p, body[0] if body else "")
        p.Party_Name = body[1] if len(body) > 1 else ""
        return
    if ch == "BIL" and body and re.fullmatch(r"[A-Z]{3,5}", body[0] or ""):
        p.Sub_Channel = body[0]
        body = body[1:]

    p.Remark_Layout = ch or "OTHER"
    refs = [s for s in body if is_ref(s)]
    p.Ref_No = refs[0] if refs else ""
    texts = text_segments(body)
    if ch in ("VPS", "VIN", "VSI", "ACH"):
        p.Party_Name = texts[0] if texts else ""
        if len(texts) > 1:
            p.extra["City"] = texts[-1]
    elif ch == "IPS":
        p.Party_Name = texts[-1] if texts else ""
    elif len(texts) >= 2:
        split_note(p, texts[0])
        p.Party_Name = texts[-1]
    elif texts:
        p.Party_Name = texts[0]


def parse_remark(remark: str) -> Parsed:
    p = Parsed()
    r = clean_cell(remark)
    if not r:
        return p
    ch = detect_channel(r)
    p.Channel_Clean = ch
    if ch == "UPI":
        parse_upi(r, p)
    else:
        parse_other(r, ch, p)
    return p


# ==========================================================================
# Matching + description
# ==========================================================================
DICT_MATCHES = {"name_exact", "name_prefix", "vpa_exact", "vpa_prefix"}


def resolve(p: Parsed, store: DictionaryStore) -> dict:
    d_pay, d_note = store.payee, store.note
    out = {"Payee_Key": "", "Payee_Label": "", "Payee_Match": "", "Payee_Role": "",
           "Payment_App": "", "Purpose_Label": "", "Purpose_Match": ""}

    # ---- WHO -----------------------------------------------------------
    name, vpa = p.Party_Name, p.Party_VPA
    handle = vpa.split("@")[0].strip() if vpa else ""
    if re.match(r"bank account", name, re.I) and p.Party_Bank:
        name = p.Party_Bank
    # Stable id used to group the same payee in the review file.
    out["Payee_Key"] = handle or name

    name_lab = name_how = None
    if name:
        name_lab = d_pay.exact(name)
        name_how = "exact" if name_lab else None
        if not name_lab:
            name_lab = d_pay.prefix(name)
            name_how = "prefix" if name_lab else None

    vpa_lab = vpa_how = None
    for cand in (vpa, handle):
        if cand and not vpa_lab:
            vpa_lab = d_pay.exact(cand)
            vpa_how = "exact" if vpa_lab else None
    for cand in (vpa, handle):
        if cand and not vpa_lab:
            vpa_lab = d_pay.prefix(cand)
            vpa_how = "prefix" if vpa_lab else None
    if vpa and not vpa_lab:
        for rx, lab in VPA_PATTERNS:
            if rx.search(vpa):
                vpa_lab, vpa_how = lab, "pattern"
                break

    vpa_is_app = bool(vpa_lab) and vpa_lab.lower() in AGGREGATOR_LABELS
    name_is_app = bool(name_lab) and name_lab.lower() in AGGREGATOR_LABELS

    # An exact hit on the UPI id (e.g. one you added in review) beats a prefix hit on the name.
    if vpa_lab and vpa_how == "exact" and not vpa_is_app and name_how != "exact":
        out["Payee_Label"], out["Payee_Match"] = vpa_lab, "vpa_exact"
    elif name_lab and not name_is_app:
        out["Payee_Label"], out["Payee_Match"] = name_lab, f"name_{name_how}"
    elif vpa_lab and not vpa_is_app:
        out["Payee_Label"], out["Payee_Match"] = vpa_lab, f"vpa_{vpa_how}"
    elif name:
        out["Payee_Label"], out["Payee_Match"] = title(name), "raw_name"
    elif vpa_lab:
        out["Payee_Label"], out["Payee_Match"] = vpa_lab, f"vpa_{vpa_how}"
    elif handle:
        out["Payee_Label"], out["Payee_Match"] = handle, "raw_vpa"

    if vpa_is_app and out["Payee_Label"] != vpa_lab:
        out["Payment_App"] = vpa_lab
    elif name_is_app and out["Payee_Label"] != name_lab:
        out["Payment_App"] = name_lab

    meta = store.meta_for(out["Payee_Label"]) if out["Payee_Label"] else {}
    out["Payee_Role"] = meta.get("Role", "")

    # ---- WHAT ------------------------------------------------------------
    if p.Fixed_Purpose:
        out["Purpose_Label"], out["Purpose_Match"] = p.Fixed_Purpose, "rule"
    else:
        note = p.User_Note or p.System_Note
        if note:
            lab, how = d_note.exact(note), "exact"
            if not lab:
                lab, how = d_note.prefix(note), "prefix"
            if not lab:
                hits = d_note.words(note)
                lab, how = (" & ".join(hits), "words") if hits else (None, None)
            if not lab and p.User_Note:
                lab, how = d_note.fuzzy(note), "fuzzy"
            if lab:
                out["Purpose_Label"], out["Purpose_Match"] = lab, how
            elif p.User_Note:
                out["Purpose_Match"] = "unmatched"
        # No note from you -> use the payee's default purpose (e.g. milkman -> Milk)
        if not out["Purpose_Label"] and not p.User_Note and meta.get("Default_Purpose"):
            out["Purpose_Label"], out["Purpose_Match"] = meta["Default_Purpose"], "payee_default"
    return out


def describe(p: Parsed, res: dict, direction: str) -> str:
    payee, purpose, role = res["Payee_Label"], res["Purpose_Label"], res["Payee_Role"]
    if not purpose and p.User_Note:
        purpose = p.User_Note
    if purpose and payee and purpose.lower() == payee.lower():
        payee = ""
    if p.Fixed_Purpose == "Interest Credited":
        return f"Interest Credited ({p.System_Note})" if p.System_Note else "Interest Credited"
    if p.Fixed_Purpose == "Cash Withdrawal":
        return "Cash Withdrawal (ATM)"
    if payee and role and role.lower() not in payee.lower():
        payee = f"{payee} ({role})"
    verb = {"Debit": "Paid to", "Credit": "Received from"}.get(direction, "Txn with")
    if payee and purpose:
        return f"{verb} {payee} - {purpose}"
    if payee:
        return f"{verb} {payee}"
    if purpose:
        return {"Debit": "Payment - ", "Credit": "Received - "}.get(direction, "") + purpose
    return ""


def direction_of(row) -> str:
    w, dep = row.get("Withdrawal_INR"), row.get("Deposit_INR")
    try:
        if not is_blank(dep) and float(dep) > 0:
            return "Credit"
        if not is_blank(w) and float(w) > 0:
            return "Debit"
    except (TypeError, ValueError):
        pass
    t = clean_cell(row.get("Txn_Type", "")).title()
    return t if t in ("Debit", "Credit") else ""


# ==========================================================================
# Enrichment
# ==========================================================================
MAIN_COLS = ["Txn_Description", "Payee_Label", "Payee_Role", "Purpose_Label", "Payment_App", "Needs_Review"]
TECH_COLS = ["Payee_Key", "Channel_Clean", "Sub_Channel", "Remark_Layout", "Party_Name", "Party_VPA",
             "Party_Bank", "Ref_No", "System_Note", "User_Note", "Payee_Match", "Purpose_Match"]
NEW_COLS = MAIN_COLS + TECH_COLS


def enrich(df: pd.DataFrame, store: DictionaryStore) -> pd.DataFrame:
    if "Transaction_Remarks" not in df.columns:
        raise SystemExit("Input has no 'Transaction_Remarks' column.")
    cache = {}
    rows = []
    for _, row in df.iterrows():
        remark = clean_cell(row["Transaction_Remarks"])
        if remark not in cache:
            p = parse_remark(remark)
            cache[remark] = (p, resolve(p, store))
        p, res = cache[remark]
        rec = {k: v for k, v in asdict(p).items() if k not in ("Fixed_Purpose", "extra")}
        rec.update(res)
        rec["Txn_Description"] = describe(p, res, direction_of(row))
        rec["Needs_Review"] = (res["Purpose_Match"] in ("unmatched", "fuzzy")
                               or (not res["Payee_Label"] and not res["Purpose_Label"]))
        rows.append(rec)
    new = pd.DataFrame(rows, index=df.index)[NEW_COLS]
    base = df.drop(columns=[c for c in NEW_COLS + ["Payee_Role"] if c in df.columns])
    return pd.concat([base, new], axis=1)


# ==========================================================================
# Review workbook (aggregated)
# ==========================================================================
PAYEE_INPUTS = ["My_Payee_Name", "Role", "Default_Purpose"]
NOTE_INPUTS = ["My_Purpose"]


def _amounts(out):
    w = pd.to_numeric(out.get("Withdrawal_INR", pd.Series(0, index=out.index)), errors="coerce").fillna(0)
    d = pd.to_numeric(out.get("Deposit_INR", pd.Series(0, index=out.index)), errors="coerce").fillna(0)
    return w, d


def _top(series: pd.Series, n=3) -> str:
    vc = series[series.ne("")].value_counts().head(n)
    return "; ".join(f"{k} ({v})" for k, v in vc.items())


def build_review(out: pd.DataFrame, store: DictionaryStore):
    w, d = _amounts(out)
    dates = pd.to_datetime(out.get("Transaction_Date"), errors="coerce") if "Transaction_Date" in out else None
    o = out.assign(_w=w, _d=d, _dt=dates)

    # ---------- Payees: group by current label if it came from the dictionary,
    #            otherwise by the raw key (each unknown person = one row) ----------
    pay = o[o["Payee_Key"].ne("") | o["Payee_Label"].ne("")].copy()
    from_dict = pay["Payee_Match"].isin(DICT_MATCHES) & ~pay["Payee_Label"].str.lower().isin(AGGREGATOR_LABELS)
    pay["_grp"] = pay["Payee_Label"].where(from_dict, "KEY::" + pay["Payee_Key"].fillna(""))
    pay["_type"] = from_dict.map({True: "dictionary", False: "not in dictionary"})

    prow = []
    for g, s in pay.groupby("_grp", sort=False):
        label = s["Payee_Label"].iloc[0]
        meta = store.meta_for(label)
        prow.append({
            "Current_Label": label,
            "Match_Type": s["_type"].iloc[0],
            "Txn_Count": len(s),
            "Total_Paid": round(s["_w"].sum(), 2),
            "Total_Received": round(s["_d"].sum(), 2),
            "First_Date": s["_dt"].min() if dates is not None else None,
            "Last_Date": s["_dt"].max() if dates is not None else None,
            "Your_Notes_Seen": _top(s["User_Note"]),
            "Names_Seen": _top(s["Party_Name"]),
            "Current_Role": meta["Role"],
            "Current_Default_Purpose": meta["Default_Purpose"],
            "Example_Remark": s["Transaction_Remarks"].iloc[0],
            "Raw_Keys": " | ".join(sorted(set(k for k in s["Payee_Key"] if k))),
        })
    payees = pd.DataFrame(prow)
    if not payees.empty:
        payees = payees.sort_values(["Txn_Count", "Total_Paid"], ascending=False).reset_index(drop=True)
        payees.insert(0, "Group_ID", [f"P{i:04d}" for i in range(1, len(payees) + 1)])
        payees.insert(4, "Cum_Coverage_%", (payees["Txn_Count"].cumsum() / payees["Txn_Count"].sum() * 100).round(1))
        for c in PAYEE_INPUTS:
            payees[c] = ""

    # ---------- Notes: group by current purpose label, unmatched notes one by one ----------
    nt = o[o["User_Note"].ne("")].copy()
    ok = nt["Purpose_Match"].isin(["exact", "prefix", "words"])
    nt["_grp"] = nt["Purpose_Label"].where(ok, "RAW::" + nt["User_Note"].map(norm))
    nt["_type"] = ok.map({True: "dictionary", False: "not in dictionary"})
    nt.loc[nt["Purpose_Match"].eq("fuzzy"), "_type"] = "fuzzy guess - check"

    nrow = []
    for g, s in nt.groupby("_grp", sort=False):
        nrow.append({
            "Current_Label": s["Purpose_Label"].iloc[0],
            "Match_Type": s["_type"].iloc[0],
            "Txn_Count": len(s),
            "Total_Amount": round((s["_w"] + s["_d"]).sum(), 2),
            "Note_Variants": " | ".join(sorted(set(s["User_Note"]))),
            "Common_Payees": _top(s["Payee_Label"]),
            "Example_Remark": s["Transaction_Remarks"].iloc[0],
        })
    notes = pd.DataFrame(nrow)
    if not notes.empty:
        notes["_o"] = notes["Match_Type"].map({"not in dictionary": 0, "fuzzy guess - check": 1, "dictionary": 2})
        notes = notes.sort_values(["_o", "Txn_Count"], ascending=[True, False]).drop(columns="_o").reset_index(drop=True)
        notes.insert(0, "Group_ID", [f"N{i:04d}" for i in range(1, len(notes) + 1)])
        notes["My_Purpose"] = ""
    return payees, notes


HOW_TO = [
    ("WHAT IS THIS FILE", ""),
    ("", "One row per UNIQUE payee (sheet Payees) or note (sheet Notes) - not per transaction."),
    ("", "Rows are sorted by how many transactions they cover. Fixing the top rows fixes most of your data."),
    ("", "Cum_Coverage_% tells you how much of your payee data is covered up to that row."),
    ("", ""),
    ("WHAT TO FILL (yellow columns only)", ""),
    ("Payees > My_Payee_Name", "Proper name. Example: J Budhle -> Jayesh Budhle. Leave blank to keep Current_Label."),
    ("Payees > Role", "Who this is (drop-down, or type your own). Example: Milkman, Family, Friend."),
    ("Payees > Default_Purpose", "What you usually pay them for. Used ONLY when you wrote no note. Example: Milk."),
    ("Notes > My_Purpose", "Clean meaning of your note. Example: chhach -> Buttermilk."),
    ("", "Blank yellow cell = no change. You do not need to fill every row."),
    ("", ""),
    ("HOW TO SAVE YOUR EDITS", ""),
    ("1", "Fill yellow cells, save this file, close Excel."),
    ("2", "Run:  python enrich_remarks.py --apply-review"),
    ("3", "Your edits go into config/remark_dictionary.csv (old copy kept in config/backup)."),
    ("4", "The statement is re-processed and a fresh review file is written."),
    ("", ""),
    ("COLUMN GUIDE", ""),
    ("Match_Type", "dictionary = already known. 'not in dictionary' = shown as raw name/UPI id, worth naming."),
    ("Raw_Keys", "The UPI ids / names this row covers. Your edit is applied to all of them."),
    ("Names_Seen / Your_Notes_Seen", "Help you recognise who this is."),
    ("", ""),
    ("LLM SUGGESTIONS (optional)", ""),
    ("1", "Close this file. Run:  python llm_suggest.py   (adds grey LLM_* columns + yellow Accept_LLM)"),
    ("2", "Check the grey suggestions. Type Y in Accept_LLM for the rows you agree with."),
    ("", "Tip: filter LLM_Confidence = high, check quickly, then fill Y down."),
    ("3", "Your own yellow cells always win over the LLM. Fill them to correct a suggestion."),
    ("4", "Run:  python enrich_remarks.py --apply-review"),
]


def _style_sheet(ws, input_cols, wide_cols=()):
    from openpyxl.styles import Alignment, Font, PatternFill
    info = PatternFill("solid", start_color="1F3864")
    inp = PatternFill("solid", start_color="FFC000")
    inp_cell = PatternFill("solid", start_color="FFF2CC")
    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        is_in = cell.value in input_cols
        cell.fill = inp if is_in else info
        cell.font = Font(bold=True, color="000000" if is_in else "FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for idx, name in enumerate(headers, start=1):
        letter = ws.cell(1, idx).column_letter
        if name in input_cols:
            ws.column_dimensions[letter].width = 24
            for r in range(2, ws.max_row + 1):
                ws.cell(r, idx).fill = inp_cell
        elif name in wide_cols:
            ws.column_dimensions[letter].width = 45
        else:
            vals = [len(str(ws.cell(r, idx).value or "")) for r in range(1, min(ws.max_row, 300) + 1)]
            ws.column_dimensions[letter].width = min(max(max(vals) + 2, 10), 30)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32


def write_review(path: Path, payees: pd.DataFrame, notes: pd.DataFrame) -> None:
    from openpyxl.styles import Font
    from openpyxl.worksheet.datavalidation import DataValidation
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="DD/MM/YYYY", date_format="DD/MM/YYYY") as xw:
        pd.DataFrame(HOW_TO, columns=["Topic", "Detail"]).to_excel(xw, sheet_name="How_To", index=False)
        payees.to_excel(xw, sheet_name="Payees", index=False)
        notes.to_excel(xw, sheet_name="Notes", index=False)
        ws = xw.book["How_To"]
        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 110
        for row in ws.iter_rows(min_row=2):
            if row[0].value and not row[1].value:
                row[0].font = Font(bold=True, size=12)
        _style_sheet(xw.book["Payees"], PAYEE_INPUTS, wide_cols=("Your_Notes_Seen", "Names_Seen", "Example_Remark"))
        _style_sheet(xw.book["Notes"], NOTE_INPUTS, wide_cols=("Note_Variants", "Common_Payees", "Example_Remark"))

        wp = xw.book["Payees"]
        hdr = [c.value for c in wp[1]]
        if "Role" in hdr and wp.max_row > 1:
            col = wp.cell(1, hdr.index("Role") + 1).column_letter
            dv = DataValidation(type="list", formula1='"' + ",".join(ROLE_SUGGESTIONS) + '"',
                                allow_blank=True, showErrorMessage=False)
            wp.add_data_validation(dv)
            dv.add(f"{col}2:{col}{wp.max_row}")
        xw.book.active = 1


def review_has_edits(path: Path) -> int:
    """How many yellow cells are filled in an existing review file."""
    if not path.exists():
        return 0
    n = 0
    for sheet, cols in (("Payees", PAYEE_INPUTS), ("Notes", NOTE_INPUTS)):
        try:
            df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False)
        except ValueError:
            continue
        for c in cols + [ACCEPT_COL]:
            if c in df:
                n += int(df[c].str.strip().ne("").sum())
    return n


ACCEPT_COL = "Accept_LLM"
ACCEPT_YES = {"y", "yes", "1", "true", "ok", "✓", "x"}


def _accepted(r) -> bool:
    return clean_cell(r.get(ACCEPT_COL)).lower() in ACCEPT_YES


def apply_review(path: Path, store: DictionaryStore) -> dict:
    """Your yellow cells win. If Accept_LLM = Y, LLM_* values fill only the yellow cells you left blank."""
    if not path.exists():
        raise SystemExit(f"Review file not found: {path}")
    summary = {"payee rows used": 0, "note rows used": 0, "of which LLM accepted": 0,
               "added": 0, "updated": 0, "labels renamed": 0}

    pay = pd.read_excel(path, sheet_name="Payees", dtype=str, keep_default_na=False)
    for _, r in pay.iterrows():
        name, role, dp = (clean_cell(r.get(c)) for c in PAYEE_INPUTS)
        used_llm = False
        if _accepted(r):
            llm = (clean_cell(r.get("LLM_Payee_Name")), clean_cell(r.get("LLM_Role")),
                   clean_cell(r.get("LLM_Default_Purpose")))
            used_llm = any(v and not mine for v, mine in zip(llm, (name, role, dp)))
            name, role, dp = name or llm[0], role or llm[1], dp or llm[2]
        if not (name or role or dp):
            continue
        summary["payee rows used"] += 1
        summary["of which LLM accepted"] += int(used_llm)
        source = "review_llm" if used_llm else "review"
        cur = clean_cell(r.get("Current_Label"))
        label = name or cur
        for key in [k.strip() for k in clean_cell(r.get("Raw_Keys")).split("|") if k.strip()]:
            res = store.upsert("payee", key, label, role, dp, source)
            if res in ("added", "updated"):
                summary[res] += 1
        if name and cur and name != cur and r.get("Match_Type") == "dictionary":
            summary["labels renamed"] += store.rename_label(cur, name, ("any", "payee"))
        store.set_meta(label, role, dp)

    notes = pd.read_excel(path, sheet_name="Notes", dtype=str, keep_default_na=False)
    for _, r in notes.iterrows():
        new = clean_cell(r.get("My_Purpose"))
        used_llm = False
        if not new and _accepted(r) and clean_cell(r.get("LLM_Purpose")):
            new, used_llm = clean_cell(r.get("LLM_Purpose")), True
        if not new:
            continue
        summary["note rows used"] += 1
        summary["of which LLM accepted"] += int(used_llm)
        cur = clean_cell(r.get("Current_Label"))
        for v in [x.strip() for x in clean_cell(r.get("Note_Variants")).split("|") if x.strip()]:
            res = store.upsert("note", v, new, source="review_llm" if used_llm else "review")
            if res in ("added", "updated"):
                summary[res] += 1
        if cur and new != cur and r.get("Match_Type") == "dictionary":
            summary["labels renamed"] += store.rename_label(cur, new, ("any", "note"))

    if summary["payee rows used"] or summary["note rows used"]:
        backup = store.save()
        archive = backup.parent / f"remark_review_applied_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(path, archive)
        summary["dictionary backup"] = str(backup)
        summary["review file archived"] = str(archive)
    return summary


# ==========================================================================
# Output workbook
# ==========================================================================
def write_output(out: pd.DataFrame, stats: pd.DataFrame, path: Path) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    # original columns, then the 6 main new columns, then technical columns (collapsed)
    orig = [c for c in out.columns if c not in NEW_COLS]
    out = out[orig + MAIN_COLS + TECH_COLS]
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="DD/MM/YYYY", date_format="DD/MM/YYYY") as xw:
        out.to_excel(xw, sheet_name="Transactions", index=False)
        stats.to_excel(xw, sheet_name="Match_Stats", index=False)
        fills = {"orig": PatternFill("solid", start_color="1F3864"),
                 "main": PatternFill("solid", start_color="375623"),
                 "tech": PatternFill("solid", start_color="7F7F7F")}
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                kind = "main" if cell.value in MAIN_COLS else "tech" if cell.value in TECH_COLS else "orig"
                if ws.title != "Transactions":
                    kind = "orig"
                cell.fill = fills[kind]
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col[:300] if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 55)
        ws = xw.book["Transactions"]
        first = len(orig) + len(MAIN_COLS) + 1
        last = first + len(TECH_COLS) - 1
        ws.column_dimensions.group(ws.cell(1, first).column_letter, ws.cell(1, last).column_letter,
                                   hidden=True, outline_level=1)


def make_stats(out: pd.DataFrame) -> pd.DataFrame:
    has_note = out["User_Note"].ne("")
    return pd.DataFrame([
        ("Total rows", len(out)),
        ("Rows with your own note", int(has_note.sum())),
        ("  note matched (exact/prefix/words)", int(out.loc[has_note, "Purpose_Match"].isin(["exact", "prefix", "words"]).sum())),
        ("  note fuzzy guess (check)", int(out["Purpose_Match"].eq("fuzzy").sum())),
        ("  note not matched", int(out["Purpose_Match"].eq("unmatched").sum())),
        ("Purpose filled from payee default", int(out["Purpose_Match"].eq("payee_default").sum())),
        ("Payee from dictionary", int(out["Payee_Match"].isin(DICT_MATCHES).sum())),
        ("Payee generic (Paytm QR / Person UPI ...)", int(out["Payee_Match"].eq("vpa_pattern").sum())),
        ("Payee shown as raw name / UPI id", int(out["Payee_Match"].isin(["raw_name", "raw_vpa"]).sum())),
        ("Payee with a Role", int(out["Payee_Role"].ne("").sum())),
        ("Rows flagged Needs_Review", int(out["Needs_Review"].sum())),
    ], columns=["Metric", "Value"])


# ==========================================================================
# Main
# ==========================================================================
def main(argv=None):
    ap = argparse.ArgumentParser(description="Extract payee / purpose from Transaction_Remarks")
    ap.add_argument("--input", default=str(DEFAULT_INPUT))
    ap.add_argument("--output", default=str(DEFAULT_OUTPUT))
    ap.add_argument("--review", default=str(DEFAULT_REVIEW), help="Aggregated review workbook")
    ap.add_argument("--dict", default=str(DEFAULT_DICT))
    ap.add_argument("--sheet", default="Transactions")
    ap.add_argument("--apply-review", action="store_true",
                    help="Copy the yellow-cell edits from the review file into the dictionary, then re-run")
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

    review_path = Path(args.review).with_suffix(".xlsx")
    store = DictionaryStore(Path(args.dict))
    log.info("Dictionary loaded: %d rows", len(store.df))

    if args.apply_review:
        s = apply_review(review_path, store)
        for k, v in s.items():
            log.info("Review applied | %-22s %s", k, v)
        if not (s["payee rows used"] or s["note rows used"]):
            log.warning("No yellow cells were filled - nothing changed in the dictionary.")

    df = pd.read_excel(args.input, sheet_name=args.sheet, dtype={"Cheque_Number": str})
    log.info("Rows read: %d", len(df))
    out = enrich(df, store)
    stats = make_stats(out)
    for m, v in stats.itertuples(index=False):
        log.info("%-45s %s", m, v)

    out_path = Path(args.output).with_suffix(".xlsx")
    write_output(out, stats, out_path)
    log.info("Saved -> %s", out_path)

    pending = 0 if args.apply_review else review_has_edits(review_path)
    if pending:
        log.warning("%s has %d filled cell(s) not yet applied - NOT overwritten. "
                    "Run with --apply-review to save them.", review_path.name, pending)
    else:
        payees, notes = build_review(out, store)
        write_review(review_path, payees, notes)
        log.info("Review file -> %s  (%d payee groups, %d note groups)", review_path, len(payees), len(notes))


if __name__ == "__main__":
    sys.exit(main())