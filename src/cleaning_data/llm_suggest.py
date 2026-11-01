"""
llm_suggest.py
--------------
Step 3 (optional helper). Uses a LOCAL LLM (Ollama + Qwen 3.5 4B) to SUGGEST
names, roles and purposes in remark_review.xlsx. Nothing leaves your laptop.

The LLM never changes your dictionary directly. It only fills grey LLM_* columns
in the review file. You approve a row by typing  Y  in the yellow Accept_LLM
column, then run:  python enrich_remarks.py --apply-review

Flow:
    ollama pull qwen3.5:4b                     # once
    python enrich_remarks.py                   # builds remark_review.xlsx
    python llm_suggest.py                      # adds LLM suggestions (close Excel first)
    (open review file, check suggestions, type Y in Accept_LLM, save, close)
    python enrich_remarks.py --apply-review    # saves approved rows to dictionary

What it works on (default):
    Payees sheet -> rows with Match_Type = "not in dictionary"   (name + role + default purpose)
                    rows with Match_Type = "dictionary" and no role yet (role + default purpose only)
    Notes sheet  -> rows "not in dictionary" or "fuzzy guess - check"     (purpose)
Rows where you already typed something in a yellow cell are skipped.
"""

from __future__ import annotations

import argparse
import difflib
import json
import logging
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd

PROJECT = Path(r"C:\Users\napatil1\Music\project\personal-finance-agent")
DEFAULT_REVIEW = PROJECT / "data" / "processed" / "remark_review.xlsx"
DEFAULT_DICT = PROJECT / "config" / "remark_dictionary.csv"
DEFAULT_MODEL = "llama3.2:3b" #qwen3.5:4b
DEFAULT_HOST = "http://localhost:11434"

# Keep in sync with enrich_remarks.ROLE_SUGGESTIONS (imported if available)
try:
    from enrich_remarks import ROLE_SUGGESTIONS
except Exception:  # pragma: no cover
    ROLE_SUGGESTIONS = ["Family", "Friend", "Self", "Milkman", "Vegetable Vendor", "Grocery Shop",
                        "Restaurant / Food", "Street Food", "Fuel Station", "Online Service",
                        "Shopping", "Travel", "Medical", "Employer", "Investment Platform",
                        "Bank / Credit Card", "Utility / Bill", "Landlord", "Other"]

PAYEE_LLM_COLS = ["LLM_Payee_Name", "LLM_Role", "LLM_Default_Purpose", "LLM_Confidence", "LLM_Reason"]
NOTE_LLM_COLS = ["LLM_Purpose", "LLM_Confidence", "LLM_Reason"]
ACCEPT_COL = "Accept_LLM"

log = logging.getLogger("llm_suggest")

# Fixed examples that show the style we want (taken from your own dictionary + your messages).
NOTE_SHOTS = [
    ("dudh", "Milk"), ("chhach", "Buttermilk"), ("bhaji", "Vegetables"), ("batata", "Potato"),
    ("naral", "Coconut"), ("tamata", "Tomato"), ("aanar", "Pomegranate"), ("bangdi", "Bangles"),
    ("chatri", "Umbrella"), ("riksha", "Rickshaw"), ("fule", "Fuel"), ("haricut", "Haircut"),
    ("vadapaw", "Vada Pav"), ("dinner family", "Dinner"), ("gole coin", "Gold Coin"),
]

SYSTEM_PROMPT = """You clean Indian bank-statement data for one person (the account owner).
Text comes from UPI / NEFT / card transaction remarks.

Facts about the data:
- The bank cuts names to about 10-15 characters (e.g. "OM ENTERPR", "VIKAS GUPT", "Yogesh bha").
- UPI ids look like "name123@okicici", "9876543210@ybl" (phone number), "paytmqr28100505" (shop QR code).
- Notes are written by the owner in English, Hindi or Marathi using English letters, often misspelled
  (dudh = milk, chhach = buttermilk, bhaji = vegetables, ani = and, naral = coconut).

Rules:
1. Never invent facts. If you cannot tell, return an empty string and confidence "low".
2. Payee name: fix CASE and complete obviously cut words only when certain
   (ENTERPR -> Enterprises, MEDICA -> Medical, GEN STOR -> General Store, BHA -> keep as is).
   Never invent a surname or full first name. A phone-number UPI id with no name -> payee_name "".
3. If an item clearly matches one of the EXISTING LABELS given, reuse that label exactly.
4. role must be one of the ALLOWED ROLES or "". Base it on the notes seen and the name
   (e.g. notes mostly "milk" -> Milkman; "petrol pump" in name -> Fuel Station).
5. default_purpose = the one short thing usually paid for (e.g. "Milk"), or "" if notes are mixed/unknown.
6. purpose for a note = short clean English label, Title Case, 1-3 words. Translate Hindi/Marathi words.
7. confidence: "high" only when obvious, "medium" when likely, "low" when guessing.
8. reason: max 12 words.
Return ONLY JSON matching the schema. Keep the same "id" for every item."""


# ==========================================================================
# Ollama client (standard library only)
# ==========================================================================
class Ollama:
    def __init__(self, host: str, model: str, num_ctx: int, timeout: int):
        self.host, self.model, self.num_ctx, self.timeout = host.rstrip("/"), model, num_ctx, timeout
        self.send_think = True       # ask the model not to "think" (faster); dropped if unsupported

    def _post(self, path, payload):
        req = urllib.request.Request(self.host + path, data=json.dumps(payload).encode("utf-8"),
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=self.timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    def check(self):
        try:
            with urllib.request.urlopen(self.host + "/api/tags", timeout=10) as r:
                tags = json.loads(r.read().decode("utf-8"))
        except Exception as e:
            raise SystemExit(f"Cannot reach Ollama at {self.host} ({e}).\n"
                             f"Start it (open the Ollama app, or run: ollama serve) and try again.")
        names = [m.get("name", "") for m in tags.get("models", [])]
        if not any(n == self.model or n.split(":")[0] == self.model and ":" not in self.model for n in names):
            raise SystemExit(f"Model '{self.model}' is not downloaded.\n"
                             f"Run:  ollama pull {self.model}\nModels you have: {names or 'none'}")

    def chat_json(self, system: str, user: str, schema: dict) -> dict:
        payload = {
            "model": self.model,
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
            "format": schema,
            "stream": False,
            "options": {"temperature": 0, "num_ctx": self.num_ctx},
        }
        if self.send_think:
            payload["think"] = False
        try:
            resp = self._post("/api/chat", payload)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "ignore")
            if self.send_think and "think" in body.lower():
                log.info("Model does not accept the 'think' option - retrying without it.")
                self.send_think = False
                return self.chat_json(system, user, schema)
            raise RuntimeError(f"Ollama HTTP {e.code}: {body[:300]}")
        content = resp.get("message", {}).get("content", "")
        content = re.sub(r"<think>.*?</think>", "", content, flags=re.S).strip()
        m = re.search(r"\{.*\}", content, flags=re.S)
        if not m:
            raise ValueError(f"No JSON in model reply: {content[:200]}")
        return json.loads(m.group(0))


# ==========================================================================
# Context from your dictionary
# ==========================================================================
def _compact(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


class DictionaryContext:
    def __init__(self, path: Path):
        d = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        d.columns = [c.strip() for c in d.columns]
        if "Field" not in d:
            d["Field"] = "any"
        d = d[d["Extracted_Info"].str.strip().ne("") & d["Refined_Label"].str.strip().ne("")]
        self.pairs = list(zip(d["Extracted_Info"].str.strip(), d["Refined_Label"].str.strip(), d["Field"]))
        self.keys_c = [_compact(k) for k, _, _ in self.pairs]
        self.labels = sorted(set(l for _, l, _ in self.pairs))

    def similar(self, text: str, n=6):
        """Dictionary rows that look like this text (helps the model reuse your labels)."""
        c = _compact(text)
        if not c:
            return []
        hits = difflib.get_close_matches(c, self.keys_c, n=n, cutoff=0.6)
        out, seen = [], set()
        for h in hits:
            i = self.keys_c.index(h)
            k, l, _ = self.pairs[i]
            if (k, l) not in seen:
                seen.add((k, l))
                out.append(f"{k} -> {l}")
        return out

    def labels_like(self, texts, n=25):
        """Existing labels close to any of the texts (keeps the prompt short)."""
        pool = {}
        lab_c = {_compact(l): l for l in self.labels}
        for t in texts:
            for h in difflib.get_close_matches(_compact(t), list(lab_c), n=5, cutoff=0.5):
                pool[lab_c[h]] = 1
        return list(pool)[:n]


# ==========================================================================
# Prompts + schemas
# ==========================================================================
PAYEE_SCHEMA = {
    "type": "object",
    "properties": {"items": {"type": "array", "items": {
        "type": "object",
        "properties": {"id": {"type": "string"}, "payee_name": {"type": "string"},
                       "role": {"type": "string"}, "default_purpose": {"type": "string"},
                       "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                       "reason": {"type": "string"}},
        "required": ["id", "payee_name", "role", "default_purpose", "confidence", "reason"]}}},
    "required": ["items"],
}
NOTE_SCHEMA = {
    "type": "object",
    "properties": {"items": {"type": "array", "items": {
        "type": "object",
        "properties": {"id": {"type": "string"}, "purpose": {"type": "string"},
                       "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                       "reason": {"type": "string"}},
        "required": ["id", "purpose", "confidence", "reason"]}}},
    "required": ["items"],
}


def _cell(r, c):
    v = r.get(c, "")
    return "" if v is None or str(v).strip().lower() in ("nan", "none") else str(v).strip()


def payee_prompt(batch, ctx: DictionaryContext) -> str:
    lines, texts = [], []
    for _, r in batch.iterrows():
        texts += [_cell(r, "Current_Label"), _cell(r, "Names_Seen")]
        item = {
            "id": _cell(r, "Group_ID"),
            "current_label": _cell(r, "Current_Label"),
            "already_in_dictionary": _cell(r, "Match_Type") == "dictionary",
            "upi_ids_or_names": _cell(r, "Raw_Keys"),
            "names_seen": _cell(r, "Names_Seen"),
            "owner_notes_seen": _cell(r, "Your_Notes_Seen"),
            "txn_count": _cell(r, "Txn_Count"),
            "paid": _cell(r, "Total_Paid"), "received": _cell(r, "Total_Received"),
            "similar_dictionary_rows": ctx.similar(_cell(r, "Current_Label")) + ctx.similar(_cell(r, "Raw_Keys").split("|")[0]),
        }
        lines.append(json.dumps(item, ensure_ascii=False))
    return ("TASK: For each payee give payee_name, role, default_purpose.\n"
            "If already_in_dictionary is true, set payee_name to current_label exactly (do not rename).\n"
            f"ALLOWED ROLES: {', '.join(ROLE_SUGGESTIONS)}\n"
            f"EXISTING LABELS (reuse if same thing): {', '.join(ctx.labels_like(texts))}\n\n"
            "ITEMS (one JSON per line):\n" + "\n".join(lines))


def note_prompt(batch, ctx: DictionaryContext) -> str:
    lines, texts = [], []
    for _, r in batch.iterrows():
        variants = _cell(r, "Note_Variants")
        texts += variants.split("|")
        item = {
            "id": _cell(r, "Group_ID"),
            "note_variants": variants,
            "paid_to": _cell(r, "Common_Payees"),
            "current_guess": _cell(r, "Current_Label"),
            "similar_dictionary_rows": ctx.similar(variants.split("|")[0]),
        }
        lines.append(json.dumps(item, ensure_ascii=False))
    shots = "; ".join(f"{a} -> {b}" for a, b in NOTE_SHOTS)
    return ("TASK: For each note give a clean purpose label.\n"
            f"EXAMPLES from the owner's dictionary: {shots}\n"
            f"EXISTING LABELS (reuse if same thing): {', '.join(ctx.labels_like(texts))}\n\n"
            "ITEMS (one JSON per line):\n" + "\n".join(lines))


# ==========================================================================
# Main logic
# ==========================================================================
YELLOW_PAYEE = ["My_Payee_Name", "Role", "Default_Purpose"]
YELLOW_NOTE = ["My_Purpose"]


def pick_rows(df, kind, scope, redo):
    if df.empty:
        return df
    yellow = YELLOW_PAYEE if kind == "payee" else YELLOW_NOTE
    user_filled = pd.Series(False, index=df.index)
    for c in yellow:
        if c in df:
            user_filled |= df[c].astype(str).str.strip().ne("")
    done = pd.Series(False, index=df.index)
    if "LLM_Confidence" in df and not redo:
        done = df["LLM_Confidence"].astype(str).str.strip().ne("")
    if kind == "payee":
        if scope == "all":
            want = pd.Series(True, index=df.index)
        else:
            no_role = df["Current_Role"].astype(str).str.strip().eq("") if "Current_Role" in df \
                else pd.Series(True, index=df.index)
            want = df["Match_Type"].eq("not in dictionary") | (df["Match_Type"].eq("dictionary") & no_role)
    else:
        want = df["Match_Type"].isin(["not in dictionary", "fuzzy guess - check"]) if scope != "all" \
            else pd.Series(True, index=df.index)
    return df[want & ~user_filled & ~done]


def run_batches(llm, rows, make_prompt, schema, ctx, batch_size, on_result):
    n_ok = n_fail = 0
    ids_left = list(rows["Group_ID"])
    i = 0
    while i < len(rows):
        batch = rows.iloc[i:i + batch_size]
        t0 = time.time()
        try:
            out = llm.chat_json(SYSTEM_PROMPT, make_prompt(batch, ctx), schema)
            got = {str(x.get("id", "")).strip(): x for x in out.get("items", [])}
        except Exception as e:
            log.warning("Batch %d-%d failed (%s)", i + 1, i + len(batch), e)
            got = {}
        missing = [g for g in batch["Group_ID"] if g not in got]
        if missing and len(batch) > 1:                      # retry the missing ones one by one
            for g in missing:
                one = batch[batch["Group_ID"].eq(g)]
                try:
                    o = llm.chat_json(SYSTEM_PROMPT, make_prompt(one, ctx), schema)
                    items = o.get("items", [])
                    if items:
                        items[0]["id"] = g
                        got[g] = items[0]
                except Exception as e:
                    log.warning("  %s failed (%s)", g, e)
        for g in batch["Group_ID"]:
            if g in got:
                on_result(g, got[g])
                n_ok += 1
            else:
                n_fail += 1
        log.info("  %d/%d done (%.1fs for this batch)", min(i + batch_size, len(rows)), len(rows), time.time() - t0)
        i += batch_size
        yield
    log.info("  ok: %d, failed: %d", n_ok, n_fail)


def clean_role(role: str) -> str:
    role = (role or "").strip()
    for r in ROLE_SUGGESTIONS:
        if role.lower() == r.lower():
            return r
    return role   # keep model's own role text; you will see it and decide


def write_sheet(ws, df_ids, results: dict, cols: list, accept_col: str):
    """Add/refresh LLM columns in an openpyxl sheet without touching your formatting."""
    from openpyxl.styles import Alignment, Font, PatternFill
    hdr = [c.value for c in ws[1]]
    for c in cols + [accept_col]:
        if c not in hdr:
            ws.cell(1, len(hdr) + 1).value = c
            hdr.append(c)
    grey = PatternFill("solid", start_color="D9D9D9")
    grey_h = PatternFill("solid", start_color="595959")
    yel = PatternFill("solid", start_color="FFC000")
    yel_c = PatternFill("solid", start_color="FFF2CC")
    conf_fill = {"high": "C6EFCE", "medium": "FFEB9C", "low": "FFC7CE"}
    gid_col = hdr.index("Group_ID") + 1
    for c in cols + [accept_col]:
        j = hdr.index(c) + 1
        h = ws.cell(1, j)
        h.fill = yel if c == accept_col else grey_h
        h.font = Font(bold=True, color="000000" if c == accept_col else "FFFFFF")
        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[h.column_letter].width = 12 if c in (accept_col, "LLM_Confidence") else 26
    for r in range(2, ws.max_row + 1):
        gid = ws.cell(r, gid_col).value
        ws.cell(r, hdr.index(accept_col) + 1).fill = yel_c
        if gid not in results:
            continue
        for c, v in results[gid].items():
            cell = ws.cell(r, hdr.index(c) + 1)
            cell.value = v
            cell.fill = grey
            if c == "LLM_Confidence" and v in conf_fill:
                cell.fill = PatternFill("solid", start_color=conf_fill[v])
    ws.auto_filter.ref = ws.dimensions


def save_wb(wb, path):
    for attempt in range(3):
        try:
            wb.save(path)
            return
        except PermissionError:
            log.error("Cannot save %s - is it open in Excel? Close it (retrying in 10s)...", path.name)
            time.sleep(10)
    raise SystemExit(f"Could not save {path}. Close Excel and run again (finished rows are kept).")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Local-LLM suggestions for remark_review.xlsx")
    ap.add_argument("--review", default=str(DEFAULT_REVIEW))
    ap.add_argument("--dict", default=str(DEFAULT_DICT))
    ap.add_argument("--model", default=DEFAULT_MODEL)
    ap.add_argument("--host", default=DEFAULT_HOST)
    ap.add_argument("--batch", type=int, default=8, help="items per LLM call (smaller = more accurate, slower)")
    ap.add_argument("--limit", type=int, default=0, help="only process the first N rows per sheet (for a test)")
    ap.add_argument("--scope", choices=["unmatched", "all"], default="unmatched")
    ap.add_argument("--sheets", choices=["both", "payees", "notes"], default="both")
    ap.add_argument("--redo", action="store_true", help="re-ask rows that already have a suggestion")
    ap.add_argument("--num-ctx", type=int, default=8192, help="model context length")
    ap.add_argument("--timeout", type=int, default=300)
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

    from openpyxl import load_workbook
    path = Path(args.review)
    if not path.exists():
        raise SystemExit(f"{path} not found. Run enrich_remarks.py first.")

    llm = Ollama(args.host, args.model, args.num_ctx, args.timeout)
    llm.check()
    ctx = DictionaryContext(Path(args.dict))
    log.info("Model: %s | dictionary rows: %d", args.model, len(ctx.pairs))

    wb = load_workbook(path)
    started = time.time()

    if args.sheets in ("both", "payees") and "Payees" in wb.sheetnames:
        df = pd.read_excel(path, sheet_name="Payees", dtype=str, keep_default_na=False)
        rows = pick_rows(df, "payee", args.scope, args.redo)
        rows = rows.head(args.limit) if args.limit else rows
        log.info("Payees to ask: %d", len(rows))
        dict_rows = set(df.loc[df["Match_Type"].eq("dictionary"), "Group_ID"])
        results = {}

        def on_payee(gid, x):
            name = (x.get("payee_name") or "").strip()
            cur = df.loc[df["Group_ID"].eq(gid), "Current_Label"].iloc[0]
            if gid in dict_rows or name.lower() == str(cur).lower():
                name = ""                                   # no rename = keep current label
            results[gid] = {"LLM_Payee_Name": name, "LLM_Role": clean_role(x.get("role")),
                            "LLM_Default_Purpose": (x.get("default_purpose") or "").strip(),
                            "LLM_Confidence": (x.get("confidence") or "").strip().lower(),
                            "LLM_Reason": (x.get("reason") or "").strip()}

        for _ in run_batches(llm, rows, payee_prompt, PAYEE_SCHEMA, ctx, args.batch, on_payee):
            write_sheet(wb["Payees"], None, results, PAYEE_LLM_COLS, ACCEPT_COL)
            save_wb(wb, path)                               # save after every batch (safe to stop)
        write_sheet(wb["Payees"], None, results, PAYEE_LLM_COLS, ACCEPT_COL)

    if args.sheets in ("both", "notes") and "Notes" in wb.sheetnames:
        df = pd.read_excel(path, sheet_name="Notes", dtype=str, keep_default_na=False)
        rows = pick_rows(df, "note", args.scope, args.redo)
        rows = rows.head(args.limit) if args.limit else rows
        log.info("Notes to ask: %d", len(rows))
        results = {}

        def on_note(gid, x):
            results[gid] = {"LLM_Purpose": (x.get("purpose") or "").strip(),
                            "LLM_Confidence": (x.get("confidence") or "").strip().lower(),
                            "LLM_Reason": (x.get("reason") or "").strip()}

        for _ in run_batches(llm, rows, note_prompt, NOTE_SCHEMA, ctx, args.batch, on_note):
            write_sheet(wb["Notes"], None, results, NOTE_LLM_COLS, ACCEPT_COL)
            save_wb(wb, path)
        write_sheet(wb["Notes"], None, results, NOTE_LLM_COLS, ACCEPT_COL)

    save_wb(wb, path)
    log.info("Done in %.0fs. Open %s, check the grey LLM_* columns, type Y in Accept_LLM, "
             "save, then run: python enrich_remarks.py --apply-review", time.time() - started, path.name)


if __name__ == "__main__":
    sys.exit(main())