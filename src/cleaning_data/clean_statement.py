"""
clean_statements.py
-------------------
Reads every bank statement .xls file in a folder, pulls out only the
transaction table, cleans it, and writes ONE combined .xlsx file.

What it removes from each file:
  * Top block  -> "DETAILED STATEMENT", Search / Advanced Search fields, etc.
  * Tail block -> "Legends Used in Account Statement" and the legend lines.
  * Empty spacer columns (e.g. column A).

What it keeps / adds:
  * Sheet "Transactions": all rows from all files, cleaned and de-duplicated.
  * Sheet "Legends":      the legend codes found in the files (for reference,
                           kept OUT of the transaction table).
  * Sheet "File_Summary": one row per input file (rows read, date range,
                           status) so you can check nothing was missed.

Usage (Windows):
    pip install pandas xlrd openpyxl lxml
    python clean_statements.py
    python clean_statements.py --input "D:\\other\\folder" --output "D:\\out\\all.xlsx"
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# Defaults
# --------------------------------------------------------------------------
DEFAULT_INPUT = r"C:\Users\napatil1\Music\project\personal-finance-agent\data\statements"
DEFAULT_OUTPUT = r"C:\Users\napatil1\Music\project\personal-finance-agent\data\processed\statements_clean.xlsx"

# Header text in the statement -> clean output column name.
# Matching is done on a normalised form (lowercase, no spaces/punctuation),
# so "S No.", "S.No", "SNo" all match "sno".
COLUMN_MAP = {
    "sno": "S_No",
    "valuedate": "Value_Date",
    "transactiondate": "Transaction_Date",
    "chequenumber": "Cheque_Number",
    "transactionremarks": "Transaction_Remarks",
    "withdrawalamountinr": "Withdrawal_INR",
    "depositamountinr": "Deposit_INR",
    "balanceinr": "Balance_INR",
}
# A row is treated as the table header if it contains at least these.
REQUIRED_HEADER_KEYS = {"sno", "transactionremarks"}

AMOUNT_COLS = ["Withdrawal_INR", "Deposit_INR", "Balance_INR"]
DATE_COLS = ["Value_Date", "Transaction_Date"]

# Legend lines look like: "1. INFT - Internal Fund Transfer"
LEGEND_RE = re.compile(r"^\s*\d+\.\s*([A-Z0-9]+)\s*-\s*(.+?)\s*$")

log = logging.getLogger("clean_statements")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def norm(text) -> str:
    """Lowercase and strip everything except letters/digits."""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def read_raw_sheets(path: Path) -> list[pd.DataFrame]:
    """
    Return every sheet as a raw DataFrame (no header, all values as text).

    Many Indian bank portals export ".xls" files that are really HTML tables.
    We check the file's first bytes and pick the right reader.
    """
    with open(path, "rb") as f:
        head = f.read(8)

    if head.startswith(b"\xd0\xcf\x11\xe0"):  # genuine Excel 97-2003 (OLE2/BIFF)
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str, engine="xlrd")
        return list(sheets.values())

    if head.startswith(b"PK"):  # actually an .xlsx renamed to .xls
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str, engine="openpyxl")
        return list(sheets.values())

    # Fall back to HTML disguised as .xls
    text = path.read_bytes().decode("utf-8", errors="ignore")
    if "<table" in text.lower():
        from io import StringIO
        return [t.astype(str).where(t.notna(), None) for t in pd.read_html(StringIO(text), header=None)]

    raise ValueError("Unrecognised file format (not BIFF .xls, .xlsx or HTML).")


def find_header_row(raw: pd.DataFrame) -> int | None:
    """Index of the row that holds the column titles, or None."""
    for idx, row in raw.iterrows():
        keys = {norm(v) for v in row.values}
        if REQUIRED_HEADER_KEYS.issubset(keys):
            return idx
    return None


def extract_legends(raw: pd.DataFrame) -> list[tuple[str, str]]:
    """Pick up 'N. CODE - Description' lines anywhere in the sheet."""
    found = []
    for val in raw.values.ravel():
        if isinstance(val, str):
            m = LEGEND_RE.match(val)
            if m:
                found.append((m.group(1), m.group(2)))
    return found


def to_amount(series: pd.Series) -> pd.Series:
    """'1,23,456.78' / '' / 'nan' -> float (blank becomes 0.0)."""
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("\u00a0", "", regex=False)
        .str.strip()
        .replace({"": None, "nan": None, "None": None, "-": None})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def to_date(series: pd.Series) -> pd.Series:
    """Parse dd/mm/yyyy (and Excel datetime strings) safely."""
    s = series.astype(str).str.strip()
    parsed = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    # Second try for cells xlrd already turned into 'YYYY-MM-DD HH:MM:SS'
    missing = parsed.isna() & s.ne("") & s.ne("nan")
    if missing.any():
        parsed.loc[missing] = pd.to_datetime(s[missing], errors="coerce", dayfirst=True)
    return parsed


def extract_table(raw: pd.DataFrame) -> pd.DataFrame | None:
    """Cut the transaction table out of one raw sheet and clean it."""
    hdr = find_header_row(raw)
    if hdr is None:
        return None

    header_cells = raw.iloc[hdr].tolist()
    keep_cols = {}  # raw column position -> clean name
    for pos, cell in enumerate(header_cells):
        key = norm(cell)
        if key in COLUMN_MAP:
            keep_cols[pos] = COLUMN_MAP[key]

    body = raw.iloc[hdr + 1 :, list(keep_cols.keys())].copy()
    body.columns = list(keep_cols.values())

    # Real transaction rows have a numeric serial number.
    # This drops the "Legends Used..." block, blank rows, footers, and any
    # repeated header rows (some exports repeat the header on each page).
    sno = pd.to_numeric(body["S_No"].astype(str).str.strip(), errors="coerce")
    body = body[sno.notna()].copy()
    body["S_No"] = sno[sno.notna()].astype(int)

    # Clean text: remove line breaks Excel uses for wrapping, squeeze spaces.
    for col in ("Transaction_Remarks", "Cheque_Number"):
        if col in body:
            body[col] = (
                body[col]
                .astype(str)
                .str.replace(r"[\r\n]+", "", regex=True)
                .str.replace(r"\s{2,}", " ", regex=True)
                .str.strip()
                .replace({"nan": "", "None": ""})
            )

    for col in DATE_COLS:
        if col in body:
            body[col] = to_date(body[col])

    for col in AMOUNT_COLS:
        if col in body:
            body[col] = to_amount(body[col])
    for col in ("Withdrawal_INR", "Deposit_INR"):
        if col in body:
            body[col] = body[col].fillna(0.0)

    return body.reset_index(drop=True)


def add_derived_columns(df: pd.DataFrame, legend_map: dict[str, str]) -> pd.DataFrame:
    """Add a few handy columns for later analysis (no data is changed)."""
    remarks = df["Transaction_Remarks"].fillna("")
    # Channel = first token of the remark, e.g. 'UPI', 'NEFT', 'ICSP', 'BIL'
    df["Txn_Channel"] = remarks.str.split(r"[/\-\s]", n=1, regex=True).str[0].str.upper()
    df["Channel_Meaning"] = df["Txn_Channel"].map(legend_map).fillna("")
    df["Txn_Type"] = df.apply(
        lambda r: "Credit" if r["Deposit_INR"] > 0 else ("Debit" if r["Withdrawal_INR"] > 0 else "Zero"),
        axis=1,
    )
    df["Net_Amount_INR"] = df["Deposit_INR"] - df["Withdrawal_INR"]
    df["Year_Month"] = df["Transaction_Date"].dt.to_period("M").astype(str)
    return df


def write_output(df: pd.DataFrame, legends: pd.DataFrame, summary: pd.DataFrame, out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="DD/MM/YYYY", date_format="DD/MM/YYYY") as xw:
        df.to_excel(xw, sheet_name="Transactions", index=False)
        legends.to_excel(xw, sheet_name="Legends", index=False)
        summary.to_excel(xw, sheet_name="File_Summary", index=False)

        from openpyxl.styles import Font, PatternFill, Alignment

        head_fill = PatternFill("solid", start_color="1F3864")
        head_font = Font(bold=True, color="FFFFFF")
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = head_fill
                cell.font = head_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col_cells in ws.columns:
                letter = col_cells[0].column_letter
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[:500])
                ws.column_dimensions[letter].width = min(max(width + 2, 10), 60)

        ws = xw.book["Transactions"]
        amount_names = set(AMOUNT_COLS) | {"Net_Amount_INR"}
        for cell in ws[1]:
            if cell.value in amount_names:
                for c in ws.iter_cols(min_col=cell.column, max_col=cell.column, min_row=2):
                    for x in c:
                        x.number_format = "#,##0.00"


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def run(input_dir: Path, output_file: Path, dedupe: bool = True) -> pd.DataFrame:
    files = sorted(p for p in input_dir.iterdir() if p.suffix.lower() == ".xls" and not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"No .xls files found in: {input_dir}")

    log.info("Found %d .xls file(s) in %s", len(files), input_dir)
    frames, legend_rows, summary = [], [], []

    for path in files:
        try:
            sheets = read_raw_sheets(path)
            tables = []
            for raw in sheets:
                legend_rows.extend(extract_legends(raw))
                t = extract_table(raw)
                if t is not None and not t.empty:
                    tables.append(t)

            if not tables:
                log.warning("  %-40s -> no transaction table found", path.name)
                summary.append({"File": path.name, "Rows": 0, "From": None, "To": None,
                                "Status": "No table found"})
                continue

            t = pd.concat(tables, ignore_index=True)
            t.insert(0, "Source_File", path.name)
            frames.append(t)
            summary.append({
                "File": path.name,
                "Rows": len(t),
                "From": t["Transaction_Date"].min(),
                "To": t["Transaction_Date"].max(),
                "Status": "OK",
            })
            log.info("  %-40s -> %d rows", path.name, len(t))
        except Exception as exc:  # keep going with the other files
            log.error("  %-40s -> FAILED: %s", path.name, exc)
            summary.append({"File": path.name, "Rows": 0, "From": None, "To": None,
                            "Status": f"Error: {exc}"})

    if not frames:
        raise SystemExit("No transactions extracted from any file. Check File_Summary / logs.")

    df = pd.concat(frames, ignore_index=True)
    rows_before = len(df)

    # Overlapping statement periods produce the same transaction twice.
    # A transaction is a duplicate if dates, remarks, amounts AND running
    # balance all match (balance makes false matches very unlikely).
    if dedupe:
        key = ["Value_Date", "Transaction_Date", "Transaction_Remarks",
               "Withdrawal_INR", "Deposit_INR", "Balance_INR"]
        df = df.drop_duplicates(subset=[c for c in key if c in df.columns], keep="first")
    dupes = rows_before - len(df)

    # Stable sort keeps the bank's original order inside the same day.
    df = df.sort_values(["Transaction_Date", "Source_File", "S_No"], kind="stable").reset_index(drop=True)
    df.insert(0, "Row_ID", range(1, len(df) + 1))

    legends = (pd.DataFrame(legend_rows, columns=["Code", "Meaning"])
               .drop_duplicates().sort_values("Code").reset_index(drop=True))
    df = add_derived_columns(df, dict(zip(legends["Code"], legends["Meaning"])))

    summary_df = pd.DataFrame(summary)
    summary_df.loc[len(summary_df)] = {
        "File": "TOTAL (after de-dup)", "Rows": len(df),
        "From": df["Transaction_Date"].min(), "To": df["Transaction_Date"].max(),
        "Status": f"{dupes} duplicate row(s) removed",
    }

    write_output(df, legends, summary_df, output_file)
    log.info("Rows read: %d | duplicates removed: %d | rows written: %d", rows_before, dupes, len(df))
    log.info("Saved -> %s", output_file)
    return df


def main(argv=None):
    ap = argparse.ArgumentParser(description="Combine and clean bank statement .xls files into one .xlsx")
    ap.add_argument("--input", default=DEFAULT_INPUT, help="Folder with .xls statements")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="Output .xlsx file path")
    ap.add_argument("--no-dedupe", action="store_true", help="Keep duplicate rows from overlapping files")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
    out = Path(args.output)
    if out.suffix.lower() != ".xlsx":
        out = out.with_suffix(".xlsx")
    run(Path(args.input), out, dedupe=not args.no_dedupe)


if __name__ == "__main__":
    sys.exit(main())