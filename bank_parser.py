"""
bank_parser.py
--------------
Parses a Bank Leumi .xls export (actually HTML-disguised-as-XLS).
Returns a list of transaction dicts, only incoming credits from tenants.

Each transaction dict:
{
    "date":        "01/04/2026",
    "ref":         "4873628",
    "description": "הוראת קבע",
    "amount":      350.0,
    "raw_detail":  "העברה מאת: אליזבט בינשטוק, 10-891-003641132 תשלום בהוראת קבע",
    "sender_name": "אליזבט בינשטוק",   # extracted from raw_detail, may be truncated
    "apt_hint":    43,                   # int if "דירה XX" found in raw_detail, else None
    "payment_month": "04/2026",          # derived from text first, then transaction date
    "petty_cash":  False,               # True when memo contains "קופה קטנה"
}
"""

import re
import pandas as pd


# ── Regex patterns ────────────────────────────────────────────────────────────
_RE_SENDER       = re.compile(r"העברה מאת:\s*(.+?)\s+\d{2}-\d{3}-\d+", re.UNICODE)
_RE_APT          = re.compile(r"דירה\s+(\d+)", re.UNICODE)
# FIX Bug 3: catch reversed "38תשלום קופה..." / "55קופה..." patterns from Bank Leumi
_RE_APT_REVERSED = re.compile(r"(\d{1,3})\s*(?:תשלום|קופה)", re.UNICODE)
_RE_CHECKIN      = re.compile(r"הפקדת שיק", re.UNICODE)

# ── Hebrew month map (used for text-based month detection) ────────────────────
_HEB_MONTHS = {
    "ינואר":  "01", "פברואר": "02", "מרץ":    "03", "אפריל":  "04",
    "מאי":    "05", "יוני":   "06", "יולי":   "07", "אוגוסט": "08",
    "ספטמבר": "09", "אוקטובר":"10", "נובמבר": "11", "דצמבר":  "12",
}


def _month_from_text(text: str, year: str) -> str | None:
    """
    Scan the MEMO portion of a bank detail line for a Hebrew month name.
    Returns 'MM/YYYY' for the first match, or None.

    Only the text AFTER the bank account number (XX-XXX-XXXXXXXXX) is
    searched, to avoid false-positives from sender names that happen to
    contain a month word (e.g. "ברייב יולי" — יולי = July AND a surname).
    """
    if not isinstance(text, str):
        return None
    m = re.search(r"\d{2}-\d{3}-\d+", text)
    search_in = text[m.end():].strip() if m else text
    for heb, mm in _HEB_MONTHS.items():
        if heb in search_in:
            return f"{mm}/{year}"
    return None


def _extract_payment_month(date_str: str) -> str | None:
    """Convert a bank date string "DD/MM/YYYY" to a payment month "MM/YYYY"."""
    if not isinstance(date_str, str):
        return None
    parts = date_str.strip().split("/")
    if len(parts) == 3:
        try:
            return f"{int(parts[1]):02d}/{parts[2]}"
        except ValueError:
            pass
    return None


def extract_comment_text(raw_detail: str) -> str:
    """
    Return the meaningful part of a bank raw_detail for use as a comment.
    Strips everything up to and including the bank account number.
    """
    if not isinstance(raw_detail, str):
        return ""
    m = re.search(r"\d{2}-\d{3}-\d+", raw_detail)
    if m:
        after = raw_detail[m.end():].strip().lstrip(",").strip()
        return after if after else raw_detail.strip()
    return raw_detail.strip()


def _extract_sender(detail: str) -> str | None:
    """Pull the sender name from 'העברה מאת: NAME BANK-BRANCH-ACC ...'"""
    if not isinstance(detail, str):
        return None
    m = _RE_SENDER.search(detail)
    if m:
        name = m.group(1).strip().rstrip(",")
        return name
    return None


def _extract_apt_hint(detail: str) -> int | None:
    """
    Pull apartment number from raw_detail. Handles two Bank Leumi patterns:
      - Forward:  "דירה 37"              (most common)
      - Reversed: "38תשלום קופה קטנה דירה"  (Bank Leumi quirk, digits before keyword)
                  "55קופה קטנה עבור דירה"

    For the reversed pattern, only searches the memo portion (after the
    bank account number XX-XXX-XXXXXXXXX) to avoid false positives.
    """
    if not isinstance(detail, str):
        return None
    # Forward pattern first — most reliable
    m = _RE_APT.search(detail)
    if m:
        return int(m.group(1))
    # Reversed pattern: search only after the account number
    acc = re.search(r"\d{2}-\d{3}-\d+", detail)
    search_in = detail[acc.end():].strip() if acc else detail
    m2 = _RE_APT_REVERSED.search(search_in)
    if m2:
        n = int(m2.group(1))
        if 1 <= n <= 60:   # sanity-check: valid apartment range for this building
            return n
    return None


def _is_incoming(row) -> bool:
    """
    Keep a row if it has a credit (זכות > 0) AND is not an explicit
    outgoing transfer ("העברה אל:").
    """
    try:
        credit = float(str(row["בזכות"]).replace(",", "")) if pd.notna(row["בזכות"]) else 0.0
    except (ValueError, TypeError):
        credit = 0.0

    if credit <= 0:
        return False

    detail = str(row.get("תאור מורחב", ""))
    if "העברה אל:" in detail:
        return False

    return True


def _resolve_payment_month(raw_detail: str, date_str: str) -> str | None:
    """
    Return the best payment month. Hebrew month name in memo takes priority.
    """
    month_from_date = _extract_payment_month(date_str)
    year = month_from_date.split("/")[1] if month_from_date else "2026"
    return _month_from_text(raw_detail, year) or month_from_date


def parse_bank_file(path: str) -> list[dict]:
    """
    Main entry point.
    Returns a list of cleaned transaction dicts ready for matching.
    """
    tables = pd.read_html(path, encoding="utf-8")

    tx_df = None
    for df in tables:
        if df.shape[0] > 2 and str(df.iloc[1, 0]).strip() == "תאריך":
            tx_df = df
            break

    if tx_df is None:
        raise ValueError("Could not locate the transactions table in the file.")

    tx_df.columns = tx_df.iloc[1].tolist()
    tx_df = tx_df.iloc[2:].reset_index(drop=True)

    col_map = {
        "תאריך":       "date",
        "תאריך ערך":   "value_date",
        "תיאור":       "description",
        "אסמכתא":      "ref",
        "בחובה":       "debit",
        "בזכות":       "credit",
        'היתרה בש"ח':  "balance",
        "תאור מורחב":  "raw_detail",
        "הערה":        "note",
    }
    tx_df = tx_df.rename(columns={k: v for k, v in col_map.items() if k in tx_df.columns})

    transactions = []
    for _, row in tx_df.iterrows():
        if pd.isna(row.get("date")):
            continue

        raw_row = {
            "תיאור":       row.get("description", ""),
            "תאור מורחב": row.get("raw_detail", ""),
            "בזכות":       row.get("credit", 0),
            "בחובה":       row.get("debit", 0),
        }

        if not _is_incoming(raw_row):
            continue

        try:
            amount = float(str(row.get("credit", "0")).replace(",", ""))
        except (ValueError, TypeError):
            amount = 0.0

        raw_detail  = str(row.get("raw_detail", "")) if pd.notna(row.get("raw_detail")) else ""
        date_s      = str(row.get("date", "")).strip()
        sender_name = _extract_sender(raw_detail)
        apt_hint    = _extract_apt_hint(raw_detail)
        payment_month = _resolve_payment_month(raw_detail, date_s)

        # FIX Bug 2: detect petty cash payments by keyword in memo
        is_petty_cash = "קופה קטנה" in raw_detail

        transactions.append({
            "date":          date_s,
            "ref":           str(row.get("ref", "")).strip(),
            "description":   str(row.get("description", "")).strip(),
            "amount":        amount,
            "raw_detail":    raw_detail,
            "sender_name":   sender_name,
            "apt_hint":      apt_hint,
            "payment_month": payment_month,
            "petty_cash":    is_petty_cash,   # NEW
        })

    return transactions


# ── Debit (expense) parsing ───────────────────────────────────────────────────

def _load_category_rules(csv_path: str = "categories.csv") -> list[dict]:
    import csv
    from pathlib import Path

    p = Path(csv_path)
    if not p.exists():
        p = Path(__file__).parent / "categories.csv"
    if not p.exists():
        raise FileNotFoundError(
            f"categories.csv not found at '{csv_path}' or next to bank_parser.py."
        )

    rules = []
    with open(p, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            col_raw = str(row.get("Column_To_Search", "")).strip()
            rules.append({
                "keyword":     str(row.get("Keyword", "")).strip(),
                "search_col":  "ext_desc" if col_raw == "Original_Ext_Desc" else "desc",
                "category":    str(row.get("Category", "")).strip(),
                "entity_name": str(row.get("Entity_Name", "")).strip(),
            })
    return rules


def _match_category(desc: str, ext_desc: str, rules: list[dict]) -> tuple[str | None, str | None]:
    for rule in rules:
        kw = rule["keyword"]
        if not kw:
            continue
        haystack = ext_desc if rule["search_col"] == "ext_desc" else desc
        if kw in (haystack or ""):
            return rule["category"], rule["entity_name"]
    return None, None


def parse_debit_file(path: str, categories_csv: str = "categories.csv") -> list[dict]:
    """Parse the same bank file but extract DEBIT (חובה) rows only."""
    rules = _load_category_rules(categories_csv)

    tables = pd.read_html(path, encoding="utf-8")
    tx_df = None
    for df in tables:
        if df.shape[0] > 2 and str(df.iloc[1, 0]).strip() == "תאריך":
            tx_df = df
            break

    if tx_df is None:
        raise ValueError("Could not locate the transactions table in the file.")

    tx_df.columns = tx_df.iloc[1].tolist()
    tx_df = tx_df.iloc[2:].reset_index(drop=True)

    col_map = {
        "תאריך":       "date",
        "תאריך ערך":   "value_date",
        "תיאור":       "description",
        "אסמכתא":      "ref",
        "בחובה":       "debit",
        "בזכות":       "credit",
        'היתרה בש"ח':  "balance",
        "תאור מורחב":  "raw_detail",
        "הערה":        "note",
    }
    tx_df = tx_df.rename(columns={k: v for k, v in col_map.items() if k in tx_df.columns})

    debits = []
    for _, row in tx_df.iterrows():
        if pd.isna(row.get("date")):
            continue

        try:
            amount = float(str(row.get("debit", "0")).replace(",", ""))
        except (ValueError, TypeError):
            amount = 0.0

        if amount <= 0:
            continue

        desc     = str(row.get("description", "")).strip() if pd.notna(row.get("description")) else ""
        ext_desc = str(row.get("raw_detail",  "")).strip() if pd.notna(row.get("raw_detail"))  else ""
        ddate_s  = str(row.get("date", "")).strip()

        category, entity_name = _match_category(desc, ext_desc, rules)

        debits.append({
            "date":          ddate_s,
            "ref":           str(row.get("ref",  "")).strip(),
            "description":   desc,
            "amount":        amount,
            "raw_detail":    ext_desc,
            "category":      category,
            "entity_name":   entity_name,
            "match_method":  "keyword" if category else "unmatched",
            "payment_month": _extract_payment_month(ddate_s),
        })

    return debits


def parse_all_rows(path: str, categories_csv: str = "categories.csv") -> list[dict]:
    """
    Return EVERY row from the bank statement without any income/expense
    filtering — credits and debits alike.
    """
    rules = _load_category_rules(categories_csv)

    tables = pd.read_html(path, encoding="utf-8")
    tx_df = None
    for df in tables:
        if df.shape[0] > 2 and str(df.iloc[1, 0]).strip() == "תאריך":
            tx_df = df
            break

    if tx_df is None:
        raise ValueError("Could not locate the transactions table in the file.")

    tx_df.columns = tx_df.iloc[1].tolist()
    tx_df = tx_df.iloc[2:].reset_index(drop=True)

    col_map = {
        "תאריך":       "date",
        "תאריך ערך":   "value_date",
        "תיאור":       "description",
        "אסמכתא":      "ref",
        "בחובה":       "debit",
        "בזכות":       "credit",
        'היתרה בש"ח':  "balance",
        "תאור מורחב":  "raw_detail",
        "הערה":        "note",
    }
    tx_df = tx_df.rename(columns={k: v for k, v in col_map.items() if k in tx_df.columns})

    rows = []
    for _, row in tx_df.iterrows():
        if pd.isna(row.get("date")):
            continue

        try:
            credit = float(str(row.get("credit", "0")).replace(",", ""))
        except (ValueError, TypeError):
            credit = 0.0
        try:
            debit = float(str(row.get("debit", "0")).replace(",", ""))
        except (ValueError, TypeError):
            debit = 0.0

        if credit <= 0 and debit <= 0:
            continue

        raw_detail = str(row.get("raw_detail", "")) if pd.notna(row.get("raw_detail")) else ""
        desc       = str(row.get("description", "")).strip() if pd.notna(row.get("description")) else ""
        date_s     = str(row.get("date", "")).strip()

        if credit > 0:
            direction     = "credit"
            amount        = credit
            sender_name   = _extract_sender(raw_detail)
            apt_hint      = _extract_apt_hint(raw_detail)
            payment_month = _resolve_payment_month(raw_detail, date_s)
            category      = None
            entity_name   = None
        else:
            direction     = "debit"
            amount        = debit
            sender_name   = None
            apt_hint      = None
            payment_month = _extract_payment_month(date_s)
            category, entity_name = _match_category(desc, raw_detail, rules)

        rows.append({
            "direction":     direction,
            "date":          date_s,
            "ref":           str(row.get("ref", "")).strip(),
            "description":   desc,
            "amount":        amount,
            "raw_detail":    raw_detail,
            "sender_name":   sender_name,
            "apt_hint":      apt_hint,
            "payment_month": payment_month,
            "category":      category,
            "entity_name":   entity_name,
        })

    return rows


if __name__ == "__main__":
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else "תנועות_בחשבון_1_5_2026.xls"
    txs = parse_bank_file(path)
    print(f"\n✓ Parsed {len(txs)} incoming transactions\n")
    for t in txs[:5]:
        print(json.dumps(t, ensure_ascii=False, indent=2))
    all_rows = parse_all_rows(path)
    print(f"\n✓ parse_all_rows: {len(all_rows)} total rows "
          f"({sum(1 for r in all_rows if r['direction']=='credit')} credit, "
          f"{sum(1 for r in all_rows if r['direction']=='debit')} debit)\n")
