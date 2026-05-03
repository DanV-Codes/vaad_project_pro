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
}
"""

import re
import pandas as pd


# ── Regex patterns ────────────────────────────────────────────────────────────
_RE_SENDER   = re.compile(r"העברה מאת:\s*(.+?)\s+\d{2}-\d{3}-\d+", re.UNICODE)
_RE_APT      = re.compile(r"דירה\s+(\d+)", re.UNICODE)
_RE_CHECKIN  = re.compile(r"הפקדת שיק", re.UNICODE)   # cheque deposits – keep


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
    """Pull apartment number from 'דירה XX' if present."""
    if not isinstance(detail, str):
        return None
    m = _RE_APT.search(detail)
    return int(m.group(1)) if m else None


def _is_incoming(row) -> bool:
    """
    Keep a row if:
      - it has a credit (זכות > 0)  AND
      - it comes from a person (העברה מאת / הפקדת שיק)
    Outgoing payments (העברה אל / fees / utilities) are discarded.
    """
    try:
        credit = float(str(row["בזכות"]).replace(",", "")) if pd.notna(row["בזכות"]) else 0.0
    except (ValueError, TypeError):
        credit = 0.0

    if credit <= 0:
        return False

    detail = str(row.get("תאור מורחב", ""))
    desc   = str(row.get("תיאור", ""))

    # Outgoing transfers start with "העברה אל:"
    if "העברה אל:" in detail:
        return False

    # Accept if it's an incoming bank transfer or a cheque deposit
    if "העברה מאת:" in detail:
        return True
    if _RE_CHECKIN.search(desc):
        return True

    return False


def parse_bank_file(path: str) -> list[dict]:
    """
    Main entry point.
    Returns a list of cleaned transaction dicts ready for matching.
    """
    # Bank Leumi exports HTML-as-XLS; read_html handles it cleanly.
    tables = pd.read_html(path, encoding="utf-8")

    # Find the transactions table: it's the one with column header "תאריך"
    tx_df = None
    for df in tables:
        # Header is in row 1 (row 0 is a merged "תנועות בחשבון" title)
        if df.shape[0] > 2 and str(df.iloc[1, 0]).strip() == "תאריך":
            tx_df = df
            break

    if tx_df is None:
        raise ValueError("Could not locate the transactions table in the file.")

    # Promote row 1 as column headers, drop rows 0 and 1
    tx_df.columns = tx_df.iloc[1].tolist()
    tx_df = tx_df.iloc[2:].reset_index(drop=True)

    # Rename columns to stable English keys internally
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
        # Skip rows where date is missing (footer rows etc.)
        if pd.isna(row.get("date")):
            continue

        # We map back to Hebrew keys so _is_incoming reuse is clean
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

        raw_detail   = str(row.get("raw_detail", "")) if pd.notna(row.get("raw_detail")) else ""
        sender_name  = _extract_sender(raw_detail)
        apt_hint     = _extract_apt_hint(raw_detail)

        transactions.append({
            "date":        str(row.get("date", "")).strip(),
            "ref":         str(row.get("ref", "")).strip(),
            "description": str(row.get("description", "")).strip(),
            "amount":      amount,
            "raw_detail":  raw_detail,
            "sender_name": sender_name,
            "apt_hint":    apt_hint,
        })

    return transactions


# ── Quick smoke-test ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else "תנועות_בחשבון_1_5_2026.xls"
    txs = parse_bank_file(path)
    print(f"\n✓ Parsed {len(txs)} incoming transactions\n")
    for t in txs[:5]:
        print(json.dumps(t, ensure_ascii=False, indent=2))
