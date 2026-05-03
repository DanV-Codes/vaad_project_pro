"""
tenant_loader.py
----------------
Loads the tenant roster CSV and returns a list of active (non-committee) tenants.

Supported CSV columns (auto-detected by alias):
    apartment_number / apartment / דירה  – integer apartment number
    payer_names / tenant_name / שם דייר  – one or more names separated by |
    fee_status / is_committee / ועד       – rows with "waived_committee"/"True"/"1" skipped
    monthly_fee / floor / notes           – optional, carried through

Each returned tenant dict:
{
    "apartment":    25,
    "tenant_name":  "כחלון שלו",           # primary display name (first in list)
    "all_names":    ["דון דז'נשבילי", "כחלון שלו", "בן לולו מרדכי"],
    "name_tokens":  [["דון", "דז'נשבילי"], ["כחלון", "שלו"], ...],
    "monthly_fee":  350,
    "floor":        6,
    "notes":        "",
}
"""

import csv
import re
from pathlib import Path


_APT_ALIASES    = {"apartment_number", "apartment", "דירה", "apt", "unit"}
_NAME_ALIASES   = {"payer_names", "tenant_name", "שם דייר", "שם", "name", "tenant", "דייר"}
_STATUS_ALIASES = {"fee_status", "is_committee", "ועד בית", "committee", "ועד", "status"}
_FEE_ALIASES    = {"monthly_fee", "fee", "תשלום"}
_FLOOR_ALIASES  = {"floor", "קומה"}
_NOTES_ALIASES  = {"notes", "הערות"}

_SKIP_VALUES = {"waived_committee", "true", "1", "yes", "כן", "ועד", "waived"}


def _detect_column(headers, aliases):
    for h in headers:
        if h.strip().lower() in {a.lower() for a in aliases}:
            return h
    return None


def _is_skip(value):
    return str(value).strip().lower() in _SKIP_VALUES


def _split_names(raw):
    parts = re.split(r"[|,;]+", raw)
    return [p.strip() for p in parts if p.strip()]


def load_tenants(path):
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Tenant roster not found: {path}")

    for encoding in ("utf-8-sig", "utf-8", "windows-1255"):
        try:
            with open(p, encoding=encoding, newline="") as f:
                f.read(1024)
            break
        except UnicodeDecodeError:
            continue
    else:
        encoding = "utf-8"

    tenants = []
    with open(p, encoding=encoding, newline="") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []

        apt_col    = _detect_column(headers, _APT_ALIASES)
        name_col   = _detect_column(headers, _NAME_ALIASES)
        status_col = _detect_column(headers, _STATUS_ALIASES)
        fee_col    = _detect_column(headers, _FEE_ALIASES)
        floor_col  = _detect_column(headers, _FLOOR_ALIASES)
        notes_col  = _detect_column(headers, _NOTES_ALIASES)

        if not apt_col:
            raise ValueError(f"No apartment column found. Headers: {headers}")
        if not name_col:
            raise ValueError(f"No name column found. Headers: {headers}")

        for row in reader:
            if status_col and _is_skip(row.get(status_col, "")):
                continue

            raw_apt  = str(row.get(apt_col, "")).strip()
            raw_name = str(row.get(name_col, "")).strip()
            if not raw_apt or not raw_name:
                continue

            try:
                apt_num = int(re.sub(r"\D", "", raw_apt))
            except ValueError:
                apt_num = raw_apt

            all_names = _split_names(raw_name)
            if not all_names:
                continue

            try:
                fee = int(str(row.get(fee_col, "0")).strip()) if fee_col else 0
            except ValueError:
                fee = 0

            try:
                floor = int(str(row.get(floor_col, "0")).strip()) if floor_col else 0
            except ValueError:
                floor = 0

            notes = str(row.get(notes_col, "")).strip() if notes_col else ""

            tenants.append({
                "apartment":   apt_num,
                "tenant_name": all_names[0],
                "all_names":   all_names,
                "name_tokens": [n.split() for n in all_names],
                "monthly_fee": fee,
                "floor":       floor,
                "notes":       notes,
            })

    if not tenants:
        raise ValueError(f"No active tenants found in {path}.")

    return tenants


def tenants_by_apt(tenants):
    return {t["apartment"]: t for t in tenants}


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "apartments.csv"
    tenants = load_tenants(path)
    print(f"\n✓ Loaded {len(tenants)} active tenants\n")
    for t in tenants[:8]:
        print(f"  Apt {t['apartment']:>3}  names: {t['all_names']}")
