"""
matcher.py
----------
Two-stage matching engine.

Stage 1 – Apartment hint (exact):
  If bank description contains "דירה XX", look up apt XX directly.
  Confidence = 1.0.

Stage 2 – Fuzzy name match:
  Match the (truncated) sender name against ALL name aliases for every tenant
  using rapidfuzz token_set_ratio. Takes the best score across all aliases.
  Confidence = score / 100.

Result dict:
{
    "transaction":  { ...bank dict... },
    "tenant":       { ...tenant dict... } or None,
    "confidence":   0.0–1.0,
    "match_method": "apt_hint" | "fuzzy_name" | "unmatched",
    "match_detail": human-readable string,
    "petty_cash":   bool,   # mirrored from transaction for convenience
}
"""

from rapidfuzz import fuzz, process

FUZZY_THRESHOLD = 72   # tunable – lower = more auto-matches, more false positives


def _best_fuzzy(sender_name, tenants):
    """
    Match sender_name against all name aliases for all tenants.
    Returns (tenant_dict, score_0_to_1, detail_string).
    """
    if not sender_name or not tenants:
        return None, 0.0, ""

    best_score  = -1
    best_tenant = None
    best_alias  = ""

    for tenant in tenants:
        for alias in tenant["all_names"]:
            score = fuzz.token_set_ratio(sender_name, alias)
            if score > best_score:
                best_score  = score
                best_tenant = tenant
                best_alias  = alias

    if best_tenant is None:
        return None, 0.0, ""

    confidence = best_score / 100.0
    detail = f"score {best_score} → {best_alias} (דירה {best_tenant['apartment']})"
    return best_tenant, confidence, detail


def _flag_mismatch(result: dict) -> None:
    """
    Sets result["amount_mismatch"] = True and result["needs_manual"] = True
    when the paid amount differs from the tenant's expected monthly_fee.

    FIX Bug 2: petty cash payments (קופה קטנה in memo) are intentionally ≠ 350.
    They must NOT be flagged as a mismatch — instead they go to manual review
    with the קופה קטנה column pre-selected (handled in app.py).
    """
    # Petty cash payments bypass the mismatch check entirely
    if result["transaction"].get("petty_cash"):
        result["amount_mismatch"] = False
        result["needs_manual"]    = True   # still goes to Tab 3, but pre-targeted to petty cash
        return

    tenant   = result.get("tenant") or {}
    expected = tenant.get("monthly_fee", 0)
    paid     = result["transaction"].get("amount", 0)
    if expected > 0 and paid != expected:
        result["amount_mismatch"] = True
        result["needs_manual"]    = True
    else:
        result["amount_mismatch"] = False
        result["needs_manual"]    = False


def match_transactions(transactions, tenants, fuzzy_threshold=FUZZY_THRESHOLD):
    """
    Returns (matched, unmatched) lists of result dicts.
    """
    by_apt = {t["apartment"]: t for t in tenants}

    matched   = []
    unmatched = []

    for tx in transactions:
        result = {
            "transaction":  tx,
            "tenant":       None,
            "confidence":   0.0,
            "match_method": "unmatched",
            "match_detail": "",
            "petty_cash":   tx.get("petty_cash", False),
        }

        # Stage 1: apartment number hint in description
        apt = tx.get("apt_hint")
        if apt is not None and apt in by_apt:
            result["tenant"]       = by_apt[apt]
            result["confidence"]   = 1.0
            result["match_method"] = "apt_hint"
            result["match_detail"] = f"דירה {apt} → {by_apt[apt]['tenant_name']}"
            _flag_mismatch(result)
            matched.append(result)
            continue

        # Stage 2: fuzzy name match across all aliases
        sender = tx.get("sender_name")
        if sender:
            tenant, confidence, detail = _best_fuzzy(sender, tenants)
            result["tenant"]       = tenant
            result["confidence"]   = confidence
            result["match_detail"] = detail

            if confidence >= (fuzzy_threshold / 100.0):
                result["match_method"] = "fuzzy_name"
                _flag_mismatch(result)
                matched.append(result)
            else:
                result["match_method"] = "unmatched"
                unmatched.append(result)
        else:
            result["match_method"] = "unmatched"
            result["match_detail"] = "no sender name extractable"
            unmatched.append(result)

    return matched, unmatched


def summarise(matched, unmatched):
    total = len(matched) + len(unmatched)
    print(f"\n{'─'*55}")
    print(f"  Matched:   {len(matched):>3}  ({len(matched)/total*100:.0f}%)")
    print(f"  Unmatched: {len(unmatched):>3}  ({len(unmatched)/total*100:.0f}%)")
    print(f"  Total:     {total:>3}")
    print(f"{'─'*55}\n")

    if matched:
        print("Matched:")
        for r in matched:
            tx = r["transaction"]
            print(
                f"  {tx['date']}  ₪{tx['amount']:>7.2f}  "
                f"[{r['match_method']:<11}]  {r['match_detail']}"
            )

    if unmatched:
        print("\nUnmatched:")
        for r in unmatched:
            tx = r["transaction"]
            print(
                f"  {tx['date']}  ₪{tx['amount']:>7.2f}  "
                f"sender={tx.get('sender_name') or '—'}  "
                f"best={r['match_detail'] or '—'}"
            )


if __name__ == "__main__":
    import sys
    from bank_parser import parse_bank_file
    from tenant_loader import load_tenants

    bank_path   = sys.argv[1] if len(sys.argv) > 1 else "תנועות_בחשבון_1_5_2026.xls"
    roster_path = sys.argv[2] if len(sys.argv) > 2 else "apartments.csv"

    transactions = parse_bank_file(bank_path)
    tenants      = load_tenants(roster_path)
    matched, unmatched = match_transactions(transactions, tenants)
    summarise(matched, unmatched)
