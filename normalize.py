"""
normalize.py — small, shared data-cleaning helpers used at every point
a value crosses the Excel <-> live-app boundary. Split out on its own
(rather than living inside excel_export.py or scraping.py) because
BOTH of those modules need it, plus contacts.py — putting it in either
one would make the other import from it for an unrelated reason.
"""

def _normalize_sales_no(value):
    """
    Excel can silently convert a Sales No. cell from text to a real
    number — its own "Number Stored as Text" auto-fix does this with
    one click, and even just re-typing a value can trigger it. Meanwhile
    a live app scrape is ALWAYS a string (Appium's .text always returns
    text, never a number). Without normalizing both sides to the same
    type, "413643" (from the app) and 413643 (from Excel) are NOT equal
    in Python — confirmed: this is exactly what caused a known, already-
    captured customer to be silently re-scraped in full instead of
    skipped, purely because of this type mismatch (nothing to do with
    ordering or position). Every place a sales_no crosses the Excel <->
    live-app boundary is normalized through this function first.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        # Excel sometimes stores a whole number as a float (413643.0) —
        # strip the trailing .0 so it matches the plain-digit string a
        # scrape would produce ("413643", not "413643.0").
        return str(int(value))
    return str(value).strip()


def _normalize_phone_digits(value):
    """
    Same problem as _normalize_sales_no(), same fix — the hidden
    WhatsApp Number column stores a plain digit string (e.g.
    "60123456789"), and Excel can just as easily silently convert that
    to a real number the same way it does Sales No. cells. Carrying
    that forward without normalizing would risk clean_phone_for_wa()
    choking on a non-string value on the next run.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
