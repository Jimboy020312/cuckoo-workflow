"""
excel_export.py — everything about turning scraped records into the
final .xlsm/.xlsx: formulas (WhatsApp message/link), formatting,
column layout, the VBA macro text, and loading back existing values
(Area, Proposed Date, carry-forward data) so re-running never wipes
out something already there.

Reads/writes filenames.OUTPUT_FILE / filenames.OUTPUT_FILE_XLSM /
filenames.CARRYFORWARD_SOURCE_FILE via the `filenames` module directly
(never `from filenames import ...`) since those values change during a
run — see filenames.py's own docstring for why that matters.
"""

import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment

import filenames
from config import ALL_COLUMNS, _COLUMN_INDEX, TEMPLATE_FILE
from areas import detect_areas
from normalize import _normalize_sales_no, _normalize_phone_digits



def _col_letter(column_name):
    """Excel column letter for a column, looked up by its header name —
    e.g. _col_letter("Proposed Date") -> "I". Keeps every formula below
    immune to column reordering."""
    from openpyxl.utils import get_column_letter
    return get_column_letter(_COLUMN_INDEX[column_name])


def _format_duration(seconds):
    """'93.4' -> '1m 33s'; '8.2' -> '8.2s'. Used for the per-customer and
    total run-time output in run()."""
    if seconds >= 60:
        minutes, secs = divmod(int(seconds), 60)
        return f"{minutes}m {secs}s"
    return f"{seconds:.1f}s"


def clean_phone_for_wa(raw):
    """
    Strips everything except digits, then converts a local Malaysian
    mobile number (leading 0, e.g. "012-345 6789") into the international
    format wa.me links need (leading 60, no separators, e.g.
    "60123456789"). A number that's already in some other international
    format (doesn't start with 0 after stripping) is left as digits-only
    — we can't safely guess a country code that isn't already there.
    Returns "" if there's nothing usable. Idempotent — running it again
    on an already-converted number ("60123456789") leaves it unchanged,
    which is what lets a carried-forward WhatsApp Number value safely
    pass back through this on the next run.
    """
    digits = re.sub(r"\D", "", raw or "")
    if not digits:
        return ""
    if digits.startswith("0"):
        digits = "60" + digits[1:]
    return digits


def message_formula(row, specialist_name, nds_id):
    """
    Excel formula for one row's WhatsApp Message cell, matching the
    copywriting/formatting shown in the reference screenshot. Single
    asterisks are WhatsApp's own bold syntax, not markdown. Column
    letters are looked up by name (see _col_letter), not hardcoded.

    specialist_name/nds_id come from the specialist's local identity
    config (see _load_or_prompt_identity) and are baked into the
    formula as literal text, not a cell reference — they're constant
    for the whole export, not something that varies per row. Any
    double-quote either one might contain is escaped to Excel's own
    "" convention so it can't break the formula string.

    Tarikh uses a fully Malay, all-caps date format with the weekday
    name — e.g. "07 OKTOBER 2026 (ISNIN)" — built with
    CHOOSE(MONTH(...)) / CHOOSE(WEEKDAY(...)) since Excel's TEXT() has
    no built-in Malay locale to draw month/weekday names from.

    Note: if you copy this CELL (Ctrl+C, not double-click) and paste
    into something like Notepad, you'll see the whole value wrapped in
    quotes. That's Excel's own clipboard behavior (CSV-style quoting)
    kicking in because the text contains commas and line breaks — it's
    not part of the cell's actual value or this formula, and it doesn't
    matter once you're using the WhatsApp Link column instead of
    copying this text by hand.

    The *Alamat:* line uses SUBSTITUTE to turn the address cell's line
    breaks back into spaces before it goes into the message — the
    address cell itself is formatted with Alt+Enter-style line breaks
    for readability in the sheet (see _format_address_for_cell), but
    since those breaks were inserted right after commas (or, for the
    postcode, right where a plain space already was), swapping each
    line break back to a single space reconstructs the exact original
    single-line address text for the customer-facing message.
    """
    sales_no_col = _col_letter("Sales No.")
    address_col = _col_letter("Installation / Service Address")
    contact_col = _col_letter("Installation / Service Contact Person")
    date_col = _col_letter("Proposed Date")
    product_col = _col_letter("Product")

    safe_name = (specialist_name or "").replace('"', '""')
    safe_nds_id = (nds_id or "").replace('"', '""')

    tarikh_part = (
        f'IFERROR(TEXT({date_col}{row},"DD") & " " & '
        f'CHOOSE(MONTH({date_col}{row}),"JANUARI","FEBRUARI","MAC","APRIL","MEI","JUN",'
        f'"JULAI","OGOS","SEPTEMBER","OKTOBER","NOVEMBER","DISEMBER") & " " & '
        f'TEXT({date_col}{row},"YYYY") & " (" & '
        f'CHOOSE(WEEKDAY({date_col}{row},2),"ISNIN","SELASA","RABU","KHAMIS","JUMAAT","SABTU","AHAD") & ")", '
        f'{date_col}{row})'
    )
    return (
        f'=IF({date_col}{row}="","","Selamat sejahtera Tuan/Puan " & UPPER({contact_col}{row}) & "," & CHAR(10) & CHAR(10) & '
        f'"Saya *{safe_name}*, CUCKOO+ Service Specialist ({safe_nds_id}). Saya memohon maaf jika saya menghubungi anda pada waktu yang tidak sesuai." & CHAR(10) & CHAR(10) & '
        f'"Saya ingin mengesahkan jika saya boleh membuat lawatan servis seperti di bawah." & CHAR(10) & CHAR(10) & '
        f'"*Tarikh:* " & {tarikh_part} & CHAR(10) & '
        f'"*Alamat:* " & SUBSTITUTE({address_col}{row},CHAR(10)," ") & CHAR(10) & '
        f'"*Produk:* " & {product_col}{row} & CHAR(10) & '
        f'"*Nombor Pesanan:* " & {sales_no_col}{row} & CHAR(10) & CHAR(10) & '
        f'"Terima kasih, sokongan dan kerjasama Tuan/Puan amat saya hargai.")'
    )


def wa_link_formula(row, phone_digits):
    """
    Excel formula for one row's WhatsApp Link cell.

    IMPORTANT LIMITATION: this can't pre-fill the message the way a
    wa.me "?text=" link normally would. Excel's own HYPERLINK() worksheet
    function hard-caps its link_location argument at 255 characters —
    if it's longer, Excel returns #VALUE! instead of opening anything.
    The encoded message (greeting + explanation + address etc.) is
    always well past that, no matter how it's trimmed, so a formula-
    based link genuinely cannot carry the pre-filled text.

    What this DOES do: opens the right WhatsApp chat directly (a plain
    "https://wa.me/<number>" link, comfortably under 255 chars), so you
    only need to copy the WhatsApp Message cell and paste it in — no
    manual number lookup or searching for the contact. Column letter is
    looked up by name (see _col_letter), not hardcoded.
    """
    date_col = _col_letter("Proposed Date")
    if not phone_digits:
        return '="No phone number found"'
    return (
        f'=IF({date_col}{row}="","",HYPERLINK('
        f'"https://wa.me/{phone_digits}",'
        f'"Open WhatsApp Chat"))'
    )


_FILTER_LINE_NUMBER_PATTERN = re.compile(r"^\d+\.\s*")


def _number_filters_text(text):
    """
    Adds "1. ", "2. ", etc. in front of each line of Filters text.
    Strips any numbering already present first (matching a leading
    "N. ") before renumbering — this is what stops repeated --resort
    runs from stacking up "1. 1. Product..." — so it's safe to call on
    text that's already numbered, freshly built, or anything in between.
    Blank/empty input passes through unchanged.
    """
    if not text:
        return text
    lines = text.split("\n")
    cleaned = [_FILTER_LINE_NUMBER_PATTERN.sub("", line) for line in lines]
    return "\n".join(f"{i}. {line}" for i, line in enumerate(cleaned, start=1))


def _format_address_for_cell(address):
    """
    Breaks a long address across multiple lines for easier reading in
    the sheet — the same effect as pressing Alt+Enter at each comma —
    WITHOUT removing anything: every comma stays exactly where it was,
    a line break is just added right after it.

    Also breaks right before the postcode, if one's found — using the
    exact same postcode detection the Area auto-fill above already
    relies on — since that's usually the one place a long address has
    no comma at all to break on naturally (e.g. "...PRESINT 11
    PUTRAJAYA 62300 WP PUTRAJAYA" runs on with no punctuation).

    Needs wrap_text on for the cell to actually show as multiple lines
    (already true for the address column via WRAPPED_COLUMNS below).

    This is applied ONLY here, at display time — detect_areas() (in the
    "Area auto-fill" block near the top of this file) still runs on the
    original, unbroken address text before this ever happens, so none
    of this affects Area detection.

    The WhatsApp Chat Message column reconstructs the original,
    single-line address from this (see message_formula's use of
    SUBSTITUTE) rather than showing these line breaks in the actual
    message text sent to the customer.
    """
    if not address:
        return address
    text = address.strip()
    # Break right after every comma — comma itself is kept, untouched.
    text = re.sub(r",\s*", ",\n", text)
    # Also break right before the postcode, if it isn't already at the
    # start of a line (e.g. wasn't already right after a comma).
    postcode_match = re.search(r"\b\d{5}\b", text)
    if postcode_match:
        start = postcode_match.start()
        if start > 0 and text[start - 1] != "\n":
            text = text[:start].rstrip() + "\n" + text[start:]
    return text


def _populate_sheet(ws, records, filters_by_sales_no, specialist_name, nds_id):
    """
    Fills in headers + rows on an already-created worksheet (either a
    fresh one, or the one already inside the macro-enabled template).
    Shared by both output paths in write_output() below so the two
    stay in sync automatically. filters_by_sales_no maps sales_no ->
    the combined multi-line Filters text (see _build_filters_lookup).
    """
    FONT = "Arial"
    # Cuckoo's own logo red is a vivid red-orange — I couldn't pull an
    # exact official hex from their site, so this is a close visual
    # match (a common vivid Korean-appliance-brand red). If you have
    # the real logo file and want an exact match, this is the one
    # value to swap.
    header_fill = PatternFill("solid", fgColor="ED1C24")

    # Short, simple values read better centered; long free text (name,
    # address, filters, message) reads better left-aligned but still
    # vertically centered — every cell in the sheet gets vertical
    # centering regardless, only horizontal centering is selective.
    CENTERED_COLUMNS = {
        _COLUMN_INDEX["Sales No."], _COLUMN_INDEX["Appointment Date"],
        _COLUMN_INDEX["Proposed Date"], _COLUMN_INDEX["WhatsApp Chat Link"],
        _COLUMN_INDEX["Product"], _COLUMN_INDEX["WhatsApp Number"],
    }
    WRAPPED_COLUMNS = {
        _COLUMN_INDEX["Installation / Service Address"],
        _COLUMN_INDEX["Filter(s)"], _COLUMN_INDEX["WhatsApp Chat Message"],
    }

    for col_idx, header in enumerate(ALL_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(
            wrap_text=True, vertical="center", horizontal="center")

    area_comment_text = (
        "Auto-filled ONLY when the script is confident (free, offline — "
        "postcode lookup + keyword matching, no AI, no paid API). Left "
        "blank whenever it isn't sure, rather than risk a wrong guess. "
        "Fill in the rest yourself, or add more detail to what's already "
        "there. Once a cell has something in it — auto-filled OR typed "
        "by hand — every future run leaves it exactly as-is, never "
        "overwrites it. Use Excel's filter on these columns to drill "
        "down: Area 1 first, then Area 2, Area 3, Area 4 as needed."
    )
    ws[f"{_col_letter('Area 1 (State)')}1"].comment = Comment(
        area_comment_text, "main.py")
    ws[f"{_col_letter('Area 2 (District / City)')}1"].comment = Comment(
        area_comment_text, "main.py")
    ws[f"{_col_letter('Area 3 (Locality)')}1"].comment = Comment(
        area_comment_text, "main.py")
    ws[f"{_col_letter('Area 4 (Street)')}1"].comment = Comment(
        area_comment_text, "main.py")

    ws[f"{_col_letter('Proposed Date')}1"].comment = Comment(
        "Type a date here (e.g. 08/08/2026 or 14 August 2026 both work) —\n"
        "it displays as \"14 August 2026\" once entered.\n"
        "The WhatsApp Message and WhatsApp Link columns fill themselves in automatically.",
        "main.py"
    )
    ws[f"{_col_letter('WhatsApp Chat Link')}1"].comment = Comment(
        "Click to open the right WhatsApp chat directly, then copy the\n"
        "WhatsApp Chat Message column and paste it in. (Excel's HYPERLINK\n"
        "function can't carry a pre-filled message this long — it caps\n"
        "links at 255 characters — so this gets you to the chat, but the\n"
        "message still needs one paste.)\n\n"
        "If this file has the WhatsApp macro installed (see write_output()'s\n"
        "docstring in main.py), double-click instead of single-\n"
        "clicking — that opens the chat with the message already filled in,\n"
        "no paste needed.",
        "main.py"
    )
    ws[f"{_col_letter('WhatsApp Number')}1"].comment = Comment(
        "Internal use only — the macro reads this to build the full "
        "pre-filled WhatsApp link, and it's also how the phone number "
        "survives into next month's export now that Mobile No. 1 isn't "
        "its own column. Don't edit or delete this column.",
        "main.py"
    )

    def set_cell(row_idx, col_idx, value):
        # clean_text is skipped for formula strings (start with "=") —
        # they're Excel syntax, not scraped text, and don't need it.
        if isinstance(value, str) and not value.startswith("="):
            value = clean_text(value)
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(
            vertical="center",
            horizontal="center" if col_idx in CENTERED_COLUMNS else "general",
            wrap_text=col_idx in WRAPPED_COLUMNS,
        )
        return cell

    for row_idx, record in enumerate(records, start=2):
        set_cell(row_idx, _COLUMN_INDEX["Sales No."],
                 record.get("sales_no", ""))
        set_cell(row_idx, _COLUMN_INDEX["Installation / Service Contact Person"],
                 record.get("install_contact_person", ""))
        set_cell(row_idx, _COLUMN_INDEX["Installation / Service Address"],
                 _format_address_for_cell(record.get("install_address", "")))

        # Area 1 -> Area 4 (State/District-City/Locality/Street) —
        # auto-filled when confident, otherwise manual (see
        # _load_existing_area_values / detect_areas). Whatever's already
        # in `record` here (existing value OR fresh auto-fill) was
        # already decided before write_output() got this far.
        set_cell(
            row_idx, _COLUMN_INDEX["Area 1 (State)"], record.get("area1", ""))
        set_cell(
            row_idx, _COLUMN_INDEX["Area 2 (District / City)"], record.get("area2", ""))
        set_cell(
            row_idx, _COLUMN_INDEX["Area 3 (Locality)"], record.get("area3", ""))
        set_cell(
            row_idx, _COLUMN_INDEX["Area 4 (Street)"], record.get("area4", ""))

        set_cell(row_idx, _COLUMN_INDEX["Appointment Date"],
                 record.get("appt_date", ""))

        # Proposed Date — typed by hand, no special fill, displayed as
        # "14 August 2026" regardless of how it was typed in.
        date_cell = set_cell(
            row_idx, _COLUMN_INDEX["Proposed Date"], record.get("proposed_date"))
        date_cell.number_format = "d mmmm yyyy"

        phone_digits = clean_phone_for_wa(record.get("install_mobile1", ""))

        link_cell = set_cell(row_idx, _COLUMN_INDEX["WhatsApp Chat Link"],
                             wa_link_formula(row_idx, phone_digits))
        link_cell.font = Font(name=FONT, size=10,
                              color="1155CC", underline="single")

        set_cell(row_idx, _COLUMN_INDEX["WhatsApp Chat Message"],
                 message_formula(row_idx, specialist_name, nds_id))

        set_cell(row_idx, _COLUMN_INDEX["Product"],
                 record.get("sales_info_product", ""))

        # Filter(s) — combined text from CCS Note data, looked up by
        # sales_no; blank if this customer had no CCS Note data captured.
        raw_filters_text = filters_by_sales_no.get(
            record.get("sales_no", ""), "")
        set_cell(row_idx, _COLUMN_INDEX["Filter(s)"],
                 _number_filters_text(raw_filters_text))

        # WhatsApp Number — hidden helper column. The macro (if installed)
        # reads this directly instead of re-deriving the phone number, and
        # it's also now the only place the phone number is persisted
        # between runs (see _load_carryforward_data). Forced to Excel's
        # text format ('@') so re-saving/re-opening the file can't quietly
        # convert it to a real number the way Sales No. cells could.
        wa_number_cell = set_cell(
            row_idx, _COLUMN_INDEX["WhatsApp Number"], phone_digits)
        wa_number_cell.number_format = "@"

    widths = {
        "Sales No.": 14, "Installation / Service Contact Person": 26,
        "Installation / Service Address": 40,
        "Area 4 (Street)": 22, "Area 3 (Locality)": 22,
        "Area 2 (District / City)": 20, "Area 1 (State)": 16,
        "Appointment Date": 12, "Proposed Date": 14,
        "WhatsApp Chat Link": 18, "WhatsApp Chat Message": 60,
        "Product": 16, "Filter(s)": 40, "WhatsApp Number": 12,
    }
    for name, w in widths.items():
        ws.column_dimensions[_col_letter(name)].width = w
    # Header row AND the first two columns (Sales No., Installation /
    # Service Contact Person) stay frozen — the freeze point is the cell
    # diagonally past both, i.e. the first cell of column C.
    ws.freeze_panes = "C2"


def _build_filters_lookup(ccs_rows):
    """
    Turns the flat list of {sales_no, cust_name, product, last_change}
    rows from get_all_ccs_note_cards() into {sales_no: combined_text},
    one "Product: date" line per filter (real line breaks, so it reads
    like pressing Alt+Enter between each one), sorted alphabetically by
    product name for a consistent read.
    """
    by_customer = {}
    for r in ccs_rows:
        by_customer.setdefault(r["sales_no"], []).append(
            (r["product"], r["last_change"]))

    lookup = {}
    for sales_no, filters in by_customer.items():
        filters_sorted = sorted(filters, key=lambda f: f[0])
        lookup[sales_no] = "\n".join(
            f"{product}: {last_change if last_change else 'not yet changed'}"
            for product, last_change in filters_sorted
        )
    return lookup


def clean_text(value):
    """
    Strips leading/trailing whitespace and collapses any run of spaces
    or tabs down to a single space — applied to every piece of scraped
    text before it reaches a cell, since the app's UI has been observed
    to hand back things like "PRESINT 11  PUTRAJAYA" (double space).
    Deliberately only touches spaces/tabs, not newlines — the Filters
    column's line breaks (one filter per line) need to survive this.
    Non-strings (numbers, dates, None) pass through untouched.
    """
    if not isinstance(value, str):
        return value
    return re.sub(r"[ \t]+", " ", value.strip())


_AREA_FIELD_TO_COLUMN = {
    "area1": "Area 1 (State)",
    "area2": "Area 2 (District / City)",
    "area3": "Area 3 (Locality)",
    "area4": "Area 4 (Street)",
}


def _load_existing_area_values():
    """
    Reads the Area 1-4 columns back from whichever export file already
    exists, keyed by sales_no -> {"area1": ..., "area2": ..., ...}.
    This is what makes anything already sitting in those columns
    (typed by hand OR auto-filled on an earlier run) survive being
    carried forward run after run instead of getting blanked out or
    silently overwritten every time the sheet is rebuilt. New customers
    (no existing row) just aren't in this dict. Returns {} if no export
    file exists yet, or if it's an older file from before these columns
    existed (in which case everything auto-fills fresh, same as a
    first-ever run).
    """
    target_file = filenames.CARRYFORWARD_SOURCE_FILE
    if not target_file:
        return {}
    wb = None
    try:
        wb = openpyxl.load_workbook(target_file, data_only=False)
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        if "Sales No." not in header_row:
            return {}
        sales_no_col = header_row.index("Sales No.") + 1

        area_cols = {}
        for field, column_name in _AREA_FIELD_TO_COLUMN.items():
            if column_name in header_row:
                area_cols[field] = header_row.index(column_name) + 1

        if not area_cols:
            return {}  # older file, from before Area 1-4 existed — nothing to carry over

        existing = {}
        for row_idx in range(2, ws.max_row + 1):
            sales_no = _normalize_sales_no(
                ws.cell(row=row_idx, column=sales_no_col).value)
            if not sales_no:
                continue
            values = {}
            for field, col_idx in area_cols.items():
                val = ws.cell(row=row_idx, column=col_idx).value
                if val:
                    values[field] = val
            if values:
                existing[sales_no] = values
        return existing
    except Exception as e:
        print(f"  !! Could not read existing Area values from {target_file} "
              f"({e}) — starting fresh for all customers.")
        return {}
    finally:
        # Always closed explicitly here, on our own terms, instead of
        # leaving it to Python's garbage collector — relying on GC timing
        # for this is what causes the harmless-but-noisy "Exception
        # ignored...ValueError: I/O operation on closed file" warning
        # some Python versions print at interpreter shutdown.
        if wb is not None:
            wb.close()


def build_vba_macro():
    """
    Builds the ready-to-paste VBA macro using the CURRENT column
    positions (same _col_letter/_COLUMN_INDEX lookups as everything
    else) — this is what makes the macro immune to going stale after a
    future column reorder, which is exactly what broke it last time (it
    was hardcoded to a fixed column for the WhatsApp Link, and drifted
    out of sync once the columns were reordered).

    Run `python main.py --show-macro` any time you need a fresh, correct
    copy — safer than trusting a copy/pasted version to still be right.
    """
    link_col_num = _COLUMN_INDEX["WhatsApp Chat Link"]
    link_col_letter = _col_letter("WhatsApp Chat Link")
    number_col_letter = _col_letter("WhatsApp Number")
    message_col_letter = _col_letter("WhatsApp Chat Message")
    return (
        "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)\n"
        "    Dim r As Long, num As String, msg As String, url As String\n"
        f"    If Target.Column <> {link_col_num} Then Exit Sub   ' column {link_col_letter} = WhatsApp Chat Link\n"
        "    r = Target.Row\n"
        "    If r < 2 Then Exit Sub\n"
        f'    num = Trim(Cells(r, "{number_col_letter}").Value)       \' hidden helper column\n'
        f'    msg = Cells(r, "{message_col_letter}").Value              \' WhatsApp Chat Message\n'
        '    If num = "" Or msg = "" Then Exit Sub\n'
        "    ' web.whatsapp.com (not wa.me) on purpose — wa.me links often\n"
        "    ' get intercepted by WhatsApp Desktop if it's installed, and\n"
        "    ' the desktop app silently drops the pre-filled text for a\n"
        "    ' chat that already exists. Routing through web.whatsapp.com\n"
        "    ' forces it into an actual browser tab, where the text\n"
        "    ' reliably shows up.\n"
        '    url = "https://web.whatsapp.com/send?phone=" & num & "&text=" & WorksheetFunction.EncodeURL(msg)\n'
        "    ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True\n"
        "    Cancel = True\n"
        "End Sub"
    )


def _load_existing_proposed_dates():
    """
    Reads the Proposed Date column back from whichever export file
    already exists, keyed by sales_no. Mirrors _load_existing_area_values
    exactly, for the exact same reason: without this, a typed-in
    Proposed Date gets silently wiped back to blank every time the
    script is run again (confirmed by testing — re-running after new
    customers appear in the app was blanking out dates already entered
    for existing ones, since a fresh scrape has no way to know what you
    typed in Excel afterward). Returns {} if no export file exists yet,
    or if it's from before this column existed.
    """
    target_file = filenames.CARRYFORWARD_SOURCE_FILE
    if not target_file:
        return {}
    wb = None
    try:
        wb = openpyxl.load_workbook(target_file, data_only=False)
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        if "Sales No." not in header_row or "Proposed Date" not in header_row:
            return {}
        sales_no_col = header_row.index("Sales No.") + 1
        date_col = header_row.index("Proposed Date") + 1

        existing = {}
        for row_idx in range(2, ws.max_row + 1):
            sales_no = _normalize_sales_no(
                ws.cell(row=row_idx, column=sales_no_col).value)
            date_val = ws.cell(row=row_idx, column=date_col).value
            if sales_no and date_val:
                existing[sales_no] = date_val
        return existing
    except Exception as e:
        print(f"  !! Could not read existing Proposed Date values from "
              f"{target_file} ({e}) — starting fresh for all customers.")
        return {}
    finally:
        if wb is not None:
            wb.close()


def write_output(records, ccs_rows=None, filters_override=None, app_order=None,
                 specialist_name="Hanis", nds_id="NDS35095"):
    """
    Writes the export with the columns in the order specified at the
    top of this section (sales_no, install_contact_person,
    install_address, Area 4-1, appt_date, proposed_date, WhatsApp Link,
    WhatsApp Message, sales_info_product, Filters, wa_number) — all in
    ONE sheet, no separate CCS Notes/Route Plan sheet.

    specialist_name/nds_id feed straight into message_formula() — see
    its docstring — so every caller should really be passing the
    values from _load_or_prompt_identity() rather than relying on
    these defaults, which only exist so this function still works if
    called without them.

    Area 1-4 auto-fill when the script is confident (free, offline
    postcode + keyword matching — see the "Area auto-fill" block near
    the top of this file), and stay blank otherwise for you to fill in
    by hand. Once any Area cell has something in it — auto-filled OR
    typed by hand — every future run preserves it exactly as-is; it's
    for filtering in Excel yourself, not something this script keeps
    trying to re-guess.

    TWO POSSIBLE OUTPUTS, chosen automatically:

    1. If TEMPLATE_FILE ("macro.xlsm") exists, this
       writes INTO a copy of it (loaded with keep_vba=True so its
       macro survives) and saves as filenames.OUTPUT_FILE_XLSM. In that file,
       double-clicking a WhatsApp Link cell runs the macro, which opens
       WhatsApp with the message already filled in — no paste needed,
       because VBA's FollowHyperlink isn't subject to the 255-character
       cap that the HYPERLINK() *formula* has.

    2. Otherwise, this falls back to a plain filenames.OUTPUT_FILE (.xlsx) with
       no macro — the WhatsApp Link column still works via single
       click, it just opens the bare chat (see wa_link_formula()'s
       docstring) rather than pre-filling the message.

    ONE-TIME SETUP for option 1 (only needs doing once, ever —
    openpyxl can't write compiled VBA itself, so this part is manual).
    IMPORTANT: if you already pasted an earlier version of this macro
    (from before a column reorder), it's now WRONG — the column
    letters below match the CURRENT layout at the time this docstring
    was last written, but the safest option is always to run
    `python main.py --show-macro` and paste whatever it prints, since
    that's generated fresh from the current column layout every time
    rather than relying on this text staying in sync.
      a. Run the script once normally so a plain cuckoo_export.xlsx
         exists with the columns/formulas already in it.
      b. Open that file in Excel. Press Alt+F11 to open the VBA editor.
      c. In the Project pane on the left, double-click the entry for
         this sheet (e.g. "Sheet1 (Cuckoo Export)") — NOT "Insert >
         Module". This must be the sheet's own code-behind so the
         double-click event actually fires.
      d. Paste in (replacing any earlier version):

           Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
               Dim r As Long, num As String, msg As String, url As String
               If Target.Column <> 9 Then Exit Sub   ' column I = WhatsApp Chat Link
               r = Target.Row
               If r < 2 Then Exit Sub
               num = Trim(Cells(r, "M").Value)       ' hidden helper column
               msg = Cells(r, "J").Value              ' WhatsApp Chat Message
               If num = "" Or msg = "" Then Exit Sub
               ' web.whatsapp.com (not wa.me) on purpose — wa.me links often
               ' get intercepted by WhatsApp Desktop if it's installed, and
               ' the desktop app silently drops the pre-filled text for a
               ' chat that already exists. Routing through web.whatsapp.com
               ' forces it into an actual browser tab, where the text
               ' reliably shows up.
               url = "https://web.whatsapp.com/send?phone=" & num & "&text=" & WorksheetFunction.EncodeURL(msg)
               ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True
               Cancel = True
           End Sub

      e. Close the VBA editor. File > Save As > "Excel Macro-Enabled
         Workbook (*.xlsm)" > save it as exactly "macro.xlsm", in the
         project's root folder (next to this script — NOT inside any
         of the per-person NDS-ID/name export folders).
      f. From then on, every run of this script detects that file and
         writes into it automatically — this setup never needs
         repeating (unless the columns change again).
    """
    ccs_rows = ccs_rows or []

    if not records:
        print("No records captured — nothing written.")
        return filenames.OUTPUT_FILE

    for r in records:
        r["sales_no"] = _normalize_sales_no(r.get("sales_no"))

    # Area 1-4: for each customer, whatever's already in the existing
    # file (typed by hand OR auto-filled on an earlier run) always wins
    # and is carried over untouched. Only a field that's genuinely never
    # been filled before gets a fresh auto-fill attempt from
    # detect_areas() — and even then, only the specific area1/2/3/4
    # fields it's confident about get filled; the rest stay blank for
    # you to fill in by hand, same as before. This has to happen BEFORE
    # sorting below, since the sort is now based on these values.
    existing_area_values = _load_existing_area_values()
    for r in records:
        existing = existing_area_values.get(r.get("sales_no", ""), {})
        auto = detect_areas(r.get("install_address", ""))
        for field in ("area1", "area2", "area3", "area4"):
            r[field] = existing.get(field) or auto.get(field, "")

    # Proposed Date: same protection as Area 1-4 above — a customer's
    # already-typed date always wins over a fresh (blank) scrape value.
    # A truly new customer has nothing here yet, so it just stays blank
    # for you to fill in, same as before.
    existing_proposed_dates = _load_existing_proposed_dates()
    for r in records:
        existing_date = existing_proposed_dates.get(r.get("sales_no", ""))
        if existing_date:
            r["proposed_date"] = existing_date

    # Sort: primarily by Area 1 -> Area 2 -> Area 3 -> Area 4, grouping
    # customers by location as closely as the auto-fill/manual data
    # allows — this is the actual point of having Area columns at all,
    # so the finished sheet should reflect it directly, not just leave
    # grouping to Excel's filter dropdowns. A blank value at any level
    # sorts AFTER a real value at that same level, so fully-categorized
    # rows cluster together and the ones still needing attention (blank
    # Area) end up together too, easy to spot.
    #
    # Within an identical Area 1-4 combination (very common — e.g. every
    # "blank/blank/blank/blank" new customer, or several people in the
    # same precinct), the tiebreaker is app_order if the caller supplied
    # one (run() does, from the order customers were actually encountered
    # scrolling through the app this run — which can genuinely differ
    # from any previous export's order) — falling back to Sales No. for
    # a caller that didn't supply one (e.g. --resort).
    app_order_lookup = {sn: idx for idx,
                        sn in enumerate(app_order)} if app_order else {}

    def _area_sort_key(r):
        def level(value):
            return (1, "") if not value else (0, value)
        sales_no = r.get("sales_no", "")
        return (
            level(r.get("area1", "")),
            level(r.get("area2", "")),
            level(r.get("area3", "")),
            level(r.get("area4", "")),
            app_order_lookup.get(sales_no, len(app_order_lookup)),
            sales_no,
        )

    records = sorted(records, key=_area_sort_key)

    filters_by_sales_no = filters_override if filters_override is not None else _build_filters_lookup(
        ccs_rows)

    if os.path.exists(TEMPLATE_FILE):
        try:
            wb = openpyxl.load_workbook(TEMPLATE_FILE, keep_vba=True)
            ws = wb.active
            # Clear out any previous run's rows before writing fresh ones,
            # but leave row 1 (headers) and the macro itself untouched.
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
            _populate_sheet(ws, records, filters_by_sales_no,
                            specialist_name, nds_id)
            wb.save(filenames.OUTPUT_FILE_XLSM)
            wb.close()
            return filenames.OUTPUT_FILE_XLSM
        except Exception as e:
            print(f"  !! Could not write into {TEMPLATE_FILE} ({e}) — "
                  f"falling back to a plain .xlsx instead.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cuckoo Export"
    _populate_sheet(ws, records, filters_by_sales_no, specialist_name, nds_id)
    wb.save(filenames.OUTPUT_FILE)
    wb.close()
    return filenames.OUTPUT_FILE




def _load_carryforward_data():
    """
    Reads whichever export file already exists and returns everything
    needed to SKIP re-scraping a customer whose sales_no is already in
    it, while still including them correctly in this run's output:

      known_sales_nos -> set of sales_no values already captured in a
                          previous run. run() uses this to skip opening
                          View Order/CCS Note for these — per Hazim's
                          explicit call, a customer's details don't
                          change once captured, so there's nothing new
                          to read there.
      carried_records -> {sales_no: {...plain fields...}}, read straight
                          from the existing file's own columns, so an
                          already-known customer's row can be reproduced
                          in the new export without the app ever being
                          touched for them this run. The phone number
                          (install_mobile1) is recovered from the hidden
                          WhatsApp Number column, since the visible
                          Mobile No. 1 column was removed — that hidden
                          column already holds the cleaned digits, and
                          re-cleaning an already-clean number through
                          clean_phone_for_wa() is a no-op.
      carried_filters -> {sales_no: filters_text}, read straight from
                          the Filter(s) column, same reasoning.

    Proposed Date and Area 1-4 are NOT included here on purpose — those
    already have their own carry-forward logic inside write_output()
    (_load_existing_area_values / _load_existing_proposed_dates), which
    reads the same file independently and applies "existing value always
    wins." Duplicating that here would just be redundant.

    Returns (set(), {}, {}) if no export file exists yet (first-ever
    run — nothing to skip, everyone gets scraped), or if the existing
    file is from too different a layout to safely reuse (missing a
    required column) — safer to re-scrape everyone than guess.
    """
    target_file = filenames.CARRYFORWARD_SOURCE_FILE
    if not target_file:
        return set(), {}, {}

    wb = openpyxl.load_workbook(target_file, data_only=False)
    try:
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        col = {name: idx + 1 for idx, name in enumerate(header_row) if name}

        required = ["Sales No.", "Appointment Date", "Installation / Service Contact Person",
                    "Installation / Service Address", "Product", "WhatsApp Number"]
        if any(name not in col for name in required):
            print(f"  !! {target_file}'s header row is missing required columns — "
                  f"can't safely skip/carry forward from it. Re-scraping everyone "
                  f"this run instead.")
            return set(), {}, {}

        known_sales_nos = set()
        carried_records = {}
        carried_filters = {}
        for row_idx in range(2, ws.max_row + 1):
            sales_no = _normalize_sales_no(
                ws.cell(row=row_idx, column=col["Sales No."]).value)
            if not sales_no:
                continue
            known_sales_nos.add(sales_no)
            carried_records[sales_no] = {
                "sales_no": sales_no,
                "appt_date": ws.cell(row=row_idx, column=col["Appointment Date"]).value or "",
                "install_contact_person": ws.cell(row=row_idx, column=col["Installation / Service Contact Person"]).value or "",
                "install_mobile1": _normalize_phone_digits(
                    ws.cell(row=row_idx, column=col["WhatsApp Number"]).value),
                "install_address": ws.cell(row=row_idx, column=col["Installation / Service Address"]).value or "",
                "sales_info_product": ws.cell(row=row_idx, column=col["Product"]).value or "",
            }
            if "Filter(s)" in col:
                filters_text = ws.cell(
                    row=row_idx, column=col["Filter(s)"]).value
                if filters_text:
                    carried_filters[sales_no] = filters_text
        return known_sales_nos, carried_records, carried_filters
    finally:
        wb.close()
