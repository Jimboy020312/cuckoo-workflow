"""
config.py — every tunable constant and static lookup table for the
Cuckoo+ scraper. This is the file to edit if you want to change a
setting (LIST_LIMIT, scroll timing, which labels map to which field,
etc.) — nothing in here is reassigned while the script is running, so
editing a value here is safe and takes effect the next time you run it.

Split out of the original single-file main.py so related settings live
together and are easy to find, without needing to hunt through the
scraping/export logic to find a constant.
"""

APP_PACKAGE = "cuckoo.doctress"
# confirmed: the list screen
APP_ACTIVITY = "cuckoo.doctress.naturalcareservicelist"

LIST_LIMIT = None   # was 3 for testing — now processes the whole list
# set True again only if something breaks and you need to see raw element data
DEBUG = False

# --- List screen (confirmed from XML) ---
LABEL_FIELD_MAP = {
    "NS No": "ns_no",
    "Sales No": "sales_no",
    "NS Date": "ns_date",
    "Status": "status",
    "Appt Date": "appt_date",
    "Product Name": "product_name",
    "Cust Name": "cust_name",
}
KEY_FIELD = "ns_no"   # unique per card — avoids re-scraping someone after scrolling
# a card missing any of these is
REQUIRED_LIST_FIELDS = set(LABEL_FIELD_MAP.values())
# probably cut off by the screen
# edge, not actually incomplete data

# --- Detail screen: Address/Contact Info tab (confirmed from XML) ---
#
# NOTE: this screen also has a "Billing Address & Contact" section and
# an "Emergency Contact" section, both fully labeled/parseable the same
# way as everything else here — but neither one is scraped anymore.
# Verified by cross-referencing every field this script could produce
# against every place it's actually read downstream: nothing in the
# export, in carry-forward, or anywhere else ever reads a billing_* or
# emergency_* field. Scraping them cost two full extra device round
# trips per newly-scraped customer (get_billing_elements() and
# get_emergency_elements() each did their own wait_for + find_elements)
# for data that went nowhere. Cut for real runtime savings, with zero
# risk to anything actually used — if this data is ever needed again,
# the original label maps are straightforward to recreate from the
# same XML captures (Doc No./Sales No./Contact Person/Tel No (Mobile
# 1)/Tel No (Mobile 2)/Tel No (Office)/Email for Billing, keyed
# "billing_*"; Contact Name/Contact Number/Relation for Emergency,
# keyed "emergency_*"), following the same pattern INSTALL_LABEL_MAP
# below still uses.

INSTALL_LABEL_MAP = {
    "Contact Person": "install_contact_person",
    "Tel No (Mobile 1)": "install_mobile1",
    "Tel No (Mobile 2)": "install_mobile2",
    "Tel No (House)": "install_house_phone",
    "Tel No (Office)": "install_office_phone",
    "Email": "install_email",
}
INSTALL_ORPHAN_FIELDS = ["install_address"]

# Only the Installation/Service section's own header needs skipping —
# pair_fields() is only ever called for THIS section now (Billing/
# Emergency are no longer scraped, see the note above), and
# get_installation_elements()'s XPath includes this header text as one
# of the TextViews it returns (it's a descendant of the section's own
# parent ViewGroup), so it has to be filtered out or it'd be
# misread as an orphan field.
HEADER_TEXTS = {"Installation/Service Address & Contact"}

# --- Detail screen: Sales Info tab (confirmed from XML) ---
SALES_INFO_LABEL_MAP = {
    "Sales No.": "sales_info_sales_no",
    "Current Stage": "current_stage",
    "Product": "sales_info_product",
    "OutRight Price": "outright_price",
    "Sales Date": "sales_date",
    "Rental Fees (Monthly)": "rental_fee_monthly",
    "Rental Processing Fees": "rental_processing_fee",
    "Application Type": "application_type",
    "Sales Status": "sales_status",
    "PO Number": "po_number",
    "Promo Code": "promo_code",
    "Rental Scheme": "rental_scheme",
}


# TEMPLATE_FILE: if this .xlsm exists (created once, manually — see
# excel_export.write_output()'s docstring for the one-time setup), the
# export writes INTO it instead, preserving its macro so the WhatsApp
# Link column can double-click-open a chat with the full message
# pre-filled, sidestepping HYPERLINK()'s 255-char limit. If it doesn't
# exist, everything still works — the export just falls back to a
# plain .xlsx with a click-to-open-chat-then-paste link instead. This
# is a SHARED file — one copy in the project's root folder, next to
# this script — NOT something that needs copying into each person's
# export folder.
TEMPLATE_FILE = "macro.xlsm"

# The actual per-run OUTPUT_FILE / OUTPUT_FILE_XLSM / CARRYFORWARD_SOURCE_FILE
# paths live in filenames.py, not here — those get REASSIGNED while the
# script runs (built fresh from the specialist's identity + month), so
# they need to live alongside the functions that mutate them rather
# than sitting here as if they were fixed settings.
IDENTITY_CONFIG_FILE = "specialist_identity.json"


WAIT_SECONDS = 10
# raised since smaller, controlled scroll steps need more of them to reach the bottom
MAX_SCROLLS = 300
MAX_STAGNANT_ROUNDS = 2
# fallback step, only used when there's no measured position to target yet (e.g. the very first scroll)
SCROLL_STEP_PERCENT = 0.32
SCROLL_REGION_TOP_FRACTION = 0.40     # stays below the pinned filter header
SCROLL_REGION_HEIGHT_FRACTION = 0.48
# How long to pause after each scroll gesture for the view to settle
# before reading it. Pulled out as its own knob (rather than a number
# buried inside scroll_down) so it's easy to try lowering during
# testing without hunting through the function — just watch for
# misreads (skipped/duplicated customers) if you push it much lower.
SCROLL_SETTLE_SECONDS = 0.6


CCS_DEDUP_FIELDS = ("product", "last_change")
# Safety cap on scrolling WITHIN one customer's CCS Note screen — same
# spirit as MAX_SCROLLS above, just a separate, smaller budget since
# this screen has far fewer cards than the full customer list does.
MAX_CCS_CARD_SCROLLS = 30

_CCS_RESERVED_LABELS = {"Last Change", "Next Change"}


ALL_COLUMNS = [
    "Sales No.", "Installation / Service Contact Person",
    "Installation / Service Address",
    "Area 1 (State)", "Area 2 (District / City)", "Area 3 (Locality)", "Area 4 (Street)",
    "Appointment Date", "Proposed Date", "WhatsApp Chat Link",
    "WhatsApp Chat Message", "Product", "Filter(s)", "WhatsApp Number",
]

_COLUMN_INDEX = {name: idx for idx, name in enumerate(ALL_COLUMNS, start=1)}
