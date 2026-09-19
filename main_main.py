# """
# Cuckoo+ Service Specialist — monthly list scraper (Appium / UiAutomator2)

# CONFIRMED SCREEN FLOW (from real XML captures)
# --------------------------------------------------
# 1. LIST screen: each customer is a card with label/value pairs (NS No,
#    Sales No, NS Date, Status, Appt Date, Product Name, Cust Name) plus
#    one button.
# 2. Tapping that button does NOT open the detail screen directly — it
#    opens a "Choose Option" POPUP with several choices (View Order, CCS
#    Note, Appointment, Contact List, Cancel Appointment, Cancel). The
#    script taps "View Order" from that popup.
# 3. That opens a "Customer Information" screen with TWO TABS:
#      - "Address/Contact Info" (shown by default) — has THREE sections:
#        Billing Address & Contact, Installation/Service Address &
#        Contact, and Emergency Contact. The first two have normal
#        label/value pairs PLUS a few fields with no visible label at
#        all (customer name, a masked ID number, and the address) —
#        these are identified by position instead of a label. Emergency
#        Contact is fully labeled, no unlabeled fields.
#      - "Sales Info" — reached by tapping its tab header. Normal
#        label/value pairs, no unlabeled fields.
# 4. One driver.back() from the detail screen returns to the list.

# HOW LABEL/VALUE PAIRING WORKS
# ---------------------------------
# On every screen here, a label sits to the LEFT of its value, and both
# share the same vertical position (y-coordinate) — but the raw reading
# order in the XML doesn't alternate label-value cleanly, and the exact
# pixel positions differ between devices/screen resolutions. So instead
# of matching on fixed x-coordinates (which broke the first time this
# was tested on a different device — a 720x1520 screen instead of the
# 1600x2560 one the fields were originally mapped from), the script
# groups elements into rows by shared y-coordinate, then within each row
# takes the leftmost text as the label and the rightmost as its value.
# This holds regardless of screen size. Unlabeled fields (no matching
# text in the known label list) are collected separately, in top-to-
# bottom order, and assigned fixed names based on where they fall in
# that section.

# ONE THING WORTH NOTING
# --------------------------
# The list screen's "Product Name" field (e.g. "CP-XN501HW") is the
# product model/SKU. The Sales Info tab's "Product" field (e.g. "XCEL")
# is a different, separate product-name value from that section of the
# app — both are genuine product names from different parts of the
# record, not a mislabeled dealer code.

# SKIP-KNOWN-CUSTOMERS NOTE
# --------------------------
# run() now skips re-scraping a customer whose sales_no is already in a
# previous export — their details don't change once captured, so
# there's nothing new to read there. Only genuinely new sales_no values
# get their View Order/CCS Note screens opened. Known customers are
# found wherever they appear while scrolling (not assumed to only be at
# the bottom) and carried forward into the new export exactly as they
# were before, with no fields re-read. See _load_carryforward_data() and
# the "skipping ... — already captured previously" check inside run().

# SPEED NOTE (known-customer scanning)
# --------------------------
# get_visible_customers_stable() waits for two consecutive matching
# reads before trusting what's on screen — a real protection against
# Android reusing list rows mid-scroll (a row can briefly show one
# customer's stale fields under another's already-updated NS No). That
# protection only matters for a card the script is about to actually
# SKIP or ACT ON. get_visible_customers_quick_or_stable() below does one
# fast read first; if every customer currently visible is either already
# handled this run or already known from a previous export, that fast
# read is trusted as-is (nothing about it is being acted on, so a stale
# field elsewhere on the card can't cause a wrong decision) and the slow
# stable protocol is skipped entirely. The moment anything new or
# incomplete shows up, it falls straight back to the full stable read
# before touching it. On a run where most customers are already known
# (the normal case after the first month), this removes a lot of the
# repeated ~0.4s settle waits that used to happen purely while scrolling
# past people who were only ever going to be skipped anyway.

# OUTPUT NOTE
# --------------------------
# The scraping/scrolling logic ABOVE this note is unchanged. The export
# step only includes the fields actually needed (sales_no,
# sales_info_product, install_contact_person, install_mobile1,
# install_address), plus proposed_date (typed by hand),
# a WhatsApp Message column (auto-fills once a date is typed, formatted
# to match the WhatsApp bold style: *Hanis*, *Tarikh:*, *Alamat:*,
# *Produk:*, *Nombor Pesanan:*), and a WhatsApp Link column — a tap-to-
# send wa.me link with that message already filled in via URL, so
# there's nothing to copy/paste by hand. This has to be .xlsx rather
# than .csv, since a plain CSV can't hold a live formula.

# Start with LIST_LIMIT = 3. Once the export looks right, set it to
# None to process the entire list.
# """

# import os
# import re
# import sys
# import time
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment
# from openpyxl.comments import Comment
# from appium import webdriver
# from appium.options.android import UiAutomator2Options
# from appium.webdriver.common.appiumby import AppiumBy
# from selenium.common.exceptions import NoSuchElementException, TimeoutException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC


# # ============================================================
# # CONFIG
# # ============================================================

# APP_PACKAGE = "cuckoo.doctress"
# # confirmed: the list screen
# APP_ACTIVITY = "cuckoo.doctress.naturalcareservicelist"

# LIST_LIMIT = None   # was 3 for testing — now processes the whole list
# # set True again only if something breaks and you need to see raw element data
# DEBUG = False

# # --- List screen (confirmed from XML) ---
# LABEL_FIELD_MAP = {
#     "NS No": "ns_no",
#     "Sales No": "sales_no",
#     "NS Date": "ns_date",
#     "Status": "status",
#     "Appt Date": "appt_date",
#     "Product Name": "product_name",
#     "Cust Name": "cust_name",
# }
# KEY_FIELD = "ns_no"   # unique per card — avoids re-scraping someone after scrolling
# # a card missing any of these is
# REQUIRED_LIST_FIELDS = set(LABEL_FIELD_MAP.values())
# # probably cut off by the screen
# # edge, not actually incomplete data

# # --- Detail screen: Address/Contact Info tab (confirmed from XML) ---
# #
# # NOTE: this screen also has a "Billing Address & Contact" section and
# # an "Emergency Contact" section, both fully labeled/parseable the same
# # way as everything else here — but neither one is scraped anymore.
# # Verified by cross-referencing every field this script could produce
# # against every place it's actually read downstream: nothing in the
# # export, in carry-forward, or anywhere else ever reads a billing_* or
# # emergency_* field. Scraping them cost two full extra device round
# # trips per newly-scraped customer (get_billing_elements() and
# # get_emergency_elements() each did their own wait_for + find_elements)
# # for data that went nowhere. Cut for real runtime savings, with zero
# # risk to anything actually used — if this data is ever needed again,
# # the original label maps are straightforward to recreate from the
# # same XML captures (Doc No./Sales No./Contact Person/Tel No (Mobile
# # 1)/Tel No (Mobile 2)/Tel No (Office)/Email for Billing, keyed
# # "billing_*"; Contact Name/Contact Number/Relation for Emergency,
# # keyed "emergency_*"), following the same pattern INSTALL_LABEL_MAP
# # below still uses.

# INSTALL_LABEL_MAP = {
#     "Contact Person": "install_contact_person",
#     "Tel No (Mobile 1)": "install_mobile1",
#     "Tel No (Mobile 2)": "install_mobile2",
#     "Tel No (House)": "install_house_phone",
#     "Tel No (Office)": "install_office_phone",
#     "Email": "install_email",
# }
# INSTALL_ORPHAN_FIELDS = ["install_address"]

# # Only the Installation/Service section's own header needs skipping —
# # pair_fields() is only ever called for THIS section now (Billing/
# # Emergency are no longer scraped, see the note above), and
# # get_installation_elements()'s XPath includes this header text as one
# # of the TextViews it returns (it's a descendant of the section's own
# # parent ViewGroup), so it has to be filtered out or it'd be
# # misread as an orphan field.
# HEADER_TEXTS = {"Installation/Service Address & Contact"}

# # --- Detail screen: Sales Info tab (confirmed from XML) ---
# SALES_INFO_LABEL_MAP = {
#     "Sales No.": "sales_info_sales_no",
#     "Current Stage": "current_stage",
#     "Product": "sales_info_product",
#     "OutRight Price": "outright_price",
#     "Sales Date": "sales_date",
#     "Rental Fees (Monthly)": "rental_fee_monthly",
#     "Rental Processing Fees": "rental_processing_fee",
#     "Application Type": "application_type",
#     "Sales Status": "sales_status",
#     "PO Number": "po_number",
#     "Promo Code": "promo_code",
#     "Rental Scheme": "rental_scheme",
# }

# OUTPUT_FILE = "cuckoo_export.xlsx"
# # Optional: if this .xlsm exists (created once, manually — see the
# # write_output()/wa_link_formula() docstrings below for the one-time
# # setup), the export writes INTO it instead, preserving its macro so
# # the WhatsApp Link column can double-click-open a chat with the full
# # message pre-filled, sidestepping HYPERLINK()'s 255-char limit. If it
# # doesn't exist, everything still works — the export just falls back
# # to the plain .xlsx with a click-to-open-chat-then-paste link instead.
# # This is a SHARED file — one copy in the project's root folder, next
# # to this script — NOT something that needs copying into each
# # person's export folder; write_output() always reads it from here
# # regardless of which NDS-ID/name subfolder this run is writing into.
# TEMPLATE_FILE = "macro.xlsm"
# OUTPUT_FILE_XLSM = "cuckoo_export.xlsm"

# # OUTPUT_FILE / OUTPUT_FILE_XLSM above are just fallback defaults —
# # run() and resort_existing_file() both overwrite them at startup with
# # paths built from the specialist's identity + the month they typed in
# # (see _apply_identity_to_filenames()), e.g.
# # "NDS35095_Hanis/NDS35095_Hanis_September_v1_export.xlsm". This is
# # what lets colleagues share this same script safely: each person's
# # export lands in their own folder rather than everyone overwriting
# # one shared file.
# IDENTITY_CONFIG_FILE = "specialist_identity.json"

# WAIT_SECONDS = 10
# # raised since smaller, controlled scroll steps need more of them to reach the bottom
# MAX_SCROLLS = 300
# MAX_STAGNANT_ROUNDS = 2
# # fallback step, only used when there's no measured position to target yet (e.g. the very first scroll)
# SCROLL_STEP_PERCENT = 0.32
# SCROLL_REGION_TOP_FRACTION = 0.40     # stays below the pinned filter header
# SCROLL_REGION_HEIGHT_FRACTION = 0.48
# # How long to pause after each scroll gesture for the view to settle
# # before reading it. Pulled out as its own knob (rather than a number
# # buried inside scroll_down) so it's easy to try lowering during
# # testing without hunting through the function — just watch for
# # misreads (skipped/duplicated customers) if you push it much lower.
# SCROLL_SETTLE_SECONDS = 0.6

# # --- Settle-time tuning telemetry ---
# # get_visible_customers_stable() waits for two consecutive matching
# # reads before trusting the screen (see its own docstring) — if the
# # very first read is already stable, that's 1 "attempt". If the screen
# # hadn't finished rendering yet, it takes 2+ attempts, each one costing
# # an extra settle_delay wait. These three counters track how often that
# # happens across a whole run, purely so SCROLL_SETTLE_SECONDS can be
# # tuned from real evidence instead of a guess — see the summary printed
# # at the end of run() (search "Screen-settle check" below).
# #
# # Why this matters even though get_visible_customers_stable() already
# # protects itself: get_visible_customers_quick_or_stable() — the FAST
# # path used for most customers in a normal run, since most are already
# # known — does exactly ONE read with no retry at all. It has no safety
# # net of its own. These counters don't measure the quick path directly
# # (it wouldn't have anything meaningful to count, by design), but they
# # ARE a proxy for it: if the careful, self-correcting path is
# # frequently needing 2+ attempts to get a stable read at the current
# # SCROLL_SETTLE_SECONDS, that means this phone's rendering is already
# # close to that time limit — which means the fast path (with no retry
# # to fall back on) is riding on a thin margin too. A run where the
# # average stays at (or very near) 1.00 is real evidence there's slack
# # to lower SCROLL_SETTLE_SECONDS a bit; an average noticeably above
# # 1.00 is a sign NOT to lower it further, and possibly to raise it.
# _STABLE_READ_TOTAL_CALLS = 0
# _STABLE_READ_TOTAL_ATTEMPTS = 0
# _STABLE_READ_GAVE_UP_COUNT = 0


# # ============================================================
# # Area auto-fill (free, offline — postcode lookup + keyword matching)
# # ============================================================
# #
# # This is a SEPARATE, self-contained block from the scraping/scrolling
# # code above and below it — it doesn't touch, call, or depend on any
# # of that. It only produces suggested values for the Area 1-4 columns,
# # which write_output() then applies with one rule: NEVER overwrite a
# # cell that already has something in it (typed by hand OR auto-filled
# # on an earlier run). Only a cell that's still genuinely blank gets a
# # new auto-filled value. Tested against a real export and a batch of
# # deliberately messy addresses before being built in here. Known
# # limitation, accepted on purpose: this will leave plenty of cells
# # blank for you to fill in by hand, especially Area 3/Area 4 on
# # addresses with no commas — that's the safe trade-off on purpose,
# # not a bug.
# #
# # POSTCODE_LOOKUP_FILE is a free, offline dataset (Malaysia postcodes
# # -> state/city, from a public GitHub repo, saved locally as
# # postcodes_my.json next to this script) — no internet call at
# # runtime, no API key, no cost. If that file is ever missing, Area 1/
# # Area 2 just won't auto-fill (everything else still works normally).

# POSTCODE_LOOKUP_FILE = "postcodes_my.json"

# # Locality-type words that can appear BEFORE the name they describe
# # (the usual Malay convention, e.g. "Taman Wangsa Melawati").
# LOCALITY_KEYWORDS = [
#     "Taman", "Kampung", "Kg", "Bandar", "Presint", "Precint",
#     "Seksyen", "Desa", "Pangsapuri", "Kondominium", "Residensi",
# ]
# # Road-type words — same "keyword before the name" pattern.
# STREET_KEYWORDS = ["Jalan", "Jln", "Lorong", "Persiaran", "Lebuh", "Lebuhraya"]

# # A match this long is almost certainly two different things run
# # together (e.g. a street name merged with the next town's name because
# # there was no comma to stop at) — confirmed by testing. Past this
# # length, the match is treated as unreliable and dropped rather than
# # risking a wrong value sitting in the sheet looking legit.
# _MAX_CONFIDENT_MATCH_CHARS = 45


# def _load_postcode_lookup(path=POSTCODE_LOOKUP_FILE):
#     """Loads the free Malaysia postcode -> (state, city) dataset. Returns
#     {} (and prints a one-time warning) if the file isn't there — Area 1/
#     Area 2 simply won't auto-fill in that case, nothing else breaks."""
#     if not os.path.exists(path):
#         print(f"  !! {path} not found next to the script — Area 1 (State) "
#               f"and Area 2 (District / City) won't auto-fill this run. "
#               f"Everything else still works normally.")
#         return {}
#     try:
#         import json
#         with open(path, "r", encoding="utf-8") as f:
#             data = json.load(f)
#         lookup = {}
#         for state in data.get("state", []):
#             for city in state.get("city", []):
#                 for pc in city.get("postcode", []):
#                     lookup[pc] = (state["name"], city["name"])
#         return lookup
#     except Exception as e:
#         print(f"  !! Could not read {path} ({e}) — Area 1/Area 2 won't "
#               f"auto-fill this run.")
#         return {}


# _POSTCODE_LOOKUP = _load_postcode_lookup()

# _LOCALITY_PATTERN = re.compile(
#     r"\b(" + "|".join(re.escape(k) for k in LOCALITY_KEYWORDS) + r")\.?\s+"
#     r"[A-Za-z0-9'\-\s]*?(?=,|\d{5}|$)",
#     re.IGNORECASE,
# )
# _STREET_PATTERN = re.compile(
#     r"\b(" + "|".join(re.escape(k) for k in STREET_KEYWORDS) + r")\.?\s+"
#     r"[A-Za-z0-9'\-/\s]+?(?=,|\d{5}|$)",
#     re.IGNORECASE,
# )


# def detect_areas(address):
#     """
#     Takes one raw address string and returns whatever it can confidently
#     work out, as {"area1": ..., "area2": ..., "area3": ..., "area4": ...}.
#     Any level it's not confident about is left as "" — by design, not
#     left for this function to guess at.

#       area1 = State       )  both from the postcode found in the address,
#       area2 = District/City) looked up in the free postcode dataset
#       area3 = Locality    -  from a keyword match (Taman/Kampung/Presint/...)
#       area4 = Street      -  from a keyword match (Jalan/Lorong/...)

#     Area 3/Area 4 are only kept if the match is short enough to be
#     trustworthy (see _MAX_CONFIDENT_MATCH_CHARS) — a long match usually
#     means it accidentally swallowed extra, unrelated text because the
#     address had no comma to stop it cleanly.
#     """
#     result = {"area1": "", "area2": "", "area3": "", "area4": ""}
#     if not address:
#         return result

#     postcode_match = re.search(r"\b(\d{5})\b", address)
#     if postcode_match:
#         postcode = postcode_match.group(1)
#         if postcode in _POSTCODE_LOOKUP:
#             state, city = _POSTCODE_LOOKUP[postcode]
#             result["area1"] = state
#             result["area2"] = city

#     locality_match = _LOCALITY_PATTERN.search(address)
#     if locality_match:
#         text = locality_match.group(0).strip().rstrip(",")
#         if len(text) <= _MAX_CONFIDENT_MATCH_CHARS:
#             result["area3"] = text

#     street_match = _STREET_PATTERN.search(address)
#     if street_match:
#         text = street_match.group(0).strip().rstrip(",")
#         if len(text) <= _MAX_CONFIDENT_MATCH_CHARS:
#             result["area4"] = text

#     return result


# # ============================================================
# # Low-level helpers
# # ============================================================

# def build_driver():
#     options = UiAutomator2Options()
#     options.platform_name = "Android"
#     options.app_package = APP_PACKAGE
#     options.app_activity = APP_ACTIVITY
#     options.no_reset = True
#     # Appium auto-terminates a session after this many seconds with NO
#     # commands sent to the device — 60s by default. The pause feature
#     # deliberately sends nothing to the phone for as long as it's
#     # paused, which is exactly what trips this: a quick test pause
#     # (a few seconds) resumes fine, but a real pause (actually using
#     # the phone for a bit) exceeds 60s and gets the session killed
#     # server-side — confirmed by testing: this is exactly what an
#     # InvalidSessionIdException right after "Resuming..." means. Set
#     # generously high (1 hour) so a genuine break doesn't trigger it;
#     # this doesn't weaken any other safety net — MAX_SCROLLS and the
#     # KeyboardInterrupt/Exception handling still catch a run that's
#     # actually stuck for an unrelated reason.
#     options.new_command_timeout = 3600
#     return webdriver.Remote("http://127.0.0.1:4723", options=options)


# def wait_for(driver, selector, timeout=WAIT_SECONDS):
#     by, value = selector
#     return WebDriverWait(driver, timeout).until(
#         EC.presence_of_element_located((by, value))
#     )


# def read_text_safe(el):
#     try:
#         return el.text.strip()
#     except Exception:
#         return ""


# def parse_bounds(bounds_str):
#     """'[x1,y1][x2,y2]' -> (x1, y1, x2, y2), or None if unparseable."""
#     if not bounds_str:
#         return None
#     m = re.match(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", bounds_str)
#     if not m:
#         return None
#     return tuple(int(g) for g in m.groups())


# def tap_element(driver, element):
#     """
#     Taps by screen coordinates instead of calling .click() directly —
#     some Appium-Python-Client / Selenium version combinations throw
#     "Wrong parameters applied for elementClick" on UiAutomator2. This
#     sidesteps that bug entirely.
#     """
#     parsed = parse_bounds(element.get_attribute("bounds"))
#     if not parsed:
#         raise RuntimeError(
#             "Could not read bounds — can't compute a tap point for this element.")
#     x1, y1, x2, y2 = parsed
#     cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
#     driver.execute_script("mobile: clickGesture", {"x": cx, "y": cy})


# def scroll_down(driver, percent=None):
#     """
#     Uses "mobile: scrollGesture" rather than "mobile: swipeGesture" —
#     swipeGesture performs a FLING, which has momentum: Android keeps
#     scrolling after the simulated finger lifts, and how far it travels
#     depends on gesture velocity in a way that's hard to predict. That
#     inconsistency was very likely why scrolling sometimes jumped past
#     several customers at once. scrollGesture instead moves a fixed,
#     controlled fraction of the given area with no fling.

#     `percent`, when given, comes from compute_scroll_percent() — a
#     measured amount based on exactly where the last already-processed
#     customer's card sits on screen right now, rather than a guessed
#     fixed distance. Falls back to SCROLL_STEP_PERCENT only when there's
#     no measurement to base it on yet (the very first scroll).

#     Note: for scrollGesture, "direction" describes which way the
#     CONTENT moves (not the simulated finger) — "down" reveals further/
#     later items in the list, which is what swipeGesture called "up".
#     """
#     if percent is None:
#         percent = SCROLL_STEP_PERCENT
#     size = driver.get_window_size()
#     width, height = size["width"], size["height"]
#     driver.execute_script("mobile: scrollGesture", {
#         "left": int(width * 0.1),
#         "top": int(height * SCROLL_REGION_TOP_FRACTION),
#         "width": int(width * 0.8),
#         "height": int(height * SCROLL_REGION_HEIGHT_FRACTION),
#         "direction": "down",
#         "percent": percent,
#     })
#     time.sleep(SCROLL_SETTLE_SECONDS)
#     # DISABLED (not deleted) — this was a defensive check for a scroll
#     # gesture accidentally landing on the list screen's pinned "NS
#     # Number."/"Sales No." search boxes and popping the keyboard, which
#     # would throw off every element's y-coordinate for the row-grouping
#     # logic elsewhere in this file. Confirmed it's never actually
#     # triggered in real runs, and this script currently only ever runs
#     # on one known device — so the extra Appium round trip on every
#     # single scroll (there are a LOT of these across a full run,
#     # especially inside the CCS Note scrolling loop) isn't buying
#     # anything right now. If this ever runs on a DIFFERENT phone
#     # (different screen size/pinned-header height — e.g. if a
#     # colleague starts using this), that's exactly the situation this
#     # was guarding against, so re-enable it first: just uncomment the
#     # line below.
#     # dismiss_keyboard_if_present(driver)


# def compute_scroll_percent(driver, target_top_y, margin_fraction=0.04):
#     """
#     Computes the exact scroll fraction needed to bring `target_top_y`
#     (a real, measured card position) up near the top of the scroll
#     region — "scroll to the line separating this card from the next,"
#     rather than a blind fixed distance.

#     A small margin is subtracted so the target lands just inside the
#     visible region instead of exactly on the edge, avoiding the
#     partially-rendered-row problem.
#     """
#     size = driver.get_window_size()
#     height = size["height"]
#     region_top = height * SCROLL_REGION_TOP_FRACTION
#     region_height = height * SCROLL_REGION_HEIGHT_FRACTION

#     desired_shift = (target_top_y - region_top) - \
#         (margin_fraction * region_height)
#     percent = desired_shift / region_height

#     # Guardrails: never scroll by ~nothing (no progress) or overshoot
#     # past the measured area (which would defeat the point of measuring).
#     return max(0.08, min(0.95, percent))


# def dismiss_keyboard_if_present(driver):
#     """Defensive safety net — closes the keyboard if anything ever
#     accidentally focuses a text field again, so it doesn't silently
#     corrupt the next read."""
#     try:
#         driver.execute_script("mobile: hideKeyboard")
#     except Exception:
#         pass  # no keyboard was showing, or the command isn't supported — fine either way


# def pair_fields(elements, known_labels=None, skip_texts=None, y_tol=8):
#     """
#     Generic label/value pairing based on RELATIVE position, not fixed
#     pixel coordinates — this makes it work across different screen
#     resolutions/devices, since it never assumes a label sits at a
#     specific x value. Elements are grouped into rows by shared
#     y-coordinate; within each row, the first element (left to right)
#     whose text matches a KNOWN label is treated as the label, and the
#     element immediately after it becomes its value. Anything sitting
#     further left than the label (e.g. a small numbered badge next to
#     each card) is simply ignored rather than mistaken for the label.

#     Returns (row, orphans):
#       row     -> {label_text: value_text} for every recognized label
#       orphans -> list of value texts (top-to-bottom order) for rows
#                  that don't match a known label — i.e. unlabeled
#                  fields like a bare address or name with no caption.
#     """
#     skip_texts = skip_texts or set()
#     entries = []
#     for el in elements:
#         parsed = parse_bounds(el.get_attribute("bounds"))
#         if not parsed:
#             continue
#         x1, y1, _, _ = parsed
#         text = read_text_safe(el)
#         if text in skip_texts:
#             continue
#         entries.append((y1, x1, text))

#     entries.sort(key=lambda e: (e[0], e[1]))

#     rows = []
#     for entry in entries:
#         placed = False
#         for row in rows:
#             if abs(row[0][0] - entry[0]) <= y_tol:
#                 row.append(entry)
#                 placed = True
#                 break
#         if not placed:
#             rows.append([entry])

#     row_dict = {}
#     orphans = []
#     for row in rows:
#         row_sorted = sorted(row, key=lambda e: e[1])  # left to right
#         row_y = row_sorted[0][0]

#         if known_labels is not None:
#             # Find the first element in this row whose text is a genuine
#             # known label — this way something sitting further left (like
#             # a small numbered badge next to each card) gets ignored
#             # instead of being mistaken for the label itself.
#             label_idx = next((i for i, e in enumerate(
#                 row_sorted) if e[2] in known_labels), None)
#             if label_idx is not None:
#                 label_text = row_sorted[label_idx][2]
#                 value_text = row_sorted[label_idx +
#                                         1][2] if label_idx + 1 < len(row_sorted) else ""
#                 row_dict[label_text] = value_text
#                 # anything before label_idx (e.g. a badge number) is just ignored
#             else:
#                 for _, _, text in row_sorted:
#                     orphans.append((row_y, text))
#         else:
#             if len(row_sorted) >= 2:
#                 row_dict[row_sorted[0][2]] = row_sorted[-1][2]
#             else:
#                 orphans.append((row_y, row_sorted[0][2]))

#     orphans.sort(key=lambda t: t[0])
#     return row_dict, [t for _, t in orphans]


# # ============================================================
# # List screen
# # ============================================================

# def group_into_rows(entries, y_tol=8):
#     """entries: list of (y1, x1, text). Returns rows: list of rows,
#     each row a list of (y1, x1, text) sorted left-to-right, rows
#     themselves sorted top-to-bottom."""
#     entries = sorted(entries, key=lambda e: (e[0], e[1]))
#     rows = []
#     for e in entries:
#         placed = False
#         for row in rows:
#             if abs(row[0][0] - e[0]) <= y_tol:
#                 row.append(e)
#                 placed = True
#                 break
#         if not placed:
#             rows.append([e])
#     rows = [sorted(r, key=lambda e: e[1]) for r in rows]
#     rows.sort(key=lambda r: r[0][0])
#     return rows


# def rows_to_fields(rows_chunk, known_labels):
#     """Same label-search-per-row logic as pair_fields, applied to an
#     already-sliced chunk of rows belonging to one customer."""
#     result = {}
#     for row in rows_chunk:
#         label_idx = next((i for i, e in enumerate(
#             row) if e[2] in known_labels), None)
#         if label_idx is not None:
#             label_text = row[label_idx][2]
#             value_text = row[label_idx + 1][2] if label_idx + \
#                 1 < len(row) else ""
#             result[label_text] = value_text
#     return result


# def get_visible_customers(driver):
#     """
#     Reads EVERY TextView and Button on screen in one global query (no
#     per-card scoped searches — those don't reliably scope on this
#     Appium/UiAutomator2 setup, which is what broke the earlier version),
#     then figures out which elements belong to which customer purely by
#     on-screen position: each "NS No" row marks where a new card starts.

#     Returns a list of {"row": {...parsed fields...}, "button": element_or_None}
#     for every customer currently visible.
#     """
#     wait_for(driver, (AppiumBy.XPATH,
#              '//android.widget.TextView[@text="NS No"]'))

#     text_elements = driver.find_elements(
#         AppiumBy.CLASS_NAME, "android.widget.TextView")
#     entries = []
#     for el in text_elements:
#         parsed = parse_bounds(el.get_attribute("bounds"))
#         if not parsed:
#             continue
#         x1, y1, _, _ = parsed
#         entries.append((y1, x1, read_text_safe(el)))
#     rows = group_into_rows(entries)

#     marker_indices = [i for i, row in enumerate(
#         rows) if any(t == "NS No" for _, _, t in row)]
#     if DEBUG:
#         print(
#             f"  [debug] found {len(rows)} row(s) total, {len(marker_indices)} 'NS No' marker(s)")

#     buttons = driver.find_elements(AppiumBy.XPATH, '//android.widget.Button')
#     button_positions = []
#     for b in buttons:
#         parsed = parse_bounds(b.get_attribute("bounds"))
#         if parsed:
#             button_positions.append((parsed[1], b))
#     button_positions.sort(key=lambda t: t[0])

#     known_labels = set(LABEL_FIELD_MAP.keys())
#     customers = []
#     for idx, start in enumerate(marker_indices):
#         end = marker_indices[idx + 1] if idx + \
#             1 < len(marker_indices) else len(rows)
#         chunk = rows[start:end]
#         card_top_y = chunk[0][0][0]
#         next_top_y = rows[marker_indices[idx + 1]][0][0] if idx + \
#             1 < len(marker_indices) else float("inf")

#         field_dict = rows_to_fields(chunk, known_labels)
#         parsed_row = {LABEL_FIELD_MAP[k]: v for k,
#                       v in field_dict.items() if k in LABEL_FIELD_MAP}

#         matching_button = next(
#             (b_el for b_y, b_el in button_positions if card_top_y <= b_y < next_top_y), None)
#         # A card missing any expected field (most commonly Cust Name,
#         # since it's the LAST field on the card) usually means it's only
#         # partially scrolled into view — its bottom hasn't fully rendered
#         # yet, not that the data is genuinely blank.
#         complete = len(field_dict) == len(LABEL_FIELD_MAP)
#         customers.append({
#             "row": parsed_row,
#             "button": matching_button,
#             "card_top_y": card_top_y,
#             "complete": complete,
#         })

#     if DEBUG:
#         for c in customers:
#             print(
#                 f"  [debug] parsed customer: {c['row']}  (button found: {c['button'] is not None})")

#     return customers


# def get_visible_customers_stable(driver, max_attempts=4, settle_delay=0.4):
#     """
#     Reads the screen repeatedly until two consecutive reads agree on
#     which NS numbers are visible. A single read can catch the view
#     mid-render — Android list rows get REUSED as you scroll, so reading
#     too early can show a row still holding the PREVIOUS customer's name
#     while its NS No has already updated to the new one. That's what was
#     causing blank names and, worse, a name attached to the wrong NS
#     number. Waiting for two matching reads in a row avoids trusting a
#     transitional, half-updated state.
#     """
#     global _STABLE_READ_TOTAL_CALLS, _STABLE_READ_TOTAL_ATTEMPTS, _STABLE_READ_GAVE_UP_COUNT

#     prev_signature = None
#     customers = []
#     _STABLE_READ_TOTAL_CALLS += 1
#     for attempt_num in range(1, max_attempts + 1):
#         customers = get_visible_customers(driver)
#         signature = tuple(c["row"].get(KEY_FIELD, "") for c in customers)
#         if signature == prev_signature and signature:
#             _STABLE_READ_TOTAL_ATTEMPTS += attempt_num
#             return customers
#         prev_signature = signature
#         time.sleep(settle_delay)
#     # Never got two matching reads in a row within max_attempts — the
#     # screen genuinely wouldn't settle this time. Counted separately
#     # from the normal attempt tally since this is a stronger signal
#     # than "needed a retry" — it's "retrying didn't even help."
#     _STABLE_READ_TOTAL_ATTEMPTS += max_attempts
#     _STABLE_READ_GAVE_UP_COUNT += 1
#     return customers


# def get_visible_customers_quick_or_stable(driver, known_sales_nos, seen_keys):
#     """
#     Speed optimization on top of get_visible_customers_stable() — see
#     the "SPEED NOTE" in this file's top docstring for the full
#     reasoning. Short version: the slow stable-read protocol only
#     matters for a card about to be SKIPPED or ACTED ON here. This does
#     one fast, single read first; if every customer currently visible is
#     either already handled this run (seen_keys) or already known from a
#     previous export (known_sales_nos), nothing about that card's other
#     fields is being trusted right now — only its identity, which is
#     what get_visible_customers_stable()'s own docstring confirms
#     updates promptly — so the fast read is used as-is. The instant
#     anything new or only-partially-rendered shows up, this falls back
#     to the full, careful stable read before touching it.
#     """
#     quick = get_visible_customers(driver)
#     if not quick:
#         return get_visible_customers_stable(driver)

#     for c in quick:
#         if not c.get("complete", True):
#             return get_visible_customers_stable(driver)
#         key = c["row"].get(KEY_FIELD)
#         if key in seen_keys:
#             continue
#         sales_no = _normalize_sales_no(c["row"].get("sales_no"))
#         if sales_no and sales_no in known_sales_nos:
#             continue
#         # Something here isn't already-handled/already-known — trust
#         # nothing from the fast read, do it properly.
#         return get_visible_customers_stable(driver)

#     return quick


# # ============================================================
# # Popup menu ("Choose Option") -> detail screen
# # ============================================================

# def select_popup_option(driver, option_text, timeout=WAIT_SECONDS):
#     # normalize-space() handles the extra leading spaces the app puts in these labels
#     xpath = f'//android.widget.TextView[normalize-space(@text)="{option_text}"]'
#     el = wait_for(driver, (AppiumBy.XPATH, xpath), timeout=timeout)
#     tap_element(driver, el)


# # ============================================================
# # Detail screen: Address/Contact Info tab
# # ============================================================

# def get_installation_elements(driver):
#     wait_for(driver, (AppiumBy.XPATH,
#              '//android.widget.TextView[@text="Installation/Service Address & Contact"]'))
#     return driver.find_elements(
#         AppiumBy.XPATH,
#         '//android.widget.TextView[@text="Installation/Service Address & Contact"]'
#         '/parent::android.view.ViewGroup//android.widget.TextView'
#     )


# def get_sales_info_elements(driver):
#     wait_for(driver, (AppiumBy.XPATH,
#              '//android.widget.TextView[@text="Current Stage"]'), timeout=8)
#     return driver.find_elements(AppiumBy.XPATH, '//android.widget.TextView')


# def read_full_detail(driver):
#     record = {}

#     # --- Address/Contact Info tab (shown by default) ---
#     # Billing Address & Contact and Emergency Contact are NOT scraped
#     # here (see the note above INSTALL_LABEL_MAP) — only the
#     # Installation/Service section is actually used downstream.
#     install_row, install_orphans = pair_fields(
#         get_installation_elements(driver),
#         known_labels=set(INSTALL_LABEL_MAP.keys()), skip_texts=HEADER_TEXTS
#     )
#     for label_text, field_name in INSTALL_LABEL_MAP.items():
#         record[field_name] = install_row.get(label_text, "")
#     for i, field_name in enumerate(INSTALL_ORPHAN_FIELDS):
#         record[field_name] = install_orphans[i] if i < len(
#             install_orphans) else ""

#     # --- Switch to Sales Info tab ---
#     sales_tab = wait_for(
#         driver, (AppiumBy.XPATH, '//android.widget.TextView[@text="Sales Info"]'))
#     tap_element(driver, sales_tab)
#     # get_sales_info_elements() below already waits/polls for "Current Stage"
#     # to appear, so no fixed sleep needed here.

#     sales_row, _ = pair_fields(
#         get_sales_info_elements(driver),
#         known_labels=set(SALES_INFO_LABEL_MAP.keys())
#     )
#     for label_text, field_name in SALES_INFO_LABEL_MAP.items():
#         record[field_name] = sales_row.get(label_text, "")

#     return record


# # ============================================================
# # Detail screen: CCS Note popup (filter/consumable change data)
# # ============================================================
# #
# # A scrollable list of "cards" — one per physical filter/consumable UNIT
# # installed, which is why the exact same product can appear several
# # times in a row (e.g. "FT-1001 Sediment Filter 8 Inch" showing up 6
# # times for one customer, because they have 6 identical units). Each
# # full card has: a product name, a small per-unit number (1, 2, 3...),
# # an interval like "(4 months)", "Last Change" -> a date, "Next Change"
# # -> a date. Only product name + Last Change date are actually kept —
# # interval and Next Change aren't part of what's needed here.
# #
# # A card is only kept if BOTH of these hold:
# #   - its product name is real text, not a bare number and not the
# #     literal label "Last Change"/"Next Change"
# #   - it has a non-empty Last Change date
# # Both checks exist because of a real, confirmed pattern in captured
# # data: alongside each genuine card, the same screen sometimes also
# # yields a partial/garbled read of it — either the standalone "Next
# # Change" label picked up on its own, or the small per-unit number
# # landing where the product name should be — and in every observed
# # case, that garbled read has a BLANK Last Change while the genuine
# # card next to it doesn't. Filtering on "has a real name AND has a
# # Last Change" reliably keeps the real entries and drops the artifacts.
# #
# # "Unique" cards: product name + Last Change date must both match for
# # two cards to be treated as duplicates and collapsed into one
# # (CCS_DEDUP_FIELDS below). The small per-unit number is deliberately
# # not part of that — this is what collapses several near-identical
# # "FT-1001 Sediment Filter 8 Inch" cards, differing only by unit
# # number, into a single row.

# CCS_DEDUP_FIELDS = ("product", "last_change")
# # Safety cap on scrolling WITHIN one customer's CCS Note screen — same
# # spirit as MAX_SCROLLS above, just a separate, smaller budget since
# # this screen has far fewer cards than the full customer list does.
# MAX_CCS_CARD_SCROLLS = 30

# _CCS_RESERVED_LABELS = {"Last Change", "Next Change"}


# def _is_valid_ccs_product_name(product):
#     if not product:
#         return False
#     if product in _CCS_RESERVED_LABELS:
#         return False
#     if product.strip().isdigit():
#         return False
#     return True


# def _group_ccs_texts_into_cards(card_boxes, text_entries):
#     """
#     Pure function — no driver access, so this is fully testable with
#     synthetic data before ever touching a real device. This is the
#     actual restructuring: instead of asking the phone "what text is
#     inside THIS card" once per card (get_ccs_note_cards() used to call
#     card_el.find_elements() in a loop — one extra device round trip per
#     visible card, every single scroll), every text element currently
#     visible across ALL cards is read in ONE bulk query, then sorted
#     into its owning card here, in Python, using nothing but on-screen
#     position — the same trick group_into_rows() already uses for the
#     main customer list.

#     card_boxes: list of (top_y, bottom_y) for each card container,
#                 in their original on-screen/document order — this
#                 order is preserved in the returned list, matching what
#                 get_ccs_note_cards() returned before.
#     text_entries: list of (y1, x1, text) for every text element
#                   currently visible in the card area, from ONE bulk
#                   read — not scoped to any particular card.

#     A text element is assigned to whichever card's [top_y, bottom_y]
#     range contains its own y1. Since real cards are stacked vertically
#     and don't overlap, this can't accidentally merge two cards' text
#     together the way a badly-tuned distance-based grouping might — it's
#     checking actual card boundaries, not guessing at spacing. A text
#     element that doesn't fall inside any card's range (e.g. stray text
#     from just above/below the visible card area) is simply dropped,
#     matching the original behavior of only reading text that was
#     inside a card element's own bounds.

#     Within each card, group_into_rows() reconstructs proper top-to-
#     bottom, left-to-right reading order from raw (y1, x1, text) tuples
#     — the exact same ordering guarantee this file already relies on
#     elsewhere (see pair_fields/get_visible_customers), rather than
#     trusting incidental element order from a query. Everything AFTER
#     that — pulling out the product name, finding the Last Change value,
#     the validity check — is character-for-character the same logic
#     get_ccs_note_cards() always used; only how the per-card text list
#     gets built has changed.
#     """
#     entries_by_card = [[] for _ in card_boxes]
#     # Check boxes top-to-bottom so a text element right on a boundary
#     # (shouldn't happen for real, non-overlapping cards, but cheap
#     # insurance) consistently resolves to the higher card rather than
#     # being ambiguous.
#     box_order = sorted(range(len(card_boxes)), key=lambda i: card_boxes[i][0])
#     for y1, x1, text in text_entries:
#         for i in box_order:
#             top_y, bottom_y = card_boxes[i]
#             # Half-open interval [top_y, bottom_y) — not <= on both ends.
#             # Two cards stacked with zero gap between them (normal for a
#             # scrollable list) can have one card's bottom_y exactly equal
#             # to the next card's top_y; with an inclusive upper bound, a
#             # text element sitting exactly on that shared line would
#             # match BOTH boxes and silently get assigned to the wrong
#             # (earlier) card. This guarantees every y-coordinate belongs
#             # to exactly one card.
#             if top_y <= y1 < bottom_y:
#                 entries_by_card[i].append((y1, x1, text))
#                 break

#     cards = []
#     for i, (top_y, _bottom_y) in enumerate(card_boxes):
#         rows = group_into_rows(entries_by_card[i])
#         texts = [t for row in rows for _, _, t in row if t]

#         if not texts:
#             continue

#         product = texts[0]
#         last_change = ""
#         for j, t in enumerate(texts):
#             if t == "Last Change" and j + 1 < len(texts):
#                 last_change = texts[j + 1]

#         if not _is_valid_ccs_product_name(product):
#             continue

#         cards.append({
#             "product": product,
#             "last_change": last_change,
#             "card_top_y": top_y,
#         })

#     return cards


# def get_ccs_note_cards(driver):
#     """Reads every currently-visible card on the CCS Note screen,
#     keeping only ones that pass the validity checks above. See
#     _group_ccs_texts_into_cards() for how the per-card split actually
#     works — this function just does the two device reads (card
#     container bounds, then every text element at once) and hands the
#     results to that pure function."""
#     wait_for(driver, (AppiumBy.XPATH,
#              '//android.widget.TextView[@text="CCS Note"]'))

#     card_elements = driver.find_elements(
#         AppiumBy.XPATH,
#         '//androidx.viewpager.widget.ViewPager//android.view.ViewGroup[@clickable="true"]'
#     )
#     card_boxes = []
#     for card_el in card_elements:
#         parsed = parse_bounds(card_el.get_attribute("bounds"))
#         if not parsed:
#             continue
#         _x1, y1, _x2, y2 = parsed
#         card_boxes.append((y1, y2))

#     if not card_boxes:
#         return []

#     # ONE bulk read for every text element currently visible across ALL
#     # cards — this replaces the old per-card find_elements() loop,
#     # which cost one extra device round trip per visible card.
#     text_elements = driver.find_elements(
#         AppiumBy.XPATH,
#         '//androidx.viewpager.widget.ViewPager//android.widget.TextView'
#     )
#     text_entries = []
#     for el in text_elements:
#         parsed = parse_bounds(el.get_attribute("bounds"))
#         if not parsed:
#             continue
#         x1, y1, _x2, _y2 = parsed
#         text = read_text_safe(el)
#         if text:
#             text_entries.append((y1, x1, text))

#     return _group_ccs_texts_into_cards(card_boxes, text_entries)


# def get_all_ccs_note_cards(driver):
#     """
#     Scrolls through the whole CCS Note screen, collecting cards as it
#     goes. Deduplicates along the way using CCS_DEDUP_FIELDS — this
#     collapses both genuinely repeated cards (same product/last change,
#     different unit number) AND the same card being seen twice after a
#     small scroll, with the same check. Finishes with a reconciliation
#     pass (see _reconcile_ccs_cards) that separates real "not yet
#     serviced" entries from parsing artifacts.
#     """
#     seen_keys = set()
#     unique_cards = []
#     scroll_count = 0
#     stagnant_rounds = 0

#     while scroll_count < MAX_CCS_CARD_SCROLLS and stagnant_rounds < MAX_STAGNANT_ROUNDS:
#         cards = get_ccs_note_cards(driver)
#         if not cards:
#             break

#         new_this_round = 0
#         for card in cards:
#             key = tuple(card[f] for f in CCS_DEDUP_FIELDS)
#             if key in seen_keys:
#                 continue
#             seen_keys.add(key)
#             unique_cards.append(card)
#             new_this_round += 1

#         stagnant_rounds = 0 if new_this_round > 0 else stagnant_rounds + 1

#         target_percent = compute_scroll_percent(
#             driver, cards[-1]["card_top_y"])
#         scroll_down(driver, percent=target_percent)
#         scroll_count += 1

#     return _reconcile_ccs_cards(unique_cards)


# def _reconcile_ccs_cards(cards):
#     """
#     Separates real "not yet serviced" entries from parsing artifacts —
#     both look identical in isolation (same product, blank Last Change),
#     so this has to look at each PRODUCT's entries together to tell them
#     apart:

#       - If a product has at least one entry WITH a Last Change date,
#         any blank-date entries for that same product are almost
#         certainly a partial/duplicate read of that same card (a
#         confirmed real pattern — see get_ccs_note_cards' docstring),
#         so they're dropped. Every distinct dated entry is kept (a
#         product can legitimately have more than one physical unit,
#         serviced on different dates).
#       - If a product has NO dated entry at all, it's kept as-is with a
#         blank date — that's a genuine filter that just hasn't had its
#         first change recorded yet, not an artifact.
#     """
#     by_product = {}
#     for card in cards:
#         by_product.setdefault(card["product"], []).append(card)

#     reconciled = []
#     for product, product_cards in by_product.items():
#         dated = [c for c in product_cards if c["last_change"]]
#         if dated:
#             reconciled.extend(dated)
#         else:
#             reconciled.append(product_cards[0])
#     return reconciled


# # ============================================================
# # Main loop: scroll + scrape until nothing new appears
# # ============================================================

# def _normalize_sales_no(value):
#     """
#     Excel can silently convert a Sales No. cell from text to a real
#     number — its own "Number Stored as Text" auto-fix does this with
#     one click, and even just re-typing a value can trigger it. Meanwhile
#     a live app scrape is ALWAYS a string (Appium's .text always returns
#     text, never a number). Without normalizing both sides to the same
#     type, "413643" (from the app) and 413643 (from Excel) are NOT equal
#     in Python — confirmed: this is exactly what caused a known, already-
#     captured customer to be silently re-scraped in full instead of
#     skipped, purely because of this type mismatch (nothing to do with
#     ordering or position). Every place a sales_no crosses the Excel <->
#     live-app boundary is normalized through this function first.
#     """
#     if value is None:
#         return ""
#     if isinstance(value, float) and value.is_integer():
#         # Excel sometimes stores a whole number as a float (413643.0) —
#         # strip the trailing .0 so it matches the plain-digit string a
#         # scrape would produce ("413643", not "413643.0").
#         return str(int(value))
#     return str(value).strip()


# def _normalize_phone_digits(value):
#     """
#     Same problem as _normalize_sales_no(), same fix — the hidden
#     WhatsApp Number column stores a plain digit string (e.g.
#     "60123456789"), and Excel can just as easily silently convert that
#     to a real number the same way it does Sales No. cells. Carrying
#     that forward without normalizing would risk clean_phone_for_wa()
#     choking on a non-string value on the next run.
#     """
#     if value is None:
#         return ""
#     if isinstance(value, float) and value.is_integer():
#         return str(int(value))
#     return str(value).strip()


# # ============================================================
# # Specialist identity + per-month export filenames
# # ============================================================
# #
# # This is what makes the script safe to hand to colleagues: the
# # WhatsApp message text used to hardcode "Hanis"/"NDS35095", and the
# # export always landed in one shared "cuckoo_export.xlsm" regardless
# # of who ran it or when. Now each person's Name + NDS Cuckoo ID is
# # asked for ONCE (cached locally — see IDENTITY_CONFIG_FILE) and used
# # both to personalize the message and to build a per-person, per-month
# # filename, so nobody's export or message text collides with anyone
# # else's.

# def _load_or_prompt_identity():
#     """
#     Asks which specialist's list this run is for — EVERY run, not just
#     the first — since this script might be run for a different person
#     entirely (e.g. covering a colleague's list), and silently reusing
#     whoever answered last would risk exporting/messaging under the
#     wrong name.

#     The last values used are read from IDENTITY_CONFIG_FILE and shown
#     as defaults (just press Enter to keep them), purely so running it
#     again and again for yourself doesn't mean retyping your own name
#     and ID every time. Whatever's answered — kept default or a new
#     value typed in — is saved back to that file so it becomes next
#     run's default.
#     """
#     import json

#     last_name, last_nds_id = "", ""
#     if os.path.exists(IDENTITY_CONFIG_FILE):
#         try:
#             with open(IDENTITY_CONFIG_FILE, "r", encoding="utf-8") as f:
#                 data = json.load(f)
#             last_name = (data.get("name") or "").strip()
#             last_nds_id = (data.get("nds_id") or "").strip()
#         except Exception:
#             pass  # missing/corrupted — just means no defaults to offer

#     name_prompt = (f"Specialist's name [{last_name}]: " if last_name
#                    else "Specialist's name (as it should appear in the WhatsApp message): ")
#     name = input(name_prompt).strip() or last_name

#     nds_prompt = (f"Specialist's NDS Cuckoo ID [{last_nds_id}]: " if last_nds_id
#                   else "Specialist's NDS Cuckoo ID (e.g. NDS35095): ")
#     nds_id = input(nds_prompt).strip() or last_nds_id

#     with open(IDENTITY_CONFIG_FILE, "w", encoding="utf-8") as f:
#         json.dump({"name": name, "nds_id": nds_id}, f)
#     return name, nds_id


# def _slugify(text):
#     """
#     Turns free text into something safe to use inside a filename —
#     keeps letters/digits, collapses everything else (spaces,
#     punctuation) into a single underscore, trims leading/trailing
#     underscores. Used for both the specialist's name and NDS ID when
#     building the export filename, since either could contain a space
#     or punctuation that Windows filenames don't like.
#     """
#     text = re.sub(r"[^A-Za-z0-9]+", "_", (text or "").strip())
#     return text.strip("_") or "Unknown"


# def _default_month_label():
#     """e.g. 'September' — today's real calendar month, used only as the
#     suggested default in _prompt_month_label(); the person can type a
#     different month if this export is actually for a different one."""
#     import datetime
#     return datetime.date.today().strftime("%B")


# def _prompt_month_label():
#     """
#     Asks which month this export is for, defaulting to the current
#     calendar month if the person just presses Enter. Asked every run
#     (not cached like the identity) since which month you're working on
#     legitimately changes far more often than your name or NDS ID.
#     """
#     default = _default_month_label()
#     entered = input(f"Which month is this export for? [{default}]: ").strip()
#     return entered or default


# # Which existing file (if any) this run should read carry-forward data
# # from — set by _apply_identity_to_filenames() before anything else
# # runs. This is DIFFERENT from OUTPUT_FILE/OUTPUT_FILE_XLSM below: the
# # version number increases by one on every run, so the file this run
# # WRITES to is never the same file it READS from.
# CARRYFORWARD_SOURCE_FILE = None


# def _apply_identity_to_filenames(name, nds_id, month_label):
#     """
#     Rebuilds the module-level OUTPUT_FILE / OUTPUT_FILE_XLSM /
#     CARRYFORWARD_SOURCE_FILE paths from the specialist's identity and
#     the month they typed in.

#     Folder layout: one shared folder per person, "<NDSID>_<Name>/",
#     holding every version for every month as flat files named
#     "<NDSID>_<Name>_<Month>_v<N>_export.xlsm" — e.g.
#     "NDS35095_Hanis/NDS35095_Hanis_September_v1_export.xlsm". Nothing
#     ever gets deleted; a rerun in the same month bumps N and adds a
#     new file alongside the old ones.

#     RESUMING AN INTERRUPTED RUN: while a run is in progress (or if it
#     got cut short — crash, Ctrl+C, LIST_LIMIT, or the scroll safety
#     cap), its file is named with an extra "_incomplete" tag, e.g.
#     "..._v3_export_incomplete.xlsm". If the highest version found here
#     for this person+month is still tagged "_incomplete", THIS run
#     reuses that EXACT same filename as both its carry-forward source
#     and its write target — no new version number — so a rerun after a
#     crash keeps filling in that same file instead of abandoning it and
#     starting a fresh version. Only once a run reaches the genuine end
#     of the app's list does _finalize_output_file() (called from run())
#     drop the "_incomplete" tag, which is what makes the NEXT deliberate
#     run pick a brand new version number instead of resuming.

#     A month with no matching files yet naturally starts fresh at v1
#     with no carry-forward source — the intended behaviour, since a
#     visit plan (Proposed Dates especially) is specific to its month.
#     """
#     global OUTPUT_FILE, OUTPUT_FILE_XLSM, CARRYFORWARD_SOURCE_FILE

#     person_folder = f"{_slugify(nds_id)}_{_slugify(name)}"
#     os.makedirs(person_folder, exist_ok=True)

#     base_no_version = f"{_slugify(nds_id)}_{_slugify(name)}_{_slugify(month_label)}"
#     complete_pattern = re.compile(
#         r"^" + re.escape(base_no_version) + r"_v(\d+)_export\.(xlsm|xlsx)$")
#     incomplete_pattern = re.compile(
#         r"^" + re.escape(base_no_version) + r"_v(\d+)_export_incomplete\.(xlsm|xlsx)$")

#     best_version = 0
#     best_file = None
#     best_is_xlsm = False
#     best_is_incomplete = False
#     for fname in os.listdir(person_folder):
#         m = incomplete_pattern.match(fname)
#         is_incomplete = m is not None
#         if not m:
#             m = complete_pattern.match(fname)
#         if not m:
#             continue
#         version = int(m.group(1))
#         is_xlsm = m.group(2) == "xlsm"
#         # Highest version wins outright. On a tie: an incomplete file
#         # beats a complete one (it's the true latest in-progress
#         # state), and among those, .xlsm beats .xlsx.
#         better = (
#             version > best_version
#             or (version == best_version and is_incomplete and not best_is_incomplete)
#             or (version == best_version and is_incomplete == best_is_incomplete
#                 and is_xlsm and not best_is_xlsm)
#         )
#         if better:
#             best_version = version
#             best_file = fname
#             best_is_xlsm = is_xlsm
#             best_is_incomplete = is_incomplete

#     if best_file and best_is_incomplete:
#         # Resume this exact file in place — same version, same name.
#         source_path = os.path.join(person_folder, best_file)
#         CARRYFORWARD_SOURCE_FILE = source_path
#         if best_is_xlsm:
#             OUTPUT_FILE_XLSM = source_path
#             OUTPUT_FILE = source_path[:-len(".xlsm")] + ".xlsx"
#         else:
#             OUTPUT_FILE = source_path
#             OUTPUT_FILE_XLSM = source_path[:-len(".xlsx")] + ".xlsm"
#         return

#     CARRYFORWARD_SOURCE_FILE = os.path.join(
#         person_folder, best_file) if best_file else None

#     next_version = best_version + 1
#     versioned_base = f"{base_no_version}_v{next_version}_export_incomplete"
#     OUTPUT_FILE = os.path.join(person_folder, f"{versioned_base}.xlsx")
#     OUTPUT_FILE_XLSM = os.path.join(person_folder, f"{versioned_base}.xlsm")


# def _finalize_output_file(written_to, reached_natural_end):
#     """
#     Called once per run, right after write_output() has saved. If the
#     run genuinely reached the end of the app's list (reached_natural_end
#     — see run()), this drops the "_incomplete" tag from the file that
#     was just written, marking it done: the NEXT run will then pick a
#     fresh version number instead of resuming it (see
#     _apply_identity_to_filenames()).

#     If the run was cut short for any reason — crash, Ctrl+C,
#     LIST_LIMIT, or the scroll safety cap — this leaves the file exactly
#     as it is (still "_incomplete"), so next time you run the script it
#     picks this same file back up and keeps filling it in, rather than
#     treating it as done and starting a new version.

#     Returns the file's final path (renamed or not) so callers can print
#     the name the person should actually go look for.
#     """
#     if not written_to or not os.path.exists(written_to):
#         return written_to

#     if not reached_natural_end:
#         print(f"Didn't reach the end of the app's list this run — "
#               f"{written_to} stays marked in-progress. Rerunning will "
#               f"continue filling in this same file rather than starting "
#               f"a new version.")
#         return written_to

#     if "_export_incomplete." not in written_to:
#         return written_to  # already finalized somehow — nothing to do

#     finalized = written_to.replace("_export_incomplete.", "_export.")
#     try:
#         if os.path.exists(finalized):
#             os.remove(finalized)
#         os.rename(written_to, finalized)
#         print(f"Reached the end of the app's list — finalized as {finalized}.")
#         return finalized
#     except OSError as e:
#         print(f"  !! Could not finalize {written_to} to {finalized} ({e}) — "
#               f"it'll still work fine, just keeps the '_incomplete' name.")
#         return written_to


# def _load_carryforward_data():
#     """
#     Reads whichever export file already exists and returns everything
#     needed to SKIP re-scraping a customer whose sales_no is already in
#     it, while still including them correctly in this run's output:

#       known_sales_nos -> set of sales_no values already captured in a
#                           previous run. run() uses this to skip opening
#                           View Order/CCS Note for these — per Hazim's
#                           explicit call, a customer's details don't
#                           change once captured, so there's nothing new
#                           to read there.
#       carried_records -> {sales_no: {...plain fields...}}, read straight
#                           from the existing file's own columns, so an
#                           already-known customer's row can be reproduced
#                           in the new export without the app ever being
#                           touched for them this run. The phone number
#                           (install_mobile1) is recovered from the hidden
#                           WhatsApp Number column, since the visible
#                           Mobile No. 1 column was removed — that hidden
#                           column already holds the cleaned digits, and
#                           re-cleaning an already-clean number through
#                           clean_phone_for_wa() is a no-op.
#       carried_filters -> {sales_no: filters_text}, read straight from
#                           the Filter(s) column, same reasoning.

#     Proposed Date and Area 1-4 are NOT included here on purpose — those
#     already have their own carry-forward logic inside write_output()
#     (_load_existing_area_values / _load_existing_proposed_dates), which
#     reads the same file independently and applies "existing value always
#     wins." Duplicating that here would just be redundant.

#     Returns (set(), {}, {}) if no export file exists yet (first-ever
#     run — nothing to skip, everyone gets scraped), or if the existing
#     file is from too different a layout to safely reuse (missing a
#     required column) — safer to re-scrape everyone than guess.
#     """
#     target_file = CARRYFORWARD_SOURCE_FILE
#     if not target_file:
#         return set(), {}, {}

#     wb = openpyxl.load_workbook(target_file, data_only=False)
#     try:
#         ws = wb.active
#         header_row = [c.value for c in ws[1]]
#         col = {name: idx + 1 for idx, name in enumerate(header_row) if name}

#         required = ["Sales No.", "Appointment Date", "Installation / Service Contact Person",
#                     "Installation / Service Address", "Product", "WhatsApp Number"]
#         if any(name not in col for name in required):
#             print(f"  !! {target_file}'s header row is missing required columns — "
#                   f"can't safely skip/carry forward from it. Re-scraping everyone "
#                   f"this run instead.")
#             return set(), {}, {}

#         known_sales_nos = set()
#         carried_records = {}
#         carried_filters = {}
#         for row_idx in range(2, ws.max_row + 1):
#             sales_no = _normalize_sales_no(
#                 ws.cell(row=row_idx, column=col["Sales No."]).value)
#             if not sales_no:
#                 continue
#             known_sales_nos.add(sales_no)
#             carried_records[sales_no] = {
#                 "sales_no": sales_no,
#                 "appt_date": ws.cell(row=row_idx, column=col["Appointment Date"]).value or "",
#                 "install_contact_person": ws.cell(row=row_idx, column=col["Installation / Service Contact Person"]).value or "",
#                 "install_mobile1": _normalize_phone_digits(
#                     ws.cell(row=row_idx, column=col["WhatsApp Number"]).value),
#                 "install_address": ws.cell(row=row_idx, column=col["Installation / Service Address"]).value or "",
#                 "sales_info_product": ws.cell(row=row_idx, column=col["Product"]).value or "",
#             }
#             if "Filter(s)" in col:
#                 filters_text = ws.cell(
#                     row=row_idx, column=col["Filter(s)"]).value
#                 if filters_text:
#                     carried_filters[sales_no] = filters_text
#         return known_sales_nos, carried_records, carried_filters
#     finally:
#         wb.close()


# # ============================================================
# # Live pause/resume — type 'p' + Enter to pause, 'r' + Enter to resume
# # ============================================================
# #
# # Different from the existing Ctrl+C handling: Ctrl+C stops the whole
# # script, and resuming means re-launching Python, re-entering the
# # identity/month prompts, and reconnecting Appium from scratch (though
# # it does correctly pick up the same in-progress file — see
# # _apply_identity_to_filenames()'s "RESUMING AN INTERRUPTED RUN" note).
# # This is for a shorter break — pause the SAME running process (Appium
# # session stays connected, nothing in memory is lost) so the phone is
# # free to use, then continue right where it left off with no restart.
# #
# # Deliberately NOT a background thread. An earlier version of this used
# # one (reading stdin in a loop via input()), and it actually worked
# # correctly for pausing/resuming — but caused a real crash: a thread
# # blocked waiting for keyboard input doesn't get cleaned up properly
# # when the main script finishes, and printed a "Fatal Python error"
# # after every run, successful or not (confirmed by testing). Polling
# # for input from the MAIN thread instead — at the same safe checkpoints
# # the loop already visits — avoids that entirely, since nothing is ever
# # left waiting on stdin when the script exits.
# _pause_requested = False
# _pause_input_buffer = ""

# try:
#     import msvcrt  # Windows only — this project runs on Windows
#     _HAS_MSVCRT = True
# except ImportError:
#     _HAS_MSVCRT = False
#     import select  # non-Windows fallback, so this doesn't break on Mac/Linux


# def _read_pending_stdin_chars():
#     """Returns whatever characters are ALREADY waiting on stdin right
#     now, without blocking — or "" if nothing's been typed yet. Windows
#     uses msvcrt.kbhit()/getwch() (the standard non-blocking keyboard
#     check on that platform); anywhere else, select() checks whether
#     stdin has data ready, then os.read() pulls the raw bytes directly
#     from the file descriptor — NOT sys.stdin.read(), which goes through
#     Python's own internal text-buffering layer. That buffering layer
#     can silently consume more bytes from the OS in one read() than it
#     hands back, so a later select() check sees the file descriptor as
#     empty even though a full line (like a typed 'r') is still sitting
#     unread inside Python's buffer — confirmed by testing: this exact
#     mismatch caused resume to never fire after a real pause. Reading
#     the raw fd directly keeps what select() sees and what actually gets
#     read in sync."""
#     chars = ""
#     if _HAS_MSVCRT:
#         while msvcrt.kbhit():
#             chars += msvcrt.getwch()
#     else:
#         while select.select([sys.stdin], [], [], 0)[0]:
#             data = os.read(sys.stdin.fileno(), 4096)
#             if not data:
#                 break
#             chars += data.decode(errors="ignore")
#     return chars


# def _poll_pause_commands():
#     """Builds up typed characters into a line buffer and checks it
#     against 'p'/'r' once Enter is pressed — same idea as input(), but
#     non-blocking. Called from _wait_if_paused() at safe points in the
#     main loop, so a 'p' typed mid-customer doesn't take effect until
#     that customer is actually finished."""
#     global _pause_requested, _pause_input_buffer
#     for ch in _read_pending_stdin_chars():
#         if ch in ("\r", "\n"):
#             cmd = _pause_input_buffer.strip().lower()
#             _pause_input_buffer = ""
#             if cmd == "p" and not _pause_requested:
#                 _pause_requested = True
#                 print("Pause requested — will pause after the customer "
#                       "currently in progress finishes (not mid-action). "
#                       "Type 'r' + Enter when you're ready to continue.")
#             elif cmd == "r" and _pause_requested:
#                 _pause_requested = False
#                 print("Resuming...")
#         else:
#             _pause_input_buffer += ch


# def _wait_if_paused():
#     """Called at safe boundaries in the main loop — right after a
#     customer is fully finished (back on the plain list screen, no
#     popup open, nothing mid-read) or after a scroll with nothing new
#     to act on. NEVER called mid-tap or mid-read, so pausing here can't
#     leave the phone in a half-navigated state."""
#     _poll_pause_commands()
#     if not _pause_requested:
#         return
#     print("\nPaused. The phone won't be touched again until you resume.\n"
#           "If you switch to a DIFFERENT app on the phone, make sure "
#           "Cuckoo+ is back in the foreground before typing 'r' — this "
#           "script taps raw screen positions, not the app specifically, "
#           "so the next action needs Cuckoo+ to actually be on screen.")
#     while _pause_requested:
#         time.sleep(0.5)
#         _poll_pause_commands()


# def run():
#     import traceback

#     name, nds_id = _load_or_prompt_identity()
#     month_label = _prompt_month_label()
#     _apply_identity_to_filenames(name, nds_id, month_label)
#     print(f"Specialist: {name} ({nds_id}) | Month: {month_label}")
#     write_target = OUTPUT_FILE_XLSM if os.path.exists(
#         TEMPLATE_FILE) else OUTPUT_FILE
#     if CARRYFORWARD_SOURCE_FILE:
#         print(f"Continuing from {CARRYFORWARD_SOURCE_FILE} — this run will "
#               f"write {write_target}.")
#     else:
#         print(f"No existing file found for {name} ({nds_id}) in {month_label} — "
#               f"starting fresh. This run will write {write_target}.")

#     known_sales_nos, carried_records, carried_filters = _load_carryforward_data()
#     if known_sales_nos:
#         print(f"Found {len(known_sales_nos)} customer(s) already captured in a "
#               f"previous export — these will be SKIPPED (their details aren't "
#               f"re-read, since they don't change), reused as-is if still found "
#               f"in the app, and REMOVED from the output if no longer found. "
#               f"Only genuinely new sales_no values get fully scraped.")

#     driver = build_driver()
#     time.sleep(3)

#     print("Tip: type 'p' + Enter anytime during this run to pause after "
#           "the current customer finishes, and 'r' + Enter to resume — "
#           "the connection to the phone stays open, nothing gets re-scraped.")

#     all_records = []
#     all_ccs_rows = []
#     seen_keys = set()
#     stagnant_rounds = 0
#     scroll_count = 0
#     run_start_time = time.time()
#     customer_durations = []
#     # Every sales_no actually confirmed present in the app this run —
#     # whether skipped (already known) or freshly scraped (new). Anyone
#     # from a previous export who's NOT in this set by the end is treated
#     # as removed from the app, PROVIDED reached_natural_end is True (see
#     # below) — otherwise removal isn't safe to trust.
#     confirmed_present_sales_nos = set()
#     # Same information as confirmed_present_sales_nos, but as a LIST in
#     # the order each sales_no was first encountered this run — a set has
#     # no order at all. This is what lets the final export follow the
#     # app's CURRENT ordering (which can genuinely change month to month)
#     # instead of always resorting alphabetically by Sales No.
#     sales_no_order = []
#     # Only True if scrolling reached the genuine end of the list (no new
#     # customers found after scrolling further) — NOT true if LIST_LIMIT
#     # cut the run short (e.g. while testing) or the safety scroll cap
#     # was hit. Removal logic below only runs when this is True, since a
#     # partial scan can't distinguish "genuinely removed from the app"
#     # from "just hasn't been scrolled to yet."
#     reached_natural_end = False

#     try:
#         while True:
#             _wait_if_paused()

#             if LIST_LIMIT and len(all_records) >= LIST_LIMIT:
#                 print(f"Reached LIST_LIMIT of {LIST_LIMIT} — stopping.")
#                 break

#             customers = get_visible_customers_quick_or_stable(
#                 driver, known_sales_nos, seen_keys)

#             next_customer = None
#             partial_customer = None
#             skipped_this_round = False
#             for c in customers:
#                 key = c["row"].get(KEY_FIELD)
#                 if not key or key in seen_keys:
#                     continue
#                 if c["button"] is None or not c.get("complete", True):
#                     # Card is likely only partially scrolled into view — its
#                     # NS No rendered (enough to be found and keyed), but its
#                     # button and/or its later fields (Cust Name is the last
#                     # field on the card, so it's usually the first casualty)
#                     # haven't fully rendered yet. Don't mark it seen; remember
#                     # it so we can scroll IT specifically into full view,
#                     # rather than treating this like "nothing new at all."
#                     if partial_customer is None:
#                         partial_customer = c
#                     continue
#                 sales_no_value = _normalize_sales_no(c["row"].get("sales_no"))
#                 if sales_no_value:
#                     # Confirmed present in the app THIS run, regardless of
#                     # whether it gets skipped (already known) or scraped
#                     # fresh (new) — this is what the end-of-run removal
#                     # check is based on. Only append to the ORDER list the
#                     # first time — the same card can be scanned again
#                     # across consecutive scroll reads before it's marked
#                     # "seen", and it must only claim one position.
#                     if sales_no_value not in confirmed_present_sales_nos:
#                         sales_no_order.append(sales_no_value)
#                     confirmed_present_sales_nos.add(sales_no_value)
#                 if sales_no_value and sales_no_value in known_sales_nos:
#                     # Already fully captured in a previous run, and a
#                     # customer's details don't change once captured — so
#                     # there's nothing new to read by opening this one. Mark
#                     # it handled and keep scanning; a genuinely new customer
#                     # could appear anywhere in the list, not necessarily
#                     # after this point, so scrolling continues normally.
#                     print(
#                         f"  [{len(confirmed_present_sales_nos)}] skipping {sales_no_value} — already captured previously")
#                     seen_keys.add(key)
#                     skipped_this_round = True
#                     continue
#                 next_customer = c
#                 break

#             if next_customer is None:
#                 if partial_customer is not None:
#                     # There IS a new customer here — it's just not fully
#                     # rendered yet. Scroll targeted at THIS card's own
#                     # position (not the generic "last customer" case below)
#                     # so it comes fully into view instead of being skipped.
#                     scroll_count += 1
#                     if scroll_count >= MAX_SCROLLS:
#                         print(
#                             "Hit the safety scroll limit — stopping to avoid an infinite loop.")
#                         break
#                     target_percent = compute_scroll_percent(
#                         driver, partial_customer["card_top_y"])
#                     if DEBUG:
#                         print(f"  [debug] {partial_customer['row'].get(KEY_FIELD)} not fully rendered yet — "
#                               f"scrolling it into view (percent={target_percent:.3f})")
#                     scroll_down(driver, percent=target_percent)
#                     continue

#                 if skipped_this_round:
#                     # We successfully handled (skipped) known customers this
#                     # round — real progress, just nothing left to actively
#                     # open in the CURRENT view. This must NOT count as
#                     # stagnant: a long run of consecutive known customers
#                     # (e.g. 50 in a row) would otherwise trip the "reached
#                     # the end" check after just a couple of rounds, long
#                     # before actually reaching the true end of the list —
#                     # which would leave later customers unscanned and, worse,
#                     # make the removal logic below wrongly think they'd
#                     # disappeared from the app.
#                     stagnant_rounds = 0
#                     scroll_count += 1
#                     if scroll_count >= MAX_SCROLLS:
#                         print(
#                             "Hit the safety scroll limit — stopping to avoid an infinite loop.")
#                         break
#                     target_percent = compute_scroll_percent(
#                         driver, customers[-1]["card_top_y"]) if customers else SCROLL_STEP_PERCENT
#                     scroll_down(driver, percent=target_percent)
#                     continue

#                 stagnant_rounds += 1
#                 if stagnant_rounds >= MAX_STAGNANT_ROUNDS:
#                     print(
#                         "No new customers found after scrolling — reached the end of the list.")
#                     # This is the ONLY point where we can trust the whole
#                     # list was actually scanned — see confirmed_present_sales_nos
#                     # / reached_natural_end usage after the loop.
#                     reached_natural_end = True
#                     break
#                 scroll_count += 1
#                 if scroll_count >= MAX_SCROLLS:
#                     print(
#                         "Hit the safety scroll limit — stopping to avoid an infinite loop.")
#                     break
#                 target_percent = compute_scroll_percent(
#                     driver, customers[-1]["card_top_y"]) if customers else SCROLL_STEP_PERCENT
#                 if DEBUG:
#                     print(
#                         f"  [debug] scrolling by measured percent={target_percent:.3f}")
#                 scroll_down(driver, percent=target_percent)
#                 continue

#             stagnant_rounds = 0
#             next_row = next_customer["row"]
#             key = next_row[KEY_FIELD]
#             seen_keys.add(key)
#             customer_start_time = time.time()
#             # Sales No. (not NS No.) is what actually identifies a
#             # customer to the specialist, and the position number here
#             # is based on confirmed_present_sales_nos — every customer
#             # confirmed present so far, skipped or not — so it tracks
#             # this customer's real position in the app's list, rather
#             # than undercounting because earlier ones were skipped.
#             display_sales_no = _normalize_sales_no(
#                 next_row.get("sales_no")) or "(no sales no)"
#             print(
#                 f"[{len(confirmed_present_sales_nos)}] Opening: {display_sales_no} ({next_row.get('cust_name', '')})")

#             try:
#                 button = next_customer["button"]
#                 if button is None:
#                     raise RuntimeError(
#                         "No 'View Order' button found near this customer's row")

#                 # --- Visit 1: View Order (sales/install details) ---
#                 try:
#                     try:
#                         tap_element(driver, button)
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: tapping row button] {type(e).__name__}: {e}")

#                     try:
#                         # select_popup_option() already waits/polls for the popup
#                         # to appear — no fixed sleep needed before it.
#                         select_popup_option(driver, "View Order")
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: selecting 'View Order' from popup] {type(e).__name__}: {e}")

#                     try:
#                         wait_for(
#                             driver, (AppiumBy.XPATH, '//android.widget.Button[@text="Customer Information"]'))
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: waiting for Customer Information screen] {type(e).__name__}: {e}")

#                     try:
#                         detail = read_full_detail(driver)
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: reading detail screen fields] {type(e).__name__}: {e}")

#                     all_records.append({**next_row, **detail})
#                 except Exception as e:
#                     print(f"  !! Skipped View Order for {key}: {e}")
#                 finally:
#                     driver.back()
#                     # get_visible_customers() below already waits/polls for "NS
#                     # No" to reappear, so no sleep needed here.

#                 # --- Visit 2: CCS Note (filter/consumable change data) ---
#                 # A fresh element lookup is required here — the `button`
#                 # WebElement from before driver.back() is stale now (the
#                 # underlying UI tree changed), so it can't just be reused for
#                 # a second tap the way it could within a single visit. This
#                 # re-find always uses the full stable read, not the quick
#                 # path — we're about to act on this exact customer, so it's
#                 # exactly the case the quick path defers to it for anyway.
#                 try:
#                     try:
#                         customers_again = get_visible_customers_stable(driver)
#                         this_customer_again = next(
#                             (c for c in customers_again if c["row"].get(KEY_FIELD) == key), None)
#                         if this_customer_again is None or this_customer_again["button"] is None:
#                             raise RuntimeError(
#                                 "Could not re-find this customer's row for CCS Note")
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: re-finding row after View Order] {type(e).__name__}: {e}")

#                     try:
#                         tap_element(driver, this_customer_again["button"])
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: tapping row button] {type(e).__name__}: {e}")

#                     try:
#                         select_popup_option(driver, "CCS Note")
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: selecting 'CCS Note' from popup] {type(e).__name__}: {e}")

#                     try:
#                         cards = get_all_ccs_note_cards(driver)
#                     except Exception as e:
#                         raise RuntimeError(
#                             f"[stage: reading CCS Note cards] {type(e).__name__}: {e}")

#                     for card in cards:
#                         all_ccs_rows.append({
#                             "sales_no": next_row.get("sales_no", ""),
#                             "cust_name": next_row.get("cust_name", ""),
#                             "product": card["product"],
#                             "last_change": card["last_change"],
#                         })
#                 except Exception as e:
#                     print(f"  !! Skipped CCS Note for {key}: {e}")
#                 finally:
#                     driver.back()
#             except Exception as e:
#                 print(f"  !! Skipped {key} entirely due to error: {e}")

#             customer_elapsed = time.time() - customer_start_time
#             customer_durations.append(customer_elapsed)
#             print(f"    ({_format_duration(customer_elapsed)})")

#     except KeyboardInterrupt:
#         # Ctrl+C mid-run. Whatever's in all_records/all_ccs_rows so far
#         # still gets saved below — reached_natural_end is correctly
#         # still False here, so the "remove missing customers" logic
#         # stays safely off, same as any other partial run.
#         print("\nInterrupted (Ctrl+C) — saving whatever was captured so "
#               "far before exiting.")
#     except Exception as e:
#         # Anything unexpected (device disconnected, a selector that no
#         # longer matches, etc.) — rather than losing every customer
#         # scraped so far, save them now and surface the full traceback
#         # so the actual problem is visible, instead of just crashing
#         # silently with nothing written.
#         print(
#             f"\n!! Unexpected error, stopping early: {type(e).__name__}: {e}")
#         print("Saving whatever was captured so far before exiting.")
#         traceback.print_exc()
#     finally:
#         # driver.quit() can itself throw — most notably if the phone
#         # was physically disconnected (USB yanked, cable fault, phone
#         # rebooted) rather than the app/script hitting a normal error.
#         # In that case there's no live session left to cleanly close,
#         # and an exception raised HERE, inside finally, would propagate
#         # straight out of run() and skip everything below — including
#         # the save — which defeats the entire point of the except
#         # clauses above. Swallowing it is safe: by this point either
#         # the loop finished normally or one of the excepts above has
#         # already handled and logged the real problem.
#         try:
#             driver.quit()
#         except Exception:
#             pass

#     total_elapsed = time.time() - run_start_time

#     if reached_natural_end:
#         # The full list was genuinely scanned this run, so anyone
#         # previously captured but NOT confirmed present now has actually
#         # been removed from the app — drop them from the export too,
#         # rather than leaving stale rows behind forever.
#         carried_to_keep = {sn: rec for sn, rec in carried_records.items()
#                            if sn in confirmed_present_sales_nos}
#         carried_filters_to_keep = {sn: txt for sn, txt in carried_filters.items()
#                                    if sn in confirmed_present_sales_nos}
#         removed = sorted(set(carried_records) - confirmed_present_sales_nos)
#         if removed:
#             print(f"Removed {len(removed)} customer(s) no longer found in "
#                   f"the app: {', '.join(removed)}")
#     else:
#         # Didn't confirm scanning the WHOLE list this run (LIST_LIMIT cut
#         # it short, or the safety scroll cap was hit) — "not seen yet" and
#         # "actually gone" can't be told apart from a partial scan, so
#         # nothing gets removed this time, to be safe.
#         carried_to_keep = carried_records
#         carried_filters_to_keep = carried_filters
#         if carried_records:
#             print("Note: this run didn't confirm scanning the full list "
#                   "(LIST_LIMIT or the scroll safety cap stopped it early), "
#                   "so no previously-captured customers were removed even if "
#                   "not seen this run.")

#     # Merge: newly-scraped customers (all_records/all_ccs_rows) plus
#     # whichever previously-captured customers are being kept (see above).
#     # These two groups never overlap by construction — a sales_no is
#     # either in known_sales_nos (skipped/carried) or it isn't (scraped
#     # fresh this run) — so a plain combine is safe.
#     merged_records = all_records + list(carried_to_keep.values())
#     merged_filters = {**carried_filters_to_keep,
#                       **_build_filters_lookup(all_ccs_rows)}

#     written_to = write_output(
#         merged_records, filters_override=merged_filters, app_order=sales_no_order,
#         specialist_name=name, nds_id=nds_id)
#     written_to = _finalize_output_file(written_to, reached_natural_end)
#     print(f"Done. Wrote {written_to} — {len(all_records)} newly-scraped "
#           f"customer(s), {len(carried_to_keep)} carried forward unchanged "
#           f"({len(merged_records)} total), {len(all_ccs_rows)} new CCS Note "
#           f"row(s) read.")
#     skip_count = len(carried_to_keep)
#     detail_time = sum(customer_durations)
#     # Exact, not an estimate — the ONLY thing timed separately per-
#     # customer is a newly-scraped customer's View Order + CCS Note visit
#     # (customer_start_time/customer_elapsed). Everything else the loop
#     # does — scrolling, the quick/stable screen reads, skip decisions —
#     # is NOT individually timed, so "everything else" is whatever's left
#     # after subtracting the one thing that IS precisely measured. This
#     # bucket isn't PURELY "time skipping known customers" — it also
#     # includes the screen-reading that happens right before a NEW
#     # customer is found too, not just before a skip — but when most of
#     # a run's customers are already-known (the normal case after month
#     # one), this bucket is realistically dominated by skip-related
#     # scrolling.
#     other_time = max(0, total_elapsed - detail_time)
#     print(f"Total time: {_format_duration(total_elapsed)}")
#     if customer_durations:
#         avg_seconds = detail_time / len(customer_durations)
#         print(f"  New customers: {len(customer_durations)} scraped, "
#               f"{_format_duration(detail_time)} total (avg {avg_seconds:.1f}s each)")
#     else:
#         print("  New customers: none scraped this run")
#     if skip_count:
#         avg_skip = other_time / skip_count
#         print(f"  Everything else — scrolling, screen reads, and skipping "
#               f"{skip_count} already-known customer(s): "
#               f"{_format_duration(other_time)} total (~{avg_skip:.1f}s per "
#               f"skipped customer, though this bucket isn't purely skip time)")
#     else:
#         print(f"  Everything else (scrolling/screen reads, no known "
#               f"customers to skip this run): {_format_duration(other_time)}")


# # ============================================================
# # Export column order:
# #   A=sales_no, B=install_contact_person, C=install_address,
# #   D-G=Area 1-4 (State/District-City/Locality/Street — auto-filled
# #   when confident, see the "Area auto-fill" block above, you fill in
# #   the rest by hand, see _load_existing_area_values), H=appt_date,
# #   I=proposed_date (typed by hand), J=WhatsApp Link, K=WhatsApp
# #   Message, L=sales_info_product, M=Filters (from CCS Note data),
# #   N=wa_number (hidden helper). Mobile No. 1 is no longer its own
# #   visible column — the phone number still flows through internally
# #   (install_mobile1 on the record) purely to build the WhatsApp
# #   Link/Message/Number columns.
# #
# # Column letters are looked up BY NAME (see _col_letter below) rather
# # than hardcoded, precisely because this project already got bitten
# # once by a hardcoded-column-position bug after a reorder — this way
# # adding/moving/removing a column later can't silently break a formula
# # again.
# # ============================================================

# ALL_COLUMNS = [
#     "Sales No.", "Installation / Service Contact Person",
#     "Installation / Service Address",
#     "Area 1 (State)", "Area 2 (District / City)", "Area 3 (Locality)", "Area 4 (Street)",
#     "Appointment Date", "Proposed Date", "WhatsApp Chat Link",
#     "WhatsApp Chat Message", "Product", "Filter(s)", "WhatsApp Number",
# ]

# _COLUMN_INDEX = {name: idx for idx, name in enumerate(ALL_COLUMNS, start=1)}


# def _col_letter(column_name):
#     """Excel column letter for a column, looked up by its header name —
#     e.g. _col_letter("Proposed Date") -> "I". Keeps every formula below
#     immune to column reordering."""
#     from openpyxl.utils import get_column_letter
#     return get_column_letter(_COLUMN_INDEX[column_name])


# def _format_duration(seconds):
#     """'93.4' -> '1m 33s'; '8.2' -> '8.2s'. Used for the per-customer and
#     total run-time output in run()."""
#     if seconds >= 60:
#         minutes, secs = divmod(int(seconds), 60)
#         return f"{minutes}m {secs}s"
#     return f"{seconds:.1f}s"


# def clean_phone_for_wa(raw):
#     """
#     Strips everything except digits, then converts a local Malaysian
#     mobile number (leading 0, e.g. "012-345 6789") into the international
#     format wa.me links need (leading 60, no separators, e.g.
#     "60123456789"). A number that's already in some other international
#     format (doesn't start with 0 after stripping) is left as digits-only
#     — we can't safely guess a country code that isn't already there.
#     Returns "" if there's nothing usable. Idempotent — running it again
#     on an already-converted number ("60123456789") leaves it unchanged,
#     which is what lets a carried-forward WhatsApp Number value safely
#     pass back through this on the next run.
#     """
#     digits = re.sub(r"\D", "", raw or "")
#     if not digits:
#         return ""
#     if digits.startswith("0"):
#         digits = "60" + digits[1:]
#     return digits


# def message_formula(row, specialist_name, nds_id):
#     """
#     Excel formula for one row's WhatsApp Message cell, matching the
#     copywriting/formatting shown in the reference screenshot. Single
#     asterisks are WhatsApp's own bold syntax, not markdown. Column
#     letters are looked up by name (see _col_letter), not hardcoded.

#     specialist_name/nds_id come from the specialist's local identity
#     config (see _load_or_prompt_identity) and are baked into the
#     formula as literal text, not a cell reference — they're constant
#     for the whole export, not something that varies per row. Any
#     double-quote either one might contain is escaped to Excel's own
#     "" convention so it can't break the formula string.

#     Tarikh uses a fully Malay, all-caps date format with the weekday
#     name — e.g. "07 OKTOBER 2026 (ISNIN)" — built with
#     CHOOSE(MONTH(...)) / CHOOSE(WEEKDAY(...)) since Excel's TEXT() has
#     no built-in Malay locale to draw month/weekday names from.

#     Note: if you copy this CELL (Ctrl+C, not double-click) and paste
#     into something like Notepad, you'll see the whole value wrapped in
#     quotes. That's Excel's own clipboard behavior (CSV-style quoting)
#     kicking in because the text contains commas and line breaks — it's
#     not part of the cell's actual value or this formula, and it doesn't
#     matter once you're using the WhatsApp Link column instead of
#     copying this text by hand.

#     The *Alamat:* line uses SUBSTITUTE to turn the address cell's line
#     breaks back into spaces before it goes into the message — the
#     address cell itself is formatted with Alt+Enter-style line breaks
#     for readability in the sheet (see _format_address_for_cell), but
#     since those breaks were inserted right after commas (or, for the
#     postcode, right where a plain space already was), swapping each
#     line break back to a single space reconstructs the exact original
#     single-line address text for the customer-facing message.
#     """
#     sales_no_col = _col_letter("Sales No.")
#     address_col = _col_letter("Installation / Service Address")
#     contact_col = _col_letter("Installation / Service Contact Person")
#     date_col = _col_letter("Proposed Date")
#     product_col = _col_letter("Product")

#     safe_name = (specialist_name or "").replace('"', '""')
#     safe_nds_id = (nds_id or "").replace('"', '""')

#     tarikh_part = (
#         f'IFERROR(TEXT({date_col}{row},"DD") & " " & '
#         f'CHOOSE(MONTH({date_col}{row}),"JANUARI","FEBRUARI","MAC","APRIL","MEI","JUN",'
#         f'"JULAI","OGOS","SEPTEMBER","OKTOBER","NOVEMBER","DISEMBER") & " " & '
#         f'TEXT({date_col}{row},"YYYY") & " (" & '
#         f'CHOOSE(WEEKDAY({date_col}{row},2),"ISNIN","SELASA","RABU","KHAMIS","JUMAAT","SABTU","AHAD") & ")", '
#         f'{date_col}{row})'
#     )
#     return (
#         f'=IF({date_col}{row}="","","Selamat sejahtera Tuan/Puan " & UPPER({contact_col}{row}) & "," & CHAR(10) & CHAR(10) & '
#         f'"Saya *{safe_name}*, CUCKOO+ Service Specialist ({safe_nds_id}). Saya memohon maaf jika saya menghubungi anda pada waktu yang tidak sesuai." & CHAR(10) & CHAR(10) & '
#         f'"Saya ingin mengesahkan jika saya boleh membuat lawatan servis seperti di bawah." & CHAR(10) & CHAR(10) & '
#         f'"*Tarikh:* " & {tarikh_part} & CHAR(10) & '
#         f'"*Alamat:* " & SUBSTITUTE({address_col}{row},CHAR(10)," ") & CHAR(10) & '
#         f'"*Produk:* " & {product_col}{row} & CHAR(10) & '
#         f'"*Nombor Pesanan:* " & {sales_no_col}{row} & CHAR(10) & CHAR(10) & '
#         f'"Terima kasih, sokongan dan kerjasama Tuan/Puan amat saya hargai.")'
#     )


# def wa_link_formula(row, phone_digits):
#     """
#     Excel formula for one row's WhatsApp Link cell.

#     IMPORTANT LIMITATION: this can't pre-fill the message the way a
#     wa.me "?text=" link normally would. Excel's own HYPERLINK() worksheet
#     function hard-caps its link_location argument at 255 characters —
#     if it's longer, Excel returns #VALUE! instead of opening anything.
#     The encoded message (greeting + explanation + address etc.) is
#     always well past that, no matter how it's trimmed, so a formula-
#     based link genuinely cannot carry the pre-filled text.

#     What this DOES do: opens the right WhatsApp chat directly (a plain
#     "https://wa.me/<number>" link, comfortably under 255 chars), so you
#     only need to copy the WhatsApp Message cell and paste it in — no
#     manual number lookup or searching for the contact. Column letter is
#     looked up by name (see _col_letter), not hardcoded.
#     """
#     date_col = _col_letter("Proposed Date")
#     if not phone_digits:
#         return '="No phone number found"'
#     return (
#         f'=IF({date_col}{row}="","",HYPERLINK('
#         f'"https://wa.me/{phone_digits}",'
#         f'"Open WhatsApp Chat"))'
#     )


# _FILTER_LINE_NUMBER_PATTERN = re.compile(r"^\d+\.\s*")


# def _number_filters_text(text):
#     """
#     Adds "1. ", "2. ", etc. in front of each line of Filters text.
#     Strips any numbering already present first (matching a leading
#     "N. ") before renumbering — this is what stops repeated --resort
#     runs from stacking up "1. 1. Product..." — so it's safe to call on
#     text that's already numbered, freshly built, or anything in between.
#     Blank/empty input passes through unchanged.
#     """
#     if not text:
#         return text
#     lines = text.split("\n")
#     cleaned = [_FILTER_LINE_NUMBER_PATTERN.sub("", line) for line in lines]
#     return "\n".join(f"{i}. {line}" for i, line in enumerate(cleaned, start=1))


# def _format_address_for_cell(address):
#     """
#     Breaks a long address across multiple lines for easier reading in
#     the sheet — the same effect as pressing Alt+Enter at each comma —
#     WITHOUT removing anything: every comma stays exactly where it was,
#     a line break is just added right after it.

#     Also breaks right before the postcode, if one's found — using the
#     exact same postcode detection the Area auto-fill above already
#     relies on — since that's usually the one place a long address has
#     no comma at all to break on naturally (e.g. "...PRESINT 11
#     PUTRAJAYA 62300 WP PUTRAJAYA" runs on with no punctuation).

#     Needs wrap_text on for the cell to actually show as multiple lines
#     (already true for the address column via WRAPPED_COLUMNS below).

#     This is applied ONLY here, at display time — detect_areas() (in the
#     "Area auto-fill" block near the top of this file) still runs on the
#     original, unbroken address text before this ever happens, so none
#     of this affects Area detection.

#     The WhatsApp Chat Message column reconstructs the original,
#     single-line address from this (see message_formula's use of
#     SUBSTITUTE) rather than showing these line breaks in the actual
#     message text sent to the customer.
#     """
#     if not address:
#         return address
#     text = address.strip()
#     # Break right after every comma — comma itself is kept, untouched.
#     text = re.sub(r",\s*", ",\n", text)
#     # Also break right before the postcode, if it isn't already at the
#     # start of a line (e.g. wasn't already right after a comma).
#     postcode_match = re.search(r"\b\d{5}\b", text)
#     if postcode_match:
#         start = postcode_match.start()
#         if start > 0 and text[start - 1] != "\n":
#             text = text[:start].rstrip() + "\n" + text[start:]
#     return text


# def _populate_sheet(ws, records, filters_by_sales_no, specialist_name, nds_id):
#     """
#     Fills in headers + rows on an already-created worksheet (either a
#     fresh one, or the one already inside the macro-enabled template).
#     Shared by both output paths in write_output() below so the two
#     stay in sync automatically. filters_by_sales_no maps sales_no ->
#     the combined multi-line Filters text (see _build_filters_lookup).
#     """
#     FONT = "Arial"
#     # Cuckoo's own logo red is a vivid red-orange — I couldn't pull an
#     # exact official hex from their site, so this is a close visual
#     # match (a common vivid Korean-appliance-brand red). If you have
#     # the real logo file and want an exact match, this is the one
#     # value to swap.
#     header_fill = PatternFill("solid", fgColor="ED1C24")

#     # Short, simple values read better centered; long free text (name,
#     # address, filters, message) reads better left-aligned but still
#     # vertically centered — every cell in the sheet gets vertical
#     # centering regardless, only horizontal centering is selective.
#     CENTERED_COLUMNS = {
#         _COLUMN_INDEX["Sales No."], _COLUMN_INDEX["Appointment Date"],
#         _COLUMN_INDEX["Proposed Date"], _COLUMN_INDEX["WhatsApp Chat Link"],
#         _COLUMN_INDEX["Product"], _COLUMN_INDEX["WhatsApp Number"],
#     }
#     WRAPPED_COLUMNS = {
#         _COLUMN_INDEX["Installation / Service Address"],
#         _COLUMN_INDEX["Filter(s)"], _COLUMN_INDEX["WhatsApp Chat Message"],
#     }

#     for col_idx, header in enumerate(ALL_COLUMNS, start=1):
#         cell = ws.cell(row=1, column=col_idx, value=header)
#         cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
#         cell.fill = header_fill
#         cell.alignment = Alignment(
#             wrap_text=True, vertical="center", horizontal="center")

#     area_comment_text = (
#         "Auto-filled ONLY when the script is confident (free, offline — "
#         "postcode lookup + keyword matching, no AI, no paid API). Left "
#         "blank whenever it isn't sure, rather than risk a wrong guess. "
#         "Fill in the rest yourself, or add more detail to what's already "
#         "there. Once a cell has something in it — auto-filled OR typed "
#         "by hand — every future run leaves it exactly as-is, never "
#         "overwrites it. Use Excel's filter on these columns to drill "
#         "down: Area 1 first, then Area 2, Area 3, Area 4 as needed."
#     )
#     ws[f"{_col_letter('Area 1 (State)')}1"].comment = Comment(
#         area_comment_text, "main.py")
#     ws[f"{_col_letter('Area 2 (District / City)')}1"].comment = Comment(
#         area_comment_text, "main.py")
#     ws[f"{_col_letter('Area 3 (Locality)')}1"].comment = Comment(
#         area_comment_text, "main.py")
#     ws[f"{_col_letter('Area 4 (Street)')}1"].comment = Comment(
#         area_comment_text, "main.py")

#     ws[f"{_col_letter('Proposed Date')}1"].comment = Comment(
#         "Type a date here (e.g. 08/08/2026 or 14 August 2026 both work) —\n"
#         "it displays as \"14 August 2026\" once entered.\n"
#         "The WhatsApp Message and WhatsApp Link columns fill themselves in automatically.",
#         "main.py"
#     )
#     ws[f"{_col_letter('WhatsApp Chat Link')}1"].comment = Comment(
#         "Click to open the right WhatsApp chat directly, then copy the\n"
#         "WhatsApp Chat Message column and paste it in. (Excel's HYPERLINK\n"
#         "function can't carry a pre-filled message this long — it caps\n"
#         "links at 255 characters — so this gets you to the chat, but the\n"
#         "message still needs one paste.)\n\n"
#         "If this file has the WhatsApp macro installed (see write_output()'s\n"
#         "docstring in main.py), double-click instead of single-\n"
#         "clicking — that opens the chat with the message already filled in,\n"
#         "no paste needed.",
#         "main.py"
#     )
#     ws[f"{_col_letter('WhatsApp Number')}1"].comment = Comment(
#         "Internal use only — the macro reads this to build the full "
#         "pre-filled WhatsApp link, and it's also how the phone number "
#         "survives into next month's export now that Mobile No. 1 isn't "
#         "its own column. Don't edit or delete this column.",
#         "main.py"
#     )

#     def set_cell(row_idx, col_idx, value):
#         # clean_text is skipped for formula strings (start with "=") —
#         # they're Excel syntax, not scraped text, and don't need it.
#         if isinstance(value, str) and not value.startswith("="):
#             value = clean_text(value)
#         cell = ws.cell(row=row_idx, column=col_idx, value=value)
#         cell.font = Font(name=FONT, size=10)
#         cell.alignment = Alignment(
#             vertical="center",
#             horizontal="center" if col_idx in CENTERED_COLUMNS else "general",
#             wrap_text=col_idx in WRAPPED_COLUMNS,
#         )
#         return cell

#     for row_idx, record in enumerate(records, start=2):
#         set_cell(row_idx, _COLUMN_INDEX["Sales No."],
#                  record.get("sales_no", ""))
#         set_cell(row_idx, _COLUMN_INDEX["Installation / Service Contact Person"],
#                  record.get("install_contact_person", ""))
#         set_cell(row_idx, _COLUMN_INDEX["Installation / Service Address"],
#                  _format_address_for_cell(record.get("install_address", "")))

#         # Area 1 -> Area 4 (State/District-City/Locality/Street) —
#         # auto-filled when confident, otherwise manual (see
#         # _load_existing_area_values / detect_areas). Whatever's already
#         # in `record` here (existing value OR fresh auto-fill) was
#         # already decided before write_output() got this far.
#         set_cell(
#             row_idx, _COLUMN_INDEX["Area 1 (State)"], record.get("area1", ""))
#         set_cell(
#             row_idx, _COLUMN_INDEX["Area 2 (District / City)"], record.get("area2", ""))
#         set_cell(
#             row_idx, _COLUMN_INDEX["Area 3 (Locality)"], record.get("area3", ""))
#         set_cell(
#             row_idx, _COLUMN_INDEX["Area 4 (Street)"], record.get("area4", ""))

#         set_cell(row_idx, _COLUMN_INDEX["Appointment Date"],
#                  record.get("appt_date", ""))

#         # Proposed Date — typed by hand, no special fill, displayed as
#         # "14 August 2026" regardless of how it was typed in.
#         date_cell = set_cell(
#             row_idx, _COLUMN_INDEX["Proposed Date"], record.get("proposed_date"))
#         date_cell.number_format = "d mmmm yyyy"

#         phone_digits = clean_phone_for_wa(record.get("install_mobile1", ""))

#         link_cell = set_cell(row_idx, _COLUMN_INDEX["WhatsApp Chat Link"],
#                              wa_link_formula(row_idx, phone_digits))
#         link_cell.font = Font(name=FONT, size=10,
#                               color="1155CC", underline="single")

#         set_cell(row_idx, _COLUMN_INDEX["WhatsApp Chat Message"],
#                  message_formula(row_idx, specialist_name, nds_id))

#         set_cell(row_idx, _COLUMN_INDEX["Product"],
#                  record.get("sales_info_product", ""))

#         # Filter(s) — combined text from CCS Note data, looked up by
#         # sales_no; blank if this customer had no CCS Note data captured.
#         raw_filters_text = filters_by_sales_no.get(
#             record.get("sales_no", ""), "")
#         set_cell(row_idx, _COLUMN_INDEX["Filter(s)"],
#                  _number_filters_text(raw_filters_text))

#         # WhatsApp Number — hidden helper column. The macro (if installed)
#         # reads this directly instead of re-deriving the phone number, and
#         # it's also now the only place the phone number is persisted
#         # between runs (see _load_carryforward_data). Forced to Excel's
#         # text format ('@') so re-saving/re-opening the file can't quietly
#         # convert it to a real number the way Sales No. cells could.
#         wa_number_cell = set_cell(
#             row_idx, _COLUMN_INDEX["WhatsApp Number"], phone_digits)
#         wa_number_cell.number_format = "@"

#     widths = {
#         "Sales No.": 14, "Installation / Service Contact Person": 26,
#         "Installation / Service Address": 40,
#         "Area 4 (Street)": 22, "Area 3 (Locality)": 22,
#         "Area 2 (District / City)": 20, "Area 1 (State)": 16,
#         "Appointment Date": 12, "Proposed Date": 14,
#         "WhatsApp Chat Link": 18, "WhatsApp Chat Message": 60,
#         "Product": 16, "Filter(s)": 40, "WhatsApp Number": 12,
#     }
#     for name, w in widths.items():
#         ws.column_dimensions[_col_letter(name)].width = w
#     # Header row AND the first two columns (Sales No., Installation /
#     # Service Contact Person) stay frozen — the freeze point is the cell
#     # diagonally past both, i.e. the first cell of column C.
#     ws.freeze_panes = "C2"


# def _build_filters_lookup(ccs_rows):
#     """
#     Turns the flat list of {sales_no, cust_name, product, last_change}
#     rows from get_all_ccs_note_cards() into {sales_no: combined_text},
#     one "Product: date" line per filter (real line breaks, so it reads
#     like pressing Alt+Enter between each one), sorted alphabetically by
#     product name for a consistent read.
#     """
#     by_customer = {}
#     for r in ccs_rows:
#         by_customer.setdefault(r["sales_no"], []).append(
#             (r["product"], r["last_change"]))

#     lookup = {}
#     for sales_no, filters in by_customer.items():
#         filters_sorted = sorted(filters, key=lambda f: f[0])
#         lookup[sales_no] = "\n".join(
#             f"{product}: {last_change if last_change else 'not yet changed'}"
#             for product, last_change in filters_sorted
#         )
#     return lookup


# def clean_text(value):
#     """
#     Strips leading/trailing whitespace and collapses any run of spaces
#     or tabs down to a single space — applied to every piece of scraped
#     text before it reaches a cell, since the app's UI has been observed
#     to hand back things like "PRESINT 11  PUTRAJAYA" (double space).
#     Deliberately only touches spaces/tabs, not newlines — the Filters
#     column's line breaks (one filter per line) need to survive this.
#     Non-strings (numbers, dates, None) pass through untouched.
#     """
#     if not isinstance(value, str):
#         return value
#     return re.sub(r"[ \t]+", " ", value.strip())


# _AREA_FIELD_TO_COLUMN = {
#     "area1": "Area 1 (State)",
#     "area2": "Area 2 (District / City)",
#     "area3": "Area 3 (Locality)",
#     "area4": "Area 4 (Street)",
# }


# def _load_existing_area_values():
#     """
#     Reads the Area 1-4 columns back from whichever export file already
#     exists, keyed by sales_no -> {"area1": ..., "area2": ..., ...}.
#     This is what makes anything already sitting in those columns
#     (typed by hand OR auto-filled on an earlier run) survive being
#     carried forward run after run instead of getting blanked out or
#     silently overwritten every time the sheet is rebuilt. New customers
#     (no existing row) just aren't in this dict. Returns {} if no export
#     file exists yet, or if it's an older file from before these columns
#     existed (in which case everything auto-fills fresh, same as a
#     first-ever run).
#     """
#     target_file = CARRYFORWARD_SOURCE_FILE
#     if not target_file:
#         return {}
#     wb = None
#     try:
#         wb = openpyxl.load_workbook(target_file, data_only=False)
#         ws = wb.active
#         header_row = [c.value for c in ws[1]]
#         if "Sales No." not in header_row:
#             return {}
#         sales_no_col = header_row.index("Sales No.") + 1

#         area_cols = {}
#         for field, column_name in _AREA_FIELD_TO_COLUMN.items():
#             if column_name in header_row:
#                 area_cols[field] = header_row.index(column_name) + 1

#         if not area_cols:
#             return {}  # older file, from before Area 1-4 existed — nothing to carry over

#         existing = {}
#         for row_idx in range(2, ws.max_row + 1):
#             sales_no = _normalize_sales_no(
#                 ws.cell(row=row_idx, column=sales_no_col).value)
#             if not sales_no:
#                 continue
#             values = {}
#             for field, col_idx in area_cols.items():
#                 val = ws.cell(row=row_idx, column=col_idx).value
#                 if val:
#                     values[field] = val
#             if values:
#                 existing[sales_no] = values
#         return existing
#     except Exception as e:
#         print(f"  !! Could not read existing Area values from {target_file} "
#               f"({e}) — starting fresh for all customers.")
#         return {}
#     finally:
#         # Always closed explicitly here, on our own terms, instead of
#         # leaving it to Python's garbage collector — relying on GC timing
#         # for this is what causes the harmless-but-noisy "Exception
#         # ignored...ValueError: I/O operation on closed file" warning
#         # some Python versions print at interpreter shutdown.
#         if wb is not None:
#             wb.close()


# def build_vba_macro():
#     """
#     Builds the ready-to-paste VBA macro using the CURRENT column
#     positions (same _col_letter/_COLUMN_INDEX lookups as everything
#     else) — this is what makes the macro immune to going stale after a
#     future column reorder, which is exactly what broke it last time (it
#     was hardcoded to a fixed column for the WhatsApp Link, and drifted
#     out of sync once the columns were reordered).

#     Run `python main.py --show-macro` any time you need a fresh, correct
#     copy — safer than trusting a copy/pasted version to still be right.
#     """
#     link_col_num = _COLUMN_INDEX["WhatsApp Chat Link"]
#     link_col_letter = _col_letter("WhatsApp Chat Link")
#     number_col_letter = _col_letter("WhatsApp Number")
#     message_col_letter = _col_letter("WhatsApp Chat Message")
#     return (
#         "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)\n"
#         "    Dim r As Long, num As String, msg As String, url As String\n"
#         f"    If Target.Column <> {link_col_num} Then Exit Sub   ' column {link_col_letter} = WhatsApp Chat Link\n"
#         "    r = Target.Row\n"
#         "    If r < 2 Then Exit Sub\n"
#         f'    num = Trim(Cells(r, "{number_col_letter}").Value)       \' hidden helper column\n'
#         f'    msg = Cells(r, "{message_col_letter}").Value              \' WhatsApp Chat Message\n'
#         '    If num = "" Or msg = "" Then Exit Sub\n'
#         "    ' web.whatsapp.com (not wa.me) on purpose — wa.me links often\n"
#         "    ' get intercepted by WhatsApp Desktop if it's installed, and\n"
#         "    ' the desktop app silently drops the pre-filled text for a\n"
#         "    ' chat that already exists. Routing through web.whatsapp.com\n"
#         "    ' forces it into an actual browser tab, where the text\n"
#         "    ' reliably shows up.\n"
#         '    url = "https://web.whatsapp.com/send?phone=" & num & "&text=" & WorksheetFunction.EncodeURL(msg)\n'
#         "    ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True\n"
#         "    Cancel = True\n"
#         "End Sub"
#     )


# def _load_existing_proposed_dates():
#     """
#     Reads the Proposed Date column back from whichever export file
#     already exists, keyed by sales_no. Mirrors _load_existing_area_values
#     exactly, for the exact same reason: without this, a typed-in
#     Proposed Date gets silently wiped back to blank every time the
#     script is run again (confirmed by testing — re-running after new
#     customers appear in the app was blanking out dates already entered
#     for existing ones, since a fresh scrape has no way to know what you
#     typed in Excel afterward). Returns {} if no export file exists yet,
#     or if it's from before this column existed.
#     """
#     target_file = CARRYFORWARD_SOURCE_FILE
#     if not target_file:
#         return {}
#     wb = None
#     try:
#         wb = openpyxl.load_workbook(target_file, data_only=False)
#         ws = wb.active
#         header_row = [c.value for c in ws[1]]
#         if "Sales No." not in header_row or "Proposed Date" not in header_row:
#             return {}
#         sales_no_col = header_row.index("Sales No.") + 1
#         date_col = header_row.index("Proposed Date") + 1

#         existing = {}
#         for row_idx in range(2, ws.max_row + 1):
#             sales_no = _normalize_sales_no(
#                 ws.cell(row=row_idx, column=sales_no_col).value)
#             date_val = ws.cell(row=row_idx, column=date_col).value
#             if sales_no and date_val:
#                 existing[sales_no] = date_val
#         return existing
#     except Exception as e:
#         print(f"  !! Could not read existing Proposed Date values from "
#               f"{target_file} ({e}) — starting fresh for all customers.")
#         return {}
#     finally:
#         if wb is not None:
#             wb.close()


# def write_output(records, ccs_rows=None, filters_override=None, app_order=None,
#                  specialist_name="Hanis", nds_id="NDS35095"):
#     """
#     Writes the export with the columns in the order specified at the
#     top of this section (sales_no, install_contact_person,
#     install_address, Area 4-1, appt_date, proposed_date, WhatsApp Link,
#     WhatsApp Message, sales_info_product, Filters, wa_number) — all in
#     ONE sheet, no separate CCS Notes/Route Plan sheet.

#     specialist_name/nds_id feed straight into message_formula() — see
#     its docstring — so every caller should really be passing the
#     values from _load_or_prompt_identity() rather than relying on
#     these defaults, which only exist so this function still works if
#     called without them.

#     Area 1-4 auto-fill when the script is confident (free, offline
#     postcode + keyword matching — see the "Area auto-fill" block near
#     the top of this file), and stay blank otherwise for you to fill in
#     by hand. Once any Area cell has something in it — auto-filled OR
#     typed by hand — every future run preserves it exactly as-is; it's
#     for filtering in Excel yourself, not something this script keeps
#     trying to re-guess.

#     TWO POSSIBLE OUTPUTS, chosen automatically:

#     1. If TEMPLATE_FILE ("macro.xlsm") exists, this
#        writes INTO a copy of it (loaded with keep_vba=True so its
#        macro survives) and saves as OUTPUT_FILE_XLSM. In that file,
#        double-clicking a WhatsApp Link cell runs the macro, which opens
#        WhatsApp with the message already filled in — no paste needed,
#        because VBA's FollowHyperlink isn't subject to the 255-character
#        cap that the HYPERLINK() *formula* has.

#     2. Otherwise, this falls back to a plain OUTPUT_FILE (.xlsx) with
#        no macro — the WhatsApp Link column still works via single
#        click, it just opens the bare chat (see wa_link_formula()'s
#        docstring) rather than pre-filling the message.

#     ONE-TIME SETUP for option 1 (only needs doing once, ever —
#     openpyxl can't write compiled VBA itself, so this part is manual).
#     IMPORTANT: if you already pasted an earlier version of this macro
#     (from before a column reorder), it's now WRONG — the column
#     letters below match the CURRENT layout at the time this docstring
#     was last written, but the safest option is always to run
#     `python main.py --show-macro` and paste whatever it prints, since
#     that's generated fresh from the current column layout every time
#     rather than relying on this text staying in sync.
#       a. Run the script once normally so a plain cuckoo_export.xlsx
#          exists with the columns/formulas already in it.
#       b. Open that file in Excel. Press Alt+F11 to open the VBA editor.
#       c. In the Project pane on the left, double-click the entry for
#          this sheet (e.g. "Sheet1 (Cuckoo Export)") — NOT "Insert >
#          Module". This must be the sheet's own code-behind so the
#          double-click event actually fires.
#       d. Paste in (replacing any earlier version):

#            Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
#                Dim r As Long, num As String, msg As String, url As String
#                If Target.Column <> 9 Then Exit Sub   ' column I = WhatsApp Chat Link
#                r = Target.Row
#                If r < 2 Then Exit Sub
#                num = Trim(Cells(r, "M").Value)       ' hidden helper column
#                msg = Cells(r, "J").Value              ' WhatsApp Chat Message
#                If num = "" Or msg = "" Then Exit Sub
#                ' web.whatsapp.com (not wa.me) on purpose — wa.me links often
#                ' get intercepted by WhatsApp Desktop if it's installed, and
#                ' the desktop app silently drops the pre-filled text for a
#                ' chat that already exists. Routing through web.whatsapp.com
#                ' forces it into an actual browser tab, where the text
#                ' reliably shows up.
#                url = "https://web.whatsapp.com/send?phone=" & num & "&text=" & WorksheetFunction.EncodeURL(msg)
#                ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True
#                Cancel = True
#            End Sub

#       e. Close the VBA editor. File > Save As > "Excel Macro-Enabled
#          Workbook (*.xlsm)" > save it as exactly "macro.xlsm", in the
#          project's root folder (next to this script — NOT inside any
#          of the per-person NDS-ID/name export folders).
#       f. From then on, every run of this script detects that file and
#          writes into it automatically — this setup never needs
#          repeating (unless the columns change again).
#     """
#     ccs_rows = ccs_rows or []

#     if not records:
#         print("No records captured — nothing written.")
#         return OUTPUT_FILE

#     for r in records:
#         r["sales_no"] = _normalize_sales_no(r.get("sales_no"))

#     # Area 1-4: for each customer, whatever's already in the existing
#     # file (typed by hand OR auto-filled on an earlier run) always wins
#     # and is carried over untouched. Only a field that's genuinely never
#     # been filled before gets a fresh auto-fill attempt from
#     # detect_areas() — and even then, only the specific area1/2/3/4
#     # fields it's confident about get filled; the rest stay blank for
#     # you to fill in by hand, same as before. This has to happen BEFORE
#     # sorting below, since the sort is now based on these values.
#     existing_area_values = _load_existing_area_values()
#     for r in records:
#         existing = existing_area_values.get(r.get("sales_no", ""), {})
#         auto = detect_areas(r.get("install_address", ""))
#         for field in ("area1", "area2", "area3", "area4"):
#             r[field] = existing.get(field) or auto.get(field, "")

#     # Proposed Date: same protection as Area 1-4 above — a customer's
#     # already-typed date always wins over a fresh (blank) scrape value.
#     # A truly new customer has nothing here yet, so it just stays blank
#     # for you to fill in, same as before.
#     existing_proposed_dates = _load_existing_proposed_dates()
#     for r in records:
#         existing_date = existing_proposed_dates.get(r.get("sales_no", ""))
#         if existing_date:
#             r["proposed_date"] = existing_date

#     # Sort: primarily by Area 1 -> Area 2 -> Area 3 -> Area 4, grouping
#     # customers by location as closely as the auto-fill/manual data
#     # allows — this is the actual point of having Area columns at all,
#     # so the finished sheet should reflect it directly, not just leave
#     # grouping to Excel's filter dropdowns. A blank value at any level
#     # sorts AFTER a real value at that same level, so fully-categorized
#     # rows cluster together and the ones still needing attention (blank
#     # Area) end up together too, easy to spot.
#     #
#     # Within an identical Area 1-4 combination (very common — e.g. every
#     # "blank/blank/blank/blank" new customer, or several people in the
#     # same precinct), the tiebreaker is app_order if the caller supplied
#     # one (run() does, from the order customers were actually encountered
#     # scrolling through the app this run — which can genuinely differ
#     # from any previous export's order) — falling back to Sales No. for
#     # a caller that didn't supply one (e.g. --resort).
#     app_order_lookup = {sn: idx for idx,
#                         sn in enumerate(app_order)} if app_order else {}

#     def _area_sort_key(r):
#         def level(value):
#             return (1, "") if not value else (0, value)
#         sales_no = r.get("sales_no", "")
#         return (
#             level(r.get("area1", "")),
#             level(r.get("area2", "")),
#             level(r.get("area3", "")),
#             level(r.get("area4", "")),
#             app_order_lookup.get(sales_no, len(app_order_lookup)),
#             sales_no,
#         )

#     records = sorted(records, key=_area_sort_key)

#     filters_by_sales_no = filters_override if filters_override is not None else _build_filters_lookup(
#         ccs_rows)

#     if os.path.exists(TEMPLATE_FILE):
#         try:
#             wb = openpyxl.load_workbook(TEMPLATE_FILE, keep_vba=True)
#             ws = wb.active
#             # Clear out any previous run's rows before writing fresh ones,
#             # but leave row 1 (headers) and the macro itself untouched.
#             if ws.max_row > 1:
#                 ws.delete_rows(2, ws.max_row - 1)
#             _populate_sheet(ws, records, filters_by_sales_no,
#                             specialist_name, nds_id)
#             wb.save(OUTPUT_FILE_XLSM)
#             wb.close()
#             return OUTPUT_FILE_XLSM
#         except Exception as e:
#             print(f"  !! Could not write into {TEMPLATE_FILE} ({e}) — "
#                   f"falling back to a plain .xlsx instead.")

#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Cuckoo Export"
#     _populate_sheet(ws, records, filters_by_sales_no, specialist_name, nds_id)
#     wb.save(OUTPUT_FILE)
#     wb.close()
#     return OUTPUT_FILE


# def resort_existing_file():
#     """
#     Preview/rebuild mode: `python main.py --resort`. Reads whichever
#     export file already exists for this specialist+month (from a
#     previous full scrape — no device/Appium needed here at all) and
#     rewrites it fresh: same sales_no sort, same Area values, same
#     proposed_date and Filters text, but formulas/formatting
#     regenerated from scratch. Useful for e.g. picking up a formatting
#     change without a full re-scrape.

#     Unlike a normal run(), this does NOT bump the version number in
#     the filename — it's a formatting refresh of the CURRENT latest
#     version, not a new capture, so it rewrites that exact file in
#     place.

#     proposed_date, Area, and the Filters text are all read back from
#     the existing file and carried over untouched (Area gets carried
#     over automatically by write_output() itself, same as a normal run
#     — see _load_existing_area_values()).
#     """
#     global OUTPUT_FILE, OUTPUT_FILE_XLSM

#     name, nds_id = _load_or_prompt_identity()
#     month_label = _prompt_month_label()
#     # This also sets CARRYFORWARD_SOURCE_FILE to the current latest
#     # version file (if any) — the OUTPUT_FILE/OUTPUT_FILE_XLSM it
#     # computes are the NEXT version, which is overridden below since
#     # --resort must not bump the version.
#     _apply_identity_to_filenames(name, nds_id, month_label)

#     target_file = CARRYFORWARD_SOURCE_FILE
#     if not target_file:
#         print(f"Couldn't find an existing export for {name} ({nds_id}) in "
#               f"{month_label} — run main.py normally at least once first.")
#         return

#     if target_file.endswith(".xlsm"):
#         OUTPUT_FILE_XLSM = target_file
#         OUTPUT_FILE = target_file[:-len(".xlsm")] + ".xlsx"
#     else:
#         OUTPUT_FILE = target_file
#         OUTPUT_FILE_XLSM = target_file[:-len(".xlsx")] + ".xlsm"

#     wb = openpyxl.load_workbook(target_file, data_only=False)
#     try:
#         ws = wb.active

#         # Read by HEADER NAME, not fixed column number — this is the actual
#         # fix for a real bug: an existing file from an older layout (before
#         # a column reorder) was being read with the CURRENT layout's fixed
#         # positions, silently pulling data from the wrong columns entirely
#         # (e.g. Filters' old position ending up mislabeled as
#         # sales_info_product). Reading by name means this can't happen
#         # again even if columns get reordered in the future.
#         header_row = [c.value for c in ws[1]]
#         col = {col_name: idx + 1 for idx,
#                col_name in enumerate(header_row) if col_name}

#         required = ["Sales No.", "Appointment Date", "Installation / Service Contact Person",
#                     "Installation / Service Address", "Proposed Date",
#                     "Product", "Filter(s)", "WhatsApp Number"]
#         missing = [name for name in required if name not in col]
#         if missing:
#             print(f"  !! {target_file}'s header row is missing {missing} — "
#                   f"it looks like it's from a very different version of this "
#                   f"script. Run main.py normally (a full scrape) to regenerate "
#                   f"it in the current layout, then --resort will work again.")
#             return

#         records = []
#         filters_by_sales_no = {}
#         for row_idx in range(2, ws.max_row + 1):
#             sales_no = _normalize_sales_no(
#                 ws.cell(row=row_idx, column=col["Sales No."]).value)
#             if not sales_no:
#                 continue
#             records.append({
#                 "sales_no": sales_no,
#                 "appt_date": ws.cell(row=row_idx, column=col["Appointment Date"]).value or "",
#                 "install_contact_person": ws.cell(row=row_idx, column=col["Installation / Service Contact Person"]).value or "",
#                 "install_mobile1": _normalize_phone_digits(
#                     ws.cell(row=row_idx, column=col["WhatsApp Number"]).value),
#                 "install_address": ws.cell(row=row_idx, column=col["Installation / Service Address"]).value or "",
#                 "proposed_date": ws.cell(row=row_idx, column=col["Proposed Date"]).value,
#                 "sales_info_product": ws.cell(row=row_idx, column=col["Product"]).value or "",
#             })
#             filters_text = ws.cell(row=row_idx, column=col["Filter(s)"]).value
#             if filters_text:
#                 filters_by_sales_no[sales_no] = filters_text
#     finally:
#         wb.close()

#     print(f"Rebuilding {len(records)} existing record(s) from {target_file} "
#           f"(no device/Appium needed for this)...")
#     write_output(records, filters_override=filters_by_sales_no,
#                  specialist_name=name, nds_id=nds_id)


# # ============================================================
# # Contact export (.vcf) — item 1 from the handoff's open list
# # ============================================================
# #
# # This does NOT push contacts to any phone directly (no USB contact
# # push, no device API call) — it just builds a standard vCard file
# # from whichever export file already exists, that Hanis or a colleague
# # can send to their OWN phone (email to self, WhatsApp to self, Google
# # Drive, whatever's easiest) and tap once to import every customer as
# # a real contact in one go, no retyping.
# #
# # Kept deliberately minimal per Hazim's call: just a name and a number,
# # no NOTE field or extra metadata — fewer fields means fewer ways a
# # phone's import screen can show something odd.

# def _vcard_escape(text):
#     """
#     Escapes the characters vCard 3.0 treats as structural — backslash,
#     comma, semicolon, and literal newlines — so a name containing any
#     of these (e.g. a comma in "SMITH, JOHN") can't corrupt the .vcf
#     file's structure on import. Order matters: backslash must be
#     escaped FIRST, or escaping the other characters would double-escape
#     the backslashes just added for them.
#     """
#     if not text:
#         return ""
#     text = text.replace("\\", "\\\\")
#     text = text.replace(",", "\\,")
#     text = text.replace(";", "\\;")
#     text = text.replace("\n", "\\n")
#     return text


# def export_contacts():
#     """
#     `python main.py --export-contacts` — reads the current export file
#     for a specialist+month (same file --resort reads; no device/Appium
#     needed) and writes a .vcf (vCard) file of every customer who has a
#     usable phone number on record.

#     Name format: "NDS | {contact person} ({sales_no})" —
#       - the "NDS |" PREFIX tags it as a work contact and clusters all
#         of these together near each other in an alphabetized contact
#         list, making bulk review/cleanup easy later;
#       - the contact person's name is what's actually read/searched day
#         to day, so it sits front and center rather than after an ID;
#       - the Sales No. in parens at the end disambiguates two customers
#         who happen to share a name, and lets you cross-reference back
#         to the Excel row if ever needed.

#     Each contact is saved as just a name and a number — no notes, no
#     extra fields — per Hazim's call to keep this minimal.

#     Customers with no usable phone number (WhatsApp Number column is
#     blank) are skipped entirely, since an empty TEL field would just
#     create a useless, unreachable contact.

#     IMPORTANT — before importing this file, DELETE any "NDS | ..."
#     contacts already on the phone from a previous import first. vCard
#     import is purely additive: it doesn't check for existing contacts,
#     so importing the same file (or a newer month's file) on top of an
#     old import without deleting first causes either duplicate contacts
#     (same customer imported twice) or stale contacts left behind (a
#     customer who's no longer in this month's list, but whose old
#     contact never gets removed just because a new file was imported).
#     Search "NDS" in Contacts (or contacts.google.com on desktop, if the
#     phone syncs to Google), select all, delete — then import the fresh
#     .vcf. Since every contact this script creates starts with the same
#     "NDS |" prefix, that search reliably catches all of them and
#     nothing else.
#     """
#     name, nds_id = _load_or_prompt_identity()
#     month_label = _prompt_month_label()
#     # Also sets CARRYFORWARD_SOURCE_FILE to the current latest version
#     # file for this person+month, if any. The OUTPUT_FILE/OUTPUT_FILE_XLSM
#     # values this also computes aren't used here — only the source lookup
#     # matters for this command.
#     _apply_identity_to_filenames(name, nds_id, month_label)

#     target_file = CARRYFORWARD_SOURCE_FILE
#     if not target_file:
#         print(f"Couldn't find an existing export for {name} ({nds_id}) in "
#               f"{month_label} — run main.py normally at least once first.")
#         return

#     wb = openpyxl.load_workbook(target_file, data_only=False)
#     try:
#         ws = wb.active
#         header_row = [c.value for c in ws[1]]
#         col = {col_name: idx + 1 for idx,
#                col_name in enumerate(header_row) if col_name}

#         required = ["Sales No.",
#                     "Installation / Service Contact Person", "WhatsApp Number"]
#         missing = [n for n in required if n not in col]
#         if missing:
#             print(f"  !! {target_file}'s header row is missing {missing} — "
#                   f"can't build contacts from it.")
#             return

#         vcards = []
#         skipped_no_phone = 0
#         for row_idx in range(2, ws.max_row + 1):
#             sales_no = _normalize_sales_no(
#                 ws.cell(row=row_idx, column=col["Sales No."]).value)
#             if not sales_no:
#                 continue

#             contact_person = (ws.cell(
#                 row=row_idx,
#                 column=col["Installation / Service Contact Person"]).value or "").strip()
#             phone_digits = _normalize_phone_digits(
#                 ws.cell(row=row_idx, column=col["WhatsApp Number"]).value)

#             if not phone_digits:
#                 skipped_no_phone += 1
#                 continue

#             display_name = (f"NDS | {contact_person} ({sales_no})"
#                             if contact_person else f"NDS | {sales_no}")
#             safe_name = _vcard_escape(display_name)
#             vcards.append(
#                 "BEGIN:VCARD\r\n"
#                 "VERSION:3.0\r\n"
#                 f"FN:{safe_name}\r\n"
#                 # N (structured name) is REQUIRED by the vCard 3.0 spec
#                 # alongside FN, even though FN alone is enough to display
#                 # a name. Android's contacts importer is lenient and
#                 # accepts FN with no N at all — this file worked fine
#                 # there. iOS's importer is strict: when it hits a vCard
#                 # missing N partway through a multi-contact file, it
#                 # chokes and silently stops, showing only whatever it
#                 # managed to parse before the failure (confirmed: this
#                 # exact symptom — only the first contact ever showing up,
#                 # across WhatsApp, Mail, and Files — is a documented iOS
#                 # vCard failure mode, not a fluke of any one app). Since
#                 # the name isn't split into first/last name components
#                 # here, the whole display name goes into N's first
#                 # component (family name) and the rest are left blank —
#                 # that's enough to satisfy the requirement without
#                 # inventing a fake name split.
#                 f"N:{safe_name};;;;\r\n"
#                 f"TEL;TYPE=CELL:{phone_digits}\r\n"
#                 "END:VCARD\r\n"
#             )
#     finally:
#         wb.close()

#     if not vcards:
#         print("No customers with a usable phone number found — nothing written.")
#         return

#     person_folder = f"{_slugify(nds_id)}_{_slugify(name)}"
#     out_path = os.path.join(
#         person_folder,
#         f"{_slugify(nds_id)}_{_slugify(name)}_{_slugify(month_label)}_contacts.vcf")
#     # newline="" so Python doesn't translate the \r\n already written
#     # above into \r\n\r\n on Windows — vCard needs exactly \r\n per line.
#     with open(out_path, "w", encoding="utf-8", newline="") as f:
#         f.write("".join(vcards))

#     print(f"Wrote {out_path} — {len(vcards)} contact(s) "
#           f"({skipped_no_phone} skipped, no phone number on file).")
#     print("Send this file to your phone (email to self, WhatsApp to self, "
#           "Google Drive, etc.) and tap it — Contacts should offer an "
#           "'Import all' option.")


# if __name__ == "__main__":
#     import sys
#     if "--resort" in sys.argv:
#         resort_existing_file()
#     elif "--show-macro" in sys.argv:
#         print(build_vba_macro())
#     elif "--export-contacts" in sys.argv:
#         export_contacts()
#     else:
#         run()
