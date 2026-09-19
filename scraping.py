"""
scraping.py — reads the phone's screen: the customer list, the popup
menu, the View Order detail tabs, and the CCS Note (filter/consumable)
cards. Everything in here is about interpreting what's currently on
screen into Python data structures; nothing here decides what to DO
with that data (that's main.py's run() loop) or writes it anywhere
(that's excel_export.py).
"""

from appium.webdriver.common.appiumby import AppiumBy

from config import (
    LABEL_FIELD_MAP, KEY_FIELD, INSTALL_LABEL_MAP, INSTALL_ORPHAN_FIELDS,
    HEADER_TEXTS, SALES_INFO_LABEL_MAP, MAX_STAGNANT_ROUNDS, WAIT_SECONDS,
    CCS_DEDUP_FIELDS, MAX_CCS_CARD_SCROLLS, _CCS_RESERVED_LABELS, DEBUG,
)
from normalize import _normalize_sales_no
from device import (
    wait_for, tap_element, parse_bounds, read_text_safe,
    scroll_down, compute_scroll_percent,
)

import time

_STABLE_READ_TOTAL_CALLS = 0
_STABLE_READ_TOTAL_ATTEMPTS = 0
_STABLE_READ_GAVE_UP_COUNT = 0


def pair_fields(elements, known_labels=None, skip_texts=None, y_tol=8):
    """
    Generic label/value pairing based on RELATIVE position, not fixed
    pixel coordinates — this makes it work across different screen
    resolutions/devices, since it never assumes a label sits at a
    specific x value. Elements are grouped into rows by shared
    y-coordinate; within each row, the first element (left to right)
    whose text matches a KNOWN label is treated as the label, and the
    element immediately after it becomes its value. Anything sitting
    further left than the label (e.g. a small numbered badge next to
    each card) is simply ignored rather than mistaken for the label.

    Returns (row, orphans):
      row     -> {label_text: value_text} for every recognized label
      orphans -> list of value texts (top-to-bottom order) for rows
                 that don't match a known label — i.e. unlabeled
                 fields like a bare address or name with no caption.
    """
    skip_texts = skip_texts or set()
    entries = []
    for el in elements:
        parsed = parse_bounds(el.get_attribute("bounds"))
        if not parsed:
            continue
        x1, y1, _, _ = parsed
        text = read_text_safe(el)
        if text in skip_texts:
            continue
        entries.append((y1, x1, text))

    entries.sort(key=lambda e: (e[0], e[1]))

    rows = []
    for entry in entries:
        placed = False
        for row in rows:
            if abs(row[0][0] - entry[0]) <= y_tol:
                row.append(entry)
                placed = True
                break
        if not placed:
            rows.append([entry])

    row_dict = {}
    orphans = []
    for row in rows:
        row_sorted = sorted(row, key=lambda e: e[1])  # left to right
        row_y = row_sorted[0][0]

        if known_labels is not None:
            # Find the first element in this row whose text is a genuine
            # known label — this way something sitting further left (like
            # a small numbered badge next to each card) gets ignored
            # instead of being mistaken for the label itself.
            label_idx = next((i for i, e in enumerate(
                row_sorted) if e[2] in known_labels), None)
            if label_idx is not None:
                label_text = row_sorted[label_idx][2]
                value_text = row_sorted[label_idx +
                                        1][2] if label_idx + 1 < len(row_sorted) else ""
                row_dict[label_text] = value_text
                # anything before label_idx (e.g. a badge number) is just ignored
            else:
                for _, _, text in row_sorted:
                    orphans.append((row_y, text))
        else:
            if len(row_sorted) >= 2:
                row_dict[row_sorted[0][2]] = row_sorted[-1][2]
            else:
                orphans.append((row_y, row_sorted[0][2]))

    orphans.sort(key=lambda t: t[0])
    return row_dict, [t for _, t in orphans]


# ============================================================
# List screen
# ============================================================

def group_into_rows(entries, y_tol=8):
    """entries: list of (y1, x1, text). Returns rows: list of rows,
    each row a list of (y1, x1, text) sorted left-to-right, rows
    themselves sorted top-to-bottom."""
    entries = sorted(entries, key=lambda e: (e[0], e[1]))
    rows = []
    for e in entries:
        placed = False
        for row in rows:
            if abs(row[0][0] - e[0]) <= y_tol:
                row.append(e)
                placed = True
                break
        if not placed:
            rows.append([e])
    rows = [sorted(r, key=lambda e: e[1]) for r in rows]
    rows.sort(key=lambda r: r[0][0])
    return rows


def rows_to_fields(rows_chunk, known_labels):
    """Same label-search-per-row logic as pair_fields, applied to an
    already-sliced chunk of rows belonging to one customer."""
    result = {}
    for row in rows_chunk:
        label_idx = next((i for i, e in enumerate(
            row) if e[2] in known_labels), None)
        if label_idx is not None:
            label_text = row[label_idx][2]
            value_text = row[label_idx + 1][2] if label_idx + \
                1 < len(row) else ""
            result[label_text] = value_text
    return result


def get_visible_customers(driver):
    """
    Reads EVERY TextView and Button on screen in one global query (no
    per-card scoped searches — those don't reliably scope on this
    Appium/UiAutomator2 setup, which is what broke the earlier version),
    then figures out which elements belong to which customer purely by
    on-screen position: each "NS No" row marks where a new card starts.

    Returns a list of {"row": {...parsed fields...}, "button": element_or_None}
    for every customer currently visible.
    """
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="NS No"]'))

    text_elements = driver.find_elements(
        AppiumBy.CLASS_NAME, "android.widget.TextView")
    entries = []
    for el in text_elements:
        parsed = parse_bounds(el.get_attribute("bounds"))
        if not parsed:
            continue
        x1, y1, _, _ = parsed
        entries.append((y1, x1, read_text_safe(el)))
    rows = group_into_rows(entries)

    marker_indices = [i for i, row in enumerate(
        rows) if any(t == "NS No" for _, _, t in row)]
    if DEBUG:
        print(
            f"  [debug] found {len(rows)} row(s) total, {len(marker_indices)} 'NS No' marker(s)")

    buttons = driver.find_elements(AppiumBy.XPATH, '//android.widget.Button')
    button_positions = []
    for b in buttons:
        parsed = parse_bounds(b.get_attribute("bounds"))
        if parsed:
            button_positions.append((parsed[1], b))
    button_positions.sort(key=lambda t: t[0])

    known_labels = set(LABEL_FIELD_MAP.keys())
    customers = []
    for idx, start in enumerate(marker_indices):
        end = marker_indices[idx + 1] if idx + \
            1 < len(marker_indices) else len(rows)
        chunk = rows[start:end]
        card_top_y = chunk[0][0][0]
        next_top_y = rows[marker_indices[idx + 1]][0][0] if idx + \
            1 < len(marker_indices) else float("inf")

        field_dict = rows_to_fields(chunk, known_labels)
        parsed_row = {LABEL_FIELD_MAP[k]: v for k,
                      v in field_dict.items() if k in LABEL_FIELD_MAP}

        matching_button = next(
            (b_el for b_y, b_el in button_positions if card_top_y <= b_y < next_top_y), None)
        # A card missing any expected field (most commonly Cust Name,
        # since it's the LAST field on the card) usually means it's only
        # partially scrolled into view — its bottom hasn't fully rendered
        # yet, not that the data is genuinely blank.
        complete = len(field_dict) == len(LABEL_FIELD_MAP)
        customers.append({
            "row": parsed_row,
            "button": matching_button,
            "card_top_y": card_top_y,
            "complete": complete,
        })

    if DEBUG:
        for c in customers:
            print(
                f"  [debug] parsed customer: {c['row']}  (button found: {c['button'] is not None})")

    return customers


def get_visible_customers_stable(driver, max_attempts=4, settle_delay=0.4):
    """
    Reads the screen repeatedly until two consecutive reads agree on
    which NS numbers are visible. A single read can catch the view
    mid-render — Android list rows get REUSED as you scroll, so reading
    too early can show a row still holding the PREVIOUS customer's name
    while its NS No has already updated to the new one. That's what was
    causing blank names and, worse, a name attached to the wrong NS
    number. Waiting for two matching reads in a row avoids trusting a
    transitional, half-updated state.
    """
    global _STABLE_READ_TOTAL_CALLS, _STABLE_READ_TOTAL_ATTEMPTS, _STABLE_READ_GAVE_UP_COUNT

    prev_signature = None
    customers = []
    _STABLE_READ_TOTAL_CALLS += 1
    for attempt_num in range(1, max_attempts + 1):
        customers = get_visible_customers(driver)
        signature = tuple(c["row"].get(KEY_FIELD, "") for c in customers)
        if signature == prev_signature and signature:
            _STABLE_READ_TOTAL_ATTEMPTS += attempt_num
            return customers
        prev_signature = signature
        time.sleep(settle_delay)
    # Never got two matching reads in a row within max_attempts — the
    # screen genuinely wouldn't settle this time. Counted separately
    # from the normal attempt tally since this is a stronger signal
    # than "needed a retry" — it's "retrying didn't even help."
    _STABLE_READ_TOTAL_ATTEMPTS += max_attempts
    _STABLE_READ_GAVE_UP_COUNT += 1
    return customers


def get_visible_customers_quick_or_stable(driver, known_sales_nos, seen_keys):
    """
    Speed optimization on top of get_visible_customers_stable() — see
    the "SPEED NOTE" in this file's top docstring for the full
    reasoning. Short version: the slow stable-read protocol only
    matters for a card about to be SKIPPED or ACTED ON here. This does
    one fast, single read first; if every customer currently visible is
    either already handled this run (seen_keys) or already known from a
    previous export (known_sales_nos), nothing about that card's other
    fields is being trusted right now — only its identity, which is
    what get_visible_customers_stable()'s own docstring confirms
    updates promptly — so the fast read is used as-is. The instant
    anything new or only-partially-rendered shows up, this falls back
    to the full, careful stable read before touching it.
    """
    quick = get_visible_customers(driver)
    if not quick:
        return get_visible_customers_stable(driver)

    for c in quick:
        if not c.get("complete", True):
            return get_visible_customers_stable(driver)
        key = c["row"].get(KEY_FIELD)
        if key in seen_keys:
            continue
        sales_no = _normalize_sales_no(c["row"].get("sales_no"))
        if sales_no and sales_no in known_sales_nos:
            continue
        # Something here isn't already-handled/already-known — trust
        # nothing from the fast read, do it properly.
        return get_visible_customers_stable(driver)

    return quick


# ============================================================
# Popup menu ("Choose Option") -> detail screen
# ============================================================

def select_popup_option(driver, option_text, timeout=WAIT_SECONDS):
    # normalize-space() handles the extra leading spaces the app puts in these labels
    xpath = f'//android.widget.TextView[normalize-space(@text)="{option_text}"]'
    el = wait_for(driver, (AppiumBy.XPATH, xpath), timeout=timeout)
    tap_element(driver, el)


# ============================================================
# Detail screen: Address/Contact Info tab
# ============================================================

def get_installation_elements(driver):
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="Installation/Service Address & Contact"]'))
    return driver.find_elements(
        AppiumBy.XPATH,
        '//android.widget.TextView[@text="Installation/Service Address & Contact"]'
        '/parent::android.view.ViewGroup//android.widget.TextView'
    )


def get_sales_info_elements(driver):
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="Current Stage"]'), timeout=8)
    return driver.find_elements(AppiumBy.XPATH, '//android.widget.TextView')


def read_full_detail(driver):
    record = {}

    # --- Address/Contact Info tab (shown by default) ---
    # Billing Address & Contact and Emergency Contact are NOT scraped
    # here (see the note above INSTALL_LABEL_MAP) — only the
    # Installation/Service section is actually used downstream.
    install_row, install_orphans = pair_fields(
        get_installation_elements(driver),
        known_labels=set(INSTALL_LABEL_MAP.keys()), skip_texts=HEADER_TEXTS
    )
    for label_text, field_name in INSTALL_LABEL_MAP.items():
        record[field_name] = install_row.get(label_text, "")
    for i, field_name in enumerate(INSTALL_ORPHAN_FIELDS):
        record[field_name] = install_orphans[i] if i < len(
            install_orphans) else ""

    # --- Switch to Sales Info tab ---
    sales_tab = wait_for(
        driver, (AppiumBy.XPATH, '//android.widget.TextView[@text="Sales Info"]'))
    tap_element(driver, sales_tab)
    # get_sales_info_elements() below already waits/polls for "Current Stage"
    # to appear, so no fixed sleep needed here.

    sales_row, _ = pair_fields(
        get_sales_info_elements(driver),
        known_labels=set(SALES_INFO_LABEL_MAP.keys())
    )
    for label_text, field_name in SALES_INFO_LABEL_MAP.items():
        record[field_name] = sales_row.get(label_text, "")

    return record


# ============================================================
# Detail screen: CCS Note popup (filter/consumable change data)
# ============================================================
#
# A scrollable list of "cards" — one per physical filter/consumable UNIT
# installed, which is why the exact same product can appear several
# times in a row (e.g. "FT-1001 Sediment Filter 8 Inch" showing up 6
# times for one customer, because they have 6 identical units). Each
# full card has: a product name, a small per-unit number (1, 2, 3...),
# an interval like "(4 months)", "Last Change" -> a date, "Next Change"
# -> a date. Only product name + Last Change date are actually kept —
# interval and Next Change aren't part of what's needed here.
#
# A card is only kept if BOTH of these hold:
#   - its product name is real text, not a bare number and not the
#     literal label "Last Change"/"Next Change"
#   - it has a non-empty Last Change date
# Both checks exist because of a real, confirmed pattern in captured
# data: alongside each genuine card, the same screen sometimes also
# yields a partial/garbled read of it — either the standalone "Next
# Change" label picked up on its own, or the small per-unit number
# landing where the product name should be — and in every observed
# case, that garbled read has a BLANK Last Change while the genuine
# card next to it doesn't. Filtering on "has a real name AND has a
# Last Change" reliably keeps the real entries and drops the artifacts.
#
# "Unique" cards: product name + Last Change date must both match for
# two cards to be treated as duplicates and collapsed into one
# (CCS_DEDUP_FIELDS below). The small per-unit number is deliberately
# not part of that — this is what collapses several near-identical
# "FT-1001 Sediment Filter 8 Inch" cards, differing only by unit
# number, into a single row.

# ============================================================
# Detail screen: CCS Note popup (filter/consumable change data)
# ============================================================
#


def _is_valid_ccs_product_name(product):
    if not product:
        return False
    if product in _CCS_RESERVED_LABELS:
        return False
    if product.strip().isdigit():
        return False
    return True


def _group_ccs_texts_into_cards(card_boxes, text_entries):
    """
    Pure function — no driver access, so this is fully testable with
    synthetic data before ever touching a real device. This is the
    actual restructuring: instead of asking the phone "what text is
    inside THIS card" once per card (get_ccs_note_cards() used to call
    card_el.find_elements() in a loop — one extra device round trip per
    visible card, every single scroll), every text element currently
    visible across ALL cards is read in ONE bulk query, then sorted
    into its owning card here, in Python, using nothing but on-screen
    position — the same trick group_into_rows() already uses for the
    main customer list.

    card_boxes: list of (top_y, bottom_y) for each card container,
                in their original on-screen/document order — this
                order is preserved in the returned list, matching what
                get_ccs_note_cards() returned before.
    text_entries: list of (y1, x1, text) for every text element
                  currently visible in the card area, from ONE bulk
                  read — not scoped to any particular card.

    A text element is assigned to whichever card's [top_y, bottom_y]
    range contains its own y1. Since real cards are stacked vertically
    and don't overlap, this can't accidentally merge two cards' text
    together the way a badly-tuned distance-based grouping might — it's
    checking actual card boundaries, not guessing at spacing. A text
    element that doesn't fall inside any card's range (e.g. stray text
    from just above/below the visible card area) is simply dropped,
    matching the original behavior of only reading text that was
    inside a card element's own bounds.

    Within each card, group_into_rows() reconstructs proper top-to-
    bottom, left-to-right reading order from raw (y1, x1, text) tuples
    — the exact same ordering guarantee this file already relies on
    elsewhere (see pair_fields/get_visible_customers), rather than
    trusting incidental element order from a query. Everything AFTER
    that — pulling out the product name, finding the Last Change value,
    the validity check — is character-for-character the same logic
    get_ccs_note_cards() always used; only how the per-card text list
    gets built has changed.
    """
    entries_by_card = [[] for _ in card_boxes]
    # Check boxes top-to-bottom so a text element right on a boundary
    # (shouldn't happen for real, non-overlapping cards, but cheap
    # insurance) consistently resolves to the higher card rather than
    # being ambiguous.
    box_order = sorted(range(len(card_boxes)), key=lambda i: card_boxes[i][0])
    for y1, x1, text in text_entries:
        for i in box_order:
            top_y, bottom_y = card_boxes[i]
            # Half-open interval [top_y, bottom_y) — not <= on both ends.
            # Two cards stacked with zero gap between them (normal for a
            # scrollable list) can have one card's bottom_y exactly equal
            # to the next card's top_y; with an inclusive upper bound, a
            # text element sitting exactly on that shared line would
            # match BOTH boxes and silently get assigned to the wrong
            # (earlier) card. This guarantees every y-coordinate belongs
            # to exactly one card.
            if top_y <= y1 < bottom_y:
                entries_by_card[i].append((y1, x1, text))
                break

    cards = []
    for i, (top_y, _bottom_y) in enumerate(card_boxes):
        rows = group_into_rows(entries_by_card[i])
        texts = [t for row in rows for _, _, t in row if t]

        if not texts:
            continue

        product = texts[0]
        last_change = ""
        for j, t in enumerate(texts):
            if t == "Last Change" and j + 1 < len(texts):
                last_change = texts[j + 1]

        if not _is_valid_ccs_product_name(product):
            continue

        cards.append({
            "product": product,
            "last_change": last_change,
            "card_top_y": top_y,
        })

    return cards


def get_ccs_note_cards(driver):
    """Reads every currently-visible card on the CCS Note screen,
    keeping only ones that pass the validity checks above. See
    _group_ccs_texts_into_cards() for how the per-card split actually
    works — this function just does the two device reads (card
    container bounds, then every text element at once) and hands the
    results to that pure function."""
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="CCS Note"]'))

    card_elements = driver.find_elements(
        AppiumBy.XPATH,
        '//androidx.viewpager.widget.ViewPager//android.view.ViewGroup[@clickable="true"]'
    )
    card_boxes = []
    for card_el in card_elements:
        parsed = parse_bounds(card_el.get_attribute("bounds"))
        if not parsed:
            continue
        _x1, y1, _x2, y2 = parsed
        card_boxes.append((y1, y2))

    if not card_boxes:
        return []

    # ONE bulk read for every text element currently visible across ALL
    # cards — this replaces the old per-card find_elements() loop,
    # which cost one extra device round trip per visible card.
    text_elements = driver.find_elements(
        AppiumBy.XPATH,
        '//androidx.viewpager.widget.ViewPager//android.widget.TextView'
    )
    text_entries = []
    for el in text_elements:
        parsed = parse_bounds(el.get_attribute("bounds"))
        if not parsed:
            continue
        x1, y1, _x2, _y2 = parsed
        text = read_text_safe(el)
        if text:
            text_entries.append((y1, x1, text))

    return _group_ccs_texts_into_cards(card_boxes, text_entries)


def get_all_ccs_note_cards(driver):
    """
    Scrolls through the whole CCS Note screen, collecting cards as it
    goes. Deduplicates along the way using CCS_DEDUP_FIELDS — this
    collapses both genuinely repeated cards (same product/last change,
    different unit number) AND the same card being seen twice after a
    small scroll, with the same check. Finishes with a reconciliation
    pass (see _reconcile_ccs_cards) that separates real "not yet
    serviced" entries from parsing artifacts.
    """
    seen_keys = set()
    unique_cards = []
    scroll_count = 0
    stagnant_rounds = 0

    while scroll_count < MAX_CCS_CARD_SCROLLS and stagnant_rounds < MAX_STAGNANT_ROUNDS:
        cards = get_ccs_note_cards(driver)
        if not cards:
            break

        new_this_round = 0
        for card in cards:
            key = tuple(card[f] for f in CCS_DEDUP_FIELDS)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            unique_cards.append(card)
            new_this_round += 1

        stagnant_rounds = 0 if new_this_round > 0 else stagnant_rounds + 1

        target_percent = compute_scroll_percent(
            driver, cards[-1]["card_top_y"])
        scroll_down(driver, percent=target_percent)
        scroll_count += 1

    return _reconcile_ccs_cards(unique_cards)


def _reconcile_ccs_cards(cards):
    """
    Separates real "not yet serviced" entries from parsing artifacts —
    both look identical in isolation (same product, blank Last Change),
    so this has to look at each PRODUCT's entries together to tell them
    apart:

      - If a product has at least one entry WITH a Last Change date,
        any blank-date entries for that same product are almost
        certainly a partial/duplicate read of that same card (a
        confirmed real pattern — see get_ccs_note_cards' docstring),
        so they're dropped. Every distinct dated entry is kept (a
        product can legitimately have more than one physical unit,
        serviced on different dates).
      - If a product has NO dated entry at all, it's kept as-is with a
        blank date — that's a genuine filter that just hasn't had its
        first change recorded yet, not an artifact.
    """
    by_product = {}
    for card in cards:
        by_product.setdefault(card["product"], []).append(card)

    reconciled = []
    for product, product_cards in by_product.items():
        dated = [c for c in product_cards if c["last_change"]]
        if dated:
            reconciled.extend(dated)
        else:
            reconciled.append(product_cards[0])
    return reconciled


# ============================================================
# Main loop: scroll + scrape until nothing new appears
# ============================================================

