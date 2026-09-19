"""
Cuckoo+ Service Specialist — monthly list scraper (Appium / UiAutomator2)

Entry point / orchestration only. The actual screen-reading logic
lives in scraping.py, the Excel/vCard writing lives in excel_export.py
and contacts.py, device control lives in device.py, and every tunable
setting lives in config.py — see each module's own docstring. This
file just runs the main scrape loop (run()), the --resort rebuild
mode, and the CLI dispatch at the bottom.

CONFIRMED SCREEN FLOW (from real XML captures)
--------------------------------------------------
1. LIST screen: each customer is a card with label/value pairs (NS No,
   Sales No, NS Date, Status, Appt Date, Product Name, Cust Name) plus
   one button.
2. Tapping that button does NOT open the detail screen directly — it
   opens a "Choose Option" POPUP with several choices (View Order, CCS
   Note, Appointment, Contact List, Cancel Appointment, Cancel). The
   script taps "View Order" from that popup.
3. That opens a "Customer Information" screen with TWO TABS:
     - "Address/Contact Info" (shown by default) — only the
       Installation/Service Address & Contact section is scraped (see
       config.py's note on why Billing/Emergency are skipped).
     - "Sales Info" — reached by tapping its tab header.
4. One driver.back() from the detail screen returns to the list.

HOW LABEL/VALUE PAIRING WORKS
---------------------------------
On every screen here, a label sits to the LEFT of its value, and both
share the same vertical position (y-coordinate) — but the raw reading
order in the XML doesn't alternate label-value cleanly, and the exact
pixel positions differ between devices/screen resolutions. So instead
of matching on fixed x-coordinates (which broke the first time this
was tested on a different device), the script groups elements into
rows by shared y-coordinate, then within each row takes the leftmost
text as the label and the rightmost as its value. See scraping.py.

SKIP-KNOWN-CUSTOMERS NOTE
--------------------------
run() skips re-scraping a customer whose sales_no is already in a
previous export — their details don't change once captured. Known
customers are found wherever they appear while scrolling (not assumed
to only be at the bottom) and carried forward exactly as they were
before. See excel_export._load_carryforward_data() and the "skipping
... — already captured previously" check below.

Start with LIST_LIMIT = 3 in config.py. Once the export looks right,
set it to None to process the entire list.
"""

import os
import time
import openpyxl

from config import (
    LIST_LIMIT, MAX_SCROLLS, MAX_STAGNANT_ROUNDS, KEY_FIELD, TEMPLATE_FILE,
    DEBUG, SCROLL_STEP_PERCENT,
)
import filenames
from filenames import (
    _load_or_prompt_identity, _prompt_month_label,
    _apply_identity_to_filenames, _finalize_output_file,
)
from normalize import _normalize_sales_no, _normalize_phone_digits
from device import (
    build_driver, scroll_down, compute_scroll_percent, _wait_if_paused,
    tap_element, wait_for,
)
from appium.webdriver.common.appiumby import AppiumBy
from scraping import (
    get_visible_customers_quick_or_stable, get_visible_customers_stable,
    select_popup_option, read_full_detail, get_all_ccs_note_cards,
)
from excel_export import (
    _load_carryforward_data, _build_filters_lookup, write_output,
    _format_duration, build_vba_macro,
)
from contacts import export_contacts

def run():
    import traceback

    name, nds_id = _load_or_prompt_identity()
    month_label = _prompt_month_label()
    _apply_identity_to_filenames(name, nds_id, month_label)
    print(f"Specialist: {name} ({nds_id}) | Month: {month_label}")
    write_target = filenames.OUTPUT_FILE_XLSM if os.path.exists(
        TEMPLATE_FILE) else filenames.OUTPUT_FILE
    if filenames.CARRYFORWARD_SOURCE_FILE:
        print(f"Continuing from {filenames.CARRYFORWARD_SOURCE_FILE} — this run will "
              f"write {write_target}.")
    else:
        print(f"No existing file found for {name} ({nds_id}) in {month_label} — "
              f"starting fresh. This run will write {write_target}.")

    known_sales_nos, carried_records, carried_filters = _load_carryforward_data()
    if known_sales_nos:
        print(f"Found {len(known_sales_nos)} customer(s) already captured in a "
              f"previous export — these will be SKIPPED (their details aren't "
              f"re-read, since they don't change), reused as-is if still found "
              f"in the app, and REMOVED from the output if no longer found. "
              f"Only genuinely new sales_no values get fully scraped.")

    driver = build_driver()
    time.sleep(3)

    print("Tip: type 'p' + Enter anytime during this run to pause after "
          "the current customer finishes, and 'r' + Enter to resume — "
          "the connection to the phone stays open, nothing gets re-scraped.")

    all_records = []
    all_ccs_rows = []
    seen_keys = set()
    stagnant_rounds = 0
    scroll_count = 0
    run_start_time = time.time()
    customer_durations = []
    # Every sales_no actually confirmed present in the app this run —
    # whether skipped (already known) or freshly scraped (new). Anyone
    # from a previous export who's NOT in this set by the end is treated
    # as removed from the app, PROVIDED reached_natural_end is True (see
    # below) — otherwise removal isn't safe to trust.
    confirmed_present_sales_nos = set()
    # Same information as confirmed_present_sales_nos, but as a LIST in
    # the order each sales_no was first encountered this run — a set has
    # no order at all. This is what lets the final export follow the
    # app's CURRENT ordering (which can genuinely change month to month)
    # instead of always resorting alphabetically by Sales No.
    sales_no_order = []
    # Only True if scrolling reached the genuine end of the list (no new
    # customers found after scrolling further) — NOT true if LIST_LIMIT
    # cut the run short (e.g. while testing) or the safety scroll cap
    # was hit. Removal logic below only runs when this is True, since a
    # partial scan can't distinguish "genuinely removed from the app"
    # from "just hasn't been scrolled to yet."
    reached_natural_end = False

    try:
        while True:
            _wait_if_paused()

            if LIST_LIMIT and len(all_records) >= LIST_LIMIT:
                print(f"Reached LIST_LIMIT of {LIST_LIMIT} — stopping.")
                break

            customers = get_visible_customers_quick_or_stable(
                driver, known_sales_nos, seen_keys)

            next_customer = None
            partial_customer = None
            skipped_this_round = False
            for c in customers:
                key = c["row"].get(KEY_FIELD)
                if not key or key in seen_keys:
                    continue
                if c["button"] is None or not c.get("complete", True):
                    # Card is likely only partially scrolled into view — its
                    # NS No rendered (enough to be found and keyed), but its
                    # button and/or its later fields (Cust Name is the last
                    # field on the card, so it's usually the first casualty)
                    # haven't fully rendered yet. Don't mark it seen; remember
                    # it so we can scroll IT specifically into full view,
                    # rather than treating this like "nothing new at all."
                    if partial_customer is None:
                        partial_customer = c
                    continue
                sales_no_value = _normalize_sales_no(c["row"].get("sales_no"))
                if sales_no_value:
                    # Confirmed present in the app THIS run, regardless of
                    # whether it gets skipped (already known) or scraped
                    # fresh (new) — this is what the end-of-run removal
                    # check is based on. Only append to the ORDER list the
                    # first time — the same card can be scanned again
                    # across consecutive scroll reads before it's marked
                    # "seen", and it must only claim one position.
                    if sales_no_value not in confirmed_present_sales_nos:
                        sales_no_order.append(sales_no_value)
                    confirmed_present_sales_nos.add(sales_no_value)
                if sales_no_value and sales_no_value in known_sales_nos:
                    # Already fully captured in a previous run, and a
                    # customer's details don't change once captured — so
                    # there's nothing new to read by opening this one. Mark
                    # it handled and keep scanning; a genuinely new customer
                    # could appear anywhere in the list, not necessarily
                    # after this point, so scrolling continues normally.
                    print(
                        f"  [{len(confirmed_present_sales_nos)}] skipping {sales_no_value} — already captured previously")
                    seen_keys.add(key)
                    skipped_this_round = True
                    continue
                next_customer = c
                break

            if next_customer is None:
                if partial_customer is not None:
                    # There IS a new customer here — it's just not fully
                    # rendered yet. Scroll targeted at THIS card's own
                    # position (not the generic "last customer" case below)
                    # so it comes fully into view instead of being skipped.
                    scroll_count += 1
                    if scroll_count >= MAX_SCROLLS:
                        print(
                            "Hit the safety scroll limit — stopping to avoid an infinite loop.")
                        break
                    target_percent = compute_scroll_percent(
                        driver, partial_customer["card_top_y"])
                    if DEBUG:
                        print(f"  [debug] {partial_customer['row'].get(KEY_FIELD)} not fully rendered yet — "
                              f"scrolling it into view (percent={target_percent:.3f})")
                    scroll_down(driver, percent=target_percent)
                    continue

                if skipped_this_round:
                    # We successfully handled (skipped) known customers this
                    # round — real progress, just nothing left to actively
                    # open in the CURRENT view. This must NOT count as
                    # stagnant: a long run of consecutive known customers
                    # (e.g. 50 in a row) would otherwise trip the "reached
                    # the end" check after just a couple of rounds, long
                    # before actually reaching the true end of the list —
                    # which would leave later customers unscanned and, worse,
                    # make the removal logic below wrongly think they'd
                    # disappeared from the app.
                    stagnant_rounds = 0
                    scroll_count += 1
                    if scroll_count >= MAX_SCROLLS:
                        print(
                            "Hit the safety scroll limit — stopping to avoid an infinite loop.")
                        break
                    target_percent = compute_scroll_percent(
                        driver, customers[-1]["card_top_y"]) if customers else SCROLL_STEP_PERCENT
                    scroll_down(driver, percent=target_percent)
                    continue

                stagnant_rounds += 1
                if stagnant_rounds >= MAX_STAGNANT_ROUNDS:
                    print(
                        "No new customers found after scrolling — reached the end of the list.")
                    # This is the ONLY point where we can trust the whole
                    # list was actually scanned — see confirmed_present_sales_nos
                    # / reached_natural_end usage after the loop.
                    reached_natural_end = True
                    break
                scroll_count += 1
                if scroll_count >= MAX_SCROLLS:
                    print(
                        "Hit the safety scroll limit — stopping to avoid an infinite loop.")
                    break
                target_percent = compute_scroll_percent(
                    driver, customers[-1]["card_top_y"]) if customers else SCROLL_STEP_PERCENT
                if DEBUG:
                    print(
                        f"  [debug] scrolling by measured percent={target_percent:.3f}")
                scroll_down(driver, percent=target_percent)
                continue

            stagnant_rounds = 0
            next_row = next_customer["row"]
            key = next_row[KEY_FIELD]
            seen_keys.add(key)
            customer_start_time = time.time()
            # Sales No. (not NS No.) is what actually identifies a
            # customer to the specialist, and the position number here
            # is based on confirmed_present_sales_nos — every customer
            # confirmed present so far, skipped or not — so it tracks
            # this customer's real position in the app's list, rather
            # than undercounting because earlier ones were skipped.
            display_sales_no = _normalize_sales_no(
                next_row.get("sales_no")) or "(no sales no)"
            print(
                f"[{len(confirmed_present_sales_nos)}] Opening: {display_sales_no} ({next_row.get('cust_name', '')})")

            try:
                button = next_customer["button"]
                if button is None:
                    raise RuntimeError(
                        "No 'View Order' button found near this customer's row")

                # --- Visit 1: View Order (sales/install details) ---
                try:
                    try:
                        tap_element(driver, button)
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: tapping row button] {type(e).__name__}: {e}")

                    try:
                        # select_popup_option() already waits/polls for the popup
                        # to appear — no fixed sleep needed before it.
                        select_popup_option(driver, "View Order")
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: selecting 'View Order' from popup] {type(e).__name__}: {e}")

                    try:
                        wait_for(
                            driver, (AppiumBy.XPATH, '//android.widget.Button[@text="Customer Information"]'))
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: waiting for Customer Information screen] {type(e).__name__}: {e}")

                    try:
                        detail = read_full_detail(driver)
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: reading detail screen fields] {type(e).__name__}: {e}")

                    all_records.append({**next_row, **detail})
                except Exception as e:
                    print(f"  !! Skipped View Order for {key}: {e}")
                finally:
                    driver.back()
                    # get_visible_customers() below already waits/polls for "NS
                    # No" to reappear, so no sleep needed here.

                # --- Visit 2: CCS Note (filter/consumable change data) ---
                # A fresh element lookup is required here — the `button`
                # WebElement from before driver.back() is stale now (the
                # underlying UI tree changed), so it can't just be reused for
                # a second tap the way it could within a single visit. This
                # re-find always uses the full stable read, not the quick
                # path — we're about to act on this exact customer, so it's
                # exactly the case the quick path defers to it for anyway.
                try:
                    try:
                        customers_again = get_visible_customers_stable(driver)
                        this_customer_again = next(
                            (c for c in customers_again if c["row"].get(KEY_FIELD) == key), None)
                        if this_customer_again is None or this_customer_again["button"] is None:
                            raise RuntimeError(
                                "Could not re-find this customer's row for CCS Note")
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: re-finding row after View Order] {type(e).__name__}: {e}")

                    try:
                        tap_element(driver, this_customer_again["button"])
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: tapping row button] {type(e).__name__}: {e}")

                    try:
                        select_popup_option(driver, "CCS Note")
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: selecting 'CCS Note' from popup] {type(e).__name__}: {e}")

                    try:
                        cards = get_all_ccs_note_cards(driver)
                    except Exception as e:
                        raise RuntimeError(
                            f"[stage: reading CCS Note cards] {type(e).__name__}: {e}")

                    for card in cards:
                        all_ccs_rows.append({
                            "sales_no": next_row.get("sales_no", ""),
                            "cust_name": next_row.get("cust_name", ""),
                            "product": card["product"],
                            "last_change": card["last_change"],
                        })
                except Exception as e:
                    print(f"  !! Skipped CCS Note for {key}: {e}")
                finally:
                    driver.back()
            except Exception as e:
                print(f"  !! Skipped {key} entirely due to error: {e}")

            customer_elapsed = time.time() - customer_start_time
            customer_durations.append(customer_elapsed)
            print(f"    ({_format_duration(customer_elapsed)})")

    except KeyboardInterrupt:
        # Ctrl+C mid-run. Whatever's in all_records/all_ccs_rows so far
        # still gets saved below — reached_natural_end is correctly
        # still False here, so the "remove missing customers" logic
        # stays safely off, same as any other partial run.
        print("\nInterrupted (Ctrl+C) — saving whatever was captured so "
              "far before exiting.")
    except Exception as e:
        # Anything unexpected (device disconnected, a selector that no
        # longer matches, etc.) — rather than losing every customer
        # scraped so far, save them now and surface the full traceback
        # so the actual problem is visible, instead of just crashing
        # silently with nothing written.
        print(
            f"\n!! Unexpected error, stopping early: {type(e).__name__}: {e}")
        print("Saving whatever was captured so far before exiting.")
        traceback.print_exc()
    finally:
        # driver.quit() can itself throw — most notably if the phone
        # was physically disconnected (USB yanked, cable fault, phone
        # rebooted) rather than the app/script hitting a normal error.
        # In that case there's no live session left to cleanly close,
        # and an exception raised HERE, inside finally, would propagate
        # straight out of run() and skip everything below — including
        # the save — which defeats the entire point of the except
        # clauses above. Swallowing it is safe: by this point either
        # the loop finished normally or one of the excepts above has
        # already handled and logged the real problem.
        try:
            driver.quit()
        except Exception:
            pass

    total_elapsed = time.time() - run_start_time

    if reached_natural_end:
        # The full list was genuinely scanned this run, so anyone
        # previously captured but NOT confirmed present now has actually
        # been removed from the app — drop them from the export too,
        # rather than leaving stale rows behind forever.
        carried_to_keep = {sn: rec for sn, rec in carried_records.items()
                           if sn in confirmed_present_sales_nos}
        carried_filters_to_keep = {sn: txt for sn, txt in carried_filters.items()
                                   if sn in confirmed_present_sales_nos}
        removed = sorted(set(carried_records) - confirmed_present_sales_nos)
        if removed:
            print(f"Removed {len(removed)} customer(s) no longer found in "
                  f"the app: {', '.join(removed)}")
    else:
        # Didn't confirm scanning the WHOLE list this run (LIST_LIMIT cut
        # it short, or the safety scroll cap was hit) — "not seen yet" and
        # "actually gone" can't be told apart from a partial scan, so
        # nothing gets removed this time, to be safe.
        carried_to_keep = carried_records
        carried_filters_to_keep = carried_filters
        if carried_records:
            print("Note: this run didn't confirm scanning the full list "
                  "(LIST_LIMIT or the scroll safety cap stopped it early), "
                  "so no previously-captured customers were removed even if "
                  "not seen this run.")

    # Merge: newly-scraped customers (all_records/all_ccs_rows) plus
    # whichever previously-captured customers are being kept (see above).
    # These two groups never overlap by construction — a sales_no is
    # either in known_sales_nos (skipped/carried) or it isn't (scraped
    # fresh this run) — so a plain combine is safe.
    merged_records = all_records + list(carried_to_keep.values())
    merged_filters = {**carried_filters_to_keep,
                      **_build_filters_lookup(all_ccs_rows)}

    written_to = write_output(
        merged_records, filters_override=merged_filters, app_order=sales_no_order,
        specialist_name=name, nds_id=nds_id)
    written_to = _finalize_output_file(written_to, reached_natural_end)
    print(f"Done. Wrote {written_to} — {len(all_records)} newly-scraped "
          f"customer(s), {len(carried_to_keep)} carried forward unchanged "
          f"({len(merged_records)} total), {len(all_ccs_rows)} new CCS Note "
          f"row(s) read.")
    skip_count = len(carried_to_keep)
    detail_time = sum(customer_durations)
    # Exact, not an estimate — the ONLY thing timed separately per-
    # customer is a newly-scraped customer's View Order + CCS Note visit
    # (customer_start_time/customer_elapsed). Everything else the loop
    # does — scrolling, the quick/stable screen reads, skip decisions —
    # is NOT individually timed, so "everything else" is whatever's left
    # after subtracting the one thing that IS precisely measured. This
    # bucket isn't PURELY "time skipping known customers" — it also
    # includes the screen-reading that happens right before a NEW
    # customer is found too, not just before a skip — but when most of
    # a run's customers are already-known (the normal case after month
    # one), this bucket is realistically dominated by skip-related
    # scrolling.
    other_time = max(0, total_elapsed - detail_time)
    print(f"Total time: {_format_duration(total_elapsed)}")
    if customer_durations:
        avg_seconds = detail_time / len(customer_durations)
        print(f"  New customers: {len(customer_durations)} scraped, "
              f"{_format_duration(detail_time)} total (avg {avg_seconds:.1f}s each)")
    else:
        print("  New customers: none scraped this run")
    if skip_count:
        avg_skip = other_time / skip_count
        print(f"  Everything else — scrolling, screen reads, and skipping "
              f"{skip_count} already-known customer(s): "
              f"{_format_duration(other_time)} total (~{avg_skip:.1f}s per "
              f"skipped customer, though this bucket isn't purely skip time)")
    else:
        print(f"  Everything else (scrolling/screen reads, no known "
              f"customers to skip this run): {_format_duration(other_time)}")



def resort_existing_file():
    """
    Preview/rebuild mode: `python main.py --resort`. Reads whichever
    export file already exists for this specialist+month (from a
    previous full scrape — no device/Appium needed here at all) and
    rewrites it fresh: same sales_no sort, same Area values, same
    proposed_date and Filters text, but formulas/formatting
    regenerated from scratch. Useful for e.g. picking up a formatting
    change without a full re-scrape.

    Unlike a normal run(), this does NOT bump the version number in
    the filename — it's a formatting refresh of the CURRENT latest
    version, not a new capture, so it rewrites that exact file in
    place.

    proposed_date, Area, and the Filters text are all read back from
    the existing file and carried over untouched (Area gets carried
    over automatically by write_output() itself, same as a normal run
    — see _load_existing_area_values()).
    """
    name, nds_id = _load_or_prompt_identity()
    month_label = _prompt_month_label()
    # This also sets filenames.CARRYFORWARD_SOURCE_FILE to the current
    # latest version file (if any) — the OUTPUT_FILE/OUTPUT_FILE_XLSM it
    # computes are the NEXT version, which is overridden below since
    # --resort must not bump the version.
    _apply_identity_to_filenames(name, nds_id, month_label)

    target_file = filenames.CARRYFORWARD_SOURCE_FILE
    if not target_file:
        print(f"Couldn't find an existing export for {name} ({nds_id}) in "
              f"{month_label} — run main.py normally at least once first.")
        return

    if target_file.endswith(".xlsm"):
        filenames.OUTPUT_FILE_XLSM = target_file
        filenames.OUTPUT_FILE = target_file[:-len(".xlsm")] + ".xlsx"
    else:
        filenames.OUTPUT_FILE = target_file
        filenames.OUTPUT_FILE_XLSM = target_file[:-len(".xlsx")] + ".xlsm"

    wb = openpyxl.load_workbook(target_file, data_only=False)
    try:
        ws = wb.active

        # Read by HEADER NAME, not fixed column number — this is the actual
        # fix for a real bug: an existing file from an older layout (before
        # a column reorder) was being read with the CURRENT layout's fixed
        # positions, silently pulling data from the wrong columns entirely
        # (e.g. Filters' old position ending up mislabeled as
        # sales_info_product). Reading by name means this can't happen
        # again even if columns get reordered in the future.
        header_row = [c.value for c in ws[1]]
        col = {col_name: idx + 1 for idx,
               col_name in enumerate(header_row) if col_name}

        required = ["Sales No.", "Appointment Date", "Installation / Service Contact Person",
                    "Installation / Service Address", "Proposed Date",
                    "Product", "Filter(s)", "WhatsApp Number"]
        missing = [name for name in required if name not in col]
        if missing:
            print(f"  !! {target_file}'s header row is missing {missing} — "
                  f"it looks like it's from a very different version of this "
                  f"script. Run main.py normally (a full scrape) to regenerate "
                  f"it in the current layout, then --resort will work again.")
            return

        records = []
        filters_by_sales_no = {}
        for row_idx in range(2, ws.max_row + 1):
            sales_no = _normalize_sales_no(
                ws.cell(row=row_idx, column=col["Sales No."]).value)
            if not sales_no:
                continue
            records.append({
                "sales_no": sales_no,
                "appt_date": ws.cell(row=row_idx, column=col["Appointment Date"]).value or "",
                "install_contact_person": ws.cell(row=row_idx, column=col["Installation / Service Contact Person"]).value or "",
                "install_mobile1": _normalize_phone_digits(
                    ws.cell(row=row_idx, column=col["WhatsApp Number"]).value),
                "install_address": ws.cell(row=row_idx, column=col["Installation / Service Address"]).value or "",
                "proposed_date": ws.cell(row=row_idx, column=col["Proposed Date"]).value,
                "sales_info_product": ws.cell(row=row_idx, column=col["Product"]).value or "",
            })
            filters_text = ws.cell(row=row_idx, column=col["Filter(s)"]).value
            if filters_text:
                filters_by_sales_no[sales_no] = filters_text
    finally:
        wb.close()

    print(f"Rebuilding {len(records)} existing record(s) from {target_file} "
          f"(no device/Appium needed for this)...")
    write_output(records, filters_override=filters_by_sales_no,
                 specialist_name=name, nds_id=nds_id)


# ============================================================
# Contact export (.vcf) — item 1 from the handoff's open list
# ============================================================
#
# This does NOT push contacts to any phone directly (no USB contact
# push, no device API call) — it just builds a standard vCard file
# from whichever export file already exists, that Hanis or a colleague
# can send to their OWN phone (email to self, WhatsApp to self, Google
# Drive, whatever's easiest) and tap once to import every customer as
# a real contact in one go, no retyping.
#
# Kept deliberately minimal per Hazim's call: just a name and a number,
# no NOTE field or extra metadata — fewer fields means fewer ways a
# phone's import screen can show something odd.




if __name__ == "__main__":
    import sys
    if "--resort" in sys.argv:
        resort_existing_file()
    elif "--show-macro" in sys.argv:
        print(build_vba_macro())
    elif "--export-contacts" in sys.argv:
        export_contacts()
    else:
        run()
