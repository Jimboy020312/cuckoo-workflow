"""
device.py — everything that directly talks to the phone (build the
Appium driver, tap/scroll/wait primitives), plus the live pause/resume
control (which isn't device-specific in itself, but exists purely to
control WHEN the device gets touched, so it lives here rather than in
its own tiny module).
"""

import os
import re
import sys
import time
from appium import webdriver
from appium.options.android import UiAutomator2Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import (
    APP_PACKAGE, APP_ACTIVITY, WAIT_SECONDS,
    SCROLL_STEP_PERCENT, SCROLL_REGION_TOP_FRACTION,
    SCROLL_REGION_HEIGHT_FRACTION, SCROLL_SETTLE_SECONDS,
)

def build_driver():
    options = UiAutomator2Options()
    options.platform_name = "Android"
    options.app_package = APP_PACKAGE
    options.app_activity = APP_ACTIVITY
    options.no_reset = True
    # Appium auto-terminates a session after this many seconds with NO
    # commands sent to the device — 60s by default. The pause feature
    # deliberately sends nothing to the phone for as long as it's
    # paused, which is exactly what trips this: a quick test pause
    # (a few seconds) resumes fine, but a real pause (actually using
    # the phone for a bit) exceeds 60s and gets the session killed
    # server-side — confirmed by testing: this is exactly what an
    # InvalidSessionIdException right after "Resuming..." means. Set
    # generously high (1 hour) so a genuine break doesn't trigger it;
    # this doesn't weaken any other safety net — MAX_SCROLLS and the
    # KeyboardInterrupt/Exception handling still catch a run that's
    # actually stuck for an unrelated reason.
    options.new_command_timeout = 3600
    return webdriver.Remote("http://127.0.0.1:4723", options=options)


def wait_for(driver, selector, timeout=WAIT_SECONDS):
    by, value = selector
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def read_text_safe(el):
    try:
        return el.text.strip()
    except Exception:
        return ""


def parse_bounds(bounds_str):
    """'[x1,y1][x2,y2]' -> (x1, y1, x2, y2), or None if unparseable."""
    if not bounds_str:
        return None
    m = re.match(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", bounds_str)
    if not m:
        return None
    return tuple(int(g) for g in m.groups())


def tap_element(driver, element):
    """
    Taps by screen coordinates instead of calling .click() directly —
    some Appium-Python-Client / Selenium version combinations throw
    "Wrong parameters applied for elementClick" on UiAutomator2. This
    sidesteps that bug entirely.
    """
    parsed = parse_bounds(element.get_attribute("bounds"))
    if not parsed:
        raise RuntimeError(
            "Could not read bounds — can't compute a tap point for this element.")
    x1, y1, x2, y2 = parsed
    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
    driver.execute_script("mobile: clickGesture", {"x": cx, "y": cy})


def scroll_down(driver, percent=None):
    """
    Uses "mobile: scrollGesture" rather than "mobile: swipeGesture" —
    swipeGesture performs a FLING, which has momentum: Android keeps
    scrolling after the simulated finger lifts, and how far it travels
    depends on gesture velocity in a way that's hard to predict. That
    inconsistency was very likely why scrolling sometimes jumped past
    several customers at once. scrollGesture instead moves a fixed,
    controlled fraction of the given area with no fling.

    `percent`, when given, comes from compute_scroll_percent() — a
    measured amount based on exactly where the last already-processed
    customer's card sits on screen right now, rather than a guessed
    fixed distance. Falls back to SCROLL_STEP_PERCENT only when there's
    no measurement to base it on yet (the very first scroll).

    Note: for scrollGesture, "direction" describes which way the
    CONTENT moves (not the simulated finger) — "down" reveals further/
    later items in the list, which is what swipeGesture called "up".
    """
    if percent is None:
        percent = SCROLL_STEP_PERCENT
    size = driver.get_window_size()
    width, height = size["width"], size["height"]
    driver.execute_script("mobile: scrollGesture", {
        "left": int(width * 0.1),
        "top": int(height * SCROLL_REGION_TOP_FRACTION),
        "width": int(width * 0.8),
        "height": int(height * SCROLL_REGION_HEIGHT_FRACTION),
        "direction": "down",
        "percent": percent,
    })
    time.sleep(SCROLL_SETTLE_SECONDS)
    # DISABLED (not deleted) — this was a defensive check for a scroll
    # gesture accidentally landing on the list screen's pinned "NS
    # Number."/"Sales No." search boxes and popping the keyboard, which
    # would throw off every element's y-coordinate for the row-grouping
    # logic elsewhere in this file. Confirmed it's never actually
    # triggered in real runs, and this script currently only ever runs
    # on one known device — so the extra Appium round trip on every
    # single scroll (there are a LOT of these across a full run,
    # especially inside the CCS Note scrolling loop) isn't buying
    # anything right now. If this ever runs on a DIFFERENT phone
    # (different screen size/pinned-header height — e.g. if a
    # colleague starts using this), that's exactly the situation this
    # was guarding against, so re-enable it first: just uncomment the
    # line below.
    # dismiss_keyboard_if_present(driver)


def compute_scroll_percent(driver, target_top_y, margin_fraction=0.04):
    """
    Computes the exact scroll fraction needed to bring `target_top_y`
    (a real, measured card position) up near the top of the scroll
    region — "scroll to the line separating this card from the next,"
    rather than a blind fixed distance.

    A small margin is subtracted so the target lands just inside the
    visible region instead of exactly on the edge, avoiding the
    partially-rendered-row problem.
    """
    size = driver.get_window_size()
    height = size["height"]
    region_top = height * SCROLL_REGION_TOP_FRACTION
    region_height = height * SCROLL_REGION_HEIGHT_FRACTION

    desired_shift = (target_top_y - region_top) - \
        (margin_fraction * region_height)
    percent = desired_shift / region_height

    # Guardrails: never scroll by ~nothing (no progress) or overshoot
    # past the measured area (which would defeat the point of measuring).
    return max(0.08, min(0.95, percent))


def dismiss_keyboard_if_present(driver):
    """Defensive safety net — closes the keyboard if anything ever
    accidentally focuses a text field again, so it doesn't silently
    corrupt the next read."""
    try:
        driver.execute_script("mobile: hideKeyboard")
    except Exception:
        pass  # no keyboard was showing, or the command isn't supported — fine either way


# ============================================================
# Live pause/resume — type 'p' + Enter to pause, 'r' + Enter to resume
# ============================================================
#
# Different from the existing Ctrl+C handling: Ctrl+C stops the whole
# script, and resuming means re-launching Python, re-entering the
# identity/month prompts, and reconnecting Appium from scratch (though
# it does correctly pick up the same in-progress file — see
# _apply_identity_to_filenames()'s "RESUMING AN INTERRUPTED RUN" note).
# This is for a shorter break — pause the SAME running process (Appium
# session stays connected, nothing in memory is lost) so the phone is
# free to use, then continue right where it left off with no restart.
#
# Deliberately NOT a background thread. An earlier version of this used
# one (reading stdin in a loop via input()), and it actually worked
# correctly for pausing/resuming — but caused a real crash: a thread
# blocked waiting for keyboard input doesn't get cleaned up properly
# when the main script finishes, and printed a "Fatal Python error"
# after every run, successful or not (confirmed by testing). Polling
# for input from the MAIN thread instead — at the same safe checkpoints
# the loop already visits — avoids that entirely, since nothing is ever
# left waiting on stdin when the script exits.
_pause_requested = False
_pause_input_buffer = ""

try:
    import msvcrt  # Windows only — this project runs on Windows
    _HAS_MSVCRT = True
except ImportError:
    _HAS_MSVCRT = False
    import select  # non-Windows fallback, so this doesn't break on Mac/Linux


def _read_pending_stdin_chars():
    """Returns whatever characters are ALREADY waiting on stdin right
    now, without blocking — or "" if nothing's been typed yet. Windows
    uses msvcrt.kbhit()/getwch() (the standard non-blocking keyboard
    check on that platform); anywhere else, select() checks whether
    stdin has data ready, then os.read() pulls the raw bytes directly
    from the file descriptor — NOT sys.stdin.read(), which goes through
    Python's own internal text-buffering layer. That buffering layer
    can silently consume more bytes from the OS in one read() than it
    hands back, so a later select() check sees the file descriptor as
    empty even though a full line (like a typed 'r') is still sitting
    unread inside Python's buffer — confirmed by testing: this exact
    mismatch caused resume to never fire after a real pause. Reading
    the raw fd directly keeps what select() sees and what actually gets
    read in sync."""
    chars = ""
    if _HAS_MSVCRT:
        while msvcrt.kbhit():
            chars += msvcrt.getwch()
    else:
        while select.select([sys.stdin], [], [], 0)[0]:
            data = os.read(sys.stdin.fileno(), 4096)
            if not data:
                break
            chars += data.decode(errors="ignore")
    return chars


def _poll_pause_commands():
    """Builds up typed characters into a line buffer and checks it
    against 'p'/'r' once Enter is pressed — same idea as input(), but
    non-blocking. Called from _wait_if_paused() at safe points in the
    main loop, so a 'p' typed mid-customer doesn't take effect until
    that customer is actually finished."""
    global _pause_requested, _pause_input_buffer
    for ch in _read_pending_stdin_chars():
        if ch in ("\r", "\n"):
            cmd = _pause_input_buffer.strip().lower()
            _pause_input_buffer = ""
            if cmd == "p" and not _pause_requested:
                _pause_requested = True
                print("Pause requested — will pause after the customer "
                      "currently in progress finishes (not mid-action). "
                      "Type 'r' + Enter when you're ready to continue.")
            elif cmd == "r" and _pause_requested:
                _pause_requested = False
                print("Resuming...")
        else:
            _pause_input_buffer += ch


def _wait_if_paused():
    """Called at safe boundaries in the main loop — right after a
    customer is fully finished (back on the plain list screen, no
    popup open, nothing mid-read) or after a scroll with nothing new
    to act on. NEVER called mid-tap or mid-read, so pausing here can't
    leave the phone in a half-navigated state."""
    _poll_pause_commands()
    if not _pause_requested:
        return
    print("\nPaused. The phone won't be touched again until you resume.\n"
          "If you switch to a DIFFERENT app on the phone, make sure "
          "Cuckoo+ is back in the foreground before typing 'r' — this "
          "script taps raw screen positions, not the app specifically, "
          "so the next action needs Cuckoo+ to actually be on screen.")
    while _pause_requested:
        time.sleep(0.5)
        _poll_pause_commands()


