"""
filenames.py — specialist identity prompts, and the per-person/per-
month export filenames.

IMPORTANT for anyone editing other modules: OUTPUT_FILE, OUTPUT_FILE_XLSM,
and CARRYFORWARD_SOURCE_FILE below are REASSIGNED while the script runs
(by _apply_identity_to_filenames(), called once per run/resort/contacts-
export). Any OTHER module that needs the current value must do
`import filenames` and read `filenames.OUTPUT_FILE_XLSM` etc. at the
point of use — NOT `from filenames import OUTPUT_FILE_XLSM`, which
would freeze in whatever value existed at import time and silently go
stale the moment this module reassigns it. This is the one real
gotcha in this whole module split; every other module in this project
was written with that in mind.
"""

import os
import re

from config import IDENTITY_CONFIG_FILE

OUTPUT_FILE = "cuckoo_export.xlsx"
OUTPUT_FILE_XLSM = "cuckoo_export.xlsm"

# ============================================================
# Specialist identity + per-month export filenames
# ============================================================
#
# This is what makes the script safe to hand to colleagues: the
# WhatsApp message text used to hardcode "Hanis"/"NDS35095", and the
# export always landed in one shared "cuckoo_export.xlsm" regardless
# of who ran it or when. Now each person's Name + NDS Cuckoo ID is
# asked for ONCE (cached locally — see IDENTITY_CONFIG_FILE) and used
# both to personalize the message and to build a per-person, per-month
# filename, so nobody's export or message text collides with anyone
# else's.


def _load_or_prompt_identity():
    """
    Asks which specialist's list this run is for — EVERY run, not just
    the first — since this script might be run for a different person
    entirely (e.g. covering a colleague's list), and silently reusing
    whoever answered last would risk exporting/messaging under the
    wrong name.

    The last values used are read from IDENTITY_CONFIG_FILE and shown
    as defaults (just press Enter to keep them), purely so running it
    again and again for yourself doesn't mean retyping your own name
    and ID every time. Whatever's answered — kept default or a new
    value typed in — is saved back to that file so it becomes next
    run's default.
    """
    import json

    last_name, last_nds_id = "", ""
    if os.path.exists(IDENTITY_CONFIG_FILE):
        try:
            with open(IDENTITY_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            last_name = (data.get("name") or "").strip()
            last_nds_id = (data.get("nds_id") or "").strip()
        except Exception:
            pass  # missing/corrupted — just means no defaults to offer

    name_prompt = (f"Specialist's name [{last_name}]: " if last_name
                   else "Specialist's name (as it should appear in the WhatsApp message): ")
    name = input(name_prompt).strip() or last_name

    nds_prompt = (f"Specialist's NDS Cuckoo ID [{last_nds_id}]: " if last_nds_id
                  else "Specialist's NDS Cuckoo ID (e.g. NDS35095): ")
    nds_id = input(nds_prompt).strip() or last_nds_id

    with open(IDENTITY_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"name": name, "nds_id": nds_id}, f)
    return name, nds_id


def _slugify(text):
    """
    Turns free text into something safe to use inside a filename —
    keeps letters/digits, collapses everything else (spaces,
    punctuation) into a single underscore, trims leading/trailing
    underscores. Used for both the specialist's name and NDS ID when
    building the export filename, since either could contain a space
    or punctuation that Windows filenames don't like.
    """
    text = re.sub(r"[^A-Za-z0-9]+", "_", (text or "").strip())
    return text.strip("_") or "Unknown"


def _default_month_label():
    """e.g. 'September' — today's real calendar month, used only as the
    suggested default in _prompt_month_label(); the person can type a
    different month if this export is actually for a different one."""
    import datetime
    return datetime.date.today().strftime("%B")


def _prompt_month_label():
    """
    Asks which month this export is for, defaulting to the current
    calendar month if the person just presses Enter. Asked every run
    (not cached like the identity) since which month you're working on
    legitimately changes far more often than your name or NDS ID.
    """
    default = _default_month_label()
    entered = input(f"Which month is this export for? [{default}]: ").strip()
    return entered or default


# Which existing file (if any) this run should read carry-forward data
# from — set by _apply_identity_to_filenames() before anything else
# runs. This is DIFFERENT from OUTPUT_FILE/OUTPUT_FILE_XLSM below: the
# version number increases by one on every run, so the file this run
# WRITES to is never the same file it READS from.
CARRYFORWARD_SOURCE_FILE = None


def _apply_identity_to_filenames(name, nds_id, month_label):
    """
    Rebuilds the module-level OUTPUT_FILE / OUTPUT_FILE_XLSM /
    CARRYFORWARD_SOURCE_FILE paths from the specialist's identity and
    the month they typed in.

    Folder layout: one shared folder per person, "<NDSID>_<Name>/",
    holding every version for every month as flat files named
    "<NDSID>_<Name>_<Month>_v<N>_export.xlsm" — e.g.
    "NDS35095_Hanis/NDS35095_Hanis_September_v1_export.xlsm". Nothing
    ever gets deleted; a rerun in the same month bumps N and adds a
    new file alongside the old ones.

    RESUMING AN INTERRUPTED RUN: while a run is in progress (or if it
    got cut short — crash, Ctrl+C, LIST_LIMIT, or the scroll safety
    cap), its file is named with an extra "_incomplete" tag, e.g.
    "..._v3_export_incomplete.xlsm". If the highest version found here
    for this person+month is still tagged "_incomplete", THIS run
    reuses that EXACT same filename as both its carry-forward source
    and its write target — no new version number — so a rerun after a
    crash keeps filling in that same file instead of abandoning it and
    starting a fresh version. Only once a run reaches the genuine end
    of the app's list does _finalize_output_file() (called from run())
    drop the "_incomplete" tag, which is what makes the NEXT deliberate
    run pick a brand new version number instead of resuming.

    A month with no matching files yet naturally starts fresh at v1
    with no carry-forward source — the intended behaviour, since a
    visit plan (Proposed Dates especially) is specific to its month.
    """
    global OUTPUT_FILE, OUTPUT_FILE_XLSM, CARRYFORWARD_SOURCE_FILE

    person_folder = f"{_slugify(nds_id)}_{_slugify(name)}"
    os.makedirs(person_folder, exist_ok=True)

    base_no_version = f"{_slugify(nds_id)}_{_slugify(name)}_{_slugify(month_label)}"
    complete_pattern = re.compile(
        r"^" + re.escape(base_no_version) + r"_v(\d+)_export\.(xlsm|xlsx)$")
    incomplete_pattern = re.compile(
        r"^" + re.escape(base_no_version) + r"_v(\d+)_export_incomplete\.(xlsm|xlsx)$")

    best_version = 0
    best_file = None
    best_is_xlsm = False
    best_is_incomplete = False
    for fname in os.listdir(person_folder):
        m = incomplete_pattern.match(fname)
        is_incomplete = m is not None
        if not m:
            m = complete_pattern.match(fname)
        if not m:
            continue
        version = int(m.group(1))
        is_xlsm = m.group(2) == "xlsm"
        # Highest version wins outright. On a tie: an incomplete file
        # beats a complete one (it's the true latest in-progress
        # state), and among those, .xlsm beats .xlsx.
        better = (
            version > best_version
            or (version == best_version and is_incomplete and not best_is_incomplete)
            or (version == best_version and is_incomplete == best_is_incomplete
                and is_xlsm and not best_is_xlsm)
        )
        if better:
            best_version = version
            best_file = fname
            best_is_xlsm = is_xlsm
            best_is_incomplete = is_incomplete

    if best_file and best_is_incomplete:
        # Resume this exact file in place — same version, same name.
        source_path = os.path.join(person_folder, best_file)
        CARRYFORWARD_SOURCE_FILE = source_path
        if best_is_xlsm:
            OUTPUT_FILE_XLSM = source_path
            OUTPUT_FILE = source_path[:-len(".xlsm")] + ".xlsx"
        else:
            OUTPUT_FILE = source_path
            OUTPUT_FILE_XLSM = source_path[:-len(".xlsx")] + ".xlsm"
        return

    CARRYFORWARD_SOURCE_FILE = os.path.join(
        person_folder, best_file) if best_file else None

    next_version = best_version + 1
    versioned_base = f"{base_no_version}_v{next_version}_export_incomplete"
    OUTPUT_FILE = os.path.join(person_folder, f"{versioned_base}.xlsx")
    OUTPUT_FILE_XLSM = os.path.join(person_folder, f"{versioned_base}.xlsm")


def _finalize_output_file(written_to, reached_natural_end):
    """
    Called once per run, right after write_output() has saved. If the
    run genuinely reached the end of the app's list (reached_natural_end
    — see run()), this drops the "_incomplete" tag from the file that
    was just written, marking it done: the NEXT run will then pick a
    fresh version number instead of resuming it (see
    _apply_identity_to_filenames()).

    If the run was cut short for any reason — crash, Ctrl+C,
    LIST_LIMIT, or the scroll safety cap — this leaves the file exactly
    as it is (still "_incomplete"), so next time you run the script it
    picks this same file back up and keeps filling it in, rather than
    treating it as done and starting a new version.

    Returns the file's final path (renamed or not) so callers can print
    the name the person should actually go look for.
    """
    if not written_to or not os.path.exists(written_to):
        return written_to

    if not reached_natural_end:
        print(f"Didn't reach the end of the app's list this run — "
              f"{written_to} stays marked in-progress. Rerunning will "
              f"continue filling in this same file rather than starting "
              f"a new version.")
        return written_to

    if "_export_incomplete." not in written_to:
        return written_to  # already finalized somehow — nothing to do

    finalized = written_to.replace("_export_incomplete.", "_export.")
    try:
        if os.path.exists(finalized):
            os.remove(finalized)
        os.rename(written_to, finalized)
        print(f"Reached the end of the app's list — finalized as {finalized}.")
        return finalized
    except OSError as e:
        print(f"  !! Could not finalize {written_to} to {finalized} ({e}) — "
              f"it'll still work fine, just keeps the '_incomplete' name.")
        return written_to


