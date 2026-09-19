"""
contacts.py — the `--export-contacts` command: reads the current
Excel export and writes a vCard (.vcf) file so customer contacts can
be imported straight into a phone's Contacts app in one go.
"""

import os
import openpyxl

import filenames
from filenames import (
    _load_or_prompt_identity, _prompt_month_label,
    _apply_identity_to_filenames, _slugify,
)
from normalize import _normalize_sales_no, _normalize_phone_digits

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

def _vcard_escape(text):
    """
    Escapes the characters vCard 3.0 treats as structural — backslash,
    comma, semicolon, and literal newlines — so a name containing any
    of these (e.g. a comma in "SMITH, JOHN") can't corrupt the .vcf
    file's structure on import. Order matters: backslash must be
    escaped FIRST, or escaping the other characters would double-escape
    the backslashes just added for them.
    """
    if not text:
        return ""
    text = text.replace("\\", "\\\\")
    text = text.replace(",", "\\,")
    text = text.replace(";", "\\;")
    text = text.replace("\n", "\\n")
    return text


def export_contacts():
    """
    `python main.py --export-contacts` — reads the current export file
    for a specialist+month (same file --resort reads; no device/Appium
    needed) and writes a .vcf (vCard) file of every customer who has a
    usable phone number on record.

    Name format: "NDS | {contact person} ({sales_no})" —
      - the "NDS |" PREFIX tags it as a work contact and clusters all
        of these together near each other in an alphabetized contact
        list, making bulk review/cleanup easy later;
      - the contact person's name is what's actually read/searched day
        to day, so it sits front and center rather than after an ID;
      - the Sales No. in parens at the end disambiguates two customers
        who happen to share a name, and lets you cross-reference back
        to the Excel row if ever needed.

    Each contact is saved as just a name and a number — no notes, no
    extra fields — per Hazim's call to keep this minimal.

    Customers with no usable phone number (WhatsApp Number column is
    blank) are skipped entirely, since an empty TEL field would just
    create a useless, unreachable contact.

    IMPORTANT — before importing this file, DELETE any "NDS | ..."
    contacts already on the phone from a previous import first. vCard
    import is purely additive: it doesn't check for existing contacts,
    so importing the same file (or a newer month's file) on top of an
    old import without deleting first causes either duplicate contacts
    (same customer imported twice) or stale contacts left behind (a
    customer who's no longer in this month's list, but whose old
    contact never gets removed just because a new file was imported).
    Search "NDS" in Contacts (or contacts.google.com on desktop, if the
    phone syncs to Google), select all, delete — then import the fresh
    .vcf. Since every contact this script creates starts with the same
    "NDS |" prefix, that search reliably catches all of them and
    nothing else.
    """
    name, nds_id = _load_or_prompt_identity()
    month_label = _prompt_month_label()
    # Also sets filenames.CARRYFORWARD_SOURCE_FILE to the current latest version
    # file for this person+month, if any. The filenames.OUTPUT_FILE/filenames.OUTPUT_FILE_XLSM
    # values this also computes aren't used here — only the source lookup
    # matters for this command.
    _apply_identity_to_filenames(name, nds_id, month_label)

    target_file = filenames.CARRYFORWARD_SOURCE_FILE
    if not target_file:
        print(f"Couldn't find an existing export for {name} ({nds_id}) in "
              f"{month_label} — run main.py normally at least once first.")
        return

    wb = openpyxl.load_workbook(target_file, data_only=False)
    try:
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        col = {col_name: idx + 1 for idx,
               col_name in enumerate(header_row) if col_name}

        required = ["Sales No.",
                    "Installation / Service Contact Person", "WhatsApp Number"]
        missing = [n for n in required if n not in col]
        if missing:
            print(f"  !! {target_file}'s header row is missing {missing} — "
                  f"can't build contacts from it.")
            return

        vcards = []
        skipped_no_phone = 0
        for row_idx in range(2, ws.max_row + 1):
            sales_no = _normalize_sales_no(
                ws.cell(row=row_idx, column=col["Sales No."]).value)
            if not sales_no:
                continue

            contact_person = (ws.cell(
                row=row_idx,
                column=col["Installation / Service Contact Person"]).value or "").strip()
            phone_digits = _normalize_phone_digits(
                ws.cell(row=row_idx, column=col["WhatsApp Number"]).value)

            if not phone_digits:
                skipped_no_phone += 1
                continue

            display_name = (f"NDS | {contact_person} ({sales_no})"
                            if contact_person else f"NDS | {sales_no}")
            safe_name = _vcard_escape(display_name)
            vcards.append(
                "BEGIN:VCARD\r\n"
                "VERSION:3.0\r\n"
                f"FN:{safe_name}\r\n"
                # N (structured name) is REQUIRED by the vCard 3.0 spec
                # alongside FN, even though FN alone is enough to display
                # a name. Android's contacts importer is lenient and
                # accepts FN with no N at all — this file worked fine
                # there. iOS's importer is strict: when it hits a vCard
                # missing N partway through a multi-contact file, it
                # chokes and silently stops, showing only whatever it
                # managed to parse before the failure (confirmed: this
                # exact symptom — only the first contact ever showing up,
                # across WhatsApp, Mail, and Files — is a documented iOS
                # vCard failure mode, not a fluke of any one app). Since
                # the name isn't split into first/last name components
                # here, the whole display name goes into N's first
                # component (family name) and the rest are left blank —
                # that's enough to satisfy the requirement without
                # inventing a fake name split.
                f"N:{safe_name};;;;\r\n"
                f"TEL;TYPE=CELL:{phone_digits}\r\n"
                "END:VCARD\r\n"
            )
    finally:
        wb.close()

    if not vcards:
        print("No customers with a usable phone number found — nothing written.")
        return

    person_folder = f"{_slugify(nds_id)}_{_slugify(name)}"
    out_path = os.path.join(
        person_folder,
        f"{_slugify(nds_id)}_{_slugify(name)}_{_slugify(month_label)}_contacts.vcf")
    # newline="" so Python doesn't translate the \r\n already written
    # above into \r\n\r\n on Windows — vCard needs exactly \r\n per line.
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        f.write("".join(vcards))

    print(f"Wrote {out_path} — {len(vcards)} contact(s) "
          f"({skipped_no_phone} skipped, no phone number on file).")
    print("Send this file to your phone (email to self, WhatsApp to self, "
          "Google Drive, etc.) and tap it — Contacts should offer an "
          "'Import all' option.")


