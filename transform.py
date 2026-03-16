#!/usr/bin/env python3
"""
Transform Team Distribution Plan calendar CSV into:
1. Time entries.csv (budget/project work)
2. Absence bookings.csv (vacations, sick leave, PDD)

Only processes 2026 data.
"""

import csv
import unicodedata
import re
from datetime import datetime, timedelta
from collections import defaultdict

CALENDAR_FILE = "Team Distribution Plan 2025 _ 2026 [All Team] - Calendar.csv"
EMAILS_FILE = "Import from TDP -_ Productive - Names x emails.csv"
TIME_ENTRIES_OUT = "Time entries - 2026.csv"
ABSENCE_BOOKINGS_OUT = "Absence bookings - 2026.csv"

# Symbols that are absences
ABSENCE_MAP = {
    "\U0001f334": "Vacation",      # 🌴
    "\U0001f637": "Sick leave",    # 😷
    "\U0001f469\u200d\U0001f393": "PDD",  # 👩‍🎓
}

# Symbols to skip entirely
SKIP_SYMBOLS = {"-", "--", ""}

# Section headers (not real people)
SECTION_HEADERS = {
    "Developers App Team", "Designers", "Delivery Managers", "Platform Unit",
    "EraSciences (Subcontractors)", "Templates", "[Template] Polish Holidays",
    "[Template] Flex Holidays", "Machine Learning Unit", "Others", "Devopsity",
}

# People to explicitly skip
SKIP_PEOPLE = {"Jędrzej Świeżewski"}

# People without emails who have 2026 data — flag with name
FLAGGED_PEOPLE = {
    "Remy Gavard", "Patryk Jedlikowski  (STX) [Arcutis]",
    "Alexander Lubeck (STX) [Arcutis]", "Jan Zoń (STX) [Arcutis]",
    "Erick Facure Giaretta", "Agustin Perez Santangelo [Jazz]",
    "Angel Escalante/Johan Rosa [Jazz]",
}


def normalize(s):
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"\[[^\]]*\]", "", s)
    return s.strip()


def load_email_map():
    email_map = {}
    with open(EMAILS_FILE, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            full_name = row[0].strip()
            email = row[4].strip()
            if email and email != "#N/A":
                email_map[normalize(full_name)] = email
    return email_map


def find_email(cal_name, email_map):
    cn = normalize(cal_name)
    # Exact match
    if cn in email_map:
        return email_map[cn]
    # First-name match
    cn_parts = cn.split()
    for en_norm, email in email_map.items():
        en_parts = en_norm.split()
        if len(cn_parts) >= 2 and len(en_parts) >= 2:
            if cn_parts[0] == en_parts[0]:
                return email
    return None


def is_single_letter_or_apostrophe(val):
    """Check if a value is a single letter (part of a multi-cell word)."""
    if len(val) == 1 and val.isalpha():
        return True
    if len(val) == 2 and val[1] == "'" and val[0].isalpha():
        return True
    return False


def reconstruct_words(row, start, end):
    """Find sequences of single-letter cells and merge them into words.
    Only merges strictly consecutive single-letter cells (no skipping weekends).
    Each Mon-Fri week spells the word independently.
    Returns a dict mapping column index -> reconstructed word for each cell in that word.
    """
    col_to_word = {}
    j = start
    while j <= end:
        if j < len(row):
            val = row[j].strip()
            if is_single_letter_or_apostrophe(val):
                # Start collecting a word — strictly consecutive cells only
                word_cols = [j]
                word_chars = [val.replace("'", "")]
                k = j + 1
                while k <= end and k < len(row):
                    next_val = row[k].strip()
                    if is_single_letter_or_apostrophe(next_val):
                        word_cols.append(k)
                        word_chars.append(next_val.replace("'", ""))
                        k += 1
                    else:
                        break
                word = "".join(word_chars).upper()
                if len(word) > 1:
                    for c in word_cols:
                        col_to_word[c] = word
                j = k
                continue
        j += 1
    return col_to_word


def is_absence(val):
    return val in ABSENCE_MAP


# Partial words that should be normalized to full budget names
WORD_NORMALIZATION = {
    "ATMO": "ATMOS",
    "ATMOS": "ATMOS",
    "NOVO": "NOVO",
    "ZURI": "ZURICH",
    "ZURIC": "ZURICH",
    "ZURICH": "ZURICH",
    "RI": "ZURICH",  # partial ZURICH from split weeks
}


def get_budget_name(val, col, col_to_word):
    """Get the budget name for a cell value."""
    if col in col_to_word:
        word = col_to_word[col]
        return WORD_NORMALIZATION.get(word, word)
    # Single apostrophe-letter that wasn't grouped into a word (e.g. lone Z')
    if is_single_letter_or_apostrophe(val):
        letter = val.replace("'", "").upper()
        # Check if this person has ZURICH or ATMOS patterns nearby
        # For safety, just map known single starts
        if letter == "Z":
            return "ZURICH"
        if letter == "A":
            return "ATMOS"
        if letter == "N":
            return "NOVO"
    return val


def is_tentative(val, col, col_to_word):
    """Check if a booking is tentative."""
    if col in col_to_word:
        return True
    if val == "?":
        return True
    return False


def main():
    email_map = load_email_map()

    with open(CALENDAR_FILE, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)

    dates_row = rows[1]

    # Find 2026 column range
    start_2026 = end_2026 = None
    date_by_col = {}
    for j in range(3, len(dates_row)):
        d = dates_row[j].strip()
        if d and "2026" in d:
            if start_2026 is None:
                start_2026 = j
            end_2026 = j
            # Parse date: M/D/YYYY -> YYYY-MM-DD
            try:
                parsed = datetime.strptime(d, "%m/%d/%Y")
                date_by_col[j] = parsed.strftime("%Y-%m-%d")
            except ValueError:
                pass

    print(f"2026 range: cols {start_2026}-{end_2026}")
    print(f"Date range: {date_by_col.get(start_2026)} to {date_by_col.get(end_2026)}")

    time_entries = []
    absence_entries = []  # Will collect (person, category, date) then merge ranges

    people_without_email = []

    for i, row in enumerate(rows):
        if i < 3:
            continue

        cal_name = row[0].strip()
        if not cal_name or cal_name in SECTION_HEADERS or cal_name in SKIP_PEOPLE:
            continue

        # Check if this person has any 2026 data
        has_data = False
        for j in range(start_2026, end_2026 + 1):
            if j < len(row):
                val = row[j].strip()
                if val and val not in SKIP_SYMBOLS:
                    has_data = True
                    break
        if not has_data:
            continue

        # Resolve person identifier
        email = find_email(cal_name, email_map)
        if email:
            person_id = email
        elif cal_name.strip() in FLAGGED_PEOPLE or any(
            cal_name.strip().startswith(fp.split()[0]) for fp in FLAGGED_PEOPLE
        ):
            person_id = f"[NEEDS EMAIL] {cal_name}"
        else:
            # No email, no 2026 data worth processing or unknown person
            continue

        # Reconstruct multi-cell words
        col_to_word = reconstruct_words(row, start_2026, end_2026)

        # Process each day
        person_absences = []  # (date_str, category)

        for j in range(start_2026, end_2026 + 1):
            if j not in date_by_col:
                continue
            if j >= len(row):
                continue

            val = row[j].strip()
            date_str = date_by_col[j]

            # Skip weekends, non-working days, empty cells
            if val in SKIP_SYMBOLS:
                continue

            # Check if absence
            if is_absence(val):
                category = ABSENCE_MAP[val]
                person_absences.append((date_str, category))
                continue

            # It's a budget entry
            budget = get_budget_name(val, j, col_to_word)
            status = "tentative" if is_tentative(val, j, col_to_word) else "confirmed"

            time_entries.append({
                "Project": "",
                "Budget": budget,
                "Deal": "",
                "Client": "",
                "Service Type": "",
                "Service": "",
                "Person": person_id,
                "Date": date_str,
                "Note": status,
                "Time (minutes)": 480,
            })

        # Merge consecutive absence days into ranges
        if person_absences:
            # Sort by date
            person_absences.sort(key=lambda x: x[0])

            range_start = person_absences[0][0]
            range_end = person_absences[0][0]
            range_cat = person_absences[0][1]

            for k in range(1, len(person_absences)):
                date_str, category = person_absences[k]
                prev_date = datetime.strptime(range_end, "%Y-%m-%d")
                curr_date = datetime.strptime(date_str, "%Y-%m-%d")
                # Allow gaps of up to 3 days (weekends) for merging
                if category == range_cat and (curr_date - prev_date).days <= 3:
                    range_end = date_str
                else:
                    absence_entries.append({
                        "Person": person_id,
                        "Absence Category": range_cat,
                        "Started On": range_start,
                        "Ended On": range_end,
                    })
                    range_start = date_str
                    range_end = date_str
                    range_cat = category

            # Don't forget the last range
            absence_entries.append({
                "Person": person_id,
                "Absence Category": range_cat,
                "Started On": range_start,
                "Ended On": range_end,
            })

    # Write Time entries CSV
    time_fields = [
        "Project", "Budget", "Deal", "Client", "Service Type",
        "Service", "Person", "Date", "Note", "Time (minutes)",
    ]
    with open(TIME_ENTRIES_OUT, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=time_fields)
        writer.writeheader()
        writer.writerows(time_entries)

    # Write Absence bookings CSV
    absence_fields = ["Person", "Absence Category", "Started On", "Ended On"]
    with open(ABSENCE_BOOKINGS_OUT, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=absence_fields)
        writer.writeheader()
        writer.writerows(absence_entries)

    # Stats
    print(f"\nTime entries: {len(time_entries)} rows written to {TIME_ENTRIES_OUT}")
    print(f"Absence bookings: {len(absence_entries)} rows written to {ABSENCE_BOOKINGS_OUT}")

    # Show flagged people
    flagged = set(e["Person"] for e in time_entries if e["Person"].startswith("[NEEDS EMAIL]"))
    flagged |= set(e["Person"] for e in absence_entries if e["Person"].startswith("[NEEDS EMAIL]"))
    if flagged:
        print(f"\nFlagged people (need email):")
        for p in sorted(flagged):
            print(f"  {p}")

    # Show unique budgets
    budgets = sorted(set(e["Budget"] for e in time_entries))
    print(f"\nUnique budgets ({len(budgets)}):")
    for b in budgets:
        print(f"  {b}")

    # Show tentative entries count
    tentative = sum(1 for e in time_entries if e["Note"] == "tentative")
    print(f"\nTentative entries: {tentative}")
    print(f"Confirmed entries: {len(time_entries) - tentative}")


if __name__ == "__main__":
    main()
