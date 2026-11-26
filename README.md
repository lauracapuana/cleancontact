# Cleancontact

A Python automation script to merge, clean, and normalize contact lists from multiple **Excel (`.xlsx`)** and **CSV (`.csv`)** files, exporting a sorted Excel spreadsheet, a **vCard (`.vcf`)** address book file, and a detailed cleanup report.

## Features
- **Multi-file selection:** Opens a system file dialog to select one or more `.xlsx` or `.csv` files in a single run.
- **Data normalization:**
  - Combines *First Name* and *Last Name* (or *Nome* and *Cognome*) and standardizes spacing and capitalization.
  - Formats 10-digit US phone numbers as `(XXX) XXX-XXXX` (and 11-digit numbers starting with `1` as `+1 (XXX) XXX-XXXX`).
  - Normalizes email addresses to lowercase without leading/trailing whitespace.
- **Automated filtering:**
  - Skips rows without a phone number.
  - Excludes fake/short numbers or numbers made of a single repeated digit (configurable).
  - Excludes test contacts (configurable).
- **Duplicate handling & merging:**
  - Removes duplicate entries sharing the same phone number (comparing digits only).
  - Merges additional phone numbers when the same person is matched by email or name.
  - Flags cases where the same phone number appears under different names in the report.
- **Multiple output formats:**
  - `contatti_uniti_puliti.xlsx` (alphabetically sorted by name).
  - `contatti_uniti_puliti.vcf` (vCard 3.0 ready for contact import).
  - `report_pulizia.txt` (summary counts, excluded entries, merged numbers, name/phone flags, and removed duplicates).

## Requirements & Setup
- Python 3.10+
- `openpyxl` for reading and writing `.xlsx` files:
  ```bash
  py -m pip install openpyxl