# Passport Check

A small desktop tool that replaces a manual, error-prone paperwork step: matching a stack of passport scans to a name list and stamping each photo with the right ID number.

Built for a real workflow — a non-technical user drags files onto two boxes and clicks one button, no command line.

## What it does

1. **PDF → photos** — splits multi-page PDF passport scans into individual images.
2. **Match + stamp** — reads the Excel name list, matches each photo to the right row (by name or passport number), and writes the ID number onto the photo in the corner.
3. Flags anything it's unsure about — unmatched photos, duplicate names, broken files — instead of silently guessing.

## Why

Doing this by hand in Excel/Finder took a long time and was easy to get wrong (wrong number on wrong photo). This automates the matching and leaves a clear report of anything it couldn't resolve, so a human only reviews the exceptions.

## Stack

Python, Tkinter (desktop UI, no extra install needed), optional drag-and-drop via `tkinterdnd2`. Bilingual UI (English/French).

## Run it

```bash
cd "check passport"
pip install openpyxl pillow pypdf tkinterdnd2
python app.py
```
Or double-click `Passport Check.bat` on Windows.

## Tests

```bash
cd "check passport"
python -m pytest tests
```

## Notes

Real client data (photos, name lists, output) is never committed — see `.gitignore`. Only the app code and tests are in this repo. See `check passport/CHANGELOG.md` for version history.
