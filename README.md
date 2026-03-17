# HAC Event Calendars

Automatic ICS calendar generation from Excel event lists.

## How it works

1. Excel files in `data/` contain event calendars
2. `scripts/excel_to_ics.py` converts them to ICS format
3. GitHub Action runs automatically on every Excel update
4. Generated ICS files are hosted via GitHub Pages

## Subscribe to calendars

Calendar subscription URLs are shared internally.

## Update events

Edit the Excel file in SharePoint (or push directly to `data/`). The GitHub Action will regenerate the ICS files automatically.
