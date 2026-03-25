# HAC Event Calendars

Automatic ICS calendar generation from Microsoft Lists via Power Automate.

## How it works

```
SharePoint Lists (PE Events / Restructuring Events)
  → Power Automate (detects changes every 5 min)
  → GitHub (data/*.json updated via API)
  → GitHub Action (generates ICS from JSON)
  → GitHub Pages (hosts ICS files)
  → Outlook / Apple Calendar (auto-sync via webcal://)
```

**End-to-end latency:** ~6–11 minutes

## Data sources

| Calendar | SharePoint List | JSON file |
|----------|----------------|-----------|
| PE & Deal Sourcing Events 2026 | PE Events 2026 | `data/pe_events_2026.json` |
| Restructuring Events 2026 | Restructuring Events 2026 | `data/restructuring_events_2026.json` |

## Subscribe to calendars

Calendar subscription URLs are shared internally.

## Update events

Edit events directly in the SharePoint Lists. Changes are automatically synced to the ICS files within ~10 minutes.

## Architecture

- **Input:** Microsoft Lists on SharePoint
- **Sync:** Power Automate (2 flows, one per calendar)
- **Storage:** `data/*.json` in this repo
- **Conversion:** `scripts/excel_to_ics.py` (supports both JSON and Excel sources)
- **Hosting:** GitHub Pages (`docs/**/*.ics`)
- **Config:** `calendars.yaml`

## Setup

See `POWER_AUTOMATE_SETUP.md` in the deploy folder of the main project for detailed setup instructions.