# FlightPath Daily Work Tracker

Source of truth for Virender Dhiman’s **FlightPath Infogain** daily work digests (draft → human review → finalize).

## Purpose

A Cursor Automation checks out this repo on a schedule (Mon–Sat **20:00 IST**), builds a **DRAFT** day summary from FlightPath work only, and emails it to **virrdhiman@gmail.com** for review. Nothing is treated as final until you approve it.

## Scope

**Include (work)**

- Flipkart Automation Suite
- HDFC / card scripts
- Flipkart gift cards
- CRM, sales sheets & dashboards
- Static websites delivered for FlightPath
- Team support, hiring, HR-assigned tasks
- VPS / platform ops for FlightPath
- Internal tools (e.g. ID card generator)

**Exclude**

- `ProductiveIT-Books` / ProductiveIT folder
- Personal projects (`personal/`, portfolio, private side work)
- Sundays (non-working)
- Leave / on-movement days when marked

## Roster model

- Work: **8 hours** (09:00–17:30 with lunch 13:00–13:30 = 30 min)
- Beyond hours: mark after 17:30 (meetings, delivery crunch)

## Files

| File | Role |
|------|------|
| [AUTOMATION.md](./AUTOMATION.md) | Prompt the scheduled agent must follow |
| [TIMELINE.md](./TIMELINE.md) | Confirmed work phases Apr–Jul 2026 |
| [log/](./log/) | Daily draft / reviewed notes (markdown) |
| [EXCLUSIONS.md](./EXCLUSIONS.md) | Folders and topics never to include |

## Attendance markers (confirmed)

- **Leave:** 14 May 2026, 20 Jun 2026, 8 Jul 2026
- **On movement:** 9–10 Jul 2026
- **Today example:** 20 Jul 2026 — Sector 8 meeting with Tarun Sir ~18:00

## Local canvas

Interactive dashboard (Cursor Canvas): open in Cursor IDE under the Programs workspace canvases folder — `flightpath-daily-work-log.canvas.tsx`.
