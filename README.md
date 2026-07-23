# FlightPath Daily Work Tracker

Source of truth for Virender Dhiman’s **FlightPath Infogain** daily work digests (draft → human review → finalize).

## Purpose

A Cursor Automation checks out this repo on a schedule (Mon–Sat **20:00 IST**), builds a **DRAFT** day summary from FlightPath work only, and emails it **from `98765dummy123@gmail.com` to `virrdhiman@gmail.com`** for review. Nothing is treated as final until you approve it.

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

- Work: **10:00–18:00** with **lunch 14:00–14:30** (30 min)
- Beyond hours: after 18:00 when needed

## Files

| File | Role |
|------|------|
| [AUTOMATION.md](./AUTOMATION.md) | Prompt the scheduled agent must follow |
| [TIMELINE.md](./TIMELINE.md) | Confirmed work phases Apr–Jul 2026 |
| [EXCLUSIONS.md](./EXCLUSIONS.md) | Folders and topics never to include |
| [export_excel.py](./export_excel.py) | **Excel export** with dashboard + date/month filters |
| [EXPORT-EXCEL.bat](./EXPORT-EXCEL.bat) | Windows one-click export menu |
| [log/](./log/) | Daily draft / reviewed notes (markdown) |
| [exports/](./exports/) | Generated `.xlsx` files (gitignored binaries optional) |

## Excel export (dashboard + date / month)

```bash
pip install -r requirements.txt
python export_excel.py                      # full period
python export_excel.py --month 2026-07      # one month
python export_excel.py --date 2026-07-20    # one day
python export_excel.py --from 2026-05-24 --to 2026-06-14
```

Or double-click `EXPORT-EXCEL.bat`.

Workbook sheets:

1. **00 How to use** — filter guide  
2. **01 Dashboard** — KPIs, month bar chart, category pie, phase table  
3. **02 Daily Log** — every task (newest date first; AutoFilter by Date / Month / Status / Category)  
4. **03 Monthly Summary** — month-wise totals  
5. **04 Day Rollup** — one row per day, **latest first** (best for date-wise review)  

## Open on this PC (local)

### Excel workbook
1. Open folder: `c:\Users\virrd\OneDrive\Desktop\Programs\flightpath-daily-work-tracker\exports\`
2. Double-click `FlightPath_Daily_Work_2026-04-08_to_2026-07-20.xlsx` (or a month/date export)
3. Or regenerate: double-click `EXPORT-EXCEL.bat` in the tracker folder

### Tracker project (Git)
```powershell
cd c:\Users\virrd\OneDrive\Desktop\Programs\flightpath-daily-work-tracker
explorer .
```
Or clone: `git clone https://github.com/virrdhiman/flightpath-daily-work-tracker.git`

### Interactive dashboard (Cursor Canvas)
1. Open Cursor on the **Programs** workspace  
2. Open file:  
   `C:\Users\virrd\.cursor\projects\c-Users-virrd-OneDrive-Desktop-Programs\canvases\flightpath-daily-work-log.canvas.tsx`  
3. Use **Open Canvas** / preview beside chat (latest dates show first)


## Attendance markers (confirmed)

- **Leave:** 14 May 2026, 20 Jun 2026, 8 Jul 2026
- **On movement:** 9–10 Jul 2026
- **Today example:** 20 Jul 2026 — Sector 8 meeting with Tarun Sir ~18:00

## Local canvas

Interactive dashboard (Cursor Canvas): open in Cursor IDE under the Programs workspace canvases folder — `flightpath-daily-work-log.canvas.tsx`.
