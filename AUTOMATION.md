# Daily digest agent instructions

Run this every Mon–Sat at **20:00 IST** (after typical meetings).

## Goal

Produce a **DRAFT** FlightPath daily work email for human review. Do **not** finalize timesheets.

## Steps

1. Read `TIMELINE.md`, `EXCLUSIONS.md`, and today’s file under `log/` if present.
2. Infer today’s FlightPath work from:
   - Commits / changes in linked work repos when available
   - Notes in `log/YYYY-MM-DD.md`
   - Confirmed timeline phases in `TIMELINE.md`
3. **Exclude** everything listed in `EXCLUSIONS.md` (especially ProductiveIT and personal work).
4. Build a day summary with:
   - Date, weekday, status (Work / Leave / On Movement)
   - Task list with approximate times and hours
   - Roster hours vs beyond-working-hours
   - Short notes
5. Email **virrdhiman@gmail.com** via Gmail (Composio):
   - Subject: `[DRAFT FOR REVIEW] FlightPath daily work — YYYY-MM-DD`
   - Body: markdown-friendly plain text or simple HTML
   - First line must say: `DRAFT — please review before finalizing`
6. Optionally write/update `log/YYYY-MM-DD.md` with the same draft content and status `draft`.
7. Stop. Do not send a “final” version and do not invent personal/ProductiveIT work.

## Email rules

- Recipient: `virrdhiman@gmail.com` only (unless instructions change).
- Always label as DRAFT for review.
- If Gmail fails, leave the draft in `log/` and report the error in the run summary.
