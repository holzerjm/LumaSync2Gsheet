# LumaSync2Gsheet

A Google Apps Script that syncs a [Luma](https://lu.ma) event's guest list into a
Google Sheet via the Luma public API, and stamps a "last updated" timestamp on a
dashboard sheet. Designed to run on a time-driven trigger for hands-off, scheduled
refreshes — no server, no CSV middle-step.

## Features

- Pulls **every approval status** you list (`approved`, `pending_approval`,
  `waitlist`, `declined`, …). Luma's `get-guests` endpoint filters by status, so
  omitting one silently drops those guests — this script requests each explicitly.
- **Preserves your column order**: it adopts whatever header row already exists in
  the destination sheet, and matches custom registration questions by their exact
  label. Reorder columns by rearranging row 1 — no code change needed.
- **Normalises Luma's answer shapes** per question type: rebuilds full LinkedIn /
  GitHub URLs from the stored paths, joins multi-select arrays, splits the bundled
  company + job-title answer, and renders the terms checkbox as `Agreed`.
- Handles **cursor pagination** so the full guest list is fetched, not just page 1.
- Writes a **last-updated timestamp** to a configurable cell on a dashboard sheet.
- **Posts a Slack update** when new registrations appear — the new + total counts
  plus the event name and date — via an Incoming Webhook (only fires when there's
  actually something new).
- **No secrets in source** — the API key, event ID, and Slack webhook live in
  Script Properties.

## Prerequisites

- A **Luma Plus** subscription on the calendar (the API is gated behind it).
- A Google account / Google Sheet.

## Setup

1. Open your Google Sheet → **Extensions ▸ Apps Script**.
2. Copy [`LumaSyncGuest2Sheet.gs`](LumaSyncGuest2Sheet.gs) into the script editor.
3. **Project Settings ▸ Script Properties** — add:
   | Property | Value |
   |---|---|
   | `LUMA_API_KEY` | Your Luma API key (Calendar → Settings → Developer → API Keys) |
   | `LUMA_EVENT_ID` | Your event ID, e.g. `evt-XXXXXXXX` |
   | `SLACK_WEBHOOK_URL` | *(optional)* Slack Incoming Webhook URL for new-registration alerts |
4. Edit the **Config** constants at the top of the script to match your sheet:
   - `SHEET_NAME` — the sheet that receives the guest rows
   - `STAMP_SHEET` / `STAMP_CELL` — where the "last updated" timestamp is written
   - `COMPANY_QUESTION_LABEL` / `JOB_TITLE_COLUMN` — your event's exact labels for
     the company/university question and the job-title column
   - `STATUSES` — the approval statuses to include
5. Run `syncLumaGuests` once to authorise the script.
6. Add a schedule: **Triggers ▸ Add Trigger ▸ `syncLumaGuests` ▸ Time-driven**
   (e.g. hourly or daily).

## Slack notifications (optional)

When `SLACK_WEBHOOK_URL` is set, every sync that finds **new** registrations posts a
message to that channel, e.g.:

> 🎉 **3 new registrations** for **Agent Day**
> 📅 Sat, Jun 27, 2026 at 9:00 AM EDT
> 👥 **128** total registered
> View on Luma

To set it up:

1. In Slack, create an **Incoming Webhook** for the target channel
   (<https://api.slack.com/messaging/webhooks>) and copy the webhook URL.
2. Add it as the `SLACK_WEBHOOK_URL` Script Property.
3. In the Apps Script editor, select **`testSlackMessage`** from the function
   dropdown and click **Run**. It posts a clearly-labelled test message (with
   sample counts) so you can confirm the webhook works before relying on the
   scheduled sync — it does **not** modify the spreadsheet. The execution log
   reports `HTTP 200` on success.

Notes:

- The message only posts when the new count is greater than zero, so empty syncs
  stay silent.
- "New" is determined by comparing the `guest_id`s already in the sheet against the
  freshly fetched list — so the **first** populated run reports everyone as new. Run
  the sync once before adding the webhook if you want to skip that initial post.
- The event name and date come from Luma's `event/get` endpoint, rendered in the
  event's own timezone.

## Column ordering

On each run the script reads row 1 of the destination sheet and writes data in that
exact order. To control the layout, set up your header row once — either by pasting
Luma's export header, or by doing a one-time **File ▸ Import** of a Luma CSV. If the
sheet has no header yet, it falls back to the built-in `COLUMNS` list.

## Fields not returned by this endpoint

Luma's `get-guests` returns guest summaries and `event_tickets`, but **not**
order-level details. These columns will be blank even though they appear in the
manual CSV export: `coupon_code`, `survey_response_rating`,
`survey_response_feedback`. Host-added fields/tags (not registration answers) are
likewise not exposed here.

> Amounts are returned in **cents** and formatted as `$0.00`. If you run a paid
> event and the figures look off by 100×, remove the `/ 100` in `money()`.

## Troubleshooting

- **Only the header row appears / no data** — usually means a status filter excluded
  your guests. Confirm `STATUSES` includes the states your registrants are in.
- **A column is blank** — run `debugLumaResponse` once and check the execution log;
  it prints the raw JSON for one guest so you can confirm the exact field/label name.

## Files

- [`LumaSyncGuest2Sheet.gs`](LumaSyncGuest2Sheet.gs) — the Apps Script.
- `README.md` — this file.
- [`LICENSE`](LICENSE) — MIT license.

## License

Released under the [MIT License](LICENSE).
