# VTC Scorecard Updater

This is a local-first starter app for updating the Vendor & Trade Council scorecard from dropped Excel files and field feedback emails.

## How to Use

1. Open `index.html` in a browser.
2. Load the master `VTC Scorecard 4.21.xlsx` workbook.
3. Drop source Excel or CSV files that contain tabs or headers for:
   - `Safety`
   - `Schedule_Adherence_Raw`
   - `Rework`
   - `Log`
   Schedule updates can be dropped as two separate workbooks at the same time. The app recognizes schedule files by headers such as `Vendor_ID`, `Vendor_Name`, `Monthly_Tasks`, and `No_Show_Count`, even if the worksheet tab names differ.
4. Review the recalculated scores.
5. Paste forwarded builder feedback into the Field feedback tab, review the parsed entry, and add it to the Log.
6. Use Export scorecard to open a printable dashboard report. Choose Save as PDF in the browser print dialog.

## Saved Data

The app saves parsed scorecard data, weights, feedback entries, and activity in the browser's local storage. Refreshing the page on the same browser will restore the latest loaded data instead of starting over.

For team-wide shared data, connect the app to Supabase using the Cloud Data panel. Browser storage remains a fallback for refreshes and offline use, but Supabase is the shared source for users on different devices.

## Supabase Setup

1. Open your Supabase project.
2. Go to SQL Editor.
3. Paste and run `supabase-schema.sql`.
4. Go to Project Settings > API.
5. Copy your Project URL and anon/public or publishable key.
6. Open the scorecard app and enter those values in Cloud Data.
7. Keep Workspace ID as `vtc-main` unless you also change the SQL policies.
8. Click Connect, then Save.
9. Click Share to copy a team link that includes the Supabase connection settings.

The app auto-saves changes to Supabase after a workbook is loaded, feedback is added, or weights change. Use Load to pull the latest shared data onto another device.

The anon key is safe to use in a browser only when Row Level Security is enabled. The included SQL enables RLS and limits anonymous access to the single `vtc-main` row.

Team members should use the Share link from the Cloud Data panel, not just the plain site URL. The plain URL opens the app, but the Share link tells the app which Supabase project/workspace to load.

## Builder Feedback Access

Builders should not need a login. The Supabase schema allows anonymous users to read active vendor, trade, category, and community dropdown data and insert widget feedback only as `needs_review`. Anonymous feedback does not affect scorecard results until a manager or admin reviews and approves it.

Because the builder widget is public, add spam protection before sharing the link broadly. Good low-cost options are a hidden honeypot field, basic rate limiting through a Supabase Edge Function, or Turnstile/reCAPTCHA if needed.

## Scoring Rules Captured

The app mirrors the current scorecard weights from the workbook:

- Safety: 25%
- Schedule: 25%
- Rework: 12.5%
- Field Log: 37.5%

The scoring logic follows the formulas in the current workbook:

- Safety starts at 100 and subtracts `Severity_Score * 10`.
- Schedule uses no-shows as the only penalty: `(Monthly_Tasks - No_Show_Count) / Monthly_Tasks * 100`.
- Rework starts at 100 and subtracts `PenaltyPoints * 5`.
- Field Log averages the `Points` column.
- Overall score is the weighted average of available weighted metrics.
- Workload is the vendor/trade share of available work in that category: vendor jobs divided by all known vendor jobs in the same trade/category.
- Vendors with the same normalized name and trade/category are merged for scoring even when the two brands use different vendor IDs.

## Hands-Off Email Path

For a true hands-off version, create a shared mailbox such as `fieldfeedback@...` and use Power Automate or Outlook rules to save forwarded messages into a OneDrive or SharePoint folder. The next version of this app can watch that folder and auto-ingest `.eml` files into the Log review queue.

Recommended email format for builders:

```text
Vendor: Vendor Name
Community: Community Name
Severity: Minor, Major, Critical, or Kudos
Notes: What happened and any context
```

The app can still parse looser emails, but this format will make automation much more reliable.
