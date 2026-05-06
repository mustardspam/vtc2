# VTC Scorecard Updater

This is a local-first starter app for updating the Vendor & Trade Council scorecard from dropped Excel files and field feedback emails.

## How to Use

1. Open `index.html` in a browser.
2. Load the master `VTC Scorecard 4.21.xlsx` workbook.
3. Drop source Excel or CSV files that contain tabs or headers for:
   - `Safety`
   - `Schedule_Adherence_Raw`
   - `Inspections`
   - `Rework`
   - `Warranty`
   - `Log`
4. Review the recalculated scores.
5. Paste forwarded builder feedback into the Field feedback tab, review the parsed entry, and add it to the Log.
6. Export the updated workbook.

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

The app auto-saves changes to Supabase after a workbook is loaded, feedback is added, or weights change. Use Load to pull the latest shared data onto another device.

The anon key is safe to use in a browser only when Row Level Security is enabled. The included SQL enables RLS and limits anonymous access to the single `vtc-main` row.

## Scoring Rules Captured

The app mirrors the current scorecard weights from the workbook:

- Safety: 25%
- Schedule: 25%
- Inspections: 0%
- Rework: 12.5%
- Warranty: 0%
- Field Log: 37.5%

The scoring logic follows the formulas in the current workbook:

- Safety starts at 100 and subtracts `Severity_Score * 10`.
- Schedule averages `Adherence_Pct` and converts it to a 0-100 score.
- Inspections average `Score_Pct`.
- Rework starts at 100 and subtracts `PenaltyPoints * 5`.
- Warranty averages `Warranty_Score`.
- Field Log averages the `Points` column.
- Overall score is the weighted average of available weighted metrics.

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
