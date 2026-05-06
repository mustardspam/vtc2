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
