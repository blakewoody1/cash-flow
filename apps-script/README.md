# Apps Script backup

`Code.gs` is a backup of the Apps Script that runs the Cash Flow backend. The live copy is bound to the Cash Flow Google Sheet (Extensions → Apps Script).

The `KEY` line is a placeholder. The real key lives only in the live script and on each device's app Settings — never commit it here.

## Restore or update the live script
1. Before pasting, copy your real key from the live script's `const KEY = '…';` line.
2. Open the Cash Flow sheet in a browser → Extensions → Apps Script → Code.gs.
3. Select all, paste this file's contents, then put your real key back on the `KEY` line.
4. Save.
5. Deploy → Manage deployments → ✎ on the existing deployment → Version: **New version** → Deploy. (Not "New deployment" — that changes the /exec URL.)
6. First time only (or after restoring into a new script): pick `setupSchedule` in the function menu and ▷ Run, then approve Google's permission prompt. That turns on the monthly job (adds month tabs ~14 months ahead; each January moves tabs and log entries older than last year into "<name> Archive" inside the "Cash Flow App" Drive folder) and a nightly copy of this repo's app files into "Cash Flow App/App code" in Drive.

## Keep this backup current
Whenever the live Code.gs changes, update this file too (with the key blanked out) and push.
