# spa-emergenyAlerts

School Emergency / Assistance Alert System — a single-page web app built on
Google Apps Script, bound to a Google Spreadsheet. Teachers raise alerts from
a mobile-first form; responders triage them on a live, auto-refreshing
dashboard.

## Files

| File | Purpose |
| --- | --- |
| `Code.gs` | Backend: `doGet`, idempotent `setup()`, cached reference data, alert CRUD with `LockService`, structured `{success, message, data}` responses |
| `Index.html` | Frontend SPA: Tailwind (CDN), teacher alert form, live dashboard with optimistic UI, 5-second polling, undo, audio cue |
| `appsscript.json` | Manifest: timezone, OAuth scopes, web-app deployment settings |

## Spreadsheet schema (created by `setup()`)

- **Data** — `ID, Timestamp, Raised By, Student Name, Location, Category, Status, Responder, Time Responded, Time Resolved`
- **Staff** — `Name, Email, Role` (reference data only — not an access-control list)
- **Student** — `Name, ID, Year` (feeds the form's predictive datalist)

## Deployment checklist

1. Create (or open) the Google Spreadsheet that will hold the data, then open
   **Extensions → Apps Script** so the script is container-bound.
2. Paste `Code.gs` into the `Code.gs` file in the editor. Create an HTML file
   named exactly `Index` and paste in `Index.html`.
3. (Optional but recommended) Enable **Project Settings → Show "appsscript.json"**
   and paste in the provided manifest so the scopes and timezone match.
4. In the editor, select the `setup` function and click **Run**. Approve the
   authorisation prompts (spreadsheet access, email identity, script storage).
   `setup()` is idempotent — safe to re-run at any time.
5. Add your students to the **Student** tab (and optionally staff to **Staff**).
   These are cached for 5 minutes, so edits appear in the form shortly after.
6. **Deploy → New deployment → Web app** with exactly:
   - **Execute as:** *User accessing the web app*
   - **Who has access:** *Anyone within [your domain]*
7. Open the web-app URL. Each user authorises once on first visit; their email
   is then captured automatically for `Raised By` / `Responder`.
8. Share the URL with staff. Dashboard users should click **Enable sound 🔔**
   once per session to unlock the new-alert beep (browsers block audio until
   the page has been interacted with).

> After editing the code later, use **Deploy → Manage deployments → Edit →
> New version** so the existing URL picks up the changes.
