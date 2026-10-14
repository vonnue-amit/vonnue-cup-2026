# Vonnue Cup website

A static site that shows the Vonnue Cup schedule, brackets, pools, team standings, team pages and
player pages. It reads the organiser's workbooks directly, so there is no database and no build step.

```
website/
  index.html      page shell
  config.js       settings: Google Sheets link, year, refresh interval
  styles.css      theme (dark by default, light theme from the toggle or the device setting)
  app.js          views: Overview, Schedule, Brackets, Standings, Teams, Player
  parser.js       reads the workbooks
  formula.js      fills in formula results when a workbook was saved without them
  assets/teams/   team logos (team-1.webp … team-6.webp, plus -sm versions for small badges)
  data/
    Vonnue_Cup_Knockout_Brackets.xlsx   results come from here
    Vonnue_Cup_Participants.xlsx        rosters (optional)
```

## Updating results from Google Sheets (no uploads)

1. Upload `Vonnue_Cup_Knockout_Brackets.xlsx` to Google Drive and open it with Google Sheets
   (File → Save as Google Sheets). Keep the sheet names and layout as they are.
2. Share → General access → **Anyone with the link**, role **Viewer**. You keep edit rights.
3. Copy the sheet's link and paste it into `config.js` as `googleSheet: '...'`, then upload the site
   once more.
4. From then on, record results in the Google Sheet. Open pages pick them up within the refresh
   interval in `config.js` (2 minutes by default).

If Google can't be reached, or the sheet stops looking like the Vonnue Cup workbook, the site falls
back to `data/Vonnue_Cup_Knockout_Brackets.xlsx`, so keep a reasonably recent copy there. To check
the Google version before switching over, use File → Download → Microsoft Excel in Google Sheets and
open that file with the site's "Open a newer workbook" button.

Anyone with the sheet's link can read it, including the hidden helper columns; only you can edit.
The private Claude preview page can't reach Google, so it keeps showing its snapshot.

## Updating results by uploading the workbook

1. Record results in `Vonnue_Cup_Knockout_Brackets.xlsx` as the Organiser Guide describes, and save it
   in Excel, Google Sheets (File → Download → .xlsx) or LibreOffice. Files written by a script
   (openpyxl, for example) are fine too: they carry formulas without results, and the site
   calculates those itself (`formula.js`, checked cell by cell against LibreOffice).
2. Replace `website/data/Vonnue_Cup_Knockout_Brackets.xlsx` with the saved file. Keep the file name.
3. Visitors' pages check for a new file every 2 minutes and update without a reload.

Don't rename sheets or move cells. The guide's rules apply here too, because the site finds matches,
feeders and byes through the Schedule sheet's formulas.

## Team logos

The originals are in `../team_logos`. `make-logos.py` crops each to a square and writes the web
versions to `assets/teams/` (`python3 make-logos.py`, needs Pillow). To change a logo, replace the
original and rerun it, or drop a square image in as `assets/teams/team-N.webp` and `team-N-sm.webp`.
For portrait images, the crop position is set per team at the top of the script. If a logo is
missing, the site shows the team's initials instead.

## Running it

Browsers block reading files when `index.html` is opened from disk, so serve the folder:

```
cd website
python3 -m http.server 8000
# open http://localhost:8000
```

To host it for everyone, upload the `website` folder to any static host, such as GitHub Pages,
Netlify, Cloudflare Pages, an internal web server or a SharePoint/IIS static site.

### Hosts that can't serve .xlsx files

If the .xlsx files aren't there, the site loads `data/snapshot.json`, a frozen copy of both workbooks.
After updating the workbooks, rebuild it with `npm i xlsx@0.18.5 && node make-snapshot.js` and upload it.
A snapshot doesn't change on its own, so prefer serving the .xlsx where you can.

## Handy extras

- **Open a newer workbook** (footer): preview any brackets file in your own browser without uploading it.
- **Search**: press `/` or `Ctrl K` to find a player, team, event or match ID.
- **Testing a date**: add `?now=2026-10-19T19:25` to the URL to see the site as it would look at that
  moment, with live, upcoming and "result due" states.
- The year is set in `app.js` (`CONFIG.year`), because the workbook dates have no year.
