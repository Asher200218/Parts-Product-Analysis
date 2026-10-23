# Parts · Products · Workers Analysis

A web app that runs **entirely on your own computer** for planning engineering design/check work. Enter or import products, parts (VC codes), worker abilities, orders and the hour-report hour log. The app then:

- calculates theoretical work hours,
- builds a **Schedule** (what to do next),
- runs a **Retrospective** (how past orders went),
- compares **theoretical vs actual** hours.

🔒 **100% local:** all data lives in the browser's localStorage. The server only listens on `localhost`, and Excel files are parsed with the bundled `web/vendor/xlsx.full.min.js`, so it works fully offline.

## Quick start

### Windows

1. Copy the whole folder to a local disk.
2. Double-click `Parts Analysis.exe`.
   If SmartScreen appears, click **More info → Run anyway**. The exe isn't code-signed, and it never goes online.
3. Keep the small black window open; the browser opens <http://localhost:4173>. Close the window when you're done. Your data is not lost.

If the exe is blocked, double-click `Start without exe.bat`. It starts the same app with the system's Python or PowerShell, on the same port and with the same data.

### macOS

Double-click `web/launch-mac.command` (needs Python 3; the first time, right-click → Open if macOS warns about an unidentified developer).

### Any OS (manual)

```bash
python3 -m http.server 4173 --directory web
```

Then open <http://localhost:4173>. Don't open `index.html` directly via `file://`: some browsers disable localStorage there.

## Tabs

| Tab | What it does |
|---|---|
| **Products** | Register product names (e.g. Alpha, Beta); they become worker-ability columns and order types |
| **Parts (VC)** | Each VC's difficulty (High/Low/None), repeat design hours t7, first-order design hours t8; check ratio setting |
| **Workers** | Design/check grade (A/B/C) per worker per product; leave records |
| **Orders** | Smart import of the big order sheet exported from the company system — finds the header row and maps columns |
| **Import / Export** | Drop files to auto-detect them; export Excel; full backup/restore (JSON) |
| **Analysis** | Theoretical hours summary; theory vs actual logged hours (everyone / per worker / per order) |
| **Schedule** | Unfinished orders only: sequential / by due date / close to due date; personal schedules and reassignment suggestions |
| **Retrospective** | Completed orders only: on-time rate, completion speed, worker load, standard vs actual efficiency |
| **Hour Log** | Import `HourReport_{year}_Analysis.xlsx`; apply the actual-hours basis in one click |

## Calculation rules

| Rule | Details |
|---|---|
| Design hours | Σ(VC hours in the order) × quantity coefficient: ≤5 VCs ×1, 6–10 ×1.2, >10 ×1.5 |
| First vs repeat | For the same product + VC, the order with the earliest created date uses t8; later ones use t7 |
| Special values | 0 hours → 0.03h; VC 999 → fixed 0.08h |
| Check hours | Design × check ratio (default 0.3, can be calibrated from the hour log) |
| Workdays | Monday–Friday only; annual leave is spread over 260 workdays per year as a capacity discount |
| Order no. | Sales Order + "-" + Item (leading zeros removed), e.g. `123456789-1200` |
| Open vs completed | Decided by whether "Actual finish date" is filled in: open → Schedule, completed → Retrospective |

## Data and backups

Data belongs to "this computer + this browser + `localhost:4173`". A different computer or browser, or clearing browser data, means the old data isn't visible. Make regular **full backups** (JSON) on the Import/Export tab and drop the file in to restore.

Compatibility with the earlier Chinese version:

- **Saved data:** data saved by the earlier version is converted automatically on first load.
- **Old Excel files:** the importers still recognize the older Chinese headers, sheet names and difficulty values, so old Excel files keep working.

## Folder layout

```
Parts Analysis.exe      Windows program (the web/ folder is built in)
User Guide.html         Full user guide
Quick Start.txt         Short version
Start without exe.bat   Fallback launcher if the exe is blocked
main.go, go.mod         Source of the exe launcher
web/
  index.html / app.js / analysis.js / styles.css   The app
  vendor/xlsx.full.min.js   Local Excel parser (SheetJS)
  launch-windows.bat        Windows launcher (Python or PowerShell)
  launch-mac.command        macOS launcher
  launch.ps1                PowerShell local server
  build-windows-exe.ps1     Alternative exe build with ps2exe (run on Windows)
```

## Rebuilding the exe

The exe is a small Go program (`main.go`) that bundles the `web/` folder into the exe and serves it on `localhost:4173`. After changing anything in `web/`, rebuild it (works from macOS/Linux too):

```bash
GOOS=windows GOARCH=amd64 CGO_ENABLED=0 go build -trimpath -ldflags "-s -w" -o "Parts Analysis.exe" .
```

See [`User Guide.html`](User%20Guide.html) for the recommended workflow, the blue/orange name-reconciliation colors and the FAQ.
