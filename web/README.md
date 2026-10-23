# web/ — the Parts Analysis app

The static web app itself: `index.html`, `app.js` (UI, import/export, rendering), `analysis.js` (calculations), `styles.css`, and the bundled SheetJS parser in `vendor/`. Everything runs in the browser; data is stored in localStorage. There are no network requests.

## Running it

It must be served over `http://localhost` (not opened as `file://`, where some browsers disable localStorage):

| How | Notes |
|---|---|
| `../Parts Analysis.exe` | Windows. This folder is built into the exe; serves on port 4173 and opens the browser. |
| `launch-windows.bat` | Windows, no install needed: uses Python if present, otherwise PowerShell (`launch.ps1`). |
| `launch-mac.command` | macOS, needs Python 3. Right-click → Open the first time. |
| `python3 -m http.server 4173` | Any OS, run from this folder. |

Keep using port 4173: browser data is tied to the address, so a different port shows an empty app.

`build-windows-exe.ps1` is an alternative way to make an exe (wraps `launch.ps1` with ps2exe; run it on Windows). The exe shipped in the repo root is built from `../main.go` instead.

## Excel formats

Exports (`parts.xlsx`, `workers.xlsx`, `orders.xlsx`, `products.json`):

- `parts.xlsx` → sheet `Parts`, header `Part name | Difficulty (High/Low/None)`, plus `Hours per piece (h)` if you customized hours per difficulty
- `workers.xlsx` → sheet `Worker abilities`, header `Worker name | <product1>·Design | <product1>·Check | <product2>·Design | …`
- `orders.xlsx` → sheet `Orders`, header `Product type | Parts needed | Submitted`

Importers also accept the older Chinese sheet names, headers (e.g. `零件名称`, `工人姓名`, `产品类型`, `产品·设计`/`产品·校验`) and difficulty values (`高/低/无`).

The **Parts** tab's "Import parts.xlsx" button reads by position: column 1 = VC code, column 8 = repeat design hours (t7), column 9 = first-order design hours (t8).

The **Orders** tab's "Import orders.xlsx" button finds the header row automatically and shows a column-mapping dialog. It recognizes English and Chinese keywords such as Sales Order, Item, VC, Designer, Checker, Product, Required/Due, Create On and Actual Finish.

### Retrospective Excel (strict column order, 4 columns)

The header text is ignored and row 1 is skipped:

| Col 1 | Col 2 | Col 3 | Col 4 |
|---|---|---|---|
| Product type | Parts needed (`name×N; …`) | Submitted `YYYY-MM-DD` | Actual finish date `YYYY-MM-DD` |

Date columns can be text like `2025-01-06` or real Excel dates. Use "Download retrospective Excel / template" to get a sample.

### Hour log

`HourReport_{year}_Analysis.xlsx` from hour-report is detected by its `Data (long)` sheet. The hour-report export names tasks in Chinese; `订单设计`, `订单校对` and `加班` are mapped to *Order design*, *Order check* and *Overtime*.
