# Opinion Analyzer for Google Sheets

This project is a Google Apps Script tool for reviewing Geometry Dash level placement opinions stored in a Google Sheet. It creates a dedicated `Tier Analysis` sheet, lets you choose a tier and level via dropdown, loads the corresponding opinions, and generates weighted output intended to help with placement and movement decisions.

> [!WARNING]
> The analyzer is still being tested in live use, so some results may still need manual judgment.

> [!IMPORTANT]  
> If you encounter any bugs please open an issue or contact me directly.

Credit to `Amberette/Cadrega` for inspiration and the original tier config.

## Contents

- [Setup](#setup)
- [Sheet Layout](#sheet-layout)
- [Using the Analyzer](#using-the-analyzer)
- [Updating](#updating)
- [Common Issues](#common-issues)

## Setup

This project is meant to be pushed to a spreadsheet-bound Apps Script project with [`clasp`](https://github.com/google/clasp).

### Requirements

- [Node.js](https://nodejs.org/) and `npm`
- a Google account with editor access to the target spreadsheet
- [clasp](https://github.com/google/clasp)

Install clasp:

```bash
npm install -g @google/clasp
```

Sign in with the Google account that has edit access to the spreadsheet:

```bash
clasp login
```

> [!NOTE]  
> If the Apps Script API is not enabled for your account yet, enable it here:
> [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)

### Clone the Repository

```bash
git clone https://github.com/ec560/sheets-opinion-analyzer.git
cd sheets-opinion-analyzer
```

### Bind the Script to a Spreadsheet

Open the target Google Sheet and copy the spreadsheet ID from its URL:

```text
https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit
```

Then create a script project bound to that spreadsheet:

```bash
clasp create-script --title "<YOUR_SCRIPT_NAME>" --parentId "<YOUR_SHEET_ID>"
```

> [!NOTE]  
> If the spreadsheet already has a bound Apps Script project, you will need to remove or migrate that project before creating a new one.
>
> This repository keeps its Apps Script source files in `src/`. Before running `clasp push`, make sure your `.clasp.json` is configured to use `src` as the `rootDir`.

### Push the Code

```bash
clasp push
```

After pushing, refresh the spreadsheet. A `Tier Tools` menu should appear with the following actions:

- `Setup`
- `Refresh`
- `Analyze Selected Level`
- `Scan Tier Flags`

Run `Setup` once to create the `Tier Analysis` sheet.

Setup also creates a `Tier Configuration` sheet. Paste a single formatted list into its first column: ordinary rows are tiers and any optional row containing`/` marks a split between its neighboring tiers. Split wording is not otherwise validated. Row order becomes tier order, and each pasted cell's fill and font color become the corresponding analyzer colors. Two optional fields let multiple outer tier names count toward the bottom or top placement tier. The analyzer reads and validates changes automatically.

## Sheet Layout

The analyzer expects the spreadsheet to follow a specific layout.

Each tier should have its own sheet. The sheet name is used directly by the analyzer, so it should match the tier you want to analyze. On each tier sheet, `row 1` should contain level names. Every level should occupy three columns in the following order:

```text
Player | Opinion | Reliability
```

Opinion rows begin on `row 2`.

The analyzer depends on cell colors to gather opinion data. The tier is derived from cell colors, and reliability weighting also depends on cell colors rather than text alone. By default, a black-background Fuck opinion contributes to "Insane" tier only when its text explicitly mentions both `Fuck` and `Insane` and its font is red. Tier ordering, difficulty colors, alternate colors, and split borders are managed in the `Tier Configuration` sheet. Reliability mappings are defined in [`src/config.js`](./src/config.js).

One thing to be aware of is that the tier dropdown includes every sheet in the workbook except analyzer utility sheets (`Tier Analysis`, `Tier Flags`, and `Tier Configuration`). If your spreadsheet contains unrelated sheets, they will appear in the selector as well.

## Using the Analyzer

Open the `Tier Analysis` sheet created by setup. Select a tier in cell `B1`, then select a level in cell `B2`. When a level is chosen, the analyzer clears the previous output, copies the selected level's `Player | Opinion | Reliability` data into the analysis sheet, and runs the analysis automatically.

If the source sheet changes and you want to reload the selected level, use the `Tier Tools -> Refresh`. If you only want to rerun the calculations on the currently loaded data, use `Tier Tools -> Analyze Selected Level`.

The output panel includes the selected tier and level, total weighted opinions, weighted top vote and runner-up, mean, median, outliers, standard deviation, split totals, a `Place/Move` decision, a tier distribution table, and a reliability distribution table. When applicable, it also displays `Fuck` opinion percentage and related verdict handling.

### Tier Flag Scan

Use `Tier Tools -> Scan Tier Flags` to scan a full tier sheet at once. If you run it from a tier sheet, that sheet is scanned. If you run it from `Tier Analysis` (or anywhere else), the tier selected in `B1` is scanned.

The scan writes to a `Tier Flags` sheet and only lists levels that might require attention. Levels with `0` opinions are ignored. Current flags are: low opinion count, lock alert, red or green book alert (when the analyzer's full split distribution leans away from the current tier by at least `1.5` weighted opinions), move alert, and placement alert. If the analyzer's split thresholds are met, the scan reports movement instead of book status.

## Updating

If you forked the repository, add the original repository as an upstream remote once:

```bash
git remote add upstream https://github.com/ec560/sheets-opinion-analyzer.git
```

To update your local copy and republish it to the spreadsheet:

```bash
git pull
clasp push
```

Refresh the spreadsheet after pushing.

## Common Issues

If the `Tier Tools` menu does not appear, refresh or reopen the spreadsheet and make sure `clasp push` completed successfully. It is also worth confirming that the Apps Script project is bound to the spreadsheet you meant to use.

If no levels appear after selecting a tier, make sure the selected sheet exists and that `row 1` contains non-empty level headers.

If opinions fail to load, make sure the selected level actually has data below `row 1`, and make sure the sheet follows the expected three-column `Player | Opinion | Reliability` structure.

If results look wrong, the first thing to verify is the sheet formatting. This analyzer relies on background and font colors, so incorrect colors will impact how opinions are interpreted.
