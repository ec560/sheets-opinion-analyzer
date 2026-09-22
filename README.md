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
- [Updating to the Latest Version](#updating-to-the-latest-version)
- [Common Issues](#common-issues)

## Setup

This project is meant to be pushed to a spreadsheet-bound Apps Script project with [`clasp`](https://github.com/google/clasp).

> [!IMPORTANT]
> Already have the analyzer installed and only want the newest version? Follow [Updating to the Latest Version](#updating-to-the-latest-version) instead of creating another script project.

### Requirements

- [Node.js](https://nodejs.org/) 20 or newer and `npm`
- a Google account with editor access to the target spreadsheet
- [clasp](https://github.com/google/clasp) (this guide is tested with `3.4.1`)

Install the latest version of clasp:

```bash
npm install -g @google/clasp@latest
```

Confirm that both tools are available:

```bash
node --version
clasp --version
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

This command creates `.clasp.json` and `appsscript.json` in the repository root. Leave the generated `rootDir` unchanged; clasp will include the JavaScript files inside `src/` automatically.

> [!WARNING]
> `clasp push` replaces the contents of the bound Apps Script project. For the easiest first setup, use a spreadsheet without an existing bound script. If the spreadsheet already has one, back it up and migrate it before continuing.

### Push the Code

```bash
clasp push
```

After pushing, refresh or reopen the spreadsheet. A `Tier Tools` menu should appear with the following actions:

- `Setup`
- `Refresh`
- `Analyze Selected Level`
- `Lock/Unlock Selected Level`
- `Scan Tier Flags`
- `Toggle Opinion Validation`

Run `Tier Tools -> Setup` once. Google may ask you to review and approve the script's spreadsheet permissions the first time it runs. When setup finishes, the spreadsheet should contain both `Tier Analysis` and `Tier Configuration` sheets.

The generated `Tier Configuration` sheet already contains a complete default tier list, so no additional configuration is required for the default setup. To customize it, replace the formatted list in its first column: ordinary rows are tiers and any optional row containing `/` marks a split between its neighboring tiers. Split wording is not otherwise validated. Row order becomes tier order, and each pasted cell's fill and font color become the corresponding analyzer colors. Two optional fields let multiple outer tier names count toward the bottom or top placement tier. The analyzer reads and validates changes automatically.

## Sheet Layout

The analyzer expects the spreadsheet to follow a specific layout.

Each tier should have its own sheet. The sheet name is used directly by the analyzer, so it should match the tier you want to analyze. On each tier sheet, `row 1` should contain level names. Every level header should be merged across exactly three columns in the following order:

```text
Player | Opinion | Reliability
```

Opinion rows begin on `row 2`.

For example, merge `A1:C1` and place the level name in that merged cell. Starting on row 2, put player names in column A, opinions in column B, and reliability values in column C. The tier sheet name must match a tier in `Tier Configuration`, such as `Beginner`, before it appears in the analyzer dropdown.

Row 1 headings merged across any width other than three columns are treated as section labels and excluded from level dropdowns and flag scans. A three-column, empty heading is also excluded when the following headers restart alphabetically and then continue in ascending order. This filters headings such as `Medium Tier`, `Classic`, and `Platformer` by structure and placement rather than by name, so an actual level may still use one of those names.

The analyzer depends on cell colors to gather opinion data. The tier is derived from cell colors, and reliability weighting also depends on cell colors rather than text alone. By default, a black-background Fuck opinion contributes to "Insane" tier only when its text explicitly mentions both `Fuck` and `Insane` and its font is red. Tier ordering, difficulty colors, alternate colors, and split borders are managed in the `Tier Configuration` sheet. Reliability mappings are defined in [`src/config.js`](./src/config.js).

The tier dropdown includes existing sheets whose names exactly match tier rows in `Tier Configuration`, plus the special `Pending` and `Fuck` sheets. Split labels, analyzer utility sheets, and unrelated tabs are excluded. If no configuration sheet exists, the default tier names are used. The list follows workbook tab order and refreshes when you reopen the workbook, edit the tier configuration, or use `Refresh`. An invalid or removed tier selection clears its level and analysis output. Invalid tier configuration disables tier selection until corrected.

## Using the Analyzer

Open the `Tier Analysis` sheet created by setup. Select a tier in cell `B1`, then select a level in cell `B2`. When a level is chosen, the analyzer clears the previous output, copies the selected level's `Player | Opinion | Reliability` data into the analysis sheet, and runs the analysis automatically.

If the source sheet changes and you want to reload the selected level, use the `Tier Tools -> Refresh`. If you only want to rerun the calculations on the currently loaded data, use `Tier Tools -> Analyze Selected Level`.

### Locking a Level

To lock a level, open its tier sheet, select any cell in the level's three-column `Player | Opinion | Reliability` block, and use `Tier Tools -> Lock/Unlock Selected Level`. The full three-column block changes to a black background. Player names become white, while opinion and reliability text inherit their former background colors. This is a visual formatting convention and does not prevent editors from changing cells.

Use the same menu action again to unlock the level and reverse the transformation. Black-background Fuck opinions and reliability cells preserve their original text colors through the lock and unlock cycle.

The output panel includes the selected tier and level, total weighted opinions, weighted top vote and runner-up, mean, median, outliers, standard deviation, split totals, a `Place/Move` decision, a tier distribution table, and a reliability distribution table. When applicable, it also displays `Fuck` opinion percentage and related verdict handling.

### Tier Flag Scan

Use `Tier Tools -> Scan Tier Flags` to scan a full tier sheet at once. If you run it from an eligible tier sheet, that sheet is scanned. If you run it from `Tier Analysis` (or any unrelated tab), the eligible tier selected in `B1` is scanned. Flag scans and level locking use the same tier eligibility check as the dropdown.

The scan writes to a `Tier Flags` sheet and only lists levels that might require attention. Established tiers ignore levels with `0` opinions and use low opinion count, lock alert, red or green book alert (when the analyzer's full split distribution leans away from the current tier by at least `1.5` weighted opinions), and move alert. If the analyzer's split thresholds are met, the scan reports movement instead of book status.

Pending scans use a separate flag priority: placement alert; `requires natural opinion to place` when the latest counted opinion is placeable and its reliability text contains `AREDL`, `UDL`, `External`, `Extrapolated`, or `GDDL` (for example, `Somewhat (GDDL op)`); no opinions; high opinion count at `20` or more raw opinions; an exact `+2.75` near-placement split; `close to placement` for splits of `+2` or more; low reliability when no counted opinion has green or blue reliability; and low opinion count below `3`. The natural-opinion and split flags display the current split in the Difference column. Only the highest-priority applicable flag is shown, while level rows retain their normal alphabetical order.

### Validate Opinion Colors

Open a tier sheet and use `Tier Tools -> Toggle Opinion Validation` to check opinion fill colors across all its levels, regardless of reliability. Usernames with recognized opinion fills receive a strong green highlight and readable black text; unrecognized fills remain unshaded. Configured tier colors, configured split colors, and the special black Fuck/locked fill are recognized. Red, black, missing, or unrecognized reliability does not exclude an opinion from validation. This checks fill recognition, not opinion wording, font colors, or whether an opinion contributes to analysis. The analyzer's separate counted-opinion shading stays unchanged.

The action uses the same tier eligibility check as the dropdown and the same level detection as the analyzer. It runs on the active tier sheet only; running it from `Tier Analysis` or an unrelated tab prompts you to open a tier sheet instead of using `B1`.

Run the same action again to hide the highlights. Temporary conditional formatting preserves the original cell formatting and existing rules; opinion and reliability cells are untouched. Each tier toggles independently, and all spreadsheet viewers see the highlights. The state persists when you reopen the workbook. Existing highlights can also be removed after renaming a tier or changing its configuration.

Highlights are a snapshot. After correcting colors or editing opinions, toggle off and back on to revalidate. Completely blank rows are ignored and blank gaps retain their original row positions. Merged username cells are skipped to avoid shading adjacent cells or extra rows; the completion message reports skipped rows and unrecognized fills. An empty sheet or one with no recognized opinion fills displays a message without adding formatting.

If the previous counted-opinion overlay is still enabled when updating, the first toggle removes it and the next enables validation. Buttons assigned to the old `toggleTierCountedOpinions` function continue to work.

## Updating to the Latest Version

The spreadsheet does not update itself when this repository changes. To install the latest analyzer code, update the local repository, push it to Apps Script, and then refresh the spreadsheet.

Before updating, open the repository in a terminal and check for local changes:

```bash
cd sheets-opinion-analyzer
git status
```

If `git status` shows changes you want to keep, commit or stash them before pulling. If you cloned this repository directly, update from `origin`:

```bash
git pull --ff-only origin main
clasp push
```

If you forked the repository, add the original repository as an upstream remote once:

```bash
git remote add upstream https://github.com/ec560/sheets-opinion-analyzer.git
```

Then use `upstream`, rather than your fork's potentially older `origin`, whenever you want the latest analyzer version:

```bash
git pull --ff-only upstream main
clasp push
```

After `clasp push` succeeds, refresh or reopen the spreadsheet so the latest menu and code are loaded. If the update changes the analyzer layout or setup behavior, run `Tier Tools -> Setup` again. This rebuilds the `Tier Analysis` sheet but does not modify any source tier sheets.

If `clasp push` reports that you are not authorized, run `clasp login` again and repeat the push. Keep the `.clasp.json` created during initial setup, because it identifies the Apps Script project that receives future updates.

To update clasp itself when needed:

```bash
npm install -g @google/clasp@latest
clasp --version
```

## Common Issues

If the `Tier Tools` menu does not appear, refresh or reopen the spreadsheet and make sure `clasp push` completed successfully. It is also worth confirming that the Apps Script project is bound to the spreadsheet you meant to use.

If no levels appear after selecting a tier, make sure the selected sheet exists and that `row 1` contains non-empty level headers.

If opinions fail to load, make sure the selected level actually has data below `row 1`, and make sure the sheet follows the expected three-column `Player | Opinion | Reliability` structure.

If results look wrong, the first thing to verify is the sheet formatting. This analyzer relies on background and font colors, so incorrect colors will impact how opinions are interpreted.
