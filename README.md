# opinion-analyzer
Analyzes opinions from major opinion sheets and prints relevant statistics to assist with placement/movement of levels in their respective difficulty categories.

Credit to Amberette/Cadrega for inspiration and basic config structures.

## Setup

After you fork the repository and clone it to your local system, navigate to the root folder
### On Linux (or WSL):

```bash
npm install -g @google/clasp
```
---

### Authenticate

Choose the **Google account associated with the spreadsheet** you will use. If necessary, use an alternate Google account.

```bash
clasp login
```
Ensure you are only logged into the google account you want to use. Having multiple signed-in accounts may cause linking issues

---

### Get Your Spreadsheet ID

Open your Google Sheet and copy the ID from the URL:

```
https://docs.google.com/spreadsheets/d/<YOUR_SHEET_ID>/edit
```

Copy the string of text between `/d/` and `/edit`.

---

### Create the Apps Script Project (Bound to the Sheet)

Choose a unique script name (any name works).

```bash
clasp create-script --title "<YOUR_SCRIPT_NAME>" --parentId <YOUR_SHEET_ID>
```

If you see an error like:

```
"User has not enabled the Apps Script API."
```

Enable the Apps Script API for your account with the provided link, then run the command again.

---

### Push the code

```bash
clasp push
```
The files will now be available in the associated apps script dashboard to your sheet.

---

### Activate in Google Sheets

Refresh your Google Sheet; You should now see a new **"Tier Tools"** menu button at the top of the sheet.

![Tier Tools Menu](https://github.com/user-attachments/assets/b0fdafca-0fc4-4884-80cf-ed9698763271)

Pressing the "Setup" button from Tier Tools will create the **Tier Analysis** sheet.

---

## Updating to Latest Version

If you forked the repository you must run the following command in your root folder first:

```bash
git remote add upstream https://github.com/ec560/sheets-opinion-analyzer.git
```

1. Pull the latest repository changes:

```bash
git pull
```

2. Push updates to your Google Sheet:

```bash
clasp push
```

3. Refresh your spreadsheet.

## Using the sheet

### Example output
<img width="1200" height="774" alt="image" src="https://github.com/user-attachments/assets/91aa483e-26b3-474f-b1fa-6cccd709663d" />
