# opinion-analyzer
Analyzes opinions from major opinion sheets and prints relevant statistics to assist with placement/movement of levels in their respective difficulty categories

## Setup

After you fork the repository and clone it to your local system, navigate to the root folder
On a linux-based system (or WSL), paste the following commands and follow the steps in order:

npm i -g @google/clasp (tested with version 3.2.0)

clasp login - choose the google account associated with the sheet you will be using, and authenticate it to your account (use an alternative google account if necessary)

Choose a unique script name, it does not matter what you call it specifically.
To get your sheet ID, Copy the part between /d/ and /edit
clasp create-script --title "<YOUR_SCRIPT_NAME>" --parentId <YOUR_SHEET_ID> - (it may print an error such as "User has not enabled the Apps Script API." If it does, enable that for the respective account and then run the command again)

clasp push

Refresh your google sheet. You should see the "Tier Tools" button on the end.
<img width="707" height="55" alt="image" src="https://github.com/user-attachments/assets/b0fdafca-0fc4-4884-80cf-ed9698763271" />
Click the button to reveal more options, and then click "Setup" to trigger the Tier Analysis sheet creation.

## Using the sheet

### Example output
<img width="1200" height="774" alt="image" src="https://github.com/user-attachments/assets/91aa483e-26b3-474f-b1fa-6cccd709663d" />
