This project uses a Google Sheet to manually keep track of your finances within a single account and to forecast your balances based on both recurring transactions and manually entered transactions.

You can make a copy of the spreadsheet here to get started: [Checking Estimator](https://docs.google.com/spreadsheets/d/11YmGiyljsHIa93hsYmDxWZ_L_ruKUmPMRlufNl3WJCA/copy)

Set up instructions:
1. Change the name of the tab that is labeled with the format 'Mmm yyyy' to the three letter month and four digit year you'd like to start at. The name is important, be sure to capitalize only the first letter and only use the standard three letter abbreviation such as 'Jan 2026' for January 2026.
2. On the 'Setup' tab, enter the name of your starting month in cell C2 in the same format as the previous step: 'Mmm yyy'.
3. Change the name of the tab that is labeled with the format 'yyyy' to the four digit year you'd like to start with.
4. On your first monthly tab, enter the starting balance of your account in the green box near the top where it says 'Starting balance:'.
5. Authorize the script automation by selecting the 'Sheet automation' menu and then the 'Enable automation' option. You will then follow the Google prompts to authorize the script to run.
6. On the 'Recurring Transactions' tab, set up any recurring transactions you'd like to start with each month. Enter only the day of the month, not full date, for each transaction. Place the recurring transactions on the 'Transactions' tab by selecting the month you'd like to add them to in the blue box on the right and then click the checkbox next to the 'Process Month' text. Note, if you do not already have a new tab created for the month you've selected, a new tab will be created automatically.
7. You can add any non-recurring transactions to the 'Transactions' tab before or after the recurring transactions have been added.
8. If you would like to reduce clutter, you can use the tab box on the right side of the 'Recurring Transactions' tab to archive any older data you'd like to remove from the 'Transactions' tab. Select the month you'd like to archive and click the checkbox next to 'Archive Month' to move the older transactions to the 'Archived Transactions' tab. Note that only reconciled transactions will be moved, so make sure you've entered actual amounts for all old transactions before you run this function. You can run it multiple times, if necessary. Note that the tab or the archived month will be hidden when you run this function.
9. The 'Transactions' tab has a few automations that are triggered with checkboxes. They should be fairly self-explanatory, so go ahead and give them a try.

You may wonder why I chose to use checkboxes to trigger automations rather than buttons. This was primarily to support running this spreadsheet from the Google Sheets mobile app, which does not support running functions from buttons at this time. Some of the automation for the dates was also added to make this spreadsheet function better from the mobile app as entering dates can be a bit tedious in the app.
