# ReplicateETFGASheets
## About
replicateETFGASSheets uses the idea of Direct Indexing to replicate the constituents of an ETF or any file of constituents using Alpaca API and Google Sheets. Similar products are offered by brokers with fees however this app if you choose to run it will not have fees unless Google and Alpaca renege. I strongly recommend to try with Paper-trading first.

This is still a Work-In-Progress, contributions will be welcome, ask any questions in [Discussions](https://github.com/ProgramComputer/ReplicateETFGASheets/discussions).

## Info


A rebalance trigger is fired every 3 hours and constituents updated everyday at 9:30 ET.
## Disclaimer:
There can be significant tax, financial, and legal consequences related to this repository and its use.
Signficant losses can occur. Trading equities is always dangerous and should be approached with extreme caution.
By using this repository, you accept all liability or consequences related to the code provided here.

## Steps
1. Move spreadsheets to Google Drive and convert to Google Sheets
2. Copy Scripts to Apps Script Project in Google
3. Add Google Sheets Service to Apps Script Project and assign identifier "Sheets"
4. Assign the scripts - "setup" to "setup"  and "orders" to "Submit Orders" in "Create New Orders" sheet; "updateFills":"Refresh Sheet" in "View Order Fills" sheet,"updateSheet":"Refresh Sheet" in "Account & Portfolio" sheet
5. Enter Alpaca API keys in Account & Portfolio sheet and enter a .csv or .xlsv file link of the ETF holdings
6. Run Setup
7. Try other functions
## TODO
* Use https://github.com/SheetJS/sheetjs and migrate away from Google
* Create a Docker container to run offline
