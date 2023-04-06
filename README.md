# ReplicateETFGASheets
##Info
A rebalance trigger is fired everyday at 9:30 ET.
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
## Contributions
are welcome
