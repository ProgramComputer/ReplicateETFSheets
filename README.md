# ReplicateETFGASheets
## About
replicateETFGASSheets uses the idea of [Direct Indexing](https://www.investopedia.com/direct-indexing-5205141) to replicate the constituents of an ETF or any file of constituents using Alpaca API and Google Sheets. Similar products are offered by brokers with fees however this app if you choose to run it will not have fees unless Google and Alpaca renege. 

<strong>I strongly recommend to try with Paper-trading first.</strong> The rebalance trigger will be costly tax-wise. It can be changed in the triggers portion in Google Apps Script menu or at [line](https://github.com/ProgramComputer/ReplicateETFGASheets/blob/ce416aaa726d3f84a7f1965643e7c8f81bbdd04a/Code.gs#L13) before execution.

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
4. Assign the scripts to the images - "setup" to "setup"  and "orders" to "Submit Orders" in "Create New Orders" sheet; "updateFills":"Refresh Sheet" in "View Order Fills" sheet,"updateSheet":"Refresh Sheet" in "Account & Portfolio" sheet
5. Enter Alpaca API keys in Account & Portfolio sheet and enter a .csv or .xlsv file link of the ETF holdings
6. Run Setup
7. Try other functions
## TODO
* Use https://github.com/SheetJS/sheetjs and migrate away from Google
* Create a Docker container to run offline
* Develop Better Tax-Loss harvesting
