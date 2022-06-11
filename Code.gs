function setup(){
  var name  = getHoldings();
  var triggers = ScriptApp.getProjectTriggers();
  var shouldCreateTrigger = true;
  triggers.forEach(function (trigger) {
  
    if(trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK && trigger.getHandlerFunction() === "setup") {
      shouldCreateTrigger = false; 
    }
  });
  
  if(shouldCreateTrigger) {
    ScriptApp.newTrigger("setup").timeBased().atHour(9).nearMinute(30).everyDays(1).inTimezone("America/New_York").create()
  }
  shouldCreateTrigger = true;
   triggers.forEach(function (trigger) {
  
    if(trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK && trigger.getHandlerFunction() === "rebalance") {
      shouldCreateTrigger = false; 
    }
  });
  
  if(shouldCreateTrigger &&  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders")
.getRange("F4").getValue()) {
    ScriptApp.newTrigger("rebalance").timeBased().everyHours(3).inTimezone("America/New_York").create()
  }


var holdingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name +" Holdings")




var foundTickers = holdingsSheet.getRange(1,1,holdingsSheet.getLastRow(),holdingsSheet.getLastColumn()).createTextFinder("(?:^Holding [Tt][Ii][cC][kK][eE][rR]|^[Tt][Ii][cC][kK][eE][rR])$").useRegularExpression(true).findNext();

var foundWeights = holdingsSheet.getRange(1,1,holdingsSheet.getLastRow(),holdingsSheet.getLastColumn()).createTextFinder("^Weight$").useRegularExpression(true).findNext();

tickersA1 = foundTickers.getA1Notation();
weightsA1 = foundWeights.getA1Notation();
var ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders")

ordersSheet.getRange("Q9:Q").clear();
ordersSheet.getRange("R9:R").clear();
ordersSheet.getRange("Q9:Q").setNumberFormat("#0.00#%");
ordersSheet.getRange("R9:R").setNumberFormat("#0.00#%")
ordersSheet.getRange("Q9").setFormula("=indirect(cell(\"address\",Offset(Indirect(Index(CELL(\"address\",INDEX('Account & Portfolio'!$A$17:$A,MATCH(trim(I9),'Account & Portfolio'!$A$17:$A,0))))),0,2)))/indirect(cell(\"address\",Offset(Indirect(Index(CELL(\"address\",INDEX('Account & Portfolio'!$C$17:$C,MATCH(\"total\",'Account & Portfolio'!$C$17:$C,0))))),1,0)))")
  ordersSheet.getRange("R9").setFormula("=Q9-(J9/100)");
  ordersSheet.getRange("Q9").copyTo(ordersSheet.getRange("Q10:Q"))
  ordersSheet.getRange("R9").copyTo(ordersSheet.getRange("R10:R"))

var batchGetValues = Sheets.Spreadsheets.Values.batchGet(holdingsSheet.getParent().getId(),{ranges:[name +" Holdings!"+tickersA1[0]+(parseInt(tickersA1[1])+1)+":"+tickersA1[0],name +" Holdings!"+weightsA1[0]+(parseInt(weightsA1[1])+1)+":"+weightsA1[0]]})
// console.log(batchGetValues.valueRanges[0].values)
// console.log(batchGetValues.valueRanges[1].values)
 const batchGetValuesMerged = batchGetValues.valueRanges[0].values.map((item,i) => [item[0].trim(),batchGetValues.valueRanges[1].values[i][0].trim()]); 
 const assets = JSON.stringify(getAssets())
 const filteredBatchGetValuesMerged = batchGetValuesMerged.filter(o => assets.includes(o[0] ))

const firstArrayBatchGet = []
const secondArrayBatchGet = []
 for (const element of filteredBatchGetValuesMerged) {
    firstArrayBatchGet.push([element[0]]);
    secondArrayBatchGet.push([element[1]]);
}


//buy starts here
ordersSheet.getRange("A9:G").clearContent()




ordersSheet.getRange("A9:A" + (8 + firstArrayBatchGet.length)).setValues(firstArrayBatchGet)
ordersSheet.getRange("B9:B"+ (8 + secondArrayBatchGet.length)).setValues(secondArrayBatchGet)
ordersSheet.getRange("C9:C"+ (8 + secondArrayBatchGet.length)).setValue("market")
ordersSheet.getRange("D9:D"+ (8 + secondArrayBatchGet.length)).setValue("day")


//  let symbolRange = Sheets.newValueRange();
//       symbolRange.range = "Create New Orders!A9:A";
//       symbolRange.values = batchGetValues.valueRanges[0].values;
//        let qtyRange = Sheets.newValueRange();
//       qtyRange.range = "Create New Orders!B9:B";
//       qtyRange.values = batchGetValues.valueRanges[1].values
//        let typeRange = Sheets.newValueRange();
//       typeRange.range = "Create New Orders!C9:C";
//       typeRange.values = [["market"]]
//        let tifRange = Sheets.newValueRange();
//       tifRange.range = "Create New Orders!D9:D";
//       tifRange.values = [["day"]]

//       let batchUpdateRequest = Sheets.newBatchUpdateValuesRequest();
//       batchUpdateRequest.data = [symbolRange,qtyRange,typeRange,tifRange];
//       batchUpdateRequest.valueInputOption = 'RAW';
// const result = Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest,holdingsSheet.getParent().getId()
//         )
//         console.log(result)

//clear A to G from 9 down to getMaxRows in Create New Orders
//copy add orders


//sell starts here
ordersSheet.getRange("I9:O").clearContent()
ordersSheet.getRange("I9:I" + (8 + firstArrayBatchGet.length)).setValues(firstArrayBatchGet)
ordersSheet.getRange("J9:J"+ (8 + secondArrayBatchGet.length)).setValues(secondArrayBatchGet)
ordersSheet.getRange("K9:K"+ (8 + secondArrayBatchGet.length)).setValue("market")
ordersSheet.getRange("L9:L"+ (8 + secondArrayBatchGet.length)).setValue("day")

//add conditional format rules
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders");

var rangeG = sheet.getRangeList(["G1","G6:G"])

rangeG = rangeG.getRanges()
var rangeO = sheet.getRangeList(["O1","O6:O"])
rangeO = rangeO.getRanges()

sheet.clearConditionalFormatRules();
var rule1 = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith("{ id")
    .setBackground("#B7E1CD")
    .setRanges(rangeG.concat(rangeO))
    .build();
var rule2 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Status")
    .setBackground("#FFF7CC")
    .setRanges(rangeG.concat(rangeO))
    .build();
var rule3 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Order not sent, already have desired quantity.")
    .setBackground("#FCE8B2")
    .setRanges(rangeG.concat(rangeO))
    .build();
var rule4 = SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty()
    .setBackground("#F4C7C3")
    .setRanges(rangeG.concat(rangeO))
    .build(); 
var rules = sheet.getConditionalFormatRules();
rules.push(rule1,rule2,rule3,rule4)
sheet.setConditionalFormatRules(rules);

sheet.getRange("G9:G").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
sheet.getRange("O9:O").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);




 
var percents = 
    sheet.getRange("B9:B").getValues().filter(String).map(function(elem){return elem.toString()})
        var minimumDollars = 0
      var  lowestPercentage = 100
        percents.forEach((key) => {
        
            
           var percent = parseFloat(key)
          
            if (percent < lowestPercentage){
                lowestPercentage = percent
            }
            minimumDollars+=Math.ceil(1/(percent/100))
       
        })
        //reduce the total minimum dollars to ensure smallest equity meets minimum notional requirement of $1
        smallestInvestment = minimumDollars*(lowestPercentage/100)
        if (smallestInvestment > 1){
            minimumDollars /= smallestInvestment
        }
        
sheet.getRange("B6").setValue(Math.ceil(minimumDollars))
}
function orders(){
   var side =     SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders").getRange("B3").getValue();
   console.log(side)
   if(side == "Rebalance"){
     rebalance()
   }
   else if(side == "Buy"){
buy()
   }
    else if(side == "Sell"){
     sell()
   }

}
function buy(){
  clearOrders();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders");
  var market_value = parseFloat(getAccount().long_market_value) + parseFloat(getAccount().short_market_value);
  var percent = sheet.getRange("F3").getValue();
  var rebalance = sheet.getRange("F4").getValue();
  var extendedHours = sheet.getRange("F5").getValue();
  
  
  var symbols = {
    "buy": sheet.getRange("A9:A").getValues().filter(String).map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("I9:I").getValues().filter(String).map(function(elem){return elem.toString()})
  }
  var qtys = {
    "buy": sheet.getRange("B9:B").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("J9:J").getValues().map(function(elem){return elem.toString()})
  }
  var types = {
    "buy": sheet.getRange("C9:C").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("K9:K").getValues().map(function(elem){return elem.toString()})
  }
  var tifs = {
    "buy": sheet.getRange("D9:D").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("L9:L").getValues().map(function(elem){return elem.toString()})
  }
  var limits = {
    "buy": sheet.getRange("E9:E").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("M9:M").getValues().map(function(elem){return elem.toString()})
  }
  var stops = {
    "buy": sheet.getRange("F9:F").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("N9:N").getValues().map(function(elem){return elem.toString()})
  }
  for(var i = 0; i < symbols.buy.length; i++) {
      if(symbols.buy[i] != ""){
         var account = getAccount()
        market_value = parseFloat(account.long_market_value) + parseFloat(account.short_market_value);
      sheet.getRange("G"+parseFloat(9+i)).setValue("submitting...");
      
      var qty = parseFloat(qtys.buy[i].toString().trim());
      var sym = symbols.buy[i].toString().trim()
      var side = "buy"
      //var diffPercent = parseFloat(getPosition(sym).market_value)/market_value - (qty)/100
      var targetBuy = parseFloat(sheet.getRange("B4").getValue()) * (qty/100);
      var position = getPosition(sym)


     // if(percent) {
        //targetBuy = (parseFloat(position.market_value) - market_value * (qty/100) )/((qty/100)-1)
     // }
     // if(rebalance) {
        var position_qty
        if(isNaN(parseFloat(position.qty))) position_qty = 0;
        else position_qty = parseFloat(position.qty)
    
          if(/*diffPercent >= 0 ||*/ targetBuy < 1){
           sheet.getRange("G"+parseFloat(9+i)).setValue("skipping...")
              continue;

        }
        qty = Math.abs(qty)
    //  }
      
      if(qty == 0){
        sheet.getRange("G"+parseFloat(9+i)).setValue("Order not sent, already have desired quantity.")
      }
      else {
        var b_resp = submitOrder(sym,null,side,types.buy[i].toString().trim(),tifs.buy[i].toString().trim(),limits.buy[i].toString().trim(),stops.buy[i].toString().trim(),extendedHours,targetBuy)
                //console.log(b_resp)

       
        sheet.getRange("G"+parseFloat(9+i)).setValue(truncateJson(b_resp))
      }
    }
 }
 SpreadsheetApp.getActive().toast("Buy completed.", "⚠️ Alert"); 
  updateSheet()
}
function sell(){
    clearOrders();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders");
  var market_value = parseFloat(getAccount().long_market_value) + parseFloat(getAccount().short_market_value);
  var percent = sheet.getRange("F3").getValue();
  var rebalance = sheet.getRange("F4").getValue();
  var extendedHours = sheet.getRange("F5").getValue();
  
  
  var symbols = {
    "buy": sheet.getRange("A9:A").getValues().filter(String).map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("I9:I").getValues().filter(String).map(function(elem){return elem.toString()})
  }
  var qtys = {
    "buy": sheet.getRange("B9:B").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("J9:J").getValues().map(function(elem){return elem.toString()})
  }
  var types = {
    "buy": sheet.getRange("C9:C").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("K9:K").getValues().map(function(elem){return elem.toString()})
  }
  var tifs = {
    "buy": sheet.getRange("D9:D").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("L9:L").getValues().map(function(elem){return elem.toString()})
  }
  var limits = {
    "buy": sheet.getRange("E9:E").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("M9:M").getValues().map(function(elem){return elem.toString()})
  }
  var stops = {
    "buy": sheet.getRange("F9:F").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("N9:N").getValues().map(function(elem){return elem.toString()})
  }
    for(var i = 0; i < symbols.sell.length; i++) {
    var account = getAccount()
    market_value = parseFloat(account.long_market_value) + parseFloat(account.short_market_value);
    if(symbols.sell[i] != "") {
      sheet.getRange("O"+parseFloat(9+i)).setValue("submitting...");
      
      var qty = parseFloat(qtys.sell[i].toString().trim());
      var sym = symbols.sell[i].toString().trim();
      var side = "sell"
      var targetSell = 0.00;
      var position = getPosition(sym)

      var diffPercent = parseFloat(position.market_value)/market_value - (qty)/100
    //  if(percent) { 
        targetSell = Math.abs((parseFloat(position.market_value) - parseFloat(market_value) * (qty/100) )/((qty/100)-1))
     // }
     // if(rebalance) {
        var position_qty
        if(isNaN(parseFloat(position.qty))) position_qty = 0;
        else position_qty = parseFloat(position.qty)
        
        
        
        if(diffPercent <= 0){
           sheet.getRange("O"+parseFloat(9+i)).setValue("skipping...")
          targetSell = parseFloat(sheet.getRange("B4").getValue())*(qty/100)

        }
       
   //   }
      
      if(qty == 0){
        sheet.getRange("O"+parseFloat(9+i)).setValue("Order not sent, already have desired quantity.")
      }
      else {
        var s_resp = submitOrder(sym,null,side,types.sell[i].toString().trim(),tifs.sell[i].toString().trim(),limits.sell[i].toString().trim(),stops.sell[i].toString().trim(),extendedHours,targetSell)
        //console.log(s_resp)
        sheet.getRange("O"+parseFloat(9+i)).setValue(truncateJson(s_resp))

      }
    }
 
  }

}
function rebalance(){
 clearOrders()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders");
  var market_value = parseFloat(getAccount().long_market_value) + parseFloat(getAccount().short_market_value);
  var percent = sheet.getRange("F3").getValue();
  var rebalance = sheet.getRange("F4").getValue();
  var extendedHours = sheet.getRange("F5").getValue();
  var soldAmount = 0;
  
  var symbols = {
    "buy": sheet.getRange("A9:A").getValues().filter(String).map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("I9:I").getValues().filter(String).map(function(elem){return elem.toString()})
  }
  var qtys = {
    "buy": sheet.getRange("B9:B").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("J9:J").getValues().map(function(elem){return elem.toString()})
  }
  var types = {
    "buy": sheet.getRange("C9:C").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("K9:K").getValues().map(function(elem){return elem.toString()})
  }
  var tifs = {
    "buy": sheet.getRange("D9:D").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("L9:L").getValues().map(function(elem){return elem.toString()})
  }
  var limits = {
    "buy": sheet.getRange("E9:E").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("M9:M").getValues().map(function(elem){return elem.toString()})
  }
  var stops = {
    "buy": sheet.getRange("F9:F").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("N9:N").getValues().map(function(elem){return elem.toString()})
  }
  
  
 
  
  symbolsLength = symbols.buy.length > symbols.sell.length ? symbols.buy.length : symbols.sell.length
  for(var i = 0; i < symbols.sell.length; i++) {
    var account = getAccount()
    market_value = parseFloat(account.long_market_value) + parseFloat(account.short_market_value);
    if(symbols.sell[i] != "") {
      sheet.getRange("O"+parseFloat(9+i)).setValue("submitting...");
      
      var qty = parseFloat(qtys.sell[i].toString().trim());
      var sym = symbols.sell[i].toString().trim();
      var side = "sell"
      var targetSell = 0.00;
      var position = getPosition(sym)

      var diffPercent = parseFloat(position.market_value)/market_value - (qty)/100
    //  if(percent) { 
        targetSell = Math.abs((parseFloat(position.market_value) - parseFloat(market_value) * (qty/100) )/((qty/100)-1))
     // }
     // if(rebalance) {
        var position_qty
        if(isNaN(parseFloat(position.qty))) position_qty = 0;
        else position_qty = parseFloat(position.qty)
        
        
        
        if(diffPercent <= 0){
           sheet.getRange("O"+parseFloat(9+i)).setValue("skipping...")
              continue;

        }
       
   //   }
      
      if(qty == 0){
        sheet.getRange("O"+parseFloat(9+i)).setValue("Order not sent, already have desired quantity.")
      }
      else {
        var s_resp = submitOrder(sym,null,side,types.sell[i].toString().trim(),tifs.sell[i].toString().trim(),limits.sell[i].toString().trim(),stops.sell[i].toString().trim(),extendedHours,targetSell)
        //console.log(s_resp)
        soldAmount += s_resp.notional;
        sheet.getRange("O"+parseFloat(9+i)).setValue(truncateJson(s_resp))

      }
    }

  }
 for(var i = 0; i < symbols.buy.length; i++) {
      if(symbols.buy[i] != ""){
         var account = getAccount()
        market_value = parseFloat(account.long_market_value) + parseFloat(account.short_market_value);
      sheet.getRange("G"+parseFloat(9+i)).setValue("submitting...");
      
      var qty = parseFloat(qtys.buy[i].toString().trim());
      var sym = symbols.buy[i].toString().trim()
      var side = "buy"
      var diffPercent = parseFloat(getPosition(sym).market_value)/market_value - (qty)/100
      var targetBuy = 0.00;
      var position = getPosition(sym)


     // if(percent) {
        targetBuy = (parseFloat(position.market_value) - market_value * (qty/100) )/((qty/100)-1)
     // }
     // if(rebalance) {
        var position_qty
        if(isNaN(parseFloat(position.qty))) position_qty = 0;
        else position_qty = parseFloat(position.qty)
    
          if(diffPercent >= 0 || targetBuy > soldAmount || targetBuy < 1){
           sheet.getRange("G"+parseFloat(9+i)).setValue("skipping...")
              continue;

        }
        qty = Math.abs(qty)
    //  }
      
      if(qty == 0){
        sheet.getRange("G"+parseFloat(9+i)).setValue("Order not sent, already have desired quantity.")
      }
      else {
        var b_resp = submitOrder(sym,null,side,types.buy[i].toString().trim(),tifs.buy[i].toString().trim(),limits.buy[i].toString().trim(),stops.buy[i].toString().trim(),extendedHours,targetBuy)
                //console.log(b_resp)

        soldAmount -= b_resp.notional;
        sheet.getRange("G"+parseFloat(9+i)).setValue(truncateJson(b_resp))
      }
    }
 }
 SpreadsheetApp.getActive().toast("Rebalance completed.", "⚠️ Alert"); 
  updateSheet()
}



//https://alpaca.markets/learn/google-spreadsheet-to-manage-your-stocks-using-api/
var PositionRowStart = 17;

// Submit a request to the Alpaca API
function _request(path,params,data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account & Portfolio")
  var isPaper = sheet.getRange("E3").getValue()
  var apiKey = sheet.getRange("E4").getValue()
  var apiSecret = sheet.getRange("E5").getValue()
  if (!apiKey || !apiSecret) {
    sheet.getRange("E2").setValue("Please input an API key and secret key.")
    throw "Please input an API key and secret key."
  } else {
    sheet.getRange("E2").setValue("")
  }
  
  var headers = {
    "APCA-API-KEY-ID": apiKey,
    "APCA-API-SECRET-KEY": apiSecret,
  };
  
  var paper_live = isPaper ? "https://paper-api.alpaca.markets" : "https://api.alpaca.markets"
  var endpoint = (data ? "https://data.alpaca.markets" : paper_live);
  var options = {
    "headers": headers,
  };
  var url = endpoint + path;
  if (params) {
    if (params.qs) {
      var kv = [];
      for (var k in params.qs) {
        kv.push(k + "=" + encodeURIComponent(params.qs[k]));
      }
      url += "?" + kv.join("&");
      delete params.qs
    }
    for (var k in params) {
      options[k] = params[k];
    }
  }
  var response = UrlFetchApp.fetch(url, options);
  

  var json = response.getContentText();

  var data;
  try{
    data = JSON.parse(json);
  }
  catch(err) {
    data = err;
  }
  //console.log(data)
  return data;
}

/*
 * Alpaca API methods
 */
function getAccount() {
  return _request("/v2/account",{
    method: "GET",
    muteHttpExceptions: true
  });
}

function listOrders() {
  return _request("/v2/orders",{
    method: "GET",
    muteHttpExceptions: true
  });
}

function listPositions() {
  return _request("/v2/positions",{
    method: "GET",
    muteHttpExceptions: true
  });
}

function getPosition(sym) {
  return _request(("/v2/positions/" + sym),{
    method: "GET",
    muteHttpExceptions: true
  });
}
// uses latest trades as price point
function getPrice(sym) {
  return _request(("/v2/stocks/"+sym+"/trades/latest"), {
    method: "GET",
   
    muteHttpExceptions: true
  }, true).trade.p
}

function getAssets(){

  return _request(("/v2/assets"), {
    method: "GET",
    muteHttpExceptions: true
  }, false)

}

function clearPosition(sym) {
  return _request(('/v2/positions/' + sym), {
    method: "DELETE",
    muteHttpExceptions: true
  })
}
  
function clearPositions() {
  return _request('/v2/positions', {
    method: "DELETE",
    muteHttpExceptions: true
  })
}

function clearOrders() {
  return _request('/v2/orders', {
    method: "DELETE",
    muteHttpExceptions: true
  })
}

function listFillActivities(date) {
  qs = date ? {"date": date.toISOString()} : null
  return _request('/v2/account/activities/FILL', {
    method: "GET",
    qs: qs,
    muteHttpExceptions: true
  })
}

// Submit an order to the Alpaca API
function submitOrder(symbol, qty=null, side, type, tif, limit, stop, extendedHours, notional=null) {
  var payload = {
    symbol: symbol,
    side: side,
    notional: notional,
    qty: qty,
    type: type,
    time_in_force: tif,
    extended_hours: extendedHours,
  };
  if( qty == null){
    delete payload.qty;
  }
    if( notional == null){
    delete payload.notional;
  }
  if (limit) {
    payload.limit_price = limit;
  }
  if (stop) {
    payload.stop_price = stop;
  }
  return _request("/v2/orders", {
    method: "POST",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

// Format JSON responses for display
function truncateJson(json) {
  var out_str = "{ "
  for(var key in json) {
    if(json.hasOwnProperty(key)){
      out_str += (key + ": " + json[key] + "; ");
    }
  }
  out_str += "}"
  return out_str;
}

// Delete the order specified by the field value
function deleteOrderById() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account & Portfolio");
  var id = sheet.getRange("J9").getValue();
  
  var resp = _request(("/v2/orders/" + id),{
    method: "DELETE",
    muteHttpExceptions: true
  });
  if(resp.message && resp.message == "Empty JSON string") resp = "Order Sent";
  sheet.getRange("J10").setValue(resp);
  updateSheet();
}

// Create a new order from the Create New Orders sheet
function orderFromSheet() {
  clearOrders()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders");
  var portfolio_value = parseFloat(getAccount().portfolio_value);
  var percent = sheet.getRange("F3").getValue();
  var rebalance = sheet.getRange("F4").getValue();
  var extendedHours = sheet.getRange("F5").getValue();
  
  var symbols = {
    "buy": sheet.getRange("A9:A").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("I9:I").getValues().map(function(elem){return elem.toString()})
  }
  var qtys = {
    "buy": sheet.getRange("B9:B").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("J9:J").getValues().map(function(elem){return elem.toString()})
  }
  var types = {
    "buy": sheet.getRange("C9:C").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("K9:K").getValues().map(function(elem){return elem.toString()})
  }
  var tifs = {
    "buy": sheet.getRange("D9:D").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("L9:L").getValues().map(function(elem){return elem.toString()})
  }
  var limits = {
    "buy": sheet.getRange("E9:E").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("M9:M").getValues().map(function(elem){return elem.toString()})
  }
  var stops = {
    "buy": sheet.getRange("F9:F").getValues().map(function(elem){return elem.toString()}),
    "sell": sheet.getRange("N9:N").getValues().map(function(elem){return elem.toString()})
  }
  
   if(rebalance) {
    var positions = listPositions()
    positions.forEach(function(position){
      if(symbols.buy.indexOf(position.symbol) == -1) {
        clearPosition(position.symbol);
      }
    })
  }

  
  symbolsLength = symbols.buy.length > symbols.sell.length ? symbols.buy.length : symbols.sell.length
  for(var i = 0; i < symbolsLength; i++) {
    if(symbols.buy[i] != ""){
      sheet.getRange("G"+parseFloat(9+i)).setValue("submitting...");
      
      var qty = parseFloat(qtys.buy[i].toString().trim());
      var sym = symbols.buy[i].toString().trim()
      var side = "buy"
      if(percent) {
        qty = portfolio_value / getPrice(sym) * qty / 100

      }
      if(rebalance) {
        var position_qty
        if(isNaN(parseFloat(getPosition(sym).qty))) position_qty = 0;
        else position_qty = parseFloat(getPosition(sym).qty)
        
        qty -= position_qty
        side = (qty < 0 ? "sell" : "buy")
        qty = Math.abs(qty)
      }
      
      if(qty == 0){
        sheet.getRange("G"+parseFloat(9+i)).setValue("Order not sent, already have desired quantity.")
      }
      else {
        var b_resp = submitOrder(sym,qty,side,types.buy[i].toString().trim(),tifs.buy[i].toString().trim(),limits.buy[i].toString().trim(),stops.buy[i].toString().trim())
        sheet.getRange("G"+parseFloat(9+i)).setValue(truncateJson(b_resp))
      }
    }
    if(symbols.sell[i] != "") {
      sheet.getRange("O"+parseFloat(9+i)).setValue("submitting...");
      
      var qty = qtys.sell[i].toString().trim();
      var sym = symbols.sell[i].toString().trim();
      var side = "sell"
      if(percent) {
        qty = portfolio_value / getPrice(sym) * qty / 100
      }
      if(rebalance) {
        var position_qty
        if(isNaN(parseFloat(getPosition(sym).qty))) position_qty = 0;
        else position_qty = parseFloat(getPosition(sym).qty)
        
        qty = (-1 * qty) - position_qty
        side = (qty < 0 ? "sell" : "buy")
        qty = Math.abs(qty)
      }
      
      if(qty == 0){
        sheet.getRange("O"+parseFloat(9+i)).setValue("Order not sent, already have desired quantity.")
      }
      else {
        var s_resp = submitOrder(sym,qty,side,types.sell[i].toString().trim(),tifs.sell[i].toString().trim(),limits.sell[i].toString().trim(),stops.sell[i].toString().trim(),extendedHours)
        sheet.getRange("O"+parseFloat(9+i)).setValue(truncateJson(s_resp))
      }
    }
  }
  updateSheet()
}

// Clear existing positions from the spreadsheet so they can be updated
function wipePositions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account & Portfolio");
  var rowIdx = PositionRowStart;
  while (true) {
    var symbol = sheet.getRange("A" + rowIdx).getValue();
    if (!symbol) {
      break;
    }
    rowIdx++;
  }
  var rows = rowIdx - PositionRowStart;
  if (rows > 0) {
    sheet.deleteRows(PositionRowStart, rows);
  }
}

// Update the Open Positions & Orders Sheet
function updateSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account & Portfolio");
  var account = getAccount()
  
  sheet.getRange("B9").setValue(account.id)
  sheet.getRange("B10").setValue(account.buying_power)
  sheet.getRange("B11").setValue(account.cash)
  sheet.getRange("B12").setValue(account.portfolio_value)
  sheet.getRange("B13").setValue(account.status)
  
  sheet.getRange("B10:B12").setNumberFormat("$#,##0.00")
  
  

  // Updating orders and positions can take a bit of time - avoid trying to do it twice at once.
  if (sheet.getRange("B15").getValue() == "Updating") {
    return
  }
  sheet.getRange("B15").setValue("Updating...")
  wipePositions();
  var positions = listPositions()
  
  var endIdx = null
  if (positions.length > 0) {
    positions.sort(function(a, b) { return a.symbol < b.symbol ? -1 : 1 });
    positions.forEach(function(position, i) {
      var rowIdx = PositionRowStart + i;
      sheet.getRange("A" + rowIdx).setValue(position.symbol);
      sheet.getRange("B" + rowIdx).setValue(position.qty);
      sheet.getRange("C" + rowIdx).setValue(position.market_value);
      sheet.getRange("D" + rowIdx).setValue(position.cost_basis);
      sheet.getRange("E" + rowIdx).setValue(position.unrealized_pl);
      sheet.getRange("F" + rowIdx).setValue(position.unrealized_plpc);
      sheet.getRange("G" + rowIdx).setValue(position.current_price);
    });
    endIdx = PositionRowStart + positions.length - 1;
    sheet.getRange("B" + PositionRowStart + ":B" + endIdx).setNumberFormat("###0.00");
    sheet.getRange("C" + PositionRowStart + ":C" + endIdx).setNumberFormat("$#,##0.00");
    sheet.getRange("D" + PositionRowStart + ":D" + endIdx).setNumberFormat("$#,##0.00");
    sheet.getRange("E" + PositionRowStart + ":E" + endIdx).setNumberFormat("$#,##0.00");
    sheet.getRange("F" + PositionRowStart + ":F" + endIdx).setNumberFormat("0.00%");
    sheet.getRange("G" + PositionRowStart + ":G" + endIdx).setNumberFormat("$#,##0.00");

    sheet.getRange("C" + (endIdx + 1)).setValue("total")
    sheet.getRange("D" + (endIdx + 1)).setValue("total")
    sheet.getRange("E" + (endIdx + 1)).setValue("total")
    sheet.getRange("F" + (endIdx + 1)).setValue("average")
    sheet.getRange("G" + (endIdx + 1)).setValue("median")
    
    sheet.getRange("C" + (endIdx + 2)).setFormula("=sum(C" + PositionRowStart + ":C" + endIdx + ")")
    sheet.getRange("D" + (endIdx + 2)).setFormula("=sum(D" + PositionRowStart + ":D" + endIdx + ")")
    sheet.getRange("E" + (endIdx + 2)).setFormula("=sum(E" + PositionRowStart + ":E" + endIdx + ")")
    sheet.getRange("F" + (endIdx + 2)).setFormula("=average(F" + PositionRowStart + ":F" + endIdx + ")")
    sheet.getRange("G" + (endIdx + 2)).setFormula("=median(G" + PositionRowStart + ":G" + endIdx + ")")
  }
  sheet.getRange("B15").setValue("")
  var orders = listOrders()
  if(orders.length > 0) {
    sheet.getRange("J15").setValue("Updating...")
    orders.sort(function(a, b) { return a.symbol < b.symbol ? -1 : 1 })
    orders.forEach(function(order, i) {
      var rowIdx = PositionRowStart + i;
      var price = getPrice(order.symbol)
      var filled_qty_str = order.filled_qty + " / " + order.qty
      sheet.getRange("I" + rowIdx).setValue(order.symbol);
      sheet.getRange("J" + rowIdx).setValue(filled_qty_str);
      sheet.getRange("K" + rowIdx).setValue(order.filled_avg_price);
      sheet.getRange("L" + rowIdx).setValue(order.type);
      sheet.getRange("M" + rowIdx).setValue(order.limit_price);
      sheet.getRange("N" + rowIdx).setValue(order.stop_price);
      sheet.getRange("O" + rowIdx).setValue(price);
      sheet.getRange("P" + rowIdx).setValue(order.time_in_force);
      sheet.getRange("Q" + rowIdx).setValue(order.id);
    });
    sheet.getRange("J15").setValue("")
  }
}

// Clear existing order fills so the table can be updated
function clearFills() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("View Order Fills")
  var rowIdx = 9
  while (true) {
    var symbol = sheet.getRange("A" + rowIdx).getValue()
    if (!symbol) {
      break
    }
    rowIdx++
  }
  var rows = rowIdx - 9
  if (rows > 0) {
    sheet.deleteRows(9, rows) 
  }
}

// Update order fills table
function updateFills() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("View Order Fills")
  sheet.getRange("C7").setValue("Updating...")
  clearFills()
  var date = sheet.getRange("E4").getValue()
  var fills = listFillActivities(date)
  if (fills.length > 0) {
    var rowIdx = 9
    fills.forEach(function(fill, i) {
      sheet.getRange("A"+rowIdx).setValue(fill.symbol)
      sheet.getRange("B"+rowIdx).setValue(fill.side)
      sheet.getRange("C"+rowIdx).setValue(fill.price)
      sheet.getRange("D"+rowIdx).setValue(fill.qty)
      sheet.getRange("E"+rowIdx).setValue(fill.transaction_time)
      rowIdx++
    })
  }
  sheet.getRange("C7").setValue("")
}
