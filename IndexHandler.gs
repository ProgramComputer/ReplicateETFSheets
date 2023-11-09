    
function writeCSVDataToSheet(data) {
  var name = data.getBlob().getName()
  var contents = Utilities.parseCsv(data);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var itt = ss.getSheetByName(name +' Holdings')
  itt.clear();
   if (!itt) {

   ss.insertSheet(name +' Holdings');
}
 
  ss.getSheetByName(name +' Holdings').getRange(1, 1, contents.length, contents[0].length).setValues(contents);
  return ss.getSheetByName(name +' Holdings').getName();
}

//add parameters for custom urls such as isCsv and url 
function getHoldings(){
var url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create New Orders").getRange("F2").getValue()


var  response = UrlFetchApp.fetch(url);

  
 var sheetName = response.getBlob().getName();
  if(response.getBlob().getName().match(/csv/) != null){
   writeCSVDataToSheet(response);
  SpreadsheetApp.getActive().toast("The CSV file was successfully imported into " + sheetName + ".", "⚠️ Alert"); 

  }
  else if(response.getBlob().getName().match(/xlsx/) != null){
    writeXLSVDataToSheet(response)
  SpreadsheetApp.getActive().toast("The XLSV file was successfully imported into " + sheetName + ".", "⚠️ Alert"); 
  }
  else
  {
      SpreadsheetApp.getActive().toast("The CSV file was unsuccussfully imported.", "⚠️ Alert"); 

  }

return response.getBlob().getName()
}


function writeXLSVDataToSheet(data) {    
  

  const excelFile = data.getBlob();
  

 
  let config = {
    title: excelFile.getName(),
    parent: "Direct Indexing",
    mimeType: MimeType.GOOGLE_SHEETS
  };
  let spreadsheet = Drive.Files.insert(config, excelFile);
 var source =  SpreadsheetApp.openById(spreadsheet.getId()).getActiveSheet();
   var itt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(excelFile.getName()+" Holdings")
   itt.clear()
   if (itt) {

   SpreadsheetApp.getActiveSpreadsheet().deleteSheet(itt);
}
 const destination = SpreadsheetApp.getActiveSpreadsheet();

  const newSheet =  source.copyTo(destination) 
 newSheet.setName(excelFile.getName()+" Holdings")
  Drive.Files.trash(spreadsheet.id)
}