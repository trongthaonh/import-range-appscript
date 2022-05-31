/* Copyright (C) 2022 Thao Nguyen (https://github.com/trongthaonh) - All Rights Reserved */


function onOpen() {
  let spreadsheet = SpreadsheetApp.getActive();
  let menuItems = [
    {name: 'New project', functionName: 'MAKEACOPY'},
    {name: 'Fetch data', functionName: 'MULTIPLERANGE'}
  ];
  spreadsheet.addMenu('AWS', menuItems);
}


/**
 * Clone template-servers-info file when creating a new project
 */
function MAKEACOPY(){
  var destFolder = DriveApp.getFolderById("16eX5JpU_ug958dA4tKBW0USP5de7b8JM"); // "Project List" folder
  output = DriveApp.getFileById("1GlK1TIQP5wPr82l9LEchmxRFJUZrvGJCH-fb8IO6h5A").makeCopy("[New project name]-servers-info", destFolder);

  insertToList(output.getUrl())
  openUrl(output.getUrl());
}


/**
 * Auto generate the script for importing all projects server info into one sheet.
 */
function MULTIPLERANGE() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sourceSheet = ss.getSheetByName('Links');
  let targetSheetName = "Overall (Auto updates, don't edit)";

  // Get range from the top-left cell at (2,2) with two columns, and the number of rows is the last row's index - 1
  let lastRow = sourceSheet.getLastRow();
  let sourceRange = sourceSheet.getRange(2,2,lastRow-1,2);
  let data = sourceRange.getValues();

  // Build import multiple range formula.
  let importRangeArray = [];
  data.forEach(function(row) {
   importRangeArray.push('IMPORTRANGE("' + row[1] + '","Info!B2:L50")');
  });

  // Set formula to the selected cell
  let formula = '=QUERY({' + importRangeArray.join(';') + '},"select * where Col1 is not null")';
  let targetSheet = ss.getSheetByName(targetSheetName);
  let cell = targetSheet.getRange("B2");
  cell.setFormula(formula);

  // Logger.log(formula);
}

/**
 * Insert a new file to List of projects ("Links" sheet).
 */
function insertToList(url){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName('Links');

  // Set the order number to the first cell of the new last row.
  targetSheet.getRange(targetSheet.getLastRow() + 1, 1).setFormula("=row()-1");

  // Insert name of project and url to the last row.
  // getRange(targetSheet.getLastRow(), 2, 1, 2) => getRange(row, column, numRows, numColumns)
  targetSheet.getRange(targetSheet.getLastRow(), 2, 1, 2).setValues([["New project name", url]]);

  // Insert formula to get the permission from new created file. 
  // You can remove it after approving the permisison to access this file.
  let formula = '=IMPORTRANGE("' + url + '","Info!B2:L50")';
  targetSheet.getRange(targetSheet.getLastRow(), 4).setFormula(formula);
}


/**
 * Open a URL in a new tab.
 */
function openUrl(url){
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}
