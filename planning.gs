var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName("POD_2019_Scope")
var destinationSheet = ss.getSheetByName("EXP")
var jiraLink = "=HYPERLINK(\"https://hootsuite.atlassian.net/browse/"
var colIssueType, colSummary, colIssueKey, colParentLink, colDescription, colPriority
var initiativesMap = []
var destinationSheetHeaderRows = 2
var productAccelerationInit = "SSF: Performant & Reliable Product Experience"
var productAccelerationDelivable = "Strong Social Foundation (SSF)"
var priorities = [{priority: "Blocker", prio: "Top 30%"}, {priority: "Critical", prio: "Mid 30%"}, {priority: "Major", prio: "Bottom 30%"}];
var priorities = [{team: "PCRE", owner: "Alister West"}, {team: "PCP", prio: "Matt Moore"}, {team: "PBTD", prio: "Matt Moore"}];
var owner = "Lei Guo"

function onOpen() {
  var menuEntries = [{name: "Sync Bets", functionName: "syncBets"}, {name: "Collapse Bets", functionName: "collapseBets"}, {name: "Expand Bets", functionName: "expandBets"}];
  ss.addMenu("Commands", menuEntries);
}

function syncBets() {
  var sheets = ss.getSheets();
  if( sourceSheet == null || destinationSheet == null) {
    SpreadsheetApp.getUi().alert("No source sheet present!");
    return
  }

  //Loop through source sheet and populate the destination sheet
  populateDestinationSheet();  
}

function collapseBets (){
  destinationSheet.collapseAllRowGroups()
}

function expandBets() {
  destinationSheet.expandAllRowGroups()
}

function populateDestinationSheet() {
  var rawData = sourceSheet.getDataRange().getValues();
    
  //Get dependency column
  colIssueType = rawData[0].indexOf("Hierarchy");
  colSummary = rawData[0].indexOf("Title");
  colIssueKey = rawData[0].indexOf("Issue key")
  colEffort = rawData[0].indexOf("Story Points")
  colStart = rawData[0].indexOf("Scheduled start")
  colEnd = rawData[0].indexOf("Scheduled end")
  
  //Delete everything except top two header rows
  if(destinationSheet.getLastRow() > destinationSheetHeaderRows) {
    destinationSheet.deleteRows(destinationSheetHeaderRows+1, destinationSheet.getLastRow()-destinationSheetHeaderRows)
  }
  
  //Only proceed when all column exist
  if(colIssueType != -1 && colSummary != -1 && colIssueKey != -1 && colEffort != -1 && colStart != -1 && colEnd != -1){
    flushInitiatives(rawData)
    destinationSheet.setFrozenColumns(1);
    collapseBets()
  }
}

function flushInitiatives(rawData){
  for(var i = 1; i < rawData.length; i++){
    var values = []
    values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")', rawData[i][colIssueKey], "", productAccelerationInit, productAccelerationDelivable, "", rawData[i][colEffort], "", "", "", "", rawData[i][colStart], rawData[i][colEnd], owner]);
    destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
    if(rawData[i][colIssueType].toString().toLowerCase() == "initiative"){
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontSize(10);
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontWeight("Bold")
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).shiftRowGroupDepth(-1)
    }else if(rawData[i][colIssueType].toString().toLowerCase() == "epic"){
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontSize(8);
      if(destinationSheet.getRowGroupDepth(destinationSheet.getLastRow()-1) == 0) {
         destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).shiftRowGroupDepth(1)
      }
    }
    SpreadsheetApp.flush();  
  }
}