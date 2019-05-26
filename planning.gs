var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName("POD 2019 EPICs and Initiatives (Jira)")
var destinationSheet = ss.getSheetByName("POD")
var jiraLink = "=HYPERLINK(\"https://hootsuite.atlassian.net/browse/"
var colIssueType, colSummary, colIssueKey, colParentLink
var initiativesMap = []
var destinationSheetHeaderRows = 2

function onOpen() {
  var menuEntries = [{name: "Sync Bets", functionName: "syncBets"}, {name: "Collapse Bets", functionName: "collapseBets"}, {name: "Expand Bets", functionName: "expandBets"}];
  ss.addMenu("Commands", menuEntries);
}

function syncBets() {
  var sheets = ss.getSheets();
  if( sourceSheet == null || destinationSheet == null) {
    exit;
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
  colIssueType = rawData[0].indexOf("Issue Type");
  colSummary = rawData[0].indexOf("Summary");
  colIssueKey = rawData[0].indexOf("Issue key")
  colParentLink = rawData[0].indexOf("Custom field (Parent Link)")
  
  //Delete everything except top two header rows
  if(destinationSheet.getLastRow() > destinationSheetHeaderRows) {
    destinationSheet.deleteRows(3, destinationSheet.getLastRow()-destinationSheetHeaderRows)
  }
  
  //Only proceed when all column exist
  if(colIssueType != -1 && colSummary != -1 && colIssueKey != -1 && colParentLink != -1){
    flushInitiatives(rawData)
    flushEPICs(rawData)
    collapseBets()
  }
}

function flushInitiatives(rawData){
    var values = [];

    for(var i = 1; i < rawData.length; i++){
      if(rawData[i][colIssueType].toString().toLowerCase() == "initiative"){
        values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")']);
        initiativesMap.push({issuekey:rawData[i][colIssueKey], summary:rawData[i][colSummary]})
      }
    }
    if(values != undefined && values.length > 0) {
      destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
      SpreadsheetApp.flush();  
    }
}

function flushEPICs(rawData){

    for(var i = 1; i < rawData.length; i++){
      if(rawData[i][colIssueType].toString().toLowerCase() == "epic"){
        destinationData = destinationSheet.getDataRange().getValues()
        parentRow = -1
        if ( rawData[i][colParentLink] != undefined && rawData[i][colParentLink].length > 0){
          //Insert after parent initiative and start a group
          for(var j = destinationSheetHeaderRows; j<destinationData.length;j++){
            var parentInit = initiativesMap.filter(function (init){return init.issuekey === rawData[i][colParentLink]})
            if(parentInit != undefined && parentInit.length == 1 && destinationData[j][0] == parentInit[0].summary ){ 
              parentRow = j+1;
              break
            }
          }
          destinationSheet = destinationSheet.insertRowAfter(parentRow)
          var values = []
          values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")']);
          destinationSheet.getRange(parentRow+1, 1, 1, 1).setValues(values);
          destinationSheet.getRange(parentRow+1, 1, 1, 1).shiftRowGroupDepth(1)
        }
      }
    }
    SpreadsheetApp.flush();  
}
