var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName("POD 2019 EPICs and Initiatives (Jira)")
var destinationSheet = ss.getSheetByName("POD")
var jiraLink = "=HYPERLINK(\"https://hootsuite.atlassian.net/browse/"
var colIssueType, colSummary, colIssueKey, colParentLink, colDescription, colPriority
var initiativesMap = []
var destinationSheetHeaderRows = 2
var productAccelerationInit = "SSF: Performant & Reliable Product Experience"
var productAccelerationDelivable = "Strong Social Foundation (SSF)"
var priorities = [{priority: "Blocker", prio: "Top 30%"}, {priority: "Critical", prio: "Mid 30%"}, {priority: "Major", prio: "Bottom 30%"}];
var effort = "X-Large"
var start = "Q3 '19"
var end = "Q4 '19"
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
  colIssueType = rawData[0].indexOf("Issue Type");
  colSummary = rawData[0].indexOf("Summary");
  colIssueKey = rawData[0].indexOf("Issue key")
  colParentLink = rawData[0].indexOf("Custom field (Parent Link)")
  colDescription = rawData[0].indexOf("Description")
  colPriority = rawData[0].indexOf("Priority")
  
  //Delete everything except top two header rows
  if(destinationSheet.getLastRow() > destinationSheetHeaderRows) {
    destinationSheet.deleteRows(destinationSheetHeaderRows+1, destinationSheet.getLastRow()-destinationSheetHeaderRows)
  }
  
  //Only proceed when all column exist
  if(colIssueType != -1 && colSummary != -1 && colIssueKey != -1 && colParentLink != -1 && colPriority != -1){
    flushInitiatives(rawData)
    flushEPICs(rawData)
    collapseBets()
  }
}

function flushInitiatives(rawData){
    var values = [];

    for(var i = 1; i < rawData.length; i++){
      if(rawData[i][colIssueType].toString().toLowerCase() == "initiative"){
        var priority = priorities.filter(function (prio) { return prio.priority === rawData[i][colPriority]})
        if(priority != undefined && priority.length == 1){
          values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")', rawData[i][colIssueKey], "", productAccelerationInit, productAccelerationDelivable, priority[0].prio, effort, "", "", "", "", start, end, owner]);
          initiativesMap.push({issuekey:rawData[i][colIssueKey], summary:rawData[i][colSummary]})
        }
      }
    }
    if(values != undefined && values.length > 0) {
      destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
      destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setFontSize(10);
      destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setFontWeight("Bold");
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
          if(parentRow > destinationSheetHeaderRows) {
            //var group = destinationSheet.getRowGroupAt(parentRow, 1)
            //Add a row to the bottom of group
            //var groupSize = group.getRange() != undefined ? group.getRange().getNumRows() : 0
            destinationSheet = destinationSheet.insertRowAfter(parentRow)
            var values = []
            var priority = priorities.filter(function (prio) { return prio.priority === rawData[i][colPriority]})
            if(priority != undefined && priority.length == 1){
              values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")', rawData[i][colIssueKey], '', productAccelerationInit, productAccelerationDelivable, priority[0].prio, effort, '', '', '', "", start, end, owner]);
              destinationSheet.getRange(parentRow+1, 1, values.length, values[0].length).setValues(values);
              destinationSheet.getRange(parentRow+1, 1, 1, values[0].length).shiftRowGroupDepth(1)
              destinationSheet.getRange(parentRow+1, 1, 1, values[0].length).setFontSize(8);
            }
          }
        }
      }
    }
    SpreadsheetApp.flush();  
}