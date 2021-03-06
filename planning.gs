var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName("POD_2019_Scope")
var destinationSheet = ss.getSheetByName("EXP")
var jiraLink = "=HYPERLINK(\"https://hootsuite.atlassian.net/browse/"
var colIssueType, colSummary, colIssueKey, colParentLink, colDescription, colPriority, colOwner
var initiativesMap = []
var destinationSheetHeaderRows = 2
var productAccelerationInit = "SSF: Performant & Reliable Product Experience"
var productAccelerationDelivable = "Strong Social Foundation (SSF)"
var priorities = [{priority: "Blocker", prio: "Top 30%"}, {priority: "Critical", prio: "Mid 30%"}, {priority: "Major", prio: "Bottom 30%"}];
var owners = [{team: "PCRE", owner: "Alister West"}, {team: "PCP", owner: "Matt Moore"}, {team: "PBTD", owner: "Matt Moore"}];

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
  colOwner = rawData[0].indexOf("Teams")

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
  var dependencyType, dependency, dependencyTLDR = ""
  for(var i = 1; i < rawData.length; i++){
    var values = []
    if(rawData[i][colIssueType].toString().toLowerCase() == "initiative"){
      values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")', rawData[i][colIssueKey], "", productAccelerationInit, productAccelerationDelivable, "", rawData[i][colEffort], "", "", "", "", rawData[i][colStart], rawData[i][colEnd], "Lei Guo"]);
      destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontSize(10);
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontWeight("Bold")
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setBackgroundColor("White")
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).shiftRowGroupDepth(-1)
    }else if(rawData[i][colIssueType].toString().toLowerCase() == "epic"){
      var owner = owners.filter(function (teams) {return teams.team === rawData[i][colOwner]});
      switch(rawData[i][colIssueKey]) {
        case "PCRE-762":
          dependencyType = "Has a Dependency"
          dependency = "Platform"
          dependencyTLDR = '=HYPERLINK("#rangeid=1667158727","Platform ! Legacy Platform: PHP 7.x upgrade")'
          break
        case "PCRE-1814":
          dependencyType = "Is Depended On"
          dependency = "P+C/Engage"
          dependencyTLDR = '=HYPERLINK("#rangeid=1117211100","Engage!Expand SLO Coverage")'
          break
        case "PBTD-1150":
          dependencyType = "Is Depended On"
          dependency = "Platform"
          dependencyTLDR = '=HYPERLINK("#rangeid=1787665534","PLAT22")'
          break
        default:
          dependencyType = ""
          dependency = ""
          dependencyTLDR = ""
          break
      }
      values.push([jiraLink + rawData[i][colIssueKey] + '", "' + rawData[i][colSummary] + '")', rawData[i][colIssueKey], "", productAccelerationInit, productAccelerationDelivable, "", rawData[i][colEffort], dependencyType, dependency, dependencyTLDR, "", rawData[i][colStart], rawData[i][colEnd], owner[0].owner]);
      destinationSheet.getRange(destinationSheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontSize(8);
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setFontWeight("Normal")
      destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).setBackgroundColor("White")
      if(destinationSheet.getRowGroupDepth(destinationSheet.getLastRow()-1) == 0) {
         destinationSheet.getRange(destinationSheet.getLastRow(), 1, 1, values[0].length).shiftRowGroupDepth(1)
      }
    }
    SpreadsheetApp.flush();  
  }
}