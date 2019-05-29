var ss = SpreadsheetApp.getActiveSpreadsheet();
var dependencyName = ["Dept / Portfolio Dependency", "Dept / Portfolio Dependencies"]
var priority = "Priority for Portfolio"
var effort = "Effort / Estimate"
var dependencyType = ["Bet Type", "Bet Dependency Type"]
var dependency = "Dependency Tl;Dr (or details or Bet IDs for dependent bets)"
var unlikely = ["Unlikely in H2?", "Unlikely in H2"]

var portfolios = ["Engage", "P+C", "Measure", "PIF", "Product Growth", "Promote", "Platform"];
var dependencies = ["Engage", "Plan+Create", "Measure", "PIF", "Product Growth", "Promote", "POD", "Platform", "Security"];

function onOpen() {
  var menuEntries = [{name: "Update Dependency", functionName: "updateDependency"}];
  ss.addMenu("Commands", menuEntries);
}

function updateDependency() {
  var sheets = ss.getSheets();
  var dependencySheet = ss.getSheetByName(dependencyName[0]);
  if( dependencySheet != null) {
    ss.deleteSheet(dependencySheet);
  }
  dependencySheet = ss.insertSheet(dependencyName[0]);
  
  var values = [];
  values.push(["Portfolio", "Bet", "Depend On", priority, "Dependency Type", dependency, "Unlikely in H2"]);
  dependencySheet.getRange(1, 1, 1, values[0].length).setValues(values);
  dependencySheet.getRange(1, 1, 1, values[0].length).setFontWeight("Bold");
  dependencySheet.setFrozenRows(1);
  SpreadsheetApp.flush();    

  portfolios.forEach(findDependencies);
}

function findDependencies(portfolio) {
  var portfolioSheet = ss.getSheetByName(portfolio);
  var dependencySheet = ss.getSheetByName(dependencyName[0]);
  //For each sheet, find the dependent column, copy value into rowdata  
  var rawData = portfolioSheet.getDataRange().getValues();
  var values = [];
    
  //Get dependency column
  var colDependencyName = rawData[1].indexOf(dependencyName[0]) || rawData[1].indexOf(dependencyName[1]);
  var colPriority = rawData[1].indexOf(priority)
  var colDependencyType = rawData[1].indexOf(dependencyType[0]) || rawData[1].indexOf(dependencyType[1])
  var colDependency = rawData[1].indexOf(dependency)
  var colUnlikely = rawData[1].indexOf(unlikely[0]) || rawData[1].indexOf(unlikely[1])
  
  if (colDependencyName != -1 && colPriority != -1 && colDependencyType != -1 && colDependency != -1 && colUnlikely != -1) {
    for(var i = 2; i < rawData.length; i++){
      if(rawData[i][0].toString().toLowerCase() == "DEPENDENCIES (generated)".toLowerCase() || rawData[i][0].toString().toLowerCase() == "BELOW THIS LINE IS GENERATED".toLowerCase()){
        break;
      }
      if(rawData[i][colDependencyName] != undefined && rawData[i][colDependencyName].length > 0 && rawData[i][colDependencyName].toString().toLowerCase() != "n/a"){
        values.push([portfolio, rawData[i][0], rawData[i][colDependencyName], rawData[i][colPriority], rawData[i][colDependencyType], rawData[i][colDependency], rawData[i][colUnlikely]]);
      }
    }
  }else{
    SpreadsheetApp.getUi().alert("One of the required column is missing");
  }
  
  if(values != undefined && values.length > 0) {
    dependencySheet.getRange(dependencySheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();  
  }
}
