var ss = SpreadsheetApp.getActiveSpreadsheet();
var dependencyName = "Dept / Portfolio Dependencies"

var portfolios = ["Engage", "P+C", "Measure", "PIF", "Product Growth", "Promote", "Platform"];
var dependencies = ["Engage", "Plan+Create", "Measure", "PIF", "Product Growth", "Promote", "POD", "Platform", "Security"];

function onOpen() {
  var menuEntries = [{name: "Update Dependency", functionName: "updateDependency"}];
  ss.addMenu("Commands", menuEntries);
}

function updateDependency() {
  var sheets = ss.getSheets();
  var dependencySheet = ss.getSheetByName(dependencyName);
  if( dependencySheet != null) {
    ss.deleteSheet(dependencySheet);
  }
  dependencySheet = ss.insertSheet(dependencyName);
  
  var values = [];
  values.push(["Portfolio", "Bet", "Depend On"]);
  dependencySheet.getRange(1, 1, 1, 3).setValues(values);
  dependencySheet.getRange(1, 1, 1, 3).setFontWeight("Bold");
  dependencySheet.setFrozenRows(1);
  SpreadsheetApp.flush();    

  portfolios.forEach(findDependencies);
}

function findDependencies(portfolio) {
  var portfolioSheet = ss.getSheetByName(portfolio);
  var dependencySheet = ss.getSheetByName(dependencyName);
  //For each sheet, find the dependent column, copy value into rowdata  
  var rawData = portfolioSheet.getDataRange().getValues();
  var values = [];
    
  //Get dependency column
  var col = rawData[1].indexOf(dependencyName);
  if (col != -1) {
    for(var i = 2; i < rawData.length; i++){
      if(rawData[i][0].toString().toLowerCase() == "DEPENDENCIES (generated)".toLowerCase() || rawData[i][0].toString().toLowerCase() == "BELOW THIS LINE IS GENERATED".toLowerCase()){
        break;
      }
      if(rawData[i][col] != undefined && rawData[i][col].length > 0 && rawData[i][col].toString().toLowerCase() != "n/a"){
        values.push([portfolio, rawData[i][0], rawData[i][col]]);
      }
    }
  }
  
  if(values != undefined && values.length > 0) {
    dependencySheet.getRange(dependencySheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();  
  }
}
