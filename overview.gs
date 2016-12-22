function generateOverview() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  
  var overviewData = getDataValues("Overview");
  var end = 0;
  for(var i = 3; i < overviewData.length; i++) {
    if(overviewData[i][0].trim() == "No Due Date"){
      end = i + 1;
    }
  }
  
  var currentData = getDataValues("Current");
  var contextCol = currentData[0].indexOf("Context");
  var scopeCol = currentData[0].indexOf("Scope");
  var contexts = [];
  var scopes = [];
  
  sa.getSheetByName("Overview").getRange(end+1, 1, overviewData.length, overviewData[0].length).clear();
  
  for(i = 1; i<currentData.length; i++) {
    contexts.push(currentData[i][contextCol]);
    scopes.push(currentData[i][scopeCol]);
  }
  
  contexts.sort();
  scopes.sort();
  
  var contextVal = "";
  var scopeVal = "";
  var context = [];
  var scope = [];
  
  for(i = 0; i < contexts.length; i++) {
    if(contexts[i] != contextVal) {
      context.push("    " + contexts[i]);
      contextVal = contexts[i];
    }
    
    if(scopes[i] != scopeVal) {
      scope.push("    " + scopes[i]);
      scopeVal = scopes[i];
    }
  }
  
  context.sort();
  context.push("    no context");
  context.reverse();
  context.push("By Context");
  context.reverse();
  
  scope.sort();
  scope.push("    no scope");
  scope.reverse();
  scope.push("By Scope");
  scope.reverse();
  
  context = context.concat(scope);
  var complete = [];
  for(i = 0; i < context.length; i++) {
    complete.push([context[i]]);
  }
  sa.getSheetByName("Overview").getRange(end+1, 1, complete.length, 1).setValues(complete);
  
  overviewData = getDataValues("Overview");
  var countFormula = "=COUNTA(Current!E:E)-1"
  var contextFormula = "=COUNTIF(Current!H:H, \"\=\"&TRIM(A";
  var scopeFormula = "=COUNTIF(Current!I:I, \"\=\"&TRIM(A";
  var rngLen = currentData.length+1
  var noContext = "=COUNTIF(Current!H2:H" + currentData.length.toString() + ", \"\")";
  var noScope = "=COUNTIF(Current!I2:I" + currentData.length.toString() + ", \"\")";
  
  complete = [];
  var overview = sa.getSheetByName("Overview");
  var frmtRng = "";
  for(i = end; i < overviewData.length; i++) {
    var j = i+1
    if(overviewData[i][0] == "By Context") {
      var group = "context";
      complete.push([countFormula]);
      overview.getRange(i+1, 1, 1, 2).setBackgroundRGB(210, 210, 210);
      overview.getRange(i+1, 1, 1, 2).setFontWeight("bold");
      overview.getRange(i+1, 1, 1, 2).setBorder(true, true, true, true, true, true);
    } else if(overviewData[i][0] == "By Scope") {
      group = "scope";
      complete.push([countFormula]);
      overview.getRange(i+1, 1, 1, 2).setBackgroundRGB(210, 210, 210);
      overview.getRange(i+1, 1, 1, 2).setFontWeight("bold");
      overview.getRange(i+1, 1, 1, 2).setBorder(true, true, true, true, true, true);
    } else if(overviewData[i][0].trim() == "no context") {
      complete.push([noContext]);
       overview.getRange(i+1, 1).setFontWeight("bold");
      overview.getRange(i+1, 1, 1, 2).setBorder(true, true, true, true, true, true);
    } else if(overviewData[i][0].trim() == "no scope") {
      complete.push([noScope]);
      overview.getRange(i+1, 1).setFontWeight("bold");
      overview.getRange(i+1, 1, 1, 2).setBorder(true, true, true, true, true, true);
    } else if(group == "context") {
      complete.push([contextFormula + j + "))"]);
      overview.getRange(i+1, 1).setFontWeight("bold");
      overview.getRange(i+1, 1, 1, 2).setBorder(true, true, true, true, true, true);
    } else if(group == "scope"){
      complete.push([scopeFormula + j + "))"]);
      overview.getRange(i+1, 1).setFontWeight("bold");
      overview.getRange(i+1, 1, 1, 2).setBorder(true, true, true, true, true, true);
    }
  }
  
  sa.getSheetByName("Overview").getRange(end+1, 2, complete.length, 1).setFormulas(complete);
  
  var rng = sa.getSheetByName("Overview").getDataRange();
  for(i = end; i <= rng.getLasRow; i++) {
    
  }
  var helper = 1;
}
