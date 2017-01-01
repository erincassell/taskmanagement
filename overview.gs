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

function summarizeMonth() {
  moveComplete();
  
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var complete = sa.getSheetByName("Complete");
  var deleted = sa.getSheetByName("Deleted");
  var overview = sa.getSheetByName("Overview");
  
  var newMont = new Date();
  var newMonth = new Date(2017, 00, 01);
  if(newMonth.getMonth() == 0) {
    var month = 11;
    var year = newMonth.getFullYear() - 1;
  } else {
    var month = newMonth.getMonth();
    var year = newMonth.getFullYear();
  }
  
  var completeData = complete.getDataRange().getValues();
  var deletedData = deleted.getDataRange().getValues();
  var overviewData = overview.getDataRange().getValues();
  
  var dueCol = completeData[0].indexOf("Due");
  var doneCol = completeData[0].indexOf("Date Completed");
  var countedCol = completeData[0].indexOf("Counted");
  var completeLength = completeData[0].length;
  
  completeData.splice(0, 1);
  
  var i = 0;
  var totalComplete = 0;
  var totalDeleted = 0;
  while(i < completeData.length) {
    var helper = completeData[i];
    if(completeData[i][countedCol] == "" && (completeData[i][dueCol] != "" || completeData[i][doneCol] != "")) {
      if(completeData[i][dueCol] == "") {
        completeData[i][dueCol] = new Date(1970, 0, 1);
      }
      
      if(completeData[i][doneCol] == "") {
        completeData[i][doneCol] = new Date(1970, 0, 1);
      }
                     
      if(completeData[i][dueCol].getMonth() == month || completeData[i][doneCol].getMonth() == month) {
        totalComplete++;
      }
    completeData[i][countedCol] = "X";
    }
    i++;
  }
  
  dueCol = deletedData[0].indexOf("Due");
  doneCol = deletedData[0].indexOf("Date Deleted");
  countedCol = deletedData[0].indexOf("Counted");
  var deleteLength = deletedData[0].length;

  deletedData.splice(0, 1);

  var i = 0;
  while(i < deletedData.length) {
    var helper = deletedData[i];
    if(deletedData[i][countedCol] == "" && (deletedData[i][dueCol] != "" || deletedData[i][doneCol] != "")) {
      if(deletedData[i][dueCol] == "") {
        deletedData[i][dueCol] = new Date(1970, 0, 1);
      }
      
      if(deletedData[i][doneCol] == "") {
        deletedData[i][doneCol] = new Date(1970, 0, 1);
      }
                     
      if(deletedData[i][dueCol].getMonth() == month || deletedData[i][doneCol].getMonth() == month) {
        totalDeleted++;
      }
    deletedData[i][countedCol] = "X";
    }
    i++;
  }

  var monthCol = overviewData[3].indexOf("Month");
  var compCol = overviewData[3].indexOf("Completed");
  var delCol = overviewData[3].indexOf("Deleted");
  
  var putData = [[(month + 1).toString() + "/" + year.toString(), totalComplete, totalDeleted]];

  var tableData = overview.getRange(4, monthCol+1, 50, 3).getValues();
  for(var j = 0; j < tableData.length; j++) {
    if(tableData[j][0] == 0) {
      break;
    }
  }
  
  overview.getRange(j+4, monthCol+1, 1, 3).setValues(putData);
  overview.getRange(j+4, monthCol+1, 1, 3).setBorder(true, true, true, true, true, true);
  
  complete.getRange(2, 1, completeData.length, completeLength).setValues(completeData);
  deleted.getRange(2, 1, deletedData.length, deleteLength).setValues(deletedData);
}
