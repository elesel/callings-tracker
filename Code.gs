// See https://github.com/elesel/callings-tracker
var VERSION = '0.7';
var ABOUT_URL = 'https://github.com/elesel/callings-tracker';

var NAME_PARSER_FNF = /^(.+)\s+(\S+)$/;
var NAME_PARSER_LNF = /^(\S+?),\s+(.+)$/;
var DEFERRED_ON_CHANGE = 'deferred_on_change';

// Initialize sheets
var config = {};
var sheets = {
  pendingCallings: { name: "Callings - Pending" },
  currentCallings: { name: "Callings - Current" },
  archivedCallings: { name: "Callings - Archive" },
  units: { name: "Units" },
  leaders: { name: "Leaders" },
  members: { name: "Members" },
  positions: { name: "Positions" },
  lifecycles: { name: "Lifecycles" },
  actions: { name: "Actions" }
};

// Add references to sheets
for (var property in sheets) {
  if (sheets.hasOwnProperty(property)) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheets[property].name);
    
    // Add object reference
    sheets[property].ref = sheet;
    
    // Add top row
    var topRow = sheet.getFrozenRows() + 1;
    sheets[property].topRow = topRow;
    
    // Add column map
    var map = {};
    var columns = sheet.getRange(topRow - 1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var c = 0; c < columns.length; c++) {
      var cBase1 = c + 1;
      var value = columns[c];
      if (value) {
        if (value in map) {
          // Multiple columns with the same name--handle as array
          if (Array.isArray(map[value])) {
            map[value].push(cBase1);
          } else {
            map[value] = [map[value], cBase1];
          }
        } else {
          map[value] = cBase1;
        }
      }
    }
    sheets[property].columns = map;
  }
}
Logger.log("sheets= " + JSON.stringify(sheets));

function showProgressMessage(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message);
}

function onChange(e) {
  Logger.log("In onChange()");
  handleEvent(e);
}

function onEdit(e) {
  Logger.log("In onEdit()");
  handleEvent(e);
}

function handleEvent(e) {
  Logger.log("In handleEvent()");
  Logger.log("e= " + JSON.stringify(e));
  
  // Look up sheet
  var sheet = null;
  if (e.range) {
    for (var property in sheets) {
      if (sheets.hasOwnProperty(property)) {
        if (isSameSheet(e.range.getSheet(), sheets[property].ref)) {
          sheet = sheets[property];
          break;
        }
      }
    }
  }
  
  try {
    if (! sheet) {
      Logger.log("No sheet");
      // Let's assume everything changed :(
      // TODO: Be smart about what validations to add
      showProgressMessage("Sorting configuration worksheets");
      sortUnits_();
      sortLeaders_();
      sortLifecycles_();
      addValidations();
      return;
    }
    
    // Fire appropriate update functions depending on the sheet
    // TODO: Be more selective about which validations to add
    var refreshValidations = false;
    if (sheet === sheets.pendingCallings || sheet === sheets.currentCallings) {
      updateCallingStatus(sheet, e.range.getRowIndex(), e.range.getNumRows());
    } else if (sheet === sheets.units) {
      sortUnits_();
      refreshValidations = true;
    } else if (sheet === sheets.leaders) {
      sortLeaders_();
      refreshValidations = true;
    } else if (sheet === sheets.positions) {
      refreshValidations = true;
    }
    Logger.log("Done with sheet-specific triggers");
    
    // Refresh validations if necessary
    if (refreshValidations) {
      addValidations();
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Callings')
    .addItem('Sort callings', 'sortAllCallings')
    .addItem('Organize callings', 'organizeCallings')
    .addItem('Update calling status', 'updateAllCallingStatus')
    .addItem('Print pending callings', 'printPendingCallings')
    .addItem('Format member list', 'formatMembers')
    .addItem('About', 'showAbout')
    .addToUi();
};

function getMembers_() {
  if (config['members']) {
    return config['members'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.members, "Lookup name");
  
  config['members'] = list;
  return list;
}

function sortUnits_() {
  var sheet = sheets.units.ref;
  sheet.sort(sheets.units.columns["Abbreviation"]);
  sheet.sort(sheets.units.columns["Visible"]);
}

function getUnits_() {
  if (config['units']) {
    return config['units'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.units, "Abbreviation", function(value, row, sheet){
    return row[sheet.columns["Visible"] - 1];
  });
  
  config['units'] = list;
  return list;
}

function sortLeaders_() {
  var sheet = sheets.leaders.ref;
  sheet.sort(sheets.leaders.columns["Name"]);
  sheet.sort(sheets.leaders.columns["Visible"]);
}

function getLeaders_() {
  if (config['leaders']) {
    return config['leaders'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.leaders, "Name", function(value, row, sheet){
    return row[sheet.columns["Visible"] - 1];
  });
  
  config['leaders'] = list;
  return list;
}

function getPositions_() {
  if (config['positions']) {
    return config['positions'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.positions, "Name", function(value, row, sheet){
    return row[sheet.columns["Visible"] - 1];
  });
  
  config['positions'] = list;
  return list;
}

function getPositionLifecycles_() {
  if (config['positionLifecycles']) {
    return config['positionLifecycles'];
  }
  
  // Grab copy of all data
  var sheet = sheets.positions;
  var allData = getAllData_(sheet);
  
  // Determine lookup names
  var positions = {};
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var name = row[sheet.columns["Name"] - 1];
    var lifecycle = row[sheet.columns["Lifecycle"] - 1];
    if (name && lifecycle) {
      positions[name] = lifecycle;
    }
  }
  
  config['positionLifecycles'] = positions;
  return positions;
}

function sortLifecycles_() {
  var sheet = sheets.lifecycles.ref;
  sheet.sort(sheets.lifecycles.columns["Name"]);
}

function getLifecycleNames_() {
  if (config['lifecycles']) {
    return config['lifecycles'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.lifecycles, "Name");
  
  config['lifecycles'] = list;
  return list;
}

function getLifecycleActions_() {
  if (config['lifecycleActions']) {
    return config['lifecycleActions'];
  }
  
  // Get all data
  var sheet = sheets.lifecycles;
  var allData = getAllData_(sheet);
  
  // Get/check column counts
  var columnColumns = sheet.columns["Column"];
  var actionColumns = sheet.columns["Action"];
  if (columnColumns.length != actionColumns.length) {
    throw new Error("Mismatch in number of Column and Action columns on " + sheets.lifecycles.name + " sheet");
  }
  
  // Get lifecycles
  var lifecycles = {};
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var name = row[sheet.columns["Name"] - 1];
    
    // Only get active ones
    var actions = [];
    if (name) {
      // Add default action
      // TODO: Ensure Default action column exists
      var action = row[sheet.columns["Default action"] - 1];
      actions.push({ column: null, action: action });
      
      // Add other actions
      for (var i = 0; i < columnColumns.length; i++) {
        var column = row[columnColumns[i] + 1];
        var action = row[actionColumns[i] + 1];
        if (column && action) {
          actions.push({ column: column, action: action });
        }
      }
      lifecycles[name] = actions;
    }
  }
  
  config['lifecycleActions'] = lifecycles;
  return lifecycles;
}

function getActionNames_() {
  if (config['actionNames']) {
    return config['actionNames'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.actions, "Name");
    
  config['actionNames'] = list;
  return list;
}

function getActionSheets_() {
  if (config['actionSheets']) {
    return config['actionSheets'];
  }
  
  // Grab copy of all data
  var sheet = sheets.actions;
  var allData = getAllData_(sheet);
  
  // Determine lookup names
  var actions = {};
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var name = row[sheet.columns["Name"] - 1];
    var sheetName = row[sheet.columns["Sheet"] - 1];
    if (name && sheetName) {
      // Translate sheet name into sheets object reference
      for (var property in sheets) {
        if (sheets.hasOwnProperty(property)) {
          if (sheetName == sheets[property].name) {
            actions[name] = sheets[property];
          }
        }
      }
    }
  }
  
  config['actionSheets'] = actions;
  return actions;
}

function getAllData_(sheet) {
  var startRow = sheet.topRow;
  return sheet.ref.getRange(startRow, 1, sheet.ref.getLastRow() - sheet.topRow + 1, sheet.ref.getLastColumn()).getValues();
}

function getLookupValues_(sheet, columnName, validationFunction) {
  var allData = getAllData_(sheet);
  
  // Get lookup values
  var list = [];
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var value = row[sheet.columns[columnName] - 1];
    // Validate
    if (typeof validationFunction == 'function') {
      if (validationFunction(value, row, sheet)) {
        list.push(value);
      }
    } else if (value) {
      list.push(value);
    }   
  }
  
  return list;
}

function getCallingSheetNames_() {
  var list = [];
  list.push(sheets.pendingCallings.name);
  list.push(sheets.currentCallings.name);
  list.push(sheets.archivedCallings.name);
  return list;
}

function addValidations() {
  var sheet;
  
  showProgressMessage("Refreshing validations");
  
  // Add members lists to pending callings sheet
  var membersRule = SpreadsheetApp.newDataValidation().requireValueInList(getMembers_()).setAllowInvalid(true).build();
  sheet = sheets.pendingCallings;
  sheet.ref.getRange(sheet.topRow, sheet.columns["Name"], sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(membersRule);
    
  // Add units lists to pending callings sheet
  var unitsRule = SpreadsheetApp.newDataValidation().requireValueInList(getUnits_()).setAllowInvalid(true).build();
  sheet = sheets.pendingCallings;
  sheet.ref.getRange(sheet.topRow, sheet.columns["Unit"], sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(unitsRule);
    
  // Add positions lists to pending callings sheet
  var positionsRule = SpreadsheetApp.newDataValidation().requireValueInList(getPositions_()).setAllowInvalid(true).build();
  sheet = sheets.pendingCallings;
  sheet.ref.getRange(sheet.topRow, sheet.columns["Position"], sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(positionsRule);
  
  // Add leaders lists to pending callings sheet
  var leadersRule = SpreadsheetApp.newDataValidation().requireValueInList(getLeaders_()).setAllowInvalid(true).build();
  sheet = sheets.pendingCallings;
  sheet.ref.getRange(sheet.topRow, sheet.columns["Set apart by"], sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(leadersRule);
  
  // Add lifecycles list to positions sheet
  var lifecyclesRule = SpreadsheetApp.newDataValidation().requireValueInList(getLifecycleNames_()).setAllowInvalid(true).build();
  sheet = sheets.positions;
  sheet.ref.getRange(sheet.topRow, sheet.columns["Lifecycle"], sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(lifecyclesRule);
  
  // Add positions columns lists to lifecycle sheet
  var positionsColumnsRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(sheets.pendingCallings.columns)).setAllowInvalid(true).build();
  sheet = sheets.lifecycles;
  sheet.columns['Column'].forEach(function(c){
    sheet.ref.getRange(sheet.topRow, c, sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(positionsColumnsRule);
  });
  
  // Add action names lists to lifecycle sheet
  var actionsRule = SpreadsheetApp.newDataValidation().requireValueInList(getActionNames_()).setAllowInvalid(true).build();
  sheet = sheets.lifecycles;
  sheet.columns['Action'].forEach(function(c){
    sheet.ref.getRange(sheet.topRow, c, sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(actionsRule);
  });
  
  // Add sheet name lists to actions sheet
  var callingSheetsRule = SpreadsheetApp.newDataValidation().requireValueInList(getCallingSheetNames_()).setAllowInvalid(true).build();
  sheet = sheets.actions;
  sheet.ref.getRange(sheet.topRow, sheet.columns["Sheet"], sheet.ref.getMaxRows() - sheet.topRow + 1, 1).setDataValidation(callingSheetsRule);
}

function updateAllCallingStatus() {
  // Update calling status
  [sheets.pendingCallings, sheets.currentCallings].forEach(function(sheet){
    updateCallingStatus(sheet, sheet.topRow, sheet.ref.getLastRow() - sheet.topRow + 1);
  });
}

function updateCallingStatus(sheet, startRow, numRows) {
  showProgressMessage("Updating calling status on " + sheet.name + " worksheet");
                      
  // Skip empty sheets
  if (startRow < sheet.topRow || numRows < 1) {
    return;
  }
  
  // Get positions and lifecycles
  var lifecycleActions = getLifecycleActions_();
  var positionLifecycles = getPositionLifecycles_();
  var positionNames = getPositions_();
  var actionNames = getActionNames_();
  
  // Grab copy of all data
  var allData = sheet.ref.getRange(startRow, 1, numRows, sheet.ref.getLastColumn()).getValues();
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var positionName = row[sheet.columns["Position"] - 1];
    if (positionName) {
      var lifecycle = lifecycleActions[positionLifecycles[positionName]];
      var action = 'Unknown';
      if (lifecycle) {
        lifecycle.some(function(columnAction){
          if (! columnAction.column) {
            action = columnAction.action;
          } else if (row[sheet.columns[columnAction.column] - 1]) {
            action = columnAction.action;
          } else {
            return true;
          }
        });
      }
      
      // Change values only if we have to
      var status = row[sheet.columns["Status"] - 1];
      var sid = row[sheet.columns["SID"] - 1];
      var pid = row[sheet.columns["PID"] - 1];
      var actionIndex = actionNames.indexOf(action);
      var positionIndex = positionNames.indexOf(positionName);
      if (status !== action) {
        row[sheet.columns["Status"] - 1] = action;
      }
      if (sid !== actionIndex) {
        row[sheet.columns["SID"] - 1] = actionIndex;
      }
      if (pid !== positionIndex) {
        row[sheet.columns["PID"] - 1] = positionIndex;
      }
    } else {
      row[sheet.columns["Status"] - 1] = '';
    }
  }
  
  // Write columns back
  var dataRows = allData.length;
  sheet.ref.getRange(startRow, sheet.columns["Status"], dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.columns["Status"] - 1)
  );
  sheet.ref.getRange(startRow, sheet.columns["SID"], dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.columns["SID"] - 1)
  );
  sheet.ref.getRange(startRow, sheet.columns["PID"], dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.columns["PID"] - 1)
  );
}

function organizeCallings() {
  try {
    // Sanity check the sheets
    if (sheets.pendingCallings.ref.getMaxColumns() != sheets.currentCallings.ref.getMaxColumns() || sheets.pendingCallings.ref.getMaxColumns() != sheets.archivedCallings.ref.getMaxColumns()) {
      throw new Error("Mismatch in number of columns among the callings sheets. They must be the same.");
    }
    
    // Get action to sheet mapping
    var actionSheets = getActionSheets_();
    
    // Loop through calling sheets
    var changedSheets = [];
    [sheets.pendingCallings, sheets.currentCallings].forEach(function(sheet){
      // Only continue if there's data
      var numRows = sheet.ref.getLastRow() - sheet.topRow + 1;
      if (numRows < 1) {
        return;
      }
      
      // Sort callings so that we can move contiguous rows
      updateCallingStatus(sheet, sheet.topRow, sheet.ref.getLastRow() - sheet.topRow + 1);
      sortCallings(sheet);
      
      // Iterate through all status values
      var allData = sheet.ref.getRange(sheet.topRow, sheet.columns["Status"], numRows, 1).getValues();
      var startRow = null;
      var targetSheet = null;
      var rowsMoved = 0;
      for (var r = 0; r < allData.length; r++) {
        var action = allData[r][0];
        var callingsSheet = null;
        if (action) {
          callingsSheet = actionSheets[action];
        }
          
        // Get started
        if (startRow == null) {
          startRow = r;
          targetSheet = callingsSheet;
        }
        
        // Handle changes in targetSheet
        ['previous', 'current'].forEach(function(step){
          if ((step == 'previous' && callingsSheet != targetSheet) || (step == 'current' && r == allData.length - 1)) {
            if (targetSheet != null && targetSheet != sheet) {
              // Send rows to another sheet
              var endRow = (step == 'previous' ? r - 1 : r);
              var realStartRow = startRow + sheet.topRow - rowsMoved;
              var realEndRow = endRow + sheet.topRow - rowsMoved;
              
              // Move
              Logger.log("Move row " + realStartRow + " through row " + realEndRow + " from sheet " + sheet.name + " to sheet " + targetSheet.name);
              var targetRange = moveRows(sheet, targetSheet, realStartRow, realEndRow - realStartRow + 1, targetSheet.topRow);
              if (targetRange != sheets.PendingCallings) {
                targetRange.clearDataValidations();
              }
              rowsMoved += endRow - startRow + 1;
              changedSheets.push(targetSheet);
            }
            
            // Reset to current row
            startRow = r;
            targetSheet = callingsSheet;
          }
        });
      }
    });
    changedSheets = changedSheets.filter(onlyUnique);
    
    // Sort sheets
    changedSheets.forEach(function(sheet){
      sortCallings(sheet);
    });
    
    // Add validations to pending sheet if changed
    if (sheets.pendingCallings in changedSheets) {
      addValidations();
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
}

function sortAllCallings() {
  sortCallings(sheets.pendingCallings);
  sortCallings(sheets.currentCallings);
  sortCallings(sheets.archivedCallings);
}

function sortCallings(sheet) {
  sheet.ref.sort(sheet.columns["Name"]);
  sheet.ref.sort(sheet.columns["Sustain"]);
  sheet.ref.sort(sheet.columns["PID"]);
  sheet.ref.sort(sheet.columns["SID"]);
}

function formatMembers() {
  var sheet = sheets.members;
  
  // Grab copy of all data
  var startRow = sheet.topRow;
  var allData = getAllData_(sheet);
  
  // Determine lookup names
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var fullName = row[sheet.columns["Full name"] - 1];
    var age = row[sheet.columns["Age"] - 1];
    var unit = row[sheet.columns["Unit"] - 1];
    var forcedName = row[sheet.columns["Forced name"] - 1];
    var lookupName = row[sheet.columns["Lookup name"] - 1];
    var lookupName = forcedName || reformatNameLnfToShort(fullName);
    var details = [];
    if (unit) {
      details.push(unit);
    }
    if (age) {
      details.push(age);
    }
    if (details.length > 0) {
      lookupName = lookupName + " (" + details.join(", ") + ")"; 
    }
    row[sheet.columns["Lookup name"] - 1] = lookupName;
  }
  
  // Write columns back
  var dataRows = allData.length;
  sheet.ref.getRange(startRow, sheet.columns["Lookup name"], dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.columns["Lookup name"] - 1)
  );
}

function getColumnNumber(sheet, columnName) {
  for (var r = 1; r <= sheet.getFrozenRows(); r++) {
    for (var c = 1; c <= sheet.getLastColumn(); c++) {
      if (sheet.getRange(r, c).getValue() == columnName) {
        return c;
      }
    }
  }
  return null;
}

/* Convert a provided First M. Last name into Last, First M. */
function reformatNameLNF(name) {
  if (NAME_PARSER_LNF.test(name)) {
    return name;
  }
  
  var result = NAME_PARSER_FNF.exec(name);
  if (result != null) {
    return result[2] + ", " + result[1];
  } else {
    return name;
  }
}

/* Convert a provided Last, First M. name into First M. Last */
function reformatNameFNF(name) {
  var result = NAME_PARSER_LNF.exec(name);
  if (result != null) {
    return result[2] + " " + result[1];
  } else {
    return name;
  }
}

function reformatNameLnfToShort(name) {
  var result = NAME_PARSER_LNF.exec(name);
  if (result != null) {
    var lastName = result[1];
    var firstAndMiddleNames = result[2].split(/\s+/);
    var firstName = firstAndMiddleNames.shift();
    
    if (firstAndMiddleNames) {
      var middleInitials = [];
      while (firstAndMiddleNames.length > 0) {
        var n = firstAndMiddleNames.shift();
        n.replace(/,$/, "");
        // Retain some things
        if (! n.match(/Jr\.?/)) {
          n = n.substring(0, 1) + ".";
        }
        middleInitials.push(n);
      }
      return lastName + ", " + firstName + " " + middleInitials.join(" ");
    } else {
      return lastName + ", " + firstName;
    }
  } else {
    return name;
  }
}

function onlyUnique(value, index, self) { 
    return self.indexOf(value) === index;
}

function isSameSheet(sheet1, sheet2) {
  return (sheet1 && sheet2 && sheet1.getSheetId() == sheet2.getSheetId());
}

function sliceSingleColumn(grid, startRow, numRows, columnNumber) {
  var results = [];
  var stopRow = startRow + numRows - 1;
  for (var i = startRow; i <= stopRow; i++) {
    results.push([grid[i][columnNumber]]);
  }
  return results;
}

function printPendingCallings() {
  // Get/create "temporary" document to use for the printout
  // TODO: Add support to reuse existing document
  //var doc = DocumentApp.openById('DOCUMENT_ID_GOES_HERE');
  // Create new one
  var documentName = SpreadsheetApp.getActiveSpreadsheet().getName() + "-Printable";
  var doc = DocumentApp.create(documentName);
  var body = doc.getBody();
  
  // Set page margins
  var margin = 36;
  doc.getBody()
    .setMarginTop(margin)
    .setMarginBottom(margin)
    .setMarginLeft(margin)
    .setMarginRight(margin);
  
  // Create header and footer
  var header = doc.addHeader();
  header.appendParagraph('Pending Callings - Confidential');
  header.appendParagraph('Printed at ' + formatDate(Date()));
  header.appendHorizontalRule();
  var footer = doc.addFooter();
  footer.appendHorizontalRule();
  footer.appendParagraph('Confidential');
  
  // Create styles
  var sectionHeaderStyle = {};
  sectionHeaderStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  sectionHeaderStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
  sectionHeaderStyle[DocumentApp.Attribute.BOLD] = true;
  
  // Create section for write-ins
  body.appendParagraph('Nominations').setAttributes(sectionHeaderStyle);
  var table = body.appendTable([
    ['Position', 'Name', 'Unit', 'BP', 'TR', 'SP', 'HC'],
    ['______________________________', '____________________________', '_____', '☐', '☐', '    /    /    ', '    /    /    '],
    ['______________________________', '____________________________', '_____', '☐', '☐', '    /    /    ', '    /    /    '],
    ['______________________________', '____________________________', '_____', '☐', '☐', '    /    /    ', '    /    /    '],
    ['______________________________', '____________________________', '_____', '☐', '☐', '    /    /    ', '    /    /    ']
  ]);
  setTableStyle(table);
  table.setColumnWidth(0, inchesToPoints(2.125));
  table.setColumnWidth(1, inchesToPoints(2));
  table.setColumnWidth(2, inchesToPoints(.375));
  table.setColumnWidth(3, inchesToPoints(.25));
  table.setColumnWidth(4, inchesToPoints(.25));
  table.setColumnWidth(5, inchesToPoints(.675));
  table.setColumnWidth(6, inchesToPoints(.675));
  centerTableColumn(table, 3);
  centerTableColumn(table, 4);
  centerTableColumn(table, 5);
  centerTableColumn(table, 6);
  
  // Sort pending sheet
  var sheet = sheets.pendingCallings;
  sortCallings(sheets.pendingCallings);
  
  // Get data and iterate through it
  var lastAction = null;
  var table = null;
  var allData = sheet.ref.getRange(sheet.topRow, 1, sheet.ref.getLastRow() - sheet.topRow + 1, sheet.ref.getLastColumn()).getValues();
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var action = row[sheet.columns["Status"] - 1];
    if (action != lastAction) {
      // Start new table
      body.appendParagraph(action).setAttributes(sectionHeaderStyle);
      table = body.appendTable([
        ['Position', 'Name', 'Unit', 'BP', 'TR', 'SP', 'HC', 'Interview', 'Sustain', 'Set Apart']
      ]);
      setTableStyle(table);
      table.setColumnWidth(0, inchesToPoints(1.75));
      table.setColumnWidth(1, inchesToPoints(1.375));
      table.setColumnWidth(2, inchesToPoints(.375));
      table.setColumnWidth(3, inchesToPoints(.25));
      table.setColumnWidth(4, inchesToPoints(.25));
      table.setColumnWidth(5, inchesToPoints(.675));
      table.setColumnWidth(6, inchesToPoints(.675));
      table.setColumnWidth(7, inchesToPoints(.675));
      table.setColumnWidth(8, inchesToPoints(.675));
      table.setColumnWidth(9, inchesToPoints(.675));
      centerTableColumn(table, 3);
      centerTableColumn(table, 4);
      centerTableColumn(table, 5);
      centerTableColumn(table, 6);
      centerTableColumn(table, 7);
      centerTableColumn(table, 8);
      centerTableColumn(table, 9);
    }
    // Add row to the table
    var tableRow = table.appendTableRow();
    var position = row[sheet.columns["Position"] - 1] || '________________________';
    var name = row[sheet.columns["Name"] - 1] || '___________________';
    var unit = row[sheet.columns["Unit"] - 1] || '_____';
    var bishop = (row[sheet.columns["Bishop"] - 1] ? '☒' : '☐');
    var templeRecommend = row[sheet.columns["TR"] - 1] || '☐';
    var stakePresidency = formatDate(row[sheet.columns["Stake presidency"] - 1]) || '    /    /    ';
    var highCouncil = formatDate(row[sheet.columns["High council"] - 1]) || '    /    /    ';
    var interview = formatDate(row[sheet.columns["Interview"] - 1]) || '    /    /    ';
    var sustain = formatDate(row[sheet.columns["Sustain"] - 1]) || '    /    /    ';
    var setApart = formatDate(row[sheet.columns["Set Apart"] - 1]) || '    /    /    ';
    tableRow.appendTableCell(position);
    tableRow.appendTableCell(name);
    tableRow.appendTableCell(unit);
    tableRow.appendTableCell(bishop);
    tableRow.appendTableCell(templeRecommend);
    tableRow.appendTableCell(stakePresidency);
    tableRow.appendTableCell(highCouncil);
    tableRow.appendTableCell(interview);
    tableRow.appendTableCell(sustain);
    tableRow.appendTableCell(setApart);
    setTableRowStyle(tableRow);
    centerTableRowColumn(tableRow, 3);
    centerTableRowColumn(tableRow, 4);
    centerTableRowColumn(tableRow, 5);
    centerTableRowColumn(tableRow, 6);
    centerTableRowColumn(tableRow, 7);
    centerTableRowColumn(tableRow, 8);
    centerTableRowColumn(tableRow, 9);
    
    lastAction = action;
  }
  
  // Send to user
  doc.saveAndClose();
  var pdfUrl = 'https://docs.google.com/document/d/' + doc.getId() + '/export?format=pdf';
  SpreadsheetApp.getUi().alert('Printable version is located at ' + pdfUrl);
  
  // Delete
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}

function setTableStyle(table) {
  var gridHeaderStyle = {};
  gridHeaderStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  gridHeaderStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  gridHeaderStyle[DocumentApp.Attribute.BOLD] = true;
  
  table.setBorderWidth(0);

  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    setTableRowStyle(row);
    if (r == 0) {
      table.getRow(0).editAsText().setAttributes(gridHeaderStyle);
    }
  }
}

function setTableRowStyle(row) {
  var gridNormalStyle = {};
  gridNormalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  gridNormalStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  gridNormalStyle[DocumentApp.Attribute.BOLD] = false;
  
  row.editAsText().setAttributes(gridNormalStyle);
  row.setMinimumHeight(inchesToPoints(.25));
  for (var c = 0; c < row.getNumCells(); c++) {
    var cell = row.getCell(c);
    cell.setPaddingLeft(0);
    cell.setPaddingRight(0);
    cell.setPaddingTop(0);
    cell.setPaddingBottom(0);
  }
}

function centerTableColumn(table, columnNumber) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    centerTableRowColumn(row, columnNumber);
  }
}

function centerTableRowColumn(row, columnNumber) {
  var cell = row.getCell(columnNumber);
  cell.getChild(0).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}
  
function inchesToPoints(inches) {
  return inches * 72;
}

function formatDate(date) {
  if (date instanceof Date) {
    return (date.getMonth() + 1).toString() + '/' + date.getDate().toString() + '/' + date.getFullYear().toString().substr(2, 2);
  } else {
    return date;
  }
}

function moveRows(sourceSheet, targetSheet, startRow, numRows, targetRow) {
  var numColumns = sourceSheet.ref.getMaxColumns();
  var sourceRange = sourceSheet.ref.getRange(startRow, 1, numRows, numColumns);
  
  targetSheet.ref.insertRows(targetSheet.topRow, numRows);
  var targetRange = targetSheet.ref.getRange(targetSheet.topRow, 1, numRows, numColumns);
  sourceRange.moveTo(targetRange);
  sourceSheet.ref.deleteRows(startRow, numRows);
  
  return targetRange;
}

function showAbout() {
  SpreadsheetApp.getUi().alert('callings-tracker v' + VERSION + '\n' + ABOUT_URL);
}
