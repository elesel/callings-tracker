/*
TODO
Speed up organizeCallings()
Print pending callings
Add members lookup
Switch from "Action" term to "Status" in code
Handle renaming positions, lifecycles
Highlight next action, and/or gray out irrelevant actions
Add check for column existence
*/

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
    sheets[property].topRow = sheet.getFrozenRows() + 1;
    
    // Add column map
    var map = {};
    var r = sheet.getFrozenRows();
    for (var c = 1; c <= sheet.getLastColumn(); c++) {
      var range = sheet.getRange(r, c);
      if (! range.isBlank()) {
        var parentValues = [];
        for (var pr = 1; pr < r; pr++) {
          var parentRange = sheet.getRange(pr, c);
          var parentValue = parentRange.getValue();
        }
            
        var value = range.getValue();
        if (value in map) {
          // Multiple columns with the same name--handle as array
          if (Array.isArray(map[value])) {
            map[value].push(c);
          } else {
            map[value] = [map[value], c];
          }
        } else {
          map[value] = c;
        }
      }
    }
    sheets[property].columns = map;
  }
}
Logger.log("sheets= " + JSON.stringify(sheets));

/*
function onChange(e) {
  CacheService.getDocumentCache().put(DEFERRED_ON_CHANGE, true);
}

function onEdit(e) {
  var cache = CacheService.getDocumentCache();
  if (cache.get(DEFERRED_ON_CHANGE)) {
    cache.remove(DEFERRED_ON_CHANGE);
  }
}
*/

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
      sortUnits_();
      sortLeaders_();
      sortPositions_();
      sortLifecycles_();
      addValidations();
      //updateAllCallingStatus();
      return;
    }
    
    // Fire appropriate update functions depending on the sheet
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
      sortPositions_();
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


/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Callings')
    .addItem('Sort callings', 'sortAllCallings')
    .addItem('Organize callings', 'organizeCallings')
    .addItem('Update calling status', 'updateAllCallingStatus')
    .addItem('Print pending callings', 'printPendingCallings')
    .addItem('Download members list', 'downloadMembers')
    .addToUi();
};

function sortUnits_() {
  var sheet = sheets.units.ref;
  sheet.sort(sheets.units.columns["Abbreviation"]);
  sheet.sort(sheets.units.columns["Visible"]);
}

function getUnits_() {
  if (config['units']) {
    return config['units'];
  }
  
  var sheet = sheets.units.ref;
  var numRows = sheet.getLastRow();
  
  var list = [];
  for (var i = sheets.units.topRow; i <= numRows; i++) {
    // Only get active ones
    if (! sheet.getRange(i, sheets.units.columns["Visible"]).isBlank()) {
      list.push(sheet.getRange(i, sheets.units.columns["Abbreviation"]).getValue());
    }
  }
  
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
  
  var sheet = sheets.leaders.ref;
  var numRows = sheet.getLastRow();
  
  var list = [];
  for (var i = sheets.leaders.topRow; i <= numRows; i++) {
    // Only get active ones
    if (! sheet.getRange(i, sheets.leaders.columns["Visible"]).isBlank()) {
      list.push(sheet.getRange(i, sheets.leaders.columns["Name"]).getValue());
    }
  }
  
  config['leaders'] = list;
  return list;
}

function sortPositions_() {
  var sheet = sheets.positions.ref;
  sheet.sort(sheets.positions.columns["Visible"]);
}

function getPositions_() {
  if (config['positions']) {
    return config['positions'];
  }
  
  var sheet = sheets.positions.ref;
  var numRows = sheet.getLastRow();
  
  var list = [];
  for (var i = sheets.positions.topRow; i <= numRows; i++) {
    // Only get active ones
    if (! sheet.getRange(i, sheets.positions.columns["Visible"]).isBlank()) {
      list.push(sheet.getRange(i, sheets.positions.columns["Name"]).getValue());
    }
  }
  
  config['positions'] = list;
  return list;
}

function getPositionLifecycles_() {
  if (config['positionLifecycles']) {
    return config['positionLifecycles'];
  }
  
  var sheet = sheets.positions.ref;
  var numRows = sheet.getLastRow();
  
  var positions = {};
  for (var i = sheets.positions.topRow; i <= numRows; i++) {
    var nameRange = sheet.getRange(i, sheets.positions.columns["Name"]);
    var lifecycleRange = sheet.getRange(i, sheets.positions.columns["Lifecycle"]);
    if (! (nameRange.isBlank() || lifecycleRange.isBlank())) {
      positions[nameRange.getValue()] = lifecycleRange.getValue();
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
  var sheet = sheets.lifecycles.ref;
  var numRows = sheet.getLastRow();
  
  var list = [];
  for (var i = sheets.lifecycles.topRow; i <= numRows; i++) {
    // Only get active ones
    var range = sheet.getRange(i, sheets.lifecycles.columns["Name"]);
    if (! range.isBlank()) {
      list.push(range.getValue());
    }
  }
  
  return list;
}

function getLifecycleActions_() {
  if (config['lifecycleActions']) {
    return config['lifecycleActions'];
  }
  
  var sheet = sheets.lifecycles.ref;
  var numRows = sheet.getLastRow();
  var numCols = sheet.getLastColumn();
  
  var lifecycles = {};
  for (var r = sheets.lifecycles.topRow; r <= numRows; r++) {
    // Only get active ones
    var range = sheet.getRange(r, sheets.lifecycles.columns["Name"]);
    var actions = [];
    if (! range.isBlank()) {
      // Add default action
      // TODO: Ensure Default action column exists
      var c = sheets.lifecycles.columns["Default action"];
      var action = sheet.getRange(r, c).getValue();
      actions.push({ column: null, action: action });
      
      // Add other actions
      var columnColumns = sheets.lifecycles.columns["Column"];
      var actionColumns = sheets.lifecycles.columns["Action"];
      if (columnColumns.length != actionColumns.length) {
        throw new Error("Mismatch in number of Column and Action columns on " + sheets.lifecycles.name + " sheet");
      }
      for (var i = 0; i < columnColumns.length; i++) {
        var column = sheet.getRange(r, columnColumns[i]).getValue();
        var action = sheet.getRange(r, actionColumns[i]).getValue();
        if (column && action) {
          actions.push({ column: column, action: action });
        }
      }
      lifecycles[range.getValue()] = actions;
    }
  }
  
  config['lifecycleActions'] = lifecycles;
  return lifecycles;
}

function getActionNames_() {
  if (config['actionNames']) {
    return config['actionNames'];
  }
  
  var sheet = sheets.actions.ref;
  var numRows = sheet.getLastRow();
  
  var list = [];
  for (var i = sheets.actions.topRow; i <= numRows; i++) {
    // Only get active ones
    var range = sheet.getRange(i, sheets.actions.columns["Name"]);
    if (! range.isBlank()) {
      list.push(range.getValue());
    }
  }
  
  config['actionNames'] = list;
  return list;
}

function getActionSheets_() {
  if (config['actionSheets']) {
    return config['actionSheets'];
  }
  
  var sheet = sheets.actions.ref;
  var numRows = sheet.getLastRow();
  
  var actions = {};
  for (var i = sheets.actions.topRow; i <= numRows; i++) {
    // Only get active ones
    var nameRange = sheet.getRange(i, sheets.actions.columns["Name"]);
    var sheetRange = sheet.getRange(i, sheets.actions.columns["Sheet"]);
    if (! (nameRange.isBlank() || sheetRange.isBlank())) {
      var sheetName = sheetRange.getValue();
      // Translate sheet name into sheets object reference
      for (var property in sheets) {
        if (sheets.hasOwnProperty(property)) {
          if (sheetName == sheets[property].name) {
            actions[nameRange.getValue()] = sheets[property];
          }
        }
      }
    }
  }
  
  config['actionSheets'] = actions;
  return actions;
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
      if (status != action) {
        row[sheet.columns["Status"] - 1] = action;
      }
      if (sid != actionIndex) {
        row[sheet.columns["SID"] - 1] = actionIndex;
      }
      if (pid != positionIndex) {
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
      
      /*
      for (var r = sheet.topRow; r <= numRows; r++) {
        var statusRange = sheet.ref.getRange(r, sheet.columns["Status"]);
        if (! statusRange.isBlank()) {
          var action = statusRange.getValue();
          // Move row to the right sheet
          var callingsSheet = actionSheets[action];
          if (callingsSheet != sheet) {
            var rangeToMove = sheet.ref.getRange(r, 1, 1, sheet.ref.getMaxColumns());
            callingsSheet.ref.insertRows(callingsSheet.topRow);
            var targetRange = callingsSheet.ref.getRange(callingsSheet.topRow, 1);
            rangeToMove.moveTo(targetRange);
            if (callingsSheet != sheets.archivedCallings) {
              targetRange.clearDataValidations();
            }
            // TODO: Remove validations from row?
            
            // Delete row from original sheet
            sheet.ref.deleteRow(r);
            r--;
            numRows--;
            
            // Record affected sheets
            changedSheets.push(sheet);
            changedSheets.push(callingsSheet);
          }    
        }
      }
      */
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
  SpreadsheetApp.getUi().alert("To be implemented...");
}

function downloadMembers() {
  SpreadsheetApp.getUi().alert("To be implemented...");
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
