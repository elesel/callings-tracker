var VERSION = '0.6';
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
    .addItem('About', 'showAbout')
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
  header.appendTable([['Pending Callings', ''],['Confidential', 'Printed at <now>']]);
  var footer = doc.addFooter();
  footer.appendTable([['Confidential']]);
  
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
  //var pdfUrl = 'https://docs.google.com/document/d/' + doc.getId() + '/export?format=pdf';
  //SpreadsheetApp.getUi().alert('Printable version is located at ' + pdfUrl);
  
  // Delete
  //DriveApp.getFileById(doc.getId()).setTrashed(true);
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

function showAbout() {
  SpreadsheetApp.getUi().alert('callings-tracker v' + VERSION + '\n' + ABOUT_URL);
}
