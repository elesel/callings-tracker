// See https://github.com/elesel/callings-tracker
"use strict";
var VERSION = '0.7.6';
var ABOUT_URL = 'https://github.com/elesel/callings-tracker';

var NAME_PARSER_FNF = /^(.+)\s+(\S+)$/;
var NAME_PARSER_LNF = /^(\S+?),\s+(.+)$/;
var NAME_PARSER_LOOKUP = /^(.+)\s+\((.+)\)$/;
var HEADER_EMPTY = '<empty>';

// Define Sheet class and methods
function Sheet(name) {
  this.name = name;
  this.ref = null;
  this.topRow = null;
  this.columns = null;
}
Sheet.prototype.exists = function(){
  try {
    this.getRef();
    return true;
  } catch (error) {
    return false;
  }
};
Sheet.prototype.getRef = function(){
  if (this.ref) {
    return this.ref;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.name);
  if (! sheet) {
    throw "Can't find a sheet named '" + this.name + "'";
  }
  this.ref = sheet;
  return this.ref;
};
Sheet.prototype.getName = function(){
  return this.name;
}
Sheet.prototype.getTopRow = function(){
  if (this.topRow) {
    return this.topRow;
  }
  var frozenRows = this.getRef().getFrozenRows();
  if (frozenRows == 0) {
    throw "Header rows on sheet '" + this.name + "' aren't frozen";
  }
  this.topRow = frozenRows + 1;
  return this.topRow;
};
Sheet.prototype.getColumns = function(){
  if (this.columns) {
    return this.columns;
  }
  
  // Create column map
  var map = {};
  var headerValues = this.getRef().getRange(1, 1, this.getTopRow() - 1, this.getRef().getLastColumn()).getValues();
  for (var c = 0; c < headerValues[0].length; c++) {
    // Create hierarchical header value
    var headers = [];
    for (var r = 0; r < headerValues.length; r++) {
      var header = headerValues[r][c];
      if (header) {
        if (header != HEADER_EMPTY) {
          headers.push(header);
        }
      } else {
        // Assume it's part of a group of merged cells, look to the left until we find a value
        for (var c2 = c; c2 >= 0; c2--) {
          var header = headerValues[r][c2];
          if (header) {
            if (header != HEADER_EMPTY) {
              headers.push(header);
              headerValues[r][c] = header;
            }
            break;
          }
        }
      }
    }
    var header = headers.join('/');
    
    // Store in map
    var cBase1 = c + 1;
    if (header) {
      if (header in map) {
        // Multiple columns with the same name--handle as array
        if (Array.isArray(map[header])) {
          map[header].push(cBase1);
        } else {
          map[header] = [map[header], cBase1];
        }
      } else {
        map[header] = cBase1;
      }
    }
  }
  
  Logger.log("columns in " + this.name + ": " + JSON.stringify(map));
  this.columns = map;
  return this.columns;
};
Sheet.prototype.hasColumn = function(name){
  return name in this.getColumns();
};
Sheet.prototype.getColumn = function(name){
  if (! this.columns) {
    this.getColumns();
  }
  
  if (name in this.columns) {
    return this.columns[name];
  } else {
    throw "Can't find column '" + name + "' on sheet '" + this.name + "'";
  }
};

// Initialize sheets
var config = {};
loadConfiguration();
var sheets = {
  pendingCallings: new Sheet("Callings - Pending"),
  currentCallings: new Sheet("Callings - Current"),
  archivedCallings: new Sheet("Callings - Archive"),
  units: new Sheet("Units"),
  leaders: new Sheet("Leaders"),
  members: new Sheet("Members"),
  positions: new Sheet("Positions"),
  lifecycles: new Sheet("Lifecycles"),
  actions: new Sheet("Actions")
};

function showProgressMessage(message) {
  Logger.log("Showing message: " + message);
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
  var eventSheet = e.range ? e.range.getSheet() : SpreadsheetApp.getActiveSheet();
  if (eventSheet) {
    var eventSheetName = eventSheet.getName();
    for (var property in sheets) {
      if (sheets.hasOwnProperty(property)) {
        if (eventSheetName == sheets[property].getName()) {
          sheet = sheets[property];
          break;
        }
      }
    }
  }
  
  try {
    if (! sheet) {
      Logger.log("No sheet");
      // Let's assume everything changed
      addValidations();
      return;
    } else {
      Logger.log('Sheet is ' + sheet.getName());
      
      // Fire appropriate update functions depending on the sheet
      if ((sheet === sheets.pendingCallings || sheet === sheets.currentCallings) && e.changeType !== 'REMOVE_ROW') {
        if (e.range) {
          var startRow = e.range.getRowIndex();
          var numRows = e.range.getNumRows();
          updateCallingStatus(sheet, startRow, numRows);
          updatePendingMembers(startRow, numRows);
        } else {
          updateCallingStatus(sheet);
          updatePendingMembers();
        }
      } 
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
    .addItem('Format member list', 'formatMembers')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Setup')
      .addItem('Reload configuration', 'reloadConfiguration')
      .addItem('Show/hide configuration sheets', 'toggleConfigurationSheets')
      .addItem('Create triggers', 'checkAndCreateTriggers')
      .addItem('Remove triggers', 'removeTriggers'))
    .addItem('About', 'showAbout')
    .addToUi();
};

/**
 * @OnlyCurrentDoc
 */
function removeTriggers() {
  // Check current triggers
  var triggers = ScriptApp.getProjectTriggers();
  var onChangeTrigger = null;
  var onEditTrigger = null;
  var triggersRemoved = 0;
  triggers.forEach(function(trigger){
    if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
      triggersRemoved++;
    }
    if (trigger.getEventType() === ScriptApp.EventType.ON_CHANGE && trigger.getHandlerFunction() === 'onChange') {
      ScriptApp.deleteTrigger(trigger);
      triggersRemoved++;
    }
  });
  
  var ui = SpreadsheetApp.getUi();
  ui.alert(triggersRemoved + ' trigger(s) removed.');  
}

/**
 * @OnlyCurrentDoc
 */
function checkAndCreateTriggers() {
  // Check current triggers
  var triggers = ScriptApp.getProjectTriggers();
  var hasOnChange = false;
  var hasOnEdit = false;
  triggers.forEach(function(trigger){
    if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === 'onEdit') {
      hasOnEdit = true;
    }
    if (trigger.getEventType() === ScriptApp.EventType.ON_CHANGE && trigger.getHandlerFunction() === 'onChange') {
      hasOnChange = true;
    }
  });
  
  // Add triggers if necessary
  var ui = SpreadsheetApp.getUi();
  if (! (hasOnEdit && hasOnChange)) {
    var response = ui.alert('Do you want to create and run triggers as yourself? Someone (but only one) needs to say yes to this question in order for the spreadsheet to function.', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      var spreadsheet = SpreadsheetApp.getActive();
      var triggersCreated = 0;
      if (! hasOnEdit) {
        ScriptApp.newTrigger('onEdit')
        .forSpreadsheet(spreadsheet)
        .onEdit()
        .create();
        triggersCreated++;
      }
      if (!hasOnChange) {
        ScriptApp.newTrigger('onChange')
        .forSpreadsheet(spreadsheet)
        .onChange()
        .create();
        triggersCreated++;
      }
      ui.alert(triggersCreated + ' trigger(s) created.');
    } 
  } else {
    ui.alert('Nothing to do! Triggers already created.');
  }
}

function toggleConfigurationSheets() {
  try {
    var configurationSheets = [
      sheets.units,
      sheets.leaders,
      sheets.members,
      sheets.positions,
      sheets.lifecycles,
      sheets.actions
    ];
    
    var doHide = null;
    configurationSheets.forEach(function(sheet){
      if (sheet.exists()) {
        var sheetRef = sheet.getRef();
        
        if (doHide == null) {
          doHide = ! sheetRef.isSheetHidden();
        }
        
        if (doHide) {
          sheetRef.hideSheet();
        } else {
          sheetRef.showSheet();
        }
      }
    });
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
}

function reloadConfiguration() {
  try {
    var cache = CacheService.getDocumentCache();
    config = {};
    addConfigurationValidations();
    addValidations();
    getLifecycleActions_();
    getPositionLifecycles_();
    getActionSheets_();
    cache.put('config', JSON.stringify(config));
    Logger.log('Stored in cache: ' + Object.getOwnPropertyNames(config).sort());
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
}

function loadConfiguration() {
  try {
    var cache = CacheService.getDocumentCache();
    var tempConfig = cache.get('config');
    if (tempConfig != null) {
      config = JSON.parse(tempConfig);
      Logger.log('Loaded from cache: ' + Object.getOwnPropertyNames(config).sort());
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
    config = {};
  }
}

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
  var sheet = sheets.units.getRef();
  sheet.sort(sheets.units.getColumn("Abbreviation"));
  sheet.sort(sheets.units.getColumn("Visible"));
}

function getUnits_() {
  if (config['units']) {
    return config['units'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.units, "Abbreviation", function(value, row, sheet){
    return row[sheet.getColumn("Visible") - 1];
  });
  
  config['units'] = list;
  return list;
}

function sortLeaders_() {
  var sheet = sheets.leaders.getRef();
  sheet.sort(sheets.leaders.getColumn("Name"));
  sheet.sort(sheets.leaders.getColumn("Visible"));
}

function getLeaders_() {
  if (config['leaders']) {
    return config['leaders'];
  }
  
  // Get values
  var list = getLookupValues_(sheets.leaders, "Name", function(value, row, sheet){
    return row[sheet.getColumn("Visible") - 1];
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
    return row[sheet.getColumn("Visible") - 1];
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
    var name = row[sheet.getColumn("Name") - 1];
    var lifecycle = row[sheet.getColumn("Lifecycle") - 1];
    if (name && lifecycle) {
      positions[name] = lifecycle;
    }
  }
  
  config['positionLifecycles'] = positions;
  return positions;
}

function sortLifecycles_() {
  var sheet = sheets.lifecycles.getRef();
  sheet.sort(sheets.lifecycles.getColumn("Name"));
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
  var columnColumns = sheet.getColumn("Column");
  var actionColumns = sheet.getColumn("Action");
  if (columnColumns.length != actionColumns.length) {
    throw "Mismatch in number of Column and Action columns on " + sheets.lifecycles.getName() + " sheet";
  }
  
  // Get lifecycles
  var lifecycles = {};
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var name = row[sheet.getColumn("Name") - 1];
    
    // Only get active ones
    var actions = [];
    if (name) {
      // Add default action
      var action = row[sheet.getColumn("Default action") - 1];
      actions.push({ column: null, action: action });
      
      // Add other actions
      for (var i = 0; i < columnColumns.length; i++) {
        var column = row[columnColumns[i] - 1];
        var action = row[actionColumns[i] - 1];
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
    var name = row[sheet.getColumn("Name") - 1];
    var sheetName = row[sheet.getColumn("Sheet") - 1];
    if (name && sheetName) {
      // Translate sheet name into sheets object reference
      for (var property in sheets) {
        if (sheets.hasOwnProperty(property)) {
          if (sheetName == sheets[property].getName()) {
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
  var startRow = sheet.getTopRow();
  return sheet.getRef().getRange(startRow, 1, sheet.getRef().getLastRow() - startRow + 1, sheet.getRef().getLastColumn()).getValues();
}

function getAllCells_(sheet) {
  var startRow = sheet.getTopRow();
  return sheet.getRef().getRange(startRow, 1, sheet.getRef().getMaxRows() - startRow + 1, sheet.getRef().getMaxColumns()).getValues();
}

function getLookupValues_(sheet, columnName, validationFunction) {
  var allData = getAllData_(sheet);
  
  // Get lookup values
  var list = [];
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var value = row[sheet.getColumn(columnName) - 1];
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
  list.push(sheets.pendingCallings.getName());
  list.push(sheets.currentCallings.getName());
  list.push(sheets.archivedCallings.getName());
  return list;
}

function addConfigurationValidations() {
  var sheet;
  
  // Sort sheets first
  if (sheets.units.exists()) {
    sortUnits_();
  }
  if (sheets.leaders.exists()) {
    sortLeaders_();
  }
  sortLifecycles_();
  
  // Add lifecycles list to positions sheet
  var lifecyclesRule = SpreadsheetApp.newDataValidation().requireValueInList(getLifecycleNames_()).setAllowInvalid(true).build();
  sheet = sheets.positions;
  sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Lifecycle"), sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(lifecyclesRule);
  
  // Add positions columns lists to lifecycle sheet
  var positionsColumnsRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(sheets.pendingCallings.getColumns())).setAllowInvalid(true).build();
  sheet = sheets.lifecycles;
  sheet.getColumn('Column').forEach(function(c){
    sheet.getRef().getRange(sheet.getTopRow(), c, sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(positionsColumnsRule);
  });
  
  // Add action names lists to lifecycle sheet
  var actionsRule = SpreadsheetApp.newDataValidation().requireValueInList(getActionNames_()).setAllowInvalid(true).build();
  sheet = sheets.lifecycles;
  sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Default action"), sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(actionsRule);
  sheet.getColumn('Action').forEach(function(c){
    sheet.getRef().getRange(sheet.getTopRow(), c, sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(actionsRule);
  });
  
  // Add sheet name lists to actions sheet
  var callingSheetsRule = SpreadsheetApp.newDataValidation().requireValueInList(getCallingSheetNames_()).setAllowInvalid(true).build();
  sheet = sheets.actions;
  sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Sheet"), sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(callingSheetsRule);
}

function addValidations() {
  var sheet;
   
  // Update member names and add validations
  updatePendingMembers();
    
  // Add units lists to pending callings sheet
  if (sheets.units.exists()) {
    var unitsRule = SpreadsheetApp.newDataValidation().requireValueInList(getUnits_()).setAllowInvalid(true).build();
    sheet = sheets.pendingCallings;
    sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Member/Unit"), sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(unitsRule);
  }
    
  // Add positions lists to pending callings sheet
  var positionsRule = SpreadsheetApp.newDataValidation().requireValueInList(getPositions_()).setAllowInvalid(true).build();
  sheet = sheets.pendingCallings;
  sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Position"), sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(positionsRule);
  
  // Add leaders lists to pending callings sheet
  if (sheets.leaders.exists()) {
    var leadersRule = SpreadsheetApp.newDataValidation().requireValueInList(getLeaders_()).setAllowInvalid(true).build();
    sheet = sheets.pendingCallings;
    sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Extend/Set apart by"), sheet.getRef().getMaxRows() - sheet.getTopRow() + 1, 1).setDataValidation(leadersRule);  
  }
}

function updateAllCallingStatus() {
  // Update calling status
  [sheets.pendingCallings, sheets.currentCallings].forEach(function(sheet){
    updateCallingStatus(sheet);
  });
}

function updateCallingStatus(sheet, startRow, numRows) {
  // Set defaults for startRow and numRows
  if (arguments.length == 1) {
    startRow = sheet.getTopRow();
    numRows = sheet.getRef().getLastRow() - sheet.getTopRow() + 1;
  }
  
  // Skip empty sheets
  if (startRow < sheet.getTopRow() || numRows < 1) {
    return;
  }
  
  // Get positions and lifecycles
  var lifecycleActions = getLifecycleActions_();
  var positionLifecycles = getPositionLifecycles_();
  var positionNames = getPositions_();
  var actionNames = getActionNames_();
  
  // Grab copy of all data
  var allData = sheet.getRef().getRange(startRow, 1, numRows, sheet.getRef().getLastColumn()).getValues();
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var memberName = row[sheet.getColumn("Member/Name") - 1];
    var positionName = row[sheet.getColumn("Position") - 1];
    var status = row[sheet.getColumn("Status") - 1];
    var newStatus = status;
    var sid = row[sheet.getColumn("SID") - 1];
    var newSid = sid;
    var pid = row[sheet.getColumn("PID") - 1];
    var newPid = pid;
    if (positionName) {
      // Find current action
      var lifecycle = lifecycleActions[positionLifecycles[positionName]];
      var newStatus = 'Unknown';
      if (lifecycle) {
        lifecycle.some(function(columnAction){
          if (! columnAction.column) {
            newStatus = columnAction.action;
          } else {
            var value = row[sheet.getColumn(columnAction.column) - 1];
            if (value && value[0] != '>') {
              newStatus = columnAction.action;
            } else {
              return true;
            }
          }
        });
      }
      
      // Look up indexes
      var newSid = actionNames.indexOf(newStatus);
      var newPid = positionNames.indexOf(positionName);
    } else if (memberName) {
      newStatus = 'Nominate';
      newSid = -1;
      newPid = '';
    } else {
      newStatus = '';
      newSid = '';
      newPid = '';
    }
    
    // Change values only if we have to
    if (status !== newStatus) {
      row[sheet.getColumn("Status") - 1] = newStatus;
    }
    if (sid !== newSid) {
      row[sheet.getColumn("SID") - 1] = newSid;
    }
    if (pid !== newPid) {
      row[sheet.getColumn("PID") - 1] = newPid;
    }
  }
  
  // Write columns back
  var dataRows = allData.length;
  sheet.getRef().getRange(startRow, sheet.getColumn("Status"), dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.getColumn("Status") - 1)
  );
  sheet.getRef().getRange(startRow, sheet.getColumn("SID"), dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.getColumn("SID") - 1)
  );
  sheet.getRef().getRange(startRow, sheet.getColumn("PID"), dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.getColumn("PID") - 1)
  );
}

function organizeCallings() {
  try {
    // Sanity check the sheets
    if (sheets.pendingCallings.getRef().getMaxColumns() != sheets.currentCallings.getRef().getMaxColumns() || sheets.pendingCallings.getRef().getMaxColumns() != sheets.archivedCallings.getRef().getMaxColumns()) {
      throw "Mismatch in number of columns among the callings sheets. They must be the same.";
    }
    
    // Get action to sheet mapping
    var actionSheets = getActionSheets_();
    
    // Loop through calling sheets
    var changedSheets = [];
    [sheets.pendingCallings, sheets.currentCallings].forEach(function(sheet){
      // Only continue if there's data
      var numRows = sheet.getRef().getLastRow() - sheet.getTopRow() + 1;
      if (numRows < 1) {
        return;
      }
      
      // Sort callings so that we can move contiguous rows
      updateCallingStatus(sheet);
      sortCallings(sheet);
      
      // Iterate through all status values
      var allData = sheet.getRef().getRange(sheet.getTopRow(), sheet.getColumn("Status"), numRows, 1).getValues();
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
            if (targetSheet != null && targetSheet.getName() != sheet.getName()) {
              // Send rows to another sheet
              var endRow = (step == 'previous' ? r - 1 : r);
              var realStartRow = startRow + sheet.getTopRow() - rowsMoved;
              var realEndRow = endRow + sheet.getTopRow() - rowsMoved;
              
              // Move
              Logger.log("Move row " + realStartRow + " through row " + realEndRow + " from sheet " + sheet.getName() + " to sheet " + targetSheet.getName());
              var targetRange = moveRows(sheet, targetSheet, realStartRow, realEndRow - realStartRow + 1, targetSheet.getTopRow());
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
  sheet.getRef().sort(sheet.getColumn("Member/Name"));
  sheet.getRef().sort(sheet.getColumn("Extend/Sustain"));
  sheet.getRef().sort(sheet.getColumn("PID"));
  sheet.getRef().sort(sheet.getColumn("SID"));
}

function formatMembers() {
  var sheet = sheets.members;
  
  // Grab copy of all data
  var startRow = sheet.getTopRow();
  var allData = getAllData_(sheet);
  
  // Determine lookup names
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var fullName = row[sheet.getColumn("Full name") - 1].trim();
    var age = row[sheet.getColumn("Age") - 1];
    var unit = sheet.hasColumn("Unit") ? row[sheet.getColumn("Unit") - 1] : null;
    var forcedName = row[sheet.getColumn("Forced name") - 1];
    var lookupName = forcedName || reformatNameLnfToShort(fullName);
    var details = [];
    if (age) {
      details.push(age);
    }
    if (unit) {
      details.push(unit);
    }
    if (details.length > 0) {
      lookupName = lookupName + " (" + details.join(", ") + ")"; 
    }
    row[sheet.getColumn("Lookup name") - 1] = lookupName;
  }
  
  // Write columns back
  var dataRows = allData.length;
  sheet.getRef().getRange(startRow, sheet.getColumn("Lookup name"), dataRows, 1).setValues(
    sliceSingleColumn(allData, 0, dataRows, sheet.getColumn("Lookup name") - 1)
  );
  
  // Refresh configuration
  reloadConfiguration();
}

function updatePendingMembers(startRow, numRows) {
  var sheet = sheets.pendingCallings;
  
  // Set defaults for startRow and numRows
  if (arguments.length == 0) {
    startRow = sheet.getTopRow();
    numRows = sheet.getRef().getMaxRows() - startRow + 1;
  }
  
  // Skip empty sheets
  if (startRow < sheet.getTopRow() || numRows < 1) {
    return;
  }
  
  // Grab copy of all data
  var allData = sheet.getRef().getRange(startRow, 1, numRows, sheet.getRef().getMaxColumns()).getValues();
  
  // Build members validation rule
  var membersRule = SpreadsheetApp.newDataValidation().requireValueInList(getMembers_()).setAllowInvalid(true).build();
  
  // Clear validations from member name column
  sheet.getRef().getRange(startRow, sheet.getColumn("Member/Name"), numRows, 1).clearDataValidations();
  
  // Process rows
  var nameChanged = false;
  var unitChanged = false;
  var rangeStartRow = null;
  var rangeStopRow = null;
  for (var r = 0; r < allData.length; r++) {
    var row = allData[r];
    var name = row[sheet.getColumn("Member/Name") - 1];
    var result = NAME_PARSER_LOOKUP.exec(name);
    if (result != null) {
      // Reduce member name cell to just member name
      row[sheet.getColumn("Member/Name") - 1] = result[1];
      nameChanged = true;
      if (sheet.hasColumn("Member/Unit")) {
        var unit = row[sheet.getColumn("Member/Unit") - 1];
        if (! unit) {
          // Add unit only if not already defined
          var ageAndUnit = result[2].split(/\s*,\s*/);
          var unit = null;
          if (ageAndUnit.length > 1) {
            unit = ageAndUnit[1];
          } else {
            unit = ageAndUnit[0];
          }
          // TODO: Validate unit better
          if (unit) {
            row[sheet.getColumn("Member/Unit") - 1] = unit;
            unitChanged = true;
          }
        }
      }
    }
    
    // Update validation pointers
    if (! name) {
      if (rangeStartRow == null) {
        rangeStartRow = r;
      }
      rangeStopRow = r;
    }
    
    // Add validation
    if (rangeStopRow != null && (r > rangeStopRow || r == allData.length - 1)) {
      sheet.getRef().getRange(startRow + rangeStartRow, sheet.getColumn("Member/Name"), rangeStopRow - rangeStartRow + 1, 1).setDataValidation(membersRule);
      rangeStartRow = null;
      rangeStopRow = null;
    }
  }
  
  // Write columns back
  var dataRows = allData.length;
  if (nameChanged) {
    sheet.getRef().getRange(startRow, sheet.getColumn("Member/Name"), dataRows, 1).setValues(
      sliceSingleColumn(allData, 0, dataRows, sheet.getColumn("Member/Name") - 1)
    );
  }
  if (unitChanged) {
    sheet.getRef().getRange(startRow, sheet.getColumn("Member/Unit"), dataRows, 1).setValues(
      sliceSingleColumn(allData, 0, dataRows, sheet.getColumn("Member/Unit") - 1)
    );
  }
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

function moveRows(sourceSheet, targetSheet, startRow, numRows, targetRow) {
  var numColumns = sourceSheet.getRef().getMaxColumns();
  var sourceRange = sourceSheet.getRef().getRange(startRow, 1, numRows, numColumns);
  
  targetSheet.getRef().insertRows(targetSheet.getTopRow(), numRows);
  var targetRange = targetSheet.getRef().getRange(targetSheet.getTopRow(), 1, numRows, numColumns);
  sourceRange.moveTo(targetRange);
  sourceSheet.getRef().deleteRows(startRow, numRows);
  
  return targetRange;
}

function isMerged(cell) {
  var source_background = cell.getBackground();
  cell.setBackground('#fffffe');
  var merged = (cell.getBackground() == '#ffffff');
  cell.setBackground(source_background);

  return merged;
}

function showAbout() {
  SpreadsheetApp.getUi().alert('Callings Tracker v' + VERSION + '\n' + ABOUT_URL);
}