// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 7 Oct 2014 19:49 GMT+1

/*
 * This file is part of Rose Task Manager.
 *
 * Copyright (C) 2014 Andrew Roberts
 * 
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later 
 * version.
 * 
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License along with 
 * this program. If not, see http://www.gnu.org/licenses/.
 */

/**
 * test.gs
 * =======
 */

function test_init() {
  Log_ = {
    functionEntryPoint: function(msg) {Logger.log(msg)},
    fine: function(msg) {Logger.log(msg)},
    finer: function(msg) {Logger.log(msg)},
    finest: function(msg) {Logger.log(msg)},
    warning: function(msg) {Logger.log(msg)},
  }
}

function test() {
  test_init()
  var a
  var config = {a:a}
  if (typeof config.a === 'undefined') throw 'STOP'  
  return
}

function test_regex1(arg) {
  test_init()
  var TEMPLATE = "a {{<a1234 - pretend HTML>1}} 2 3"
  // "\\<[^\\]]*\\>"
  // /{{[\w\s\d]+}}/g
  var a = TEMPLATE.match(/{{.*?}}/g)
  var b = a[0].replace(/<.*?>/g, '')
  return
}

// Settings Sidebar
// ----------------
//
// http://www.mcpher.com/Home/excelquirks/addons/addonsettings

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Sidebar')
}

// Emails_
// =======

function fillInTemplate() {

  var headerRow = ['P1', 'P2', 'P3']
  var activeRow = ['a', 'b', 'c']
  var rowObject = {}
  
  for (var columnIndex = 0; columnIndex < headerRow.length; columnIndex++) {
    rowObject[headerRow[columnIndex]] = activeRow[columnIndex]
  }
  
  var TEMPLATE = 'Test - {{P1}} - {{P2}} - {{P1}}'
  var result = ''
  var REGEX = /{{[\w\s\d]+}}/g
  
  result = TEMPLATE.replace(REGEX, replacer) 
  
  return

  function replacer (nextMatch) {
    var placeholderValue = nextMatch.substring(2, nextMatch.length - 2)
    return rowObject[placeholderValue]
  }

/*  
    result = TEMPLATE.replace(REGEX, )    
*/  
  
/*
  for (var columnIndex = 0; columnIndex < headerRow.length; columnIndex++) {
    
    var header = headerRow[columnIndex]
    var nextValue = activeRow[columnIndex]
    
    result = TEMPLATE.replace(REGEX, function (nextMatch) {
      var label = nextMatch.substring(2, nextMatch.length - 2)
      return nextValue
    })    
  }
*/



  return
}

// Auth
// ====

function test_auth() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
  var status = authInfo.getAuthorizationStatus()
  if (status === ScriptApp.AuthorizationStatus.NOT_REQUIRED) {
    throw "Not Required"
  }
  var url = authInfo.getAuthorizationUrl()
  return
} 

// Dump Config
// ===========

function logConfig() {

// PropertiesService.getUserProperties().deleteAllProperties()
// PropertiesService.getDocumentProperties().deleteAllProperties()
// PropertiesService.getScriptProperties().deleteAllProperties()

  Logger.log(PropertiesService.getDocumentProperties().getProperties())
  Logger.log(PropertiesService.getScriptProperties().getProperties())
  Logger.log(PropertiesService.getUserProperties().getProperties())
  
  Logger.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' project triggers.')
  
  var document = DocumentApp.getActiveDocument()
  
  if (document) {
  
    Logger.log('Current project has ' + ScriptApp.getUserTriggers(document).length + ' document triggers.')
    
  } else {
  
    Logger.log('There is no documents associated with this script, so no triggers')
  }

  var form = FormApp.getActiveForm()
  
  if (form) {
  
    Logger.log('Current project has ' + ScriptApp.getUserTriggers(form).length + ' form triggers.')
    
  } else {
  
    Logger.log('There is no forms associated with this script, so no triggers')
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  if (spreadsheet) {
  
    Logger.log('Current project has ' + ScriptApp.getUserTriggers(spreadsheet).length + ' spreadsheet triggers.')
    
  } else {
  
    Logger.log('There is no spreadsheets associated with this script, so no triggers')
  }
  
} // logConfig()

// Unit Tests
// ==========

// TODO - Lots of test code to write!!

/**
 * Main test function.
 *
 * For the neatest log output set the level to INFO.
 */

function test_roseTaskManager() {
  
  if (PRODUCTION_VERSION) {
    
    return
  }
  
  var functionName = 'test_roseTaskManager()'
  
  Log_.init(LOG_LEVEL, LOG_SHEET_ID)
  Log_.info(functionName)
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  spreadsheet.toast('Unit tests started', 
                    ROSE_TASK_MANAGER_NAME, -1)
  
  test_utilities_gs()
  test_log_gs()
  test_underscoreGS_gs()
  test_assert()    
  
  spreadsheet.toast('Unit tests finished', 
                    ROSE_TASK_MANAGER_NAME, -1)
  
} // test_roseTaskManager()

/**
 * Utils Unit tests
 * ================
 */

/**
 * Test the functions in utilities.gs.
 *
 * The task list sheet has to be active and dialog boxes will need
 * to be closed by the tester during the tests.
 */
 
function test_utilities_gs() {

  if (PRODUCTION_VERSION) {
    
    return
  }

  var functionName = 'test_utilities_gs()' 
  Log_.info(functionName)

  test_alert()
  test_getColIndexByName()
  test_setCellValue()
  test_getCellValue()
  
} // test_utilities_gs()

/**
 * Test alert()
 */
 
function test_alert() {

  if (PRODUCTION_VERSION) {
    
    return
  }

  var functionName = 'test_alert()'

  Log_.info(functionName)
  
  try {
  
    alert()
    
  } catch (error) {
  
    Log_.info('alert() Test 1 - PASSED, ERROR: ' + error.message)
  }
  
  alert('Test 2') 
  
  Log_.info('alert() Test 2 - PASSED')
  
  try {
  
    alert('Test 3', 1)
    
  } catch (error) {
  
    Log_.info('alert() Test 3 - PASSED: ERROR: ' + error.message)
  }
  
  alert('test title', 'test message that is really loooooooooooooo' + 
    'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo' + 
    'oooooong, to check that the text is not truncated.') 

  Log_.info('alert() Test 4 - PASSED')

} // test_alert()

/**
 * Test getColumnPosition_()
 *
 * The task list sheet needs to be active.
 */
 
function test_getColIndexByName() {

  if (PRODUCTION_VERSION) {
    
    return
  }

  var functionName = 'test_getColIndexByName()'
  Log_.info(functionName)
  
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
  
  try {
  
    // No sheet
    getColumnPosition_()
    
  } catch (error) {
  
    Log_.info(
             'getColumnPosition_() - TEST 1 PASSED. ERROR: ' + error.message)
  }

  try {
  
    // Non-string column name
    getColumnPosition_(sheet, 1)
    
  } catch (error) {
  
    Log_.info(
             'getColumnPosition_() - TEST 2 PASSED. ERROR: ' + error.message)
  }

  var columnIndex = getColumnPosition_(sheet, SS_COL_TIMESTAMP)

  assert_(columnIndex === 2, 
          functionName,
          'getColumnPosition_() - TEST 3 FAILED - wrong column index: ' + columnIndex)
          
  Log_.info(
           'getColumnPosition_() - TEST 3 PASSED: Read correct column')

} // test_getColIndexByName()

/**
 * Test setCellValue_()
 *
 * The task list sheet needs to be active. The test sets B2 and then 
 * returns it to it previous value.
 */
 
function test_setCellValue() {

  if (PRODUCTION_VERSION) {
    
    return
  }

  var functionName = 'test_setCellValue()'
  var functionUnderTest = 'setCellValue_()'
  
  Log.init(LOG_LEVEL, LOG_SHEET_ID)  
  Log_.info(functionName)
  
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
  
  var testCellRange = sheet.getRange('A2')
  var oldCellValue = testCellRange.getValue()
  
  try {
  
    // No sheet
    setCellValue_()
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 1 PASSED. ERROR: ' + error.message)
  }

  try {
  
    // No row index
    setCellValue_(sheet)
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 2 PASSED. ERROR: ' + error.message)
  }

  try {
  
    // No column name
    setCellValue_(sheet, 2)
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 3 PASSED. ERROR: ' + error.message)
  }

  try {
  
    // No value
    setCellValue_(sheet, 2, SS_COL_ID)
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 4 PASSED. ERROR: ' + error.message)
  }

  // Valid set
  setCellValue_(sheet, 2, SS_COL_ID, 192837456)
  
  var newCellValue = testCellRange.getValue()

  assert_(newCellValue === 192837456, 
          functionName,
          functionUnderTest + 
            ' - TEST 5 FAILED - wrong value found: ' + newCellValue)
      
  Log_.info(
           functionUnderTest + 
             ' - TEST 5 PASSED - value set correctly: ' + newCellValue)

  testCellRange.setValue(oldCellValue)

} // test_setCellValue()

/**
 * Test getCellValue_()
 *
 * The task list sheet needs to be active. The test sets B2 and then 
 * returns it to it previous value.
 */
 
/**
 * Get the value of a spreadhsheet cell
 *
 * @param {object} sheet
 * @param {number} rowIndex
 * @param {string} column name
 * @return {any} value or null
 */
  
function test_getCellValue() {

  if (PRODUCTION_VERSION) {
    
    return
  }

  var functionName = 'test_getCellValue()'
  var functionUnderTest = 'getCellValue_()'
  
  Log.init(LOG_LEVEL, LOG_SHEET_ID)  
  Log_.info(functionName)
  
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
    
  try {
  
    // No sheet
    getCellValue_()
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 1 PASSED. ERROR: ' + error.message)
  }

  try {
  
    // No row index
    getCellValue_(sheet)
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 2 PASSED. ERROR: ' + error.message)
  }

  try {
  
    // No column name
    getCellValue_(sheet, 2)
    
  } catch (error) {
  
    Log_.info(
             functionUnderTest + ' - TEST 3 PASSED. ERROR: ' + error.message)
  }

  // Valid get
  
  var testCellRange = sheet.getRange('A2')
  
  var oldCellValue = testCellRange.getValue()
  
  testCellRange.setValue(982453457)
  
  var testGetValue = getCellValue_(sheet, 2, SS_COL_ID)
  
  assert_(testGetValue === 982453457, 
          functionName,
          functionUnderTest + 
            ' - TEST 4 FAILED - wrong value found: ' + testGetValue)
      
  Log_.info(
           functionUnderTest + 
             ' - TEST 4 PASSED - value set correctly: ' + testGetValue)

  testCellRange.setValue(oldCellValue)

} // test_getCellValue()

// RoseTaskManager
// ---------------

/**
 * Test onFormSubmit()
 */
 
function test_onFormSubmit() {
  
  Log_.functionEntryPoint()
  
  var taskListRangeA1 = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
    .getRange('A1')
  
  var event = {
  
    range: taskListRangeA1
  }
  
  onFormSubmit(event)
  
} // test_onFormSubmit()

// Dump/Clear Config
// -----------------

/**
 * Clear all of the config
 */

function onClearConfig_() {

  Log_.info('Delete Local Script Properties:')  
  if (PropertiesService.getScriptProperties() !== null) {
    Log_.info(PropertiesService.getScriptProperties().deleteAllProperties())
    Log_.info('  Deleted')
  } else {
    Log_.info('  None')
  }

  Log_.info('Delete Local NPT Doc Properties:')  
  if (PropertiesService.getDocumentProperties() !== null) {
    Log_.info(PropertiesService.getDocumentProperties().deleteAllProperties())
    Log_.info('  Deleted')    
  } else {
    Log_.info('  None')
  }
 
  Log_.info('Delete Local NPT User Properties:')    
  if (PropertiesService.getUserProperties() !== null) {  
    Log_.info(PropertiesService.getUserProperties().deleteAllProperties())
    Log_.info('  Deleted')    
  } else {
    Log_.info('  None')
  }
  
  var triggers = ScriptApp.getProjectTriggers()
  Log_.info('Current project has ' + triggers.length + ' project triggers.')
  deleteTriggers(triggers)
  
  var document = DocumentApp.getActiveDocument()
  
  if (document) {
    
    triggers = ScriptApp.getUserTriggers(document)
    Log_.info('Current project has ' + triggers.length + ' user document triggers.')
    deleteTriggers(triggers)
    
  } else {
    
    Log_.info('  There is no documents associated with this script, so no triggers')
  }
  
  var form = FormApp.getActiveForm();
  
  if (form) {
    
    triggers = ScriptApp.getUserTriggers(form)
    Log_.info('Current project has ' + triggers.length + ' user form triggers.')
    deleteTriggers(triggers)
    
  } else {
    
    Log_.info('  There is no forms associated with this script, so no triggers')
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  if (spreadsheet) {
    
    triggers = ScriptApp.getUserTriggers(spreadsheet)
    Log_.info('Current project has ' + triggers.length + ' user spreadsheet triggers.')
    deleteTriggers(triggers)
    
  } else {
      
    Log_.info('  There is no spreadsheets associated with this script, so no triggers')
  }

  // Private Functions
  // -----------------
  
  function deleteTriggers(triggers) {
    triggers.forEach(function(trigger) {
      Log_.info('  Deleting trigger: ' + trigger.getUniqueId())
      ScriptApp.deleteTrigger(trigger)
    })
  }
  
} // onClearConfig_()

/**
 * Dump all of the local config
 */

function onDumpConfig_() {

  var values

  Log_.info('RTM Script Properties:')
  var properties = PropertiesService.getScriptProperties()
  if (properties !== null) {
    values = properties.getProperties()
    for (var key in values) {
      Log_.info('  ' + key + ' : ' + values[key])
    }
  } else {
    Log_.info('  None')
  }

  Log_.info('RTM Doc Properties:') 
  var properties = PropertiesService.getDocumentProperties()
  if (properties !== null) {
    values = properties.getProperties()
    for (var key in values) {
      Log_.info('  ' + key + ' : ' + values[key])
    }
  } else {
    Log_.info('  None')
  }
  
  Log_.info('RTM User Properties:')    
  var properties = PropertiesService.getUserProperties()  
  if (PropertiesService.getUserProperties() !== null) {  
    values = properties.getProperties()
    for (var key in values) {
      Log_.info('  ' + key + ' : ' + values[key])
    }
  } else {
    Log_.info('  None')
  }
  
  var triggers = ScriptApp.getProjectTriggers()
  Log_.info('Current project has ' + triggers.length + ' project triggers.')
  dumpTriggers(triggers)
  
  var document = DocumentApp.getActiveDocument()
  
  if (document) {
    
    triggers = ScriptApp.getUserTriggers(document)
    Log_.info('Current project has ' + triggers.length + ' user document triggers.')
    dumpTriggers(triggers)
    
  } else {
    
    Log_.info('  There is no documents associated with this script, so no triggers')
  }
  
  var form = FormApp.getActiveForm();
  
  if (form) {
    
    triggers = ScriptApp.getUserTriggers(form)
    Log_.info('Current project has ' + triggers.length + ' user form triggers.')
    dumpTriggers(triggers)
    
  } else {
    
    Log_.info('  There is no forms associated with this script, so no triggers')
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  if (spreadsheet) {
    
    triggers = ScriptApp.getUserTriggers(spreadsheet)
    Log_.info('Current project has ' + triggers.length + ' user spreadsheet triggers.')
    dumpTriggers(triggers)
    
  } else {
    
    Log_.info('  There is no spreadsheets associated with this script, so no triggers')
  }
  
  return
  
  // Private Functions
  
  function dumpTriggers(triggers) {
  
    triggers.forEach(function(trigger) {
      Log_.info('  ID: ' + trigger.getUniqueId())
      Log_.info('  Source ID: ' + trigger.getTriggerSourceId())
      Log_.info('  Source: ' + trigger.getTriggerSource())
      Log_.info('  Handler Function: ' + trigger.getHandlerFunction())
      Log_.info('  Event Type: ' + trigger.getEventType())    
    })
  
  } // onDumpConfig_.dumpTriggers()
    
} // onDumpConfig_()