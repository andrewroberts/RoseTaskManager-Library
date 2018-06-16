// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 14 Oct 2014 08:49 GMT+1
/* jshint asi: true */

/*
 * This file is part of Rose Task Manager.
 *
 * Copyright (C) 2015-2018 Andrew Roberts (andrewroberts.net)
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

// TODO - Re-formatting. 
// TODO - Email status alert is a bit big, create an email.gs. Probably same for 
// other functions too.
// TODO - There is a function in here with a cyclomatic complexity of 12! Would 
// like less than 10. Probably setupTaskListSheet().
 
/**
 * RoseTaskManager.gs
 * ==================
 *
 * Functions that handle script events. All externally invoked events should 
 * be processed in here: menu clicks, triggers, sheet opening, etc.
 *
 * To avoid too much information being passed up to the user, in PRODUCTION
 * they should all catch errors before displaying them to the user.
 */

var Log_

// Public event handlers
// ---------------------
//
// All external event handlers need to be top-level function calls they can't 
// be part of an object, and to ensure they are all processed similarily 
// for things like logging and error handling, they all go through 
// errorHandler_(). These can be called from custom menus, web apps, 
// triggers, etc
// 
// The main functionality of a call is in a function with the same name but 
// post-fixed with an underscore (to indicate it is private to the script)
//
// For debug, rather than production builds, lower level functions are exposed
// in the menu

//   :      [function() {},  '()',      'Failed to ', ],

var EVENT_HANDLERS = {

//                         Initial actions  Name                         onError Message                        Main Functionality
//                         ---------------  ----                         ---------------                        ------------------

  onInstall:               [function() {},  'onInstall()',              'Failed to install RTM add-on',         onInstall_],
  onEmailStatusUpdates:    [function() {},  'onEmailStatusUpdates()',   'Failed to send email status',          onEmailStatusUpdates_],
  onFormSubmit:            [function() {},  'onFormSubmit()',           'Failed to process form submission',    onFormSubmit_],
  onEdit:                  [function() {},  'onEdit()',                 'Failed to process edit',               onEdit_],
  onCalendarTrigger:       [function() {},  'onCalendarTrigger()',      'Failed to check calendar for tasks',   onCalendarTrigger_],
  onClearLog:              [function() {},  'onClearLog()',             'Failed to clear log',                  onClearLog_],  
  onUninstall:             [function() {},  'onUninstall()',            'Failed to uninstall',                  onUninstall_],  
  onDumpConfig:            [function() {},  'onDumpConfig()',           'Failed to dump config',                onDumpConfig_],  
  onClearConfig:           [function() {},  'onClearConfig()',          'Failed to clear config',               onClearConfig_],  
  onThrowError:            [function() {},  'onThrowError()',           'Deliberatley threw error',             onThrowError_],  
  onStartCalendarTrigger:  [function() {},  'onStartCalendarTrigger()', 'Failed to start calendar trigger',     onStartCalendarTrigger_],    
  onStopCalendarTrigger:   [function() {},  'onStopCalendarTrigger()',  'Failed to stop calendar trigger',      onStopCalendarTrigger_],       
  onTest:                  [function() {},  'onTest()',                  'Failed to run test',                  onTest_],      
  onShowSidebar:           [function() {},  'onShowSidebar()',           'Failed to show sidebar',              onShowSidebar_],        
  onGetSettings:           [function() {},  'onGetSettings()',           'Failed to get menu',                  onGetSettings_],        
  onSetNewTaskTemplate:    [function() {},  'onSetNewTaskTemplate()',    'Failed to set new task template',     onSetNewTaskTemplate_],          
  onSetStatusTemplate:     [function() {},  'onSetStatusTemplate()',     'Failed to set status value',          onSetStatusTemplate_],          
  onSetNewFrom:            [function() {},  'onSetNewFrom()',            'Failed to set email from',            onSetNewFrom_],          
  onDumpEventCount:        [function() {},  'onDumpEventCount()',        'Failed to dump event count',          onDumpEventCount_],          
}

// function (arg)                     {return eventHandler_(EVENT_HANDLERS., arg)}

function onInstall(arg)              {return eventHandler_(EVENT_HANDLERS.onInstall, arg)}
function onEmailStatusUpdates(arg)   {return eventHandler_(EVENT_HANDLERS.onEmailStatusUpdates, arg)}
function onFormSubmit(arg)           {return eventHandler_(EVENT_HANDLERS.onFormSubmit, arg)}
function onEdit(arg)                 {return eventHandler_(EVENT_HANDLERS.onEdit, arg)}
function onCalendarTrigger(arg)      {return eventHandler_(EVENT_HANDLERS.onCalendarTrigger, arg)}
function onClearLog(arg)             {return eventHandler_(EVENT_HANDLERS.onClearLog, arg)}
function onUninstall(arg)            {return eventHandler_(EVENT_HANDLERS.onUninstall, arg)}
function onDumpConfig(arg)           {return eventHandler_(EVENT_HANDLERS.onDumpConfig, arg)}
function onClearConfig(arg)          {return eventHandler_(EVENT_HANDLERS.onClearConfig, arg)}
function onThrowError(arg)           {return eventHandler_(EVENT_HANDLERS.onThrowError, arg)}
function onStartCalendarTrigger(arg) {return eventHandler_(EVENT_HANDLERS.onStartCalendarTrigger, arg)}
function onStopCalendarTrigger(arg)  {return eventHandler_(EVENT_HANDLERS.onStopCalendarTrigger, arg)}
function onTest(arg)                 {return eventHandler_(EVENT_HANDLERS.onTest, arg)}
function onShowSidebar(arg)          {return eventHandler_(EVENT_HANDLERS.onShowSidebar, arg)}
function onGetSettings(arg)          {return eventHandler_(EVENT_HANDLERS.onGetSettings, arg)}
function onSetNewTaskTemplate(arg)   {return eventHandler_(EVENT_HANDLERS.onSetNewTaskTemplate, arg)}
function onSetStatusTemplate(arg)    {return eventHandler_(EVENT_HANDLERS.onSetStatusTemplate, arg)}
function onSetNewFrom(arg)           {return eventHandler_(EVENT_HANDLERS.onSetNewFrom, arg)}
function onDumpEventCount(arg)       {return eventHandler_(EVENT_HANDLERS.onDumpEventCount, arg)}

function onOpen(event) {onOpen_(event)}

// Private Functions
// =================

function onCalendarTrigger_()      {Calendar_.convertEventsToTasks()}
function onStartCalendarTrigger_() {Calendar_.createTrigger()}
function onStopCalendarTrigger_()  {Calendar_.deleteTrigger(true)}
function onThrowError_()           {throw new Error('Force error throw')}
function onTest_()                 {test()}
function onDumpEventCount_()       {Calendar_.dumpEventCount()}

// General
// -------

/**
 * All external function calls should call this to ensure standard 
 * processing - logging, errors, etc - is always done.
 *
 * @param {array} config:
 *   [0] {function} prefunction
 *   [1] {string} eventName
 *   [2] {string} onErrorMessage
 *   [3] {function} mainFunction
 *
 * @parma {object} arg The argument passed to the top-level event handler
 */

/*
 {
   "values":["11/02/2017 10:57:28","1057","","","","",""],
   "namedValues":{
     "Description":[""],
     "Priority":[""],
     "Requested by":[""],
     "Contact email":[""],
     "Timestamp":[
       "11/02/2017 10:57:28"
     ],
     "Subject":["1057"],
     "Location":[""]
   },
   "range":{"columnStart":2,"rowStart":12,"rowEnd":12,"columnEnd":5},
   "source":{},
   "authMode":{},
   "triggerUid":9214999304440397000
 }
*/

function eventHandler_(config, arg) {

  try {

    // Any early functionality that is needed
    config[0]()
    
    Log_ = getLog()

    var userEmail = Session.getActiveUser().getEmail()

    initializeAssertLibrary()
    
    // Comment this out as it runs too often on all the calendar checks
//    Log_.info('Handling ' + config[1] + ' from ' + userEmail)

    return config[3](arg)
    
  } catch (error) {

    if (typeof Log_ !== 'undefined') {
      Log_.fine('Caught error: %s', error)
      Assert.handleError(error, config[2], Log_)  
    }
  }
  
  return
  
  // Private Functions
  // -----------------

  /**
   * Get different log objects for production and debug
   *
   * @return {BBLog}
   */

  function getLog() {
  
    var log
    var lock = LockService.getDocumentLock()

    if (PRODUCTION_VERSION) {
        
      var firebaseUrl = PropertiesService
        .getScriptProperties()
        .getProperty('FIREBASE_URL')
        
      var firebaseSecret = PropertiesService
        .getScriptProperties()
        .getProperty('FIREBASE_SECRET')  
        
      if (firebaseUrl !== null && firebaseSecret !== null) {
      
        var userLog = BBLog.getLog({lock: lock}); 
            
        var masterLog = BBLog.getLog({
          sheetId: null, // No GSheet
          displayUserId: BBLog.DisplayUserId.USER_KEY_FULL,
          firebaseUrl: firebaseUrl,
          firebaseSecret: firebaseSecret
        })
      
        log = {
        
          clear: function() {
            userLog.clear();
            masterLog.clear();      
          },
          
          info: function() {
            userLog.info.apply(userLog, arguments);
            masterLog.info.apply(masterLog, arguments);          
          },      
          
          warning: function() {
            userLog.warning.apply(userLog, arguments);
            masterLog.warning.apply(masterLog, arguments);          
          },
    
          severe: function() {
            userLog.severe.apply(userLog, arguments);
            masterLog.severe.apply(masterLog, arguments);          
          }
        }
        
      } else {

        log = BBLog.getLog({lock: lock});         
      }
      
      log.functionEntryPoint = function() {},
      log.fine = function() {}
      log.finer = function() {}
      log.finest = function() {}
            
    } else { // !PRODUCTION_VERSION
    
      var logOptions = {
        lock: lock,
        level: DEBUG_LOG_LEVEL,
        displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES,      
      }
    
      if (SpreadsheetApp.getActiveSpreadsheet() === null) {
        logOptions.sheetId = TEST_SPREADSHEET_ID        
      }

      log = BBLog.getLog(logOptions)      
    }
    
    return log;
    
  } // eventHandler_.getLog()

  /**
   * Initialize the Assert library
   */
   
  function initializeAssertLibrary() {
    
    Log_.functionEntryPoint()
    
    var sendErrorEmail               = false
    var calledFromInstallableTrigger = false
    var adminEmailAddress            = DEV_EMAIL_ADDRESS
    
    if (typeof arg === 'undefined') {
    
      Log_.fine('No arg')

    } else {
    
      Log_.fine('arg: ' + JSON.stringify(arg))      

      // The arg is only defined for triggers - built-in or installable,
      // but we still need to tell the difference as they have different 
      // authority
      
      calledFromInstallableTrigger = arg.hasOwnProperty('triggerUid')
      Log_.fine('calledFromInstallableTrigger: ' + calledFromInstallableTrigger)
      
      if (arg.hasOwnProperty('authMode')) {
      
        Log_.fine('arg.authMode: ' + JSON.stringify(arg.authMode))
        
        if (arg.authMode === ScriptApp.AuthMode.FULL) {
        
          if (userEmail !== '' && typeof userEmail !== 'undefined') {
            adminEmailAddress += ',' + userEmail
          }
            
          sendErrorEmail = true
          Log_.fine('Error emails will sent to ' + adminEmailAddress)
          
        } else {
        
          Log_.fine('Not sending error emails')
        }
        
      } else {
      
        Log_.fine('No authMode')
      } 
    }

    var handleError

    if (calledFromInstallableTrigger) {

      // No point trying to display the error from a installable trigger
      // as it does not run in a UI context, so throw it so the email notification
      // is sent  
      
      handleError = Assert.HandleError.THROW 
      Log_.fine('Called from installable trigger')

    } else {
    
      if (PRODUCTION_VERSION) {
      
        // The error is from a built-in trigger from the UI (onEdit() or onOpen()) so
        // tell the user
        handleError = Assert.HandleError.DISPLAY
        Log_.fine('Display errors')
        
      } else {
      
        Log_.fine('In debug mode, not called from installable trigger')
        handleError = Assert.HandleError.THROW
        Log_.fine('Throw errors')        
      }
    }

    Log_.fine('handleError: ' + handleError)
    Log_.fine('sendErrorEmail: ' + sendErrorEmail)
    Log_.fine('adminEmailAddress: ' + adminEmailAddress)

    Assert.init({
      handleError: handleError,
      sendErrorEmail: sendErrorEmail, 
      emailAddress: adminEmailAddress,
      scriptName: SCRIPT_NAME,
      scriptVersion: SCRIPT_VERSION, 
    })
    
  } // eventHandler_.initializeAssertLibrary()
  
} // eventHandler_()

/**
 * Event handler for the sheet being opened. 
 *
 * This is a special case as all it can do is create a menu whereas the 
 * usual eventHandler_() does things we don't have permission for at 
 * this stage
 */

function onOpen_(event) {
    
  if (TEST_FORCE_OPEN_ERROR) {
    throw new Error('Force onOpen() error for testing.')
  }

  var menu = SpreadsheetApp
    .getUi()
    .createMenu(SCRIPT_NAME)
    .addItem('Send status email', 'onEmailStatusUpdates')
    .addItem('Check calendar for tasks', 'onCalendarTrigger')
    
  if (event) {
  
    if (event.authMode !== ScriptApp.AuthMode.NONE) {
   
      var calendarTriggerId = PropertiesService
        .getDocumentProperties()
        .getProperty(PROPERTY_CALENDAR_TRIGGER_ID)
      
      if (calendarTriggerId === null) {
        
        menu.addItem('Enable daily calendar check', 'onStartCalendarTrigger')
        
      } else {
        
        menu.addItem('Disable daily calendar check', 'onStopCalendarTrigger');
      }  
    }
  }
    
  // TODO - Add 'display calendar' & 'display form'
  
    menu.addSeparator()
      .addItem('Settings', 'onShowSidebar')
  
  if (!PRODUCTION_VERSION) {
    
    menu
      .addSeparator()
      .addItem('Uninstall',            'onUninstall')
      .addItem('Install again',        'onInstall')
      .addItem('Clear log',            'onClearLog')
      .addItem('Run unit tests',       'test_roseTaskManager')
      .addItem('Throw an error',       'onThrowError')     
      .addItem('Dump config',          'onDumpConfig')     
      .addItem('Clear config',         'onClearConfig')
      .addItem('Run calendar trigger', 'onCalendarTrigger')  
      .addItem('Run a test',           'onTest')
  }
  
  menu.addToUi()
    
} // onOpen_()

/**
 * Installation event handler
 */

// TODO - Check that we have permission to run this
// TODO - Create formatted log sheet

function onInstall_() {

  Log_.functionEntryPoint()
  var functionName = 'onInstall_()'
  
  Log_.info('Installing...')
  
  var properties = PropertiesService.getDocumentProperties()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()    
  
  if (TEST_FORCE_INSTALL_ERROR) {
    throw new Error('Force onInstall() error for testing.')
  }
  
  // A list of functions to call to do the installation
  
  var INSTALLATION_FUNCTIONS = [
    [null,                  null,                                   'Installing ...'],
    [createRequestForm,     null,                                   'Installing ... created form'],
    [setupTaskListSheet,    null,                                   'Installing ... created task list'],
    [createCalendar,        null,                                   'Installing ... created calendar'],
    [onOpen_,               {authMode: ScriptApp.AuthMode.LIMITED}, 'Installation finished.'],
  ]

  INSTALLATION_FUNCTIONS.forEach(function(row) {
  
    if (row[0] !== null) {   
      var arg = (row[1] !== null) ? row[1] : 'undefined'
      row[0](arg)
    }
    
    if (row[2] !== null) {
      Utils_.toast(row[2])
    }
  })

  return
  
  // Private Functions
  // -----------------

  /**
   * Create the task request form and an 'on submit' trigger.
   */
   
  function createRequestForm() {
  
    Log_.functionEntryPoint()
    
    // TODO - This takes 8s!
    
    // Take a copy of the form template
    var formId = DriveApp
      .getFileById(FORM_TEMPLATE_ID)
      .makeCopy(SCRIPT_NAME + ' - ' + TASK_REQUEST_FORM_NAME)
      .getId()
    
    var form = FormApp.openById(formId)
    
    if (form === null) {    
      throw new Error('The request form could not be created.') 
    }
        
    form
      .setDestination(
        FormApp.DestinationType.SPREADSHEET,
        spreadsheet.getId()) 
      .setTitle(
        SCRIPT_NAME + ' - ' + TASK_REQUEST_FORM_NAME)

    // Set up 'on form submit' trigger.
    
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      
      if (trigger.getHandlerFunction() === 'onFormSubmit') {
        
        ScriptApp.deleteTrigger(trigger)
        Log_.warning('deleted old form submit trigger')
      }
    })

    if (!TEST_BLOCK_TRIGGERS) {

      ScriptApp
        .newTrigger('onFormSubmit')
        .forSpreadsheet(spreadsheet)
        .onFormSubmit()
        .create()
    
    } else {
    
      Log_.warning('Form submission trigger creation disabled')
    }
    
    Log_.info('Task request form and trigger created')
    
  } // onInstall.createRequestForm()
  
  /**
   * Format the task list spreadsheet so it can hold other data other
   * than the form responses.
   */
  
  // TODO - Break this into private functions.
  
  function setupTaskListSheet() {
  
    Log_.functionEntryPoint()

    var now = new Date().toString()    
    
    // The new form responses sheet has just been created so it'll 
    // be the first sheet.
    var taskListSheet = spreadsheet.getSheets()[0]
    var maxTaskRows = taskListSheet.getMaxRows()

    SpreadsheetApp.setActiveSheet(taskListSheet)

    // Wait for form to finish with responses sheet
    // --------------------------------------------

    // Rename the timestamp column. This is done before the 
    // wait on the response sheet as it'll be used again in the wait

    // Wait 30s for the responses sheet created by the form to be completed 
    // before updating the task list sheet.
    
    Log_.fine('start to wait')

    var count = 0
    var rangeG1 = taskListSheet.getRange('G1')    
    var finished = (rangeG1.getValue() === TASK_LIST_COLUMNS.NOTES)

    while (count <= 60 && !finished) {

      // Finished once the last column, TASK_LIST_COLUMNS.NOTES, has been written.
      finished = (rangeG1.getValue() === TASK_LIST_COLUMNS.NOTES)

      // Rename the 'Timestamp' column header again to force a write 
      // to the sheet, flush the sheet and wait
      taskListSheet.getRange('A1').setValue(TASK_LIST_COLUMNS.TIMESTAMP)
      SpreadsheetApp.flush()
      Utilities.sleep(500)
      count++
    }

    Log_.fine('end of wait, count: ' + count)
    
    if (count > 60) {
    
      throw new Error(
        'Timed out waiting for the form to be created, ' + 
        'please try again.')
    }

    // Rename the form's response sheet
    // --------------------------------

    var existingTaskListSheet = spreadsheet
      .getSheetByName(TASK_LIST_WORK_SHEET_NAME)

    if (existingTaskListSheet) {
    
      existingTaskListSheet.setName(TASK_LIST_WORK_SHEET_NAME + ' ' + now)

      Dialog.show(
        REGULAR_TASK_CALENDAR_NAME, 
        'An existing sheet called "' + TASK_LIST_WORK_SHEET_NAME + '" was found' + 
          'during installation. It has been renamed and a new one created',
        200)
    
      Log_.warning('Renamed existing "' + TASK_LIST_WORK_SHEET_NAME + '" sheet')    
    }

    taskListSheet.setName(TASK_LIST_WORK_SHEET_NAME)

    // Update the formatting of the form's response sheet
    // --------------------------------------------------

    // Insert the ID in the first row
    taskListSheet
      .insertColumnBefore(1)
      .getRange('A1')
      .setValue(TASK_LIST_COLUMNS.ID)

    // Insert two new columns after 'Timestamp'.
    taskListSheet
      .insertColumnsAfter(2, 2)
      .getRange('C1:D1')
      .setValues([[TASK_LIST_COLUMNS.STARTED, TASK_LIST_COLUMNS.CLOSED]])

    // Insert the three new columns after 'Priority'.
    taskListSheet
      .insertColumnsAfter(7, 3)
      .getRange(1, 8, 1, 3)
      .setValues([[TASK_LIST_COLUMNS.STATUS, 
                   TASK_LIST_COLUMNS.CATEGORY, 
                   TASK_LIST_COLUMNS.ASSIGNED_TO]])
  
    // Bold and center the headers.
    taskListSheet
      .getRange(1, 1, 1, 13)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
    
    // Resize the columns
    taskListSheet
      .setColumnWidth(1, 35) // ID
      .setColumnWidth(2, DATE_TIME_COLUMN_WIDTH) // Opened
      .setColumnWidth(3, DATE_TIME_COLUMN_WIDTH) // Started
      .setColumnWidth(4, DATE_TIME_COLUMN_WIDTH) // Closed
      .setColumnWidth(7, 95) // Priority  
      .setColumnWidth(8, 95) // Status
  
    // Create a 'fields' sheet for data validation
    // -------------------------------------------
    
    var templateSheet = SpreadsheetApp
      .openById(FIELDS_TEMPLATE_ID)
        .getSheetByName(FIELDS_TEMPLATE_NAME)

    var templateValues = templateSheet.getRange('A1:B11').getValues()

    var existingSheet = spreadsheet
      .getSheetByName(FIELDS_TEMPLATE_NAME)

    if (existingSheet) {
    
      existingSheet.setName(FIELDS_TEMPLATE_NAME + ' ' + now)
      
      Dialog.show(REGULAR_TASK_CALENDAR_NAME, 
            'An existing fields sheet was found during installation. ' + 
              'It has been renamed and a new one created')
    
      Log_.warning('renamed existing fields sheet')    
    }

    var fieldsSheet = spreadsheet.insertSheet(FIELDS_TEMPLATE_NAME, 1)

    fieldsSheet.getRange('A1:B11').setValues(templateValues)
  
    fieldsSheet.getRange('A1').setFontWeight('bold')
    fieldsSheet.getRange('A6').setFontWeight('bold')    
    
    // Set up priority column data validation and formatting 
    // -----------------------------------------------------

    // TODO - This sets it up for the first 1000, but not after that.

    var fieldsRange = fieldsSheet.getRange('A2:A4')
    var tasksRange = taskListSheet.getRange(2, 7, maxTaskRows)
    
    // TODO - This takes 3s
    
    var rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(fieldsRange)
      .build()
      
    var rules = tasksRange.getDataValidations()
    var i, j

    for (i = 0; i < rules.length; i++) {
    
      for (j = 0; j < rules[i].length; j++) {
      
        rules[i][j] = rule
      }
    }
  
    tasksRange.setDataValidations(rules)
    
    // Set up status column data validation and formatting
    // ---------------------------------------------------
      
    fieldsRange = fieldsSheet.getRange('A7:A11')
    tasksRange = taskListSheet.getRange(2, 8, maxTaskRows)

    // TODO - This sets it up for the first 1000, but not after that.

    rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(fieldsRange)
      .build()
      
    rules = tasksRange.getDataValidations()
    
    for (i = 0; i < rules.length; i++) {   
      for (j = 0; j < rules[i].length; j++) {
        rules[i][j] = rule
      }
    }
    
    tasksRange.setDataValidations(rules)
  
    // TODO - Conditional formatting.
    
    // Add meta data 
    // -------------
    //
    // This allows the location of the columns to be found even if the header
    // names (especially for different languages) are changed or the columns 
    // are moved
    
    var headers = taskListSheet
      .getRange(1, 1, 1, taskListSheet.getLastColumn())
      .getValues()[0]
      
    var columnIndex = 0
    
    for (var key in TASK_LIST_COLUMNS) {    
      if (TASK_LIST_COLUMNS.hasOwnProperty(key)) {
        MetaData_.add(taskListSheet, TASK_LIST_COLUMNS[key], columnIndex++)
      }
    }
    
    Log_.fine('Initialised meta data on' + columnIndex + 'columns')
    
    SpreadsheetApp.setActiveSheet(taskListSheet)
    Log_.info('"' + TASK_LIST_WORK_SHEET_NAME + '" sheet setup')
    
  } // onInstall_.setupTaskListSheet()
  
  /**
   * Create Calendar
   */
   
  function createCalendar() {
    
    Log_.functionEntryPoint()
    Calendar_.create()
  
  } // onInstall_.createCalendar()

  /**
   * Display instructions once the add-on is installed.
   */

  function displayInstructions() {
  
    // TODO - Use jQuery to add a link to the form that has been created.
    // TODO - Use Display library
  
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Help.html')
    
    SpreadsheetApp
      .getUi()
      .showModalDialog(htmlOutput, SCRIPT_NAME)
  
  } // onInstall.displayInstructions()
    
} // onInstall()

/** 
 * Send an email status update
 */

function onEmailStatusUpdates_() {
  
  Log_.functionEntryPoint()  

  if (TEST_FORCE_STATUS_SEND_ERROR) {
    throw new Error('Force onEmailStatusUpdates_ Error')
  }
    
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var inSpreadsheet = true
  
  if (spreadsheet === null && !PRODUCTION_VERSION) {
    inSpreadsheet = false
  }

  // Check the users row selection
    
  var taskListSheet
  var rowNumber

  if (inSpreadsheet) {
  
    taskListSheet = spreadsheet.getActiveSheet()
    
    if (taskListSheet.getName() !== TASK_LIST_WORK_SHEET_NAME) {    
      throw new Error(ERROR_SELECT_VALID_EMAIL)
    }
    
    rowNumber = taskListSheet.getActiveRange().getRowIndex()
    Log_.fine("Working on row number: " + rowNumber)
    
    if (rowNumber <= 1) {
      throw new Error(ERROR_SELECT_VALID_EMAIL)
    }  
    
  } else {
  
    spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID)
    taskListSheet = spreadsheet.getSheetByName(TASK_LIST_WORK_SHEET_NAME)
    rowNumber = 2 
  }

  var numberOfColumns = taskListSheet.getLastColumn()
  var headerRow = taskListSheet.getRange(1, 1, 1, numberOfColumns).getValues()[0]
  var activeRow = taskListSheet.getRange(rowNumber, 1, 1, numberOfColumns).getValues()[0]
  
  // Get all the user data for this row into an object
  
  var textData = {}
  
  for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {
  
    var header = headerRow[columnIndex]
    var nextValue = activeRow[columnIndex]
    
    Log_.fine('header: ' + header)
    Log_.fine('nextValue: ' + nextValue)
    
    textData[header] = nextValue.toString()
  }

//  Log_.fine('textData: ' + JSON.parse(textData))

  var emailColumnIndex = Utils_.getColumnIndex({
    sheet: taskListSheet,
    columnName: TASK_LIST_COLUMNS.CONTACT_EMAIL,
    required: true,
    headers: headerRow,
  })

  var recipient = activeRow[emailColumnIndex].trim()

  if (recipient === '') {
    throw new Error('This row has no email address')
  }

  // Get the email templates
  
  var templates = Utils_.getEmailTemplates(
    PROPERTY_STATUS_EMAIL, 
    STATUS_SUBJECT_TEMPLATE, 
    STATUS_BODY_TEMPLATE)

  // Send the status email
  
  var subject   = Utils_.fillInTemplate(templates.subject, textData)
  var htmlBody  = Utils_.fillInTemplate(templates.htmlBody, textData)
  var plainBody = Utils_.fillInTemplate(templates.plainBody, textData)
  
  var emailFrom = PropertiesService
    .getDocumentProperties()
    .getProperty(PROPERTY_EMAIL_FROM) || SCRIPT_NAME
      
  var options = {
    name: emailFrom, 
    cc: Utils_.getListAdminEmail(),
    htmlBody: htmlBody}

  if (Utils_.sendEmail(recipient, subject, plainBody, options) && inSpreadsheet) {
    Dialog.show(SCRIPT_NAME, 'Status email sent to ' + recipient, 90)
  }
  
  Log_.info('Sent email status')
  
} // onEmailStatusUpdates_()

/**
 * Event handler called when the form is submitted. Multiple submissions
 * may have been made since this was last called.
 */

// TODO - Check that the user hasn't manually entered a conflicting number. 
// Could do that in onEdit() too 

// TODO (#105) - If a manual entry is made the next auto submission doesn't get an ID.

function onFormSubmit_(event) {

  Log_.functionEntryPoint()
  
  if (Utils_.checkIfAuthorizationRequired()) {
    
    // No point continuing, an email to re-auth has already been sent to 
    // the user
    Log_.warning('Script needs to be re-authorized (on form submission)')
    return
  }

  Log_.info('Form submission')

  var functionName = 'onFormSubmit_()'
  
  if (TEST_FORCE_FORMSUBMIT_ERROR) {    
    throw new Error('Force onEmailStatusUpdates() error for testing.')
  }
    
  var taskListSheet = event.range.getSheet()
  
  if (taskListSheet.getName() !== TASK_LIST_WORK_SHEET_NAME) {
  
    Log_.warning(
      'Form submission did not come from "' + TASK_LIST_WORK_SHEET_NAME + '", ' + 
        'but "' + taskListSheet.getName() + '"')
        
    return
  }
  
  var allRows = taskListSheet.getDataRange().getValues()
  var headers = allRows[0]
  
  var columnIndex = {
    id:        getColumnIndex(TASK_LIST_COLUMNS.ID),
    timestamp: getColumnIndex(TASK_LIST_COLUMNS.TIMESTAMP, false),
    started:   getColumnIndex(TASK_LIST_COLUMNS.STARTED),
    closed:    getColumnIndex(TASK_LIST_COLUMNS.CLOSED),
    title:     getColumnIndex(TASK_LIST_COLUMNS.TITLE),
    status:    getColumnIndex(TASK_LIST_COLUMNS.STATUS),
    email:     getColumnIndex(TASK_LIST_COLUMNS.CONTACT_EMAIL)
  }

  Log_.fine("columnIndex: %s", columnIndex)

  var nextIdObject = getNextId()
  
  Log_.fine("nextIdObject: %s", nextIdObject)
  
  var nextId = nextIdObject.nextId
  var rowIndex = nextIdObject.rowIndex
  var maxRowIndex = nextIdObject.maxRowIndex
  var lastRowIndexWithId = nextIdObject.lastRowIndexWithId
  
  var countUpdatedRows = 0
    
  Log_.fine('Before processing rows, rowIndex: ' + rowIndex)
  
  // Now process the new rows
  for (; rowIndex <= maxRowIndex; rowIndex++, nextId++) {
    
    // Add the ID and status
    // ---------------------
    
    allRows[rowIndex][columnIndex.id] = nextId
    countUpdatedRows++
      
    Log_.fine('set row id to ' + nextId)
      
    allRows[rowIndex][columnIndex.status] = STATUS_NEW
      
    // Update timestamp formatting
    // ---------------------------
    //
    // Get the column numbers of the three timestamp values and update 
    // their formatting
    
    if (columnIndex.timestamp !== -1) {    
      taskListSheet
        .getRange(rowIndex + 1, columnIndex.timestamp + 1)
        .setNumberFormat(DATE_TIME_FORMAT)
    }
    
    taskListSheet
      .getRange(rowIndex + 1, columnIndex.started + 1)
      .setNumberFormat(DATE_TIME_FORMAT)
    
    taskListSheet
      .getRange(rowIndex + 1, columnIndex.closed + 1)
      .setNumberFormat(DATE_TIME_FORMAT)
    
    // Construct and send the response email to user
    // ---------------------------------------------
    
    var title = allRows[rowIndex][columnIndex.title]
    
    // Get the draft templates
    
    var templates = Utils_.getEmailTemplates(
      PROPERTY_NEW_TASK_EMAIL, 
      FORM_SUBJECT_TEMPLATE, 
      FORM_BODY_TEMPLATE)
      
    // Send the status email

    var textData = getTextDate()

    var subject   = Utils_.fillInTemplate(templates.subject, textData)
    var htmlBody  = Utils_.fillInTemplate(templates.htmlBody, textData)
    var plainBody = Utils_.fillInTemplate(templates.plainBody, textData)
    
    var emailFrom = PropertiesService
      .getDocumentProperties()
      .getProperty(PROPERTY_EMAIL_FROM) || SCRIPT_NAME
      
    var options = {
      htmlBody: htmlBody,
      name: emailFrom}

    var email = allRows[rowIndex][columnIndex.email]

    if (email.indexOf('@') === -1) {
      
      Log_.warning('no email sent to user, invalid address: ' + email)            
      
      // No user email address
      email = Utils_.getListAdminEmail()
      plainBody += '\n\nNO EMAIL SENT TO USER as no valid email address found.'
      
    } else {
      
      options.cc = Utils_.getListAdminEmail()
    }
    
    MailApp.sendEmail(email, subject, plainBody, options)
    
    Log_.info('"New task" Email sent to ' + email + ' - subject: ' + subject)
    
  } // Each new row
  
  var updatedRows = allRows.slice(lastRowIndexWithId + 1, maxRowIndex + 1)
  
  // and finally write the updated rows back to the sheet
  
  var maxColumns = allRows[0].length 
    
  Log_.fine(
    'lastRowIndexWithId: ' + lastRowIndexWithId + ', ' +
    'countUpdatedRows: ' + countUpdatedRows + ', ' +
    'maxColumns' + maxColumns + ', ' +
    'updatedRows: ' + updatedRows)
  
  if (countUpdatedRows >= 1) {    
    
    taskListSheet
      .getRange(lastRowIndexWithId + 2, 1, countUpdatedRows, maxColumns)
      .setValues(updatedRows)
    
    SpreadsheetApp.flush()
    
  } else {
    
    // Every submission triggers an onFormSubmit() but multiple
    // submissions can be processed in one, leaving nothing for later
    // triggers
    Log_.fine('Called with nothing to update')
  }
    
  return
  
  // Private Functions
  // -----------------

  /**
   * Get the index (0-based) of this column
   * 
   * @return {Number} column index
   */

  function getColumnIndex(columnName, required) {
    
    return Utils_.getColumnIndex({
      sheet: taskListSheet,
      columnName: columnName,
      required: required,
      headers: headers,
    })
    
  } // onFormSubmit_.getColumnIndex()

  /**
   * Convert the row of data containing the new task into an object
   * 
   * @return {Object} data
   */
   
  function getTextDate() {
    
    Log_.functionEntryPoint()
    
    var data = {}
    var row = allRows[rowIndex]
    var headers = allRows[0]
    
    for (var columnIndex = 0; columnIndex < headers.length; columnIndex++) {
      var header = headers[columnIndex]
      data[header] = row[columnIndex]
    }
    
    Log_.fine('data: %s', data)
    return data
          
  } // onFormSubmit_.getTextDate()

  /**
   * @return {number} the next ID to use or null
   */

  function getNextId() {
  
    Log_.functionEntryPoint()
  
    // Get a unique task id and store it in the spreadsheet. Read in all 
    // of the sheet, find the highest ID number used so far then starting 
    // from the last row work work up to find the last empty ID cell, then 
    // work down filling in each empty cell with a unique ID. This feels a bit
    // convoluted but allows for multiple form responses being written between
    // calls to this function, and avoids having to go elsewhere, like a 
    // script property, to get the index which has its own issues.
    
    var ids = ArrayLib.transpose(allRows)[columnIndex.id].filter(function(id) {  
      return typeof id === 'number'
    })
    
    var maxId = (ids.length > 0) ? Math.max.apply(null, ids) : 1
    
    var maxRowIndex = allRows.length - 1
    var lastRowIndexWithId = 0
    var nextId
    var rowIndex = maxRowIndex
    
    // Work backwards to find the last row with an ID
    for (; rowIndex >= 1; rowIndex--) {
      
      nextId = allRows[rowIndex][columnIndex.id]
      
      if (typeof nextId === 'number') {
        
        lastRowIndexWithId = rowIndex
        break
      }
    }
    
    // Step forward again into the empty id
    rowIndex++
      
    return {
      nextId: maxId + 1,
      rowIndex: rowIndex,
      maxRowIndex: maxRowIndex,
      lastRowIndexWithId: lastRowIndexWithId,
    }
  
  } // onFormSubmit_.getNextId()

} // onFormSubmit_()

/**
 * Event handler for the form being edited
 */

function onEdit_(event) {

  Log_.functionEntryPoint()
  var functionName = 'onEdit_()'
  
  if (TEST_FORCE_EDIT_ERROR) {    
    throw new Error('Force onEdit() error for testing.')
  }
  
  // Get some info about the event.
  var sheet = event.source.getActiveSheet()
  var range = event.source.getActiveRange()
  
  if (sheet.getName() !== TASK_LIST_WORK_SHEET_NAME) {
    Log_.warning('Ignoring edit: not in "' + TASK_LIST_WORK_SHEET_NAME + '" sheet')
    return
  }
  
  var columnNumber = range.getColumn()
  var rowNumber = range.getRow()
  
  Log_.fine("row:" + rowNumber + " col:" + columnNumber)
  
  // Record the "started" and "closed" date
  // --------------------------------------
  
  var statusColumnNumber = Utils_.getColumnIndex({
    sheet: sheet, 
    columnName: TASK_LIST_COLUMNS.STATUS,
    useMeta: false,
    required: false}) + 1
  
  if (columnNumber === statusColumnNumber) {
    
    var value = range.getValue()
    var now = new Date()
    
    Log_.fine("changed value = " + value)
    
    if (!getCell(TASK_LIST_COLUMNS.TIMESTAMP)) {
      
      Log_.warning('No task in this column (no started date/time)')
      return
    }
    
    // Deliberately not clearing the timestamps if the status is returned
    // to 'new' leaving the user to do that in case they need those times.
    
    if (value === STATUS_IGNORED || value === STATUS_DONE) {
      
      if (!getCell(TASK_LIST_COLUMNS.CLOSED)) {
        
        setCell(TASK_LIST_COLUMNS.CLOSED, now)
        
        if (!getCell(TASK_LIST_COLUMNS.STARTED)) {
          
          setCell(TASK_LIST_COLUMNS.STARTED, now)
        }
        
      } else {
        
        Log_.warning('There is already a date/time in closed')
      }
    } 
    
    if (value === STATUS_IN_PROGRESS) {
      
      if (!getCell(TASK_LIST_COLUMNS.STARTED)) {
       
        setCell(TASK_LIST_COLUMNS.STARTED, now)
        
      } else {
        
        Log_.warning('There is already a date/time in started')        
      }
    } 
  }
  
  // Private Functions
  // -----------------
  
  function getCell(columnName) {

    return Utils_.getCellValue(
      sheet, 
      rowNumber, 
      columnName, 
      false, // DONT_USE_META
      false) // NOT_REQUIRED
  }

  function setCell(columnName, value) {
  
    Utils_.setCellValue(
      sheet, 
      rowNumber, 
      columnName, 
      value, 
      false, // DONT_USE_META
      false) // NOT_REQUIRED
  }

} // onEdit_()

/**
* Clear the development log sheet.
*/

function onClearLog_() {

  if (PRODUCTION_VERSION) {
    return
  }
  
  Log_.clear()
  
} // onClearLog_()

/**
* Uninstall Rose Task Manager
*/

function onUninstall_() {

  Log_.functionEntryPoint()

  var docProperties = PropertiesService.getDocumentProperties()
  var userProperties = PropertiesService.getUserProperties()  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()    
  
  // Form
  // ----
  
  var formId = Utils_.getFormId()
  
  if (formId !== null) {
    
    FormApp.openById(formId).removeDestination()
    Log_.info('Unlinked form from spreadsheet')
    
    DriveApp.getFileById(formId).setTrashed(true)
    Log_.info('Deleted form')
  }
  
  // Sheets
  // ------
  
  var taskListSheet = spreadsheet.getSheetByName(TASK_LIST_WORK_SHEET_NAME)  
  var fieldsSheet = spreadsheet.getSheetByName(FIELDS_TEMPLATE_NAME)
  //    var emailSheet = spreadsheet.getSheetByName(EMAIL_TEMPLATE_NAME)
  
  if (taskListSheet && spreadsheet.getNumSheets() > 1) {
    spreadsheet.deleteSheet(taskListSheet)
    Log_.info('Deleted task lists sheet')
  }
  
  if (fieldsSheet && spreadsheet.getNumSheets() > 1) {
    spreadsheet.deleteSheet(fieldsSheet)
    Log_.info('Deleted fields sheet')    
  }
  /*  
  if (emailSheet && spreadsheet.getNumSheets() > 1) {
  
  spreadsheet.deleteSheet(emailSheet)
  Log_.info('Deleted email sheet')    
  }
  */    
  // Calendars
  // ---------
  
  var calendars = CalendarApp.getAllCalendars()
  
  calendars.forEach(function(calendar) {
    
    if (calendar.getName() === REGULAR_TASK_CALENDAR_NAME) {      
      calendar.deleteCalendar()
      Log_.info('Deleted calendar')      
    }
  })
    
  // Triggers
  // --------
  
  var triggers = ScriptApp.getProjectTriggers()
  
  triggers.forEach(function(trigger) {  
    ScriptApp.deleteTrigger(trigger)
    Log_.info('Deleted trigger')      
  })

  // Properties
  // ----------

  docProperties.deleteProperty(PROPERTY_CALENDAR_TRIGGER_ID)
  docProperties.deleteProperty(PROPERTY_CALENDAR_ID_NT)
  
  docProperties.deleteProperty(PROPERTY_STATUS_EMAIL)
  docProperties.deleteProperty(PROPERTY_NEW_TASK_EMAIL)
  docProperties.deleteProperty(PROPERTY_EMAIL_FROM)
  
  userProperties.deleteProperty(PROPERTY_LAST_AUTH_EMAIL_DATE)

  // Menu
  // ----
  
  spreadsheet.removeMenu(SCRIPT_NAME)
  
  SpreadsheetApp
    .getUi()
    .createMenu(SCRIPT_NAME)
    .addItem('Install', 'onInstall')
    .addItem('Dump config', 'onDumpConfig')
    .addItem('Clear config', 'onClearConfig')
    .addToUi()
  
  Log_.info('Deleted menu')
  
  Utils_.toast('Uninstallation finished.')
  Log_.info('Uninstallation finished')
  
} // onUninstall_()

/**
 * Private event handler for "Show Sidebar" event
 *
 * @param {Object} 
 * 
 * @return {Object} 
 */
 
function onShowSidebar_() {
  
  Log_.functionEntryPoint()
  
//  var html = HtmlService.createTemplateFromFile('SlidesMerge_Test').evaluate().setTitle(SCRIPT_NAME)

  var html = HtmlService
    .createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle(SCRIPT_NAME)
  
  SpreadsheetApp.getUi().showSidebar(html)
  Log_.info('Showing sidebar')

} // onShowSidebar_()

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent()
}

/**
 * Get the settings to display in the UI
 * 
 * @return {Object} settings
 */

function onGetSettings_() {  

  Log_.functionEntryPoint()

  var drafts = [DEFAULT_DRAFT_TEXT]

  if (!TEST_FORCE_NO_DRAFTS) {
    GmailApp.getDraftMessages().forEach(function(draft) {
      drafts.push(draft.getSubject())
    }) 
  }

  if (drafts.length === 0) {
    Log_.fine('No drafts found using defaults')
    drafts = [DEFAULT_DRAFT_TEXT]
  }

  var properties = PropertiesService.getDocumentProperties()
  
  var newTaskEmailSubject = properties.getProperty(PROPERTY_NEW_TASK_EMAIL) || ''
  
  Log_.fine('newTaskEmailSubject: ' + newTaskEmailSubject)

  var statusEmailSubject = properties.getProperty(PROPERTY_STATUS_EMAIL) || ''
  
  Log_.fine('statusEmailSubject: ' + statusEmailSubject)

  var emailFrom = properties.getProperty(PROPERTY_EMAIL_FROM) || SCRIPT_NAME
  Log_.fine('emailFrom: ' + emailFrom)

  return {
    drafts: drafts,
    newTaskEmailSubject: newTaskEmailSubject,
    statusEmailSubject: statusEmailSubject,
    emailFrom: emailFrom,
  }
  
} // onGetSettings_()

// Store settings persistently

function onSetNewFrom_(newValue)         {Utils_.setSetting(PROPERTY_EMAIL_FROM, newValue)}
function onSetNewTaskTemplate_(newValue) {Utils_.setSetting(PROPERTY_NEW_TASK_EMAIL, newValue)}
function onSetStatusTemplate_(newValue)  {Utils_.setSetting(PROPERTY_STATUS_EMAIL, newValue)}