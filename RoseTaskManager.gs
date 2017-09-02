// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 14 Oct 2014 08:49 GMT+1
/* jshint asi: true */

/*
 * This file is part of Rose Task Manager.
 *
 * Copyright (C) 2015 Andrew Roberts (andrew@roberts.net)
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
  onOpen:                  [function() {},  'onOpen()',                 'Failed to start RTM add-on',           onOpen_],
  onEmailStatusUpdates:    [function() {},  'onEmailStatusUpdates()',   'Failed to send email status',          onEmailStatusUpdates_],
  onFormSubmit:            [function() {},  'onFormSubmit()',           'Failed to process form submission',    onFormSubmit_],
  onEdit:                  [function() {},  'onEdit()',                 'Failed to process edit',               onEdit_],
  onCalendarTrigger:       [function() {},  'onCalendarTrigger()',      'Failed to check calendar for tasks',   onCalendarTrigger_],
  onClearLog:              [function() {},  'onClearLog()',             'Failed to clear log',                  onClearLog_],  
  onUninstall:             [function() {},  'onUninstall()',            'Failed to uninstall',                  onUninstall_],  
  updateTrelloBoard:       [function() {},  'updateTrelloBoard()',      'Failed to update Trello boards',       updateTrelloBoard_],  
  onDumpConfig:            [function() {},  'onDumpConfig()',           'Failed to dump config',                onDumpConfig_],  
  onClearConfig:           [function() {},  'onClearConfig()',          'Failed to clear config',               onClearConfig_],  
  onThrowError:            [function() {},  'onThrowError()',           'Deliberatley threw error',             onThrowError_],  
  onStartCalendarTrigger:  [function() {},  'onStartCalendarTrigger()', 'Failed to start calendar trigger',     onStartCalendarTrigger_],    
  onStopCalendarTrigger:   [function() {},  'onStopCalendarTrigger()',  'Failed to stop calendar trigger',      onStopCalendarTrigger_],      
}

// function (arg)                     {return eventHandler_(EVENT_HANDLERS., arg)}

function onInstall(arg)              {return eventHandler_(EVENT_HANDLERS.onInstall, arg)}
function onOpen(arg)                 {return eventHandler_(EVENT_HANDLERS.onOpen, arg)}
function onEmailStatusUpdates(arg)   {return eventHandler_(EVENT_HANDLERS.onEmailStatusUpdates, arg)}
function onFormSubmit(arg)           {return eventHandler_(EVENT_HANDLERS.onFormSubmit, arg)}
function onEdit(arg)                 {return eventHandler_(EVENT_HANDLERS.onEdit, arg)}
function onCalendarTrigger(arg)      {return eventHandler_(EVENT_HANDLERS.onCalendarTrigger, arg)}
function onClearLog(arg)             {return eventHandler_(EVENT_HANDLERS.onClearLog, arg)}
function onUninstall(arg)            {return eventHandler_(EVENT_HANDLERS.onUninstall, arg)}
function updateTrelloBoard(arg)      {return eventHandler_(EVENT_HANDLERS.updateTrelloBoard, arg)}
function onDumpConfig(arg)           {return eventHandler_(EVENT_HANDLERS.onDumpConfig, arg)}
function onClearConfig(arg)          {return eventHandler_(EVENT_HANDLERS.onClearConfig, arg)}
function onThrowError(arg)           {return eventHandler_(EVENT_HANDLERS.onThrowError, arg)}
function onStartCalendarTrigger(arg) {return eventHandler_(EVENT_HANDLERS.onStartCalendarTrigger, arg)}
function onStopCalendarTrigger(arg)  {return eventHandler_(EVENT_HANDLERS.onStopCalendarTrigger, arg)}

// Private Functions
// =================

function onCalendarTrigger_()      {Calendar_.convertEventsToTasks()}
function onStartCalendarTrigger_() {Calendar_.createTrigger()}
function onStopCalendarTrigger_()  {Calendar_.deleteTrigger(true)}
function updateTrelloBoard_()      {Trello.uploadCards()}
function onThrowError_()           {throw new Error('Force error throw')}

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
    
    Log_.info('Handling ' + config[1] + ' from ' + userEmail)

    return config[3](arg)
    
  } catch (error) {
  
    Log_.fine('Caught error: ' + error.name)
  
    if (error.name === 'AuthorizationError') {
    
      // This is a special error thrown by TrelloApp to indicate
      // that user authorization is required
      var authorizationUrl = (new TrelloApp.App()).getAuthorizationUri()
      
      Dialog.show(
        'Opening Authorization window...', 
          'Follow the instructions in this window, close ' + 
          'it and then try the action again.<script>window.open("' + authorizationUrl + '")</script>',
          100)
          
    } else {
    
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
  
      var userLog = BBLog.getLog({lock: lock}); 
      
      var firebaseUrl = PropertiesService
        .getScriptProperties()
        .getProperty('FIREBASE_URL')
        
      var firebaseSecret = PropertiesService
        .getScriptProperties()
        .getProperty('FIREBASE_SECRET')  

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
        } ,      
        
        warning: function() {
          userLog.warning.apply(userLog, arguments);
          masterLog.warning.apply(masterLog, arguments);          
        },
  
        severe: function() {
          userLog.severe.apply(userLog, arguments);
          masterLog.severe.apply(masterLog, arguments);          
        },
        
        functionEntryPoint: function() {},
        fine: function() {},
        finer: function() {},
        finest: function() {},     
      }
      
    } else { // !PRODUCTION_VERSION
    
      log = BBLog.getLog({
        lock: lock,
        level: DEBUG_LOG_LEVEL,
        displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES,
      }); 
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
 * Installation event handler
 */

// TODO - Check that we have permission to run this
// TODO - Create formatted log sheet

function onInstall_() {

  Log_.functionEntryPoint()
  var functionName = 'onInstall_()'
  
  var properties = PropertiesService.getDocumentProperties()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()    
  
  if (FORCE_INSTALL_ERROR) {
    throw new Error('Force onInstall() error for testing.')
  }
  
  Utils_.postAnalytics(functionName)

  // A list of functions to call to do the installation
  
  var INSTALLATION_FUNCTIONS = [
    [null,                'Installing ...'],
    [createRequestForm,   'Installing ... created form'],
    [setupTaskListSheet,  'Installing ... created task list'],
    [createCalendar,      'Installing ... created calendar'],
    [onOpen,              'Installation finished.'],
    [displayInstructions, null],
  ]

  INSTALLATION_FUNCTIONS.forEach(function(row) {
  
    if (row[0] !== null) {
      row[0]()
    }
    
    if (row[1] !== null) {
      Utils_.toast(row[1])
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

    if (CREATE_TRIGGERS) {

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
//    taskListSheet.getRange('A1').setValue(SS_COL_LISTED)

    // Wait 30s for the responses sheet created by the form to be completed 
    // before updating the task list sheet.
    
    Log_.fine('start to wait')

    var count = 0
    var rangeG1 = taskListSheet.getRange('G1')    
    var finished = (rangeG1.getValue() === SS_COL_NOTES)

    while (count <= 60 && !finished) {

      // Finished once the last column, SS_COL_NOTES, has been written.
      finished = (rangeG1.getValue() === SS_COL_NOTES)

      // Rename the 'Timestamp' column header again to force a write 
      // to the sheet, flush the sheet and wait
      taskListSheet.getRange('A1').setValue(SS_COL_TIMESTAMP)
      SpreadsheetApp.flush()
      Utilities.sleep(500)
      count++
    }

    Log_.fine('end of wait, count: ' + count)
    
    if (count > 60) {
    
      throw new Error('Timed out waiting for the form to be created, ' + 
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
      .setValue(SS_COL_ID)

    // Insert two new columns after 'Timestamp'.
    taskListSheet
      .insertColumnsAfter(2, 2)
      .getRange('C1:D1')
      .setValues([[SS_COL_STARTED, SS_COL_CLOSED]])

    // Insert the three new columns after 'Priority'.
    taskListSheet
      .insertColumnsAfter(7, 3)
      .getRange(1, 8, 1, 3)
      .setValues([[SS_COL_STATUS, SS_COL_CATEGORY, SS_COL_ASSIGNED_TO]])
  
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
  
    // TODO - Conditional formatting.
    
//    templateSheet
//      .getRange('A2:A4')
//      .copyFormatToRange(taskListSheet, 7, 7, 2, maxTaskRows)
  
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
    
//    templateSheet
//      .getRange('A7:A11')
//      .copyFormatToRange(taskListSheet, 8, 8, 2, maxTaskRows)
  
    // Setup the email templates sheet
    // -------------------------------

/* TODO 

    templateSheet = SpreadsheetApp
      .openById(EMAIL_TEMPLATE_ID)
        .getSheetByName(EMAIL_TEMPLATE_NAME)

    templateValues = templateSheet.getRange('A1:B8').getValues()

    existingSheet = spreadsheet.getSheetByName(EMAIL_TEMPLATE_NAME)

    if (existingSheet) {
    
      existingSheet.setName(EMAIL_TEMPLATE_NAME + ' ' + now)
      
      Dialog.show(REGULAR_TASK_CALENDAR_NAME, 
            'An existing email template sheet was found during installation. ' + 
              'It has been renamed and a new one created')
    
      Log_.warning('renamed existing email template sheet')    
    }

    var emailTemplateSheet = spreadsheet.insertSheet(EMAIL_TEMPLATE_NAME, 1)

    emailTemplateSheet.getRange('A1:B8').setValues(templateValues)
    
    emailTemplateSheet.autoResizeColumn(1)
    
    emailTemplateSheet
      .getRange(1, 1, emailTemplateSheet.getMaxRows(), 1)
      .setFontWeight('bold')
*/

    // Initialise the task ID
    // ----------------------

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
 * Event handler for the sheet being opened
 */

function onOpen_(event) {

  Log_.functionEntryPoint('event: ' + JSON.stringify(event))
  var functionName = 'onOpen_()'
  
  if (FORCE_OPEN_ERROR) {
    throw new Error('Force onOpen() error for testing.')
  }

  Utils_.postAnalytics(functionName)

  var ui = SpreadsheetApp.getUi()
  var menu = PRODUCTION_VERSION ? ui.createAddonMenu() : ui.createMenu(SCRIPT_NAME)
  menu.addItem('Send status email', 'onEmailStatusUpdates')
  menu.addItem('Check calendar for tasks', 'onCalendarTrigger')
    
  if (event) {
  
    Log_.fine('event.authMode: ' + event.authMode)
  
    if (event.authMode !== ScriptApp.AuthMode.NONE) {
   
      // Add a menu item based on properties (doesn't work in AuthMode.NONE)
      var properties = PropertiesService.getDocumentProperties()
      var calendarTriggerId = properties.getProperty(PROPERTY_CALENDAR_ID)
      
      if (calendarTriggerId === null) {
        
        menu.addItem('Enable daily calendar check', 'onStartCalendarTrigger')
        
      } else {
        
        menu.addItem('Disable daily calendar check', 'onStopCalendarTrigger');
      }
    
    } else {
    
      Log_.fine('No auth to check for calendar trigger')
    }
    
  } else {
  
    Log_.fine('No event')
  }
    
  // TODO - Add 'display calendar' & 'display form'
  
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
      .addSeparator()
//      .addItem('Display Available Trello Members', 'displayTrelloMembers')
//      .addItem('Display Available Trello Boards', 'displayTrelloBoards')
//      .addItem('Display Available Trello Lists', 'displayTrelloLists')
      .addItem('Upload tasks to Trello', 'updateTrelloBoard')
  }
  
  menu.addToUi()
    
} // onOpen_()

/**
 * Sort by status and priority
 */

// TODO - Remove these

/*

function onSortByPriortity_() {
  
  Log_.functionEntryPoint()
  
  if (FORCE_SORTBYPRIORITY_ERROR) {
    throw new Error('Force onSortByPriortity() error for testing.')
  }
  
  Utils_.postAnalytics('onSortByPriortity')
  
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
  
  var priorityIndex = Utils_.getColumnPosition(sheet, SS_COL_PRIORITY)
  
  // Sort the sheet by Status
  sheet.sort(priorityIndex, false)
    
  return true

} // onSortByPriortity_
*/
/**
 * Sort the sheet by ID
 */
/* 
function onSortById_() {
  
  Log_.functionEntryPoint()
  var functionName = 'onSortById_()'
  
  if (FORCE_SORTBYID_ERROR) {
    throw new Error('Force onSortById() error for testing.')
  }
  
  Utils_.postAnalytics(functionName)
  
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
  
  var idIndex = Utils_.getColumnPosition(sheet, SS_COL_ID)
  
  // Sort the sheet by date
  sheet.sort(idIndex, true)

} // onSortById_()
*/
/** 
 * Send an email status update
 */

// TODO - remove bool return.

function onEmailStatusUpdates_() {
  
  Log_.functionEntryPoint()
  var functionName = 'onEmailStatusUpdates_()'
  
  var SELECT_VALID_EMAIL = 'Select a row on the "' + TASK_LIST_WORK_SHEET_NAME + '" ' + 
    'sheet with a valid email address  and try again.'
  
  var FAILED_TO_SEND_STATUS = SCRIPT_NAME + ' ' + 
    'Failed to send Status email.'
  
  Utils_.postAnalytics(functionName)
  
  // Get the row number of the row in the spreadsheet that's 
  // currently active (the task list)
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var taskListSheet = spreadsheet.getActiveSheet()
  
  if (taskListSheet.getName() !== TASK_LIST_WORK_SHEET_NAME) {    
    throw new Error(SELECT_VALID_EMAIL)
  }
  
  var row = taskListSheet.getActiveRange().getRowIndex()
  
  if (row <= 1) {    
    throw new Error(SELECT_VALID_EMAIL)
  }
  
  Log_.fine("Working on row: " + row)
  
  // Retrieve the user's info 
  
  var columnNumber = Utils_.getColumnPosition(taskListSheet, SS_COL_CONTACT_EMAIL)
  var userEmail = taskListSheet.getRange(row, columnNumber).getValue()
  
  if (userEmail === "") {
    throw new Error(SELECT_VALID_EMAIL)
  }
  
  columnNumber = Utils_.getColumnPosition(taskListSheet, SS_COL_TITLE)
  var title = taskListSheet.getRange(row, columnNumber).getValue()
  
  columnNumber = Utils_.getColumnPosition(taskListSheet, SS_COL_STATUS)
  var status = taskListSheet.getRange(row, columnNumber).getValue()
  
  columnNumber = Utils_.getColumnPosition(taskListSheet, SS_COL_ID)
  var id = taskListSheet.getRange(row, columnNumber).getValue()
  
  if (status === "") {
    throw new Error('Please select a row with a Status value and try again.')
  }
  
  // Construct the update email and send it
  
  /* 
  TODO - fillTemplate.gs needs a proper review first.
  
  var emailTemplateSheet = spreadsheet.getSheetByName(EMAIL_TEMPLATE_NAME)
  
  if (!emailTemplateSheet) {
  
  throw new Error('Can not find the email template sheet. Reinstall Rose Task ' + 
  'Manager to recreate it.')
  }
  */    
  //    var subjectTemplate = emailTemplateSheet.getRange('B4').getValue()    
  var subjectTemplate = STATUS_SUBJECT_TEMPLATE
  var subjectData = {id:id, status:status}
  var subject = Utils_.fillInTemplateFromObject(subjectTemplate, subjectData)
  
  //    var bodyTemplate = emailTemplateSheet.getRange('B5').getValue()      
  var bodyTemplate = STATUS_BODY_TEMPLATE
  var bodyData = {row:id, title:title, status:status}
  var body = Utils_.fillInTemplateFromObject(bodyTemplate, bodyData)
  
  MailApp.sendEmail(userEmail, 
                    subject, 
                    body, 
                    {name:SCRIPT_NAME, cc:Utils_.getListAdminEmail()})
  
  Log_.info(
    "Email status sent to email: " + userEmail +
    "\n\nsubject: " + subject +
    "\n\nbody: " + body)
  
  Dialog.show(SCRIPT_NAME, 'Status email sent to ' + userEmail, 90)
        
  return true
  
} // onEmailStatusUpdates_()

/**
 * Event handler called when the form is submitted. Multiple submissions
 * may have been made since this was last called.
 */

// TODO - Check that the user hasn't manually entered a conflicting number. 
// Could do that in onEdit() too 

// TODO (#105) - If a manual entry is made the next auto submission doesn't get an ID.

function onFormSubmit_(event) {

  if (Utils_.checkIfAuthorizationRequired()) {
    
    // No point continuing, an email to re-auth has already been sent to 
    // the user
    Log_.warning('Script needs to be re-authorized')
    return
  }

  Log_.functionEntryPoint()
  var functionName = 'onFormSubmit_()'
  
  if (FORCE_FORMSUBMIT_ERROR) {    
    throw new Error('Force onEmailStatusUpdates() error for testing.')
  }
    
  Utils_.postAnalytics(functionName)
  
  var taskListSheet = event.range.getSheet()
  
  if (taskListSheet.getName() !== TASK_LIST_WORK_SHEET_NAME) { 
    Log_.warning('Form submission did not come from "' + TASK_LIST_WORK_SHEET_NAME + '"')
    return
  }
  
  // Rename the timestamp column header
  //    Utils_.setCellValue(taskListSheet, 1, SS_COL_TIMESTAMP, SS_COL_LISTED)
  
  var allRows = taskListSheet.getDataRange().getValues()
  
  var columnIndex = {
    id: getColumnIndex(SS_COL_ID),
    timestamp: getColumnIndex(SS_COL_TIMESTAMP),
    started: getColumnIndex(SS_COL_STARTED),
    closed: getColumnIndex(SS_COL_CLOSED),
    title: getColumnIndex(SS_COL_TITLE),
    status: getColumnIndex(SS_COL_STATUS),
    email: getColumnIndex(SS_COL_CONTACT_EMAIL)
  }
  
  var maxColumns = allRows[0].length
  
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
  
  var maxId = ids.length > 0 ? Math.max.apply(null, ids) : 1
  
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
    
  var countUpdatedRows = 0
  nextId = maxId + 1
  
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
      
      taskListSheet
      .getRange(rowIndex + 1, columnIndex.timestamp + 1)
      .setNumberFormat(DATE_TIME_FORMAT)
      
      taskListSheet
      .getRange(rowIndex + 1, columnIndex.started + 1)
      .setNumberFormat(DATE_TIME_FORMAT)
      
      taskListSheet
      .getRange(rowIndex + 1, columnIndex.closed + 1)
      .setNumberFormat(DATE_TIME_FORMAT)
      
      // Construct and send the response email to user
      // ---------------------------------------------
      
      var title = allRows[rowIndex][columnIndex.title]
      
      // TODO - Get template from sheet - need to review fillTemplate first,
      // looks like there could be overlap or it's too generic
      
      var subject = Utils_.fillInTemplateFromObject(
        FORM_SUBJECT_TEMPLATE, 
        {id:nextId, title:title})    
      
      var body = Utils_.fillInTemplateFromObject(
        FORM_BODY_TEMPLATE, 
        {id:nextId, title:title})
      
      var email = allRows[rowIndex][columnIndex.email]
      
      var options = {}
      
      if (email.indexOf('@') === -1) {
        
        Log_.warning('no email sent to user, invalid address: ' + email)            
        
        // No user email address
        email = Utils_.getListAdminEmail()
        options.name = SCRIPT_NAME
        body += '\n\nNO EMAIL SENT TO USER as no valid email address found.'
        
      } else {
        
        options = {name:SCRIPT_NAME, cc:Utils_.getListAdminEmail()}
      }
    
    MailApp.sendEmail(email, subject, body, options)
    
    Log_.info(
      'Email sent to ' + email + 
      ' - subject: ' + subject + 
      ' body: ' + body)
    
  } // Each new row
  
  var updatedRows = allRows.slice(lastRowIndexWithId + 1, maxRowIndex + 1)
  
  // and finally write the updated rows back to the sheet
  
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
    * @param {string} name
    */
  
  function getColumnIndex(ColumnName) {
    
    Log_.functionEntryPoint()
    var columnIndex = allRows[0].indexOf(ColumnName)
    
    Assert.assert(
      columnIndex >= 0, 
      'Can not find the "' + ColumnName + '" column in the "' + 
      TASK_LIST_WORK_SHEET_NAME + '" tab. This is required by the Rose Task ' + 
      'Manager add-on, could it have been renamed?')
      
    return columnIndex
    
  } // getColumnIndex()
    
} // onFormSubmit_()

/**
 * Event handler for the form being edited
 */

function onEdit_(event) {

  Log_.functionEntryPoint()
  var functionName = 'onEdit_()'
  
  if (FORCE_EDIT_ERROR) {    
    throw new Error('Force onEdit() error for testing.')
  }
  
  Utils_.postAnalytics(functionName)
  
  // Get some info about the event.
  var sheet = event.source.getActiveSheet()
  var range = event.source.getActiveRange()
  
  if (sheet.getName() !== TASK_LIST_WORK_SHEET_NAME) {
    
    Log_.warning('Not in "' + TASK_LIST_WORK_SHEET_NAME + '" sheet')
    return
  }
  
  var colIndex = range.getColumn()
  var rowIndex = range.getRow()
  
  Log_.fine("row:" + rowIndex + " col:" + colIndex)
  
  // Record the "started" and "closed" date
  // --------------------------------------
  
  if (colIndex === Utils_.getColumnPosition(sheet, SS_COL_STATUS)) {
    
    var value = range.getValue()
    var now = new Date()
    
    Log_.fine("changed value = " + value)
    
    if (!Utils_.getCellValue(sheet, rowIndex, SS_COL_TIMESTAMP)) {
      
      Dialog.show(SCRIPT_NAME, 'There is no task in this row.')
      Log_.warning('no task in this column (no started date/time)')
      return
    }
    
    // Deliberately not clearing the timestamps if the status is returned
    // to 'new' leaving the user to do that in case they need those times.
    
    if (value === STATUS_IGNORED || value === STATUS_DONE) {
      
      if (!Utils_.getCellValue(sheet, rowIndex, SS_COL_CLOSED)) {
        
        Utils_.setCellValue(sheet, rowIndex, SS_COL_CLOSED, now)
        
        if (!Utils_.getCellValue(sheet, rowIndex, SS_COL_STARTED)) {
          
          Utils_.setCellValue(sheet, rowIndex, SS_COL_STARTED, now)
        }
        
      } else {
        
        Log_.warning('There is already a date/time in closed')
      }
    } 
    
    if (value === STATUS_IN_PROGRESS) {
      
      if (!Utils_.getCellValue(sheet, rowIndex, SS_COL_STARTED)) {
        
        Utils_.setCellValue(sheet, rowIndex, SS_COL_STARTED, now)
        
      } else {
        
        Log_.warning('There is already a date/time in started')        
      }
    } 
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

  if (PRODUCTION_VERSION) {
    return
  }
    
  var properties = PropertiesService.getDocumentProperties()
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
  
  properties.deleteProperty(PROPERTY_CALENDAR_ID)

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