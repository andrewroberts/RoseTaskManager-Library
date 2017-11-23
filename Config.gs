// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint (this file) - 7 Oct 2014 19:26 GMT+1

// Code review all files - TODO
// JSHint review (see files) - TODO
// Unit Tests - TODO
// System Test (Dev) - TODO
// System Test (Prod) - TODO

/*
 * This file is part of Rose Task Manager.
 *
 * Copyright (C) 2014 - 2017 Andrew Roberts
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
 * Config.gs
 * =========
 *
 * Internal configuration settings.
 */

// TODO - Create a local type of error so that only errors thrown by the 
//   script get passed up, deeper ones could be 'softened'. throwError_() with
//   function name is used in assert.gs.
// TODO - Look at what can be brought out to a Settings window.
// TODO - Think about FINE vs INFO.
// TODO - Set default priority (?).
// TODO - Add my installing code to answer on SO (?).
// TODO - What about an error log sheet for exceptions during production.
// TODO - Create template for help, so RoseTM name can be passed in.
// TODO - Do a performance test with assertions disabled.
// TODO - Should the script know if onInstall has ran?? The menu options are
// always there. State?

// Config
// ======

var SCRIPT_NAME = 'Rose Task Manager'
var SCRIPT_VERSION = 'v1.6 (Dev)'

var REGULAR_TASK_CALENDAR_NAME = SCRIPT_NAME

var PRODUCTION_VERSION = true

// Log Library
// -----------

var DEBUG_LOG_LEVEL                  = BBLog.Level.ALL
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES = BBLog.DisplayFunctionNames.YES 

// Assert library
// --------------

var DEV_EMAIL_ADDRESS = 'andrewr1969+rtm@gmail.com'

// Testing
// -------
//
// Should all be false in production

var TEST_BLOCK_TRIGGERS             = false // Can't do this when testing the add-on
var TEST_FORCE_INSTALL_ERROR        = false
var TEST_FORCE_OPEN_ERROR           = false
var TEST_FORCE_FORMSUBMIT_ERROR     = false
var TEST_FORCE_EDIT_ERROR           = false
var TEST_FORCE_CLOCKTRIGGER_ERROR   = false
var TEST_FORCE_STATUS_SEND_ERROR    = false
var TEST_BLOCK_EMAILS               = false
var TEST_FORCE_NO_DRAFTS            = false

if (PRODUCTION_VERSION) {

  if (TEST_BLOCK_TRIGGERS ||
      TEST_FORCE_INSTALL_ERROR ||
      TEST_FORCE_OPEN_ERROR ||
      TEST_FORCE_FORMSUBMIT_ERROR ||
      TEST_FORCE_EDIT_ERROR ||
      TEST_FORCE_CLOCKTRIGGER_ERROR ||
      TEST_FORCE_STATUS_SEND_ERROR ||
      TEST_BLOCK_EMAILS ||
      TEST_FORCE_NO_DRAFTS) {
   
     throw new Error('Testing flag set in PRODUCTION VERSION')
  }
}

var TEST_SPREADSHEET_ID = '1aHLHuph3-CkgMjGjx-NhxyKRmzF95OPLUUbUJs9Zxgg'

// Constants/Enums
// ===============

var DEFAULT_DRAFT_TEXT = 'Default template'

var REENABLE_CALENDAR_TEXT = 'The "daily calendar check" has been disabled so you will not see this error again.' + 
  'You will need to re-enable it (Add-ons > Rose Task Manager > Start daily calendar check) ' + 
  'if you wish to use this feature.'

// Task list spreadsheet
// ---------------------

var TASK_LIST_WORK_SHEET_NAME = "Task List"
var DATE_TIME_COLUMN_WIDTH = 130

// TODO - publish these template files templates in case I accidentally delete them
var FORM_TEMPLATE_ID = '10SD4_a9BzspC2J_mQORMAk1TdvvQil4SrAki7qnzBQc'
var TASK_REQUEST_FORM_NAME = 'Task Request Form'

var FIELDS_TEMPLATE_ID = '1vgr_qPjt3wITaPshqpk9SyxSF-ctraKP0oSXYzUlAL0'
var FIELDS_TEMPLATE_NAME = 'Fields'

var LOG_TEMPLATE_SHEET_ID = FIELDS_TEMPLATE_ID
var LOG_SHEET_TEMPLATE_NAME = 'Log'

// EMAIL_TEMPLATE_ID = FIELDS_TEMPLATE_ID
// EMAIL_TEMPLATE_NAME = 'Email Templates'

var DATE_TIME_FORMAT = 'd MMM yyyy HH:mm:ss'

var ALERT_HEIGHT = 100
var ALERT_WIDTH = 300

// Errors
// ------

var ERROR_SELECT_VALID_EMAIL = 'Select a row on the "' + TASK_LIST_WORK_SHEET_NAME + '" ' + 
    'sheet with a valid email address  and try again.'
  
var ERROR_FAILED_TO_SEND_STATUS = SCRIPT_NAME + ' ' + 
    'Failed to send Status email.'
  
// Status Email Update Template
// ----------------------------
// 
// This is the email sent to update the person who submitted a task as to its status.
//

// TODO - Make these configurable by moving them to the spreadsheet.

var STATUS_SUBJECT_TEMPLATE = 'Task #{{ID}} - Status Update - {{Status}}'

var STATUS_BODY_TEMPLATE = 'The present status of task #{{ID}} - "{{Subject}}" is {{Status}}.'

// Form Email Template
// -------------------
//
// This is the email notification sent when a form is submitted.
//

var FORM_SUBJECT_TEMPLATE = 'Task #{{ID}} "{{Subject}}" Received'

var FORM_BODY_TEMPLATE = 'AUTO-RESPONSE. \n\nNew task #{{ID}} - "{{Subject}}" ' + 
  'entered into the ' + TASK_LIST_WORK_SHEET_NAME + ' sheet.'

// Task List spreadsheet
// ---------------------
//
// These are the required fields in the task list spreadsheet, others can be added 
// to the task list spreadsheet itself without affecting its operation.
//

// TODO - Look at a consistent way of keeping the TimeStamp cell renamed

// !!! This is also the initial order of the headers, so dont change it !!!
// To avoid language dependence the script assumes the headers, or whatever is 
// is in another languages, are laid out in this order

var TASK_LIST_COLUMNS = {
  ID            : "ID",
  TIMESTAMP     : "Timestamp",
  STARTED       : "Started",
  CLOSED        : "Closed",
  TITLE         : "Subject",
  LOCATION      : "Location",
  PRIORITY      : "Priority",
  STATUS        : "Status",
  CATEGORY      : "Category",
  ASSIGNED_TO   : "Assigned to",
  REQUESTED_BY  : "Requested by",
  CONTACT_EMAIL : "Contact email",
  NOTES         : "Description",
}

var STATUS_NEW         = "1 - New"         // New task/issue raised
var STATUS_OPEN        = "2 - Open"        // Acknowledged by Admin
var STATUS_IN_PROGRESS = "3 - In Progress" // Work in progress
var STATUS_IGNORED     = "4 - Ignored"     // Closed, but not completed
var STATUS_DONE        = "5 - Done"        // Closed and completed

var PRIORITY_LOW    = "3 - Low"
var PRIORITY_NORMAL = "2 - Normal"
var PRIORITY_HIGH   = "1 - High"

// Properties
// ----------

var PROPERTY_LAST_AUTH_EMAIL_DATE   = SCRIPT_NAME + ' last auth email date'
var PROPERTY_CALENDAR_TRIGGER_ID    = SCRIPT_NAME + ' calendar id' // actually the trigger, but kept with this name for backward compatibility
var PROPERTY_CALENDAR_ID_NT         = SCRIPT_NAME + ' calendar id (not trigger)'
var PROPERTY_STATUS_EMAIL           = SCRIPT_NAME + ' status email template'
var PROPERTY_NEW_TASK_EMAIL         = SCRIPT_NAME + ' new task email template'
var PROPERTY_EMAIL_FROM             = SCRIPT_NAME + ' email from'
var PROPERTY_CALENDAR_TRIGGER_COUNT = SCRIPT_NAME + ' calendar trigger count'

// Function Template
// -----------------

/**
 * 
 *
 * @param {Object} 
 * 
 * @return {Object} 
 */
 
function functionTemplate() {
  
  Log_.functionEntryPoint()
  
  

} // functionTemplate()
