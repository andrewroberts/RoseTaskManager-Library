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
 * Copyright (C) 2015 Andrew Roberts
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
 * config.gs
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
var SCRIPT_VERSION = 'v1.5'

var REGULAR_TASK_CALENDAR_NAME = SCRIPT_NAME

var PRODUCTION_VERSION = true

// Log Library
// -----------

var DEBUG_LOG_LEVEL                  = BBLog.Level.ALL
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES = BBLog.DisplayFunctionNames.YES 

// Assert library
// --------------

var DEV_EMAIL_ADDRESS = 'andrewr1969+rtm@gmail.com'

// Analytics
// ---------

// TODO - Can't run this when I open the script. THe UAMeasure library 
// has been removed for now. Add-ons already have some analytics builtin.
var DISABLE_ANALYTICS = true
 
// Allow this user to opt out of the anonymous analytics recorded.
var OPTOUT_OF_ANALYTICS = false

// Testing
// -------

// Can't do this when testing the add-on
var CREATE_TRIGGERS = true

// Don't need to force 'email status'as easy to trigger.
var FORCE_INSTALL_ERROR        = false
var FORCE_OPEN_ERROR           = false
var FORCE_FORMSUBMIT_ERROR     = false
var FORCE_EDIT_ERROR           = false
var FORCE_CLOCKTRIGGER_ERROR   = false

// Constants/Enums
// ===============

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

// Status Email Update Template
// ----------------------------
// 
// This is the email sent to update the person who submitted a task as to its status.
//

// TODO - Make these configurable by moving them to the spreadsheet.

var STATUS_SUBJECT_TEMPLATE = "Task #${\"id\"} - Status Update - \"${\"status\"}\""

var STATUS_BODY_TEMPLATE = 
  "We've updated the status of task #${\"row\"} - ${\"title\"}." + 
  "\n\nNew Status: ${\"status\"}." + 
  "\n\n\"1 - New\"  Listed but not yet seen by list admin." + 
  "\n\"2 - Open\"  Seen by list admin but not being worked on yet." + 
  "\n\"3 - In Progress\"  Being worked on." + 
  "\n\"4 - Ignored\"  Has not been completed but has been closed for some other reason (test task request, task superseeded, etc)." + 
  "\n\"5 - Done\"  Has been completed." 

// Form Email Template
// -------------------
//
// This is the email notification sent when a form is submitted.
//

var FORM_SUBJECT_TEMPLATE = "Task #${\"id\"} \"${\"title\"}\" Received"

var FORM_BODY_TEMPLATE = "AUTO-RESPONSE. \n\nNew task #${\"id\"} - \"${\"title\"}\" entered into the" + TASK_LIST_WORK_SHEET_NAME + " sheet."

// Task List spreadsheet
// ---------------------
//
// These are the required fields in the task list spreadsheet, others can be added 
// to the task list spreadsheet itself without affecting its operation.
//

// TODO - Look at a consistent way of keeping the TimeStamp cell renamed

var SS_COL_TIMESTAMP     = "Timestamp"
var SS_COL_STARTED       = "Started"
var SS_COL_CLOSED        = "Closed"
var SS_COL_ID            = "ID"
var SS_COL_TITLE         = "Subject"
var SS_COL_LOCATION      = "Location"
var SS_COL_PRIORITY      = "Priority"
var SS_COL_STATUS        = "Status"
var SS_COL_CATEGORY      = "Category"
var SS_COL_ASSIGNED_TO   = "Assigned to"
var SS_COL_REQUESTED_BY  = "Requested by"
var SS_COL_CONTACT_EMAIL = "Contact email"
var SS_COL_EVENT_ID      = "Event ID"
var SS_COL_NOTES         = "Description"

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

var PROPERTY_LAST_AUTH_EMAIL_DATE = SCRIPT_NAME + ' last auth email date'
var PROPERTY_CALENDAR_ID          = SCRIPT_NAME + ' calendar id'

// Google Universal Analytics
// --------------------------

if (!DISABLE_ANALYTICS) {

  var UA_CODE = 'UA-54781593-1'
  
  var UA = new cUAMeasure
    .UAMeasure (
      UA_CODE, 
      SCRIPT_NAME, 
      Session.getEffectiveUser().getEmail(), 
      OPTOUT_OF_ANALYTICS,
      SCRIPT_VERSION
    )
}                                 

// Trello
// ------

var TRELLO_BOARD_NAME = SCRIPT_NAME
var TRELLO_LIST_NAME = 'RTM Test List 1'

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
