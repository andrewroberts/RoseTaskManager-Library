// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 7 Oct 2014 19:34 GMT+1

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
 * ANY WARRANTY without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License along with 
 * this program. If not, see http://www.gnu.org/licenses/.
 */

// TODO - Review Log severity as these are utilities and the decision should be
// made higher up

// Utilities.gs
// ============

var Utils_ = {

/**
 * Get the email address of the user that authorised the script.
 */
 
getListAdminEmail: function() {
  
  Log_.functionEntryPoint()
  
  // Get the email of the user that authorised the script, probably the owner.
  return Session.getEffectiveUser().getEmail()
  
}, // Utils_.getListAdminEmail()

/**
 * Get the index (0-based) of this column
 *
 * First check if we can get it from column meta data, if not look in 
 * the header row for the column name
 *
 * @param {Object} 
 *   {Sheet} sheet
 *   {string} columnName
 *   {Array} headers [OPTIONAL, DEFAULT - got from row 1]
 *   {boolean} required [OPTIONAL, DEFAULT = true]
 *   {Boolean} useMeta [OPTIONAL, DEFAULT = true]
 *
 * @return {number} column index or -1
 */
  
getColumnIndex: function(config) {
    
  Log_.functionEntryPoint()
  
  var sheet = config.sheet
  var columnName = config.columnName
  var headers = config.headers
  
  var required
  
  if (typeof config.required === 'undefined' || config.required) {
    required = true
  } else {
    required = false
  }
  
  var useMeta
  
  if (typeof config.useMeta === 'undefined' || config.useMeta) {
    useMeta = true
  } else {
    useMeta = false
  }

  Log_.fine('columnName: %s', columnName)
  Log_.fine('headers: %s', headers)
  Log_.fine('required: %s', required)
  Log_.fine('useMeta: %s', useMeta)
  
  var columnIndex = -1
  
  if (useMeta) {
  
    // First check if we can get it from column meta data
    
    columnIndex = MetaData_.getColumnIndex(sheet, columnName)
  }
  
  if (columnIndex === -1) {
    
    // Next try from the header
    
    if (typeof headers === 'undefined') {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    }
    
    columnIndex = headers.indexOf(columnName)
    
    if (columnIndex === -1) {
    
      if (columnName === TASK_LIST_COLUMNS.TIMESTAMP) {
      
        // There may have been an old version where it was renamed to "Listed"
        columnIndex = headers.indexOf("Listed")
        
      } else if (columnName === TASK_LIST_COLUMNS.ID) {
      
        // Sometimes the 'ID' column header is accidentally deleted
        if (sheet.getRange('A1').getValue() === '') {        
          columnIndex = 0
        }
        
      } else if (columnName === TASK_LIST_COLUMNS.SUBJECT) {
      
        // These are used to be a regular error, so hard-coded (I'm that kinda guy!)
        
        columnIndex = headers.indexOf('Type of Service Request')
        
        if (columnIndex === -1) {
          columnIndex = headers.indexOf('Комментарий')
        }
      }
    }
    
    if (columnIndex !== -1) { 
    
      if (useMeta) {
    
        // Store meta data for this column in case it does get moved or renamed
        MetaData_.add(sheet, columnName, columnIndex) 
      }
      
      Log_.fine('columnIndex from headers: ' + columnIndex) 
    }
  }
  
  if (columnIndex === -1) {
    
    if (required) {
      
      Utils_.throwNoColumnError(columnName, headers)
      
    } else {
      
      Log_.warning('No "' + columnName + '" column. Found: ' + JSON.stringify(headers))
    }
  }
  
  return columnIndex
  
}, // Utils_.getColumnIndex()

/**
 * Set the value of a spreadsheet cell.
 *
 * @param {object} sheet
 * @param {number} rowNumber
 * @param {string} columnName
 * @param {object} value
 * @param {Boolean} useMeta 
 * @param {Boolean} required  
 */

setCellValue: function (sheet, rowNumber, columnName, value, useMeta, required) {
  
  Log_.functionEntryPoint()
  
  var columnIndex = Utils_.getColumnIndex({
    sheet: sheet, 
    columnName: columnName,
    useMeta: useMeta,
    required: required})
  
  if (columnIndex !== -1) {

    sheet.getRange(rowNumber, columnIndex + 1).setValue(value)
    
    Log_.info('Stored ' + value + ' in row ' + rowNumber + ', column: ' + columnName)
    
    Log_.fine(
      'Wrote value ' + value + ' ' +
        'rowNumber: ' + rowNumber + ', ' + 
        'columnName: ' + columnName,
        'columnIndex: ' + columnIndex + ' ' +
        'useMeta:' + useMeta + ' ' +
        'required:' + required)       
  }
    
}, // Utils_.setCellValue()

/**
 * Get the value of a spreadhsheet cell
 *
 * @param {object} sheet
 * @param {number} rowNumber
 * @param {string} columnName
 * @param {Boolean} useMeta
 * @param {Boolean} required 
 *
 * @return {object} value or null
 */

getCellValue: function(sheet, rowNumber, columnName, useMeta, required) {
  
  Log_.functionEntryPoint()
  
  Log_.fine('rowNumber: ' + rowNumber)
  Log_.fine('columnName: ' + columnName)
  Log_.fine('useMeta: ' + useMeta)
  Log_.fine('required: ' + required)
  
  var columnIndex = Utils_.getColumnIndex({
    sheet: sheet, 
    columnName: columnName,
    useMeta: useMeta,
    required: required})
    
  var value = null
  
  if (columnIndex !== -1) {
    
    value = sheet.getRange(rowNumber, columnIndex + 1).getValue()
    
    Log_.fine(
      'Got ' + value + ', ' +
        'rowNumber: ' + rowNumber + ' ' +
        'columnName: ' + columnName + ' ' + 
        'columnIndex: ' + columnIndex + ' ' +
        'useMeta:' + useMeta + ' ' +
        'required:' + required)
  }
    
  return value    
  
}, // Utils_.getCellValue()

/**
 * checkIfAuthorizationRequired
 * From 
 * https://developers.google.com/apps-script/reference/script/script-app#getAuthorizationInfo(AuthMode)
 * https://developers.google.com/apps-script/guides/triggers/installable#authorization
 *
 * @return {Boolean} whether auth required
 */
 
checkIfAuthorizationRequired: function() {
  
  Log_.functionEntryPoint()

  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL)
  var authRequired = false

  // Check if the actions of the trigger requires authorization that has not
  // been granted yet; if so, warn the user via email. This check is required
  // when using triggers with add-ons to maintain functional triggers.
  
  if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      
    authRequired = true  
      
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to re-authorize; the normal trigger action is not
    // conducted, since it requires authorization first. Send at most one
    // "Authorization Required" email per day to avoid spamming users

    var properties = PropertiesService.getUserProperties()
    var lastAuthEmailDate = properties.getProperty(PROPERTY_LAST_AUTH_EMAIL_DATE)
    var today = new Date().toDateString()
    
    if (lastAuthEmailDate !== today) {
    
      if (MailApp.getRemainingDailyQuota() > 0) {
      
        var html = HtmlService.createTemplateFromFile('Authorization')
        html.url = authInfo.getAuthorizationUrl()
        html.addonTitle = SCRIPT_NAME
        
        var subject = 'Authorization Required'
        var message = html.evaluate()
        
        var options = {
          name: SCRIPT_NAME,
          htmlBody: message.getContent()
        }
        
        var recipient = Session.getEffectiveUser().getEmail()
        
        MailApp.sendEmail(recipient, subject, message.getContent(), options)
        
        Log_.warning('Sent re-auth email to ' + recipient) 

      } else {
      
        Log_.warning('Unable to send "auth needed" email as run out of quota') 
      }

      // Try again in a days time
      properties.setProperty(PROPERTY_LAST_AUTH_EMAIL_DATE, today)
    }
  }
  
  return authRequired
  
}, // Utils_.checkIfAuthorizationRequired()

/**
 * Get the task request form ID
 *
 * @return {Form} form URL or null
 */
 
getFormId: function() {
  
  Log_.functionEntryPoint()
  var formUrl = SpreadsheetApp.getActive().getFormUrl()
  Log_.fine('Form url: ' + formUrl)
  
  if (formUrl === null) {
    return null
  }
  
  var formId = FormApp.openByUrl(formUrl).getId()
  Log_.fine('Form ID: ' + formId)
  return formId

}, // Utils_.getForm()

/**
 * Display spreadsheet "toast" message
 *
 * @param {String} message
 */
 
toast: function(message) {
  
  Log_.functionEntryPoint()
  
  var spreadsheet = SpreadsheetApp.getActive()
//  var DISPLAY_FOREVER = -1
  
  if (spreadsheet !== null) {
    spreadsheet.toast(message, SCRIPT_NAME)
  }

}, // Utils_.toast()

/**
 * Throw "no column" error
 *
 * @param {String} columnName
 * @param {Array} row
 * 
 * @return {Object} 
 */
 
throwNoColumnError: function(columnName, row) {
  
  Log_.functionEntryPoint()

  throw new Error('Can not find the "' + columnName + '" column in the "' + 
    TASK_LIST_WORK_SHEET_NAME + '" tab. This is required by the add-on, ' + 
    'check it\'s not been renamed? Did find: ' + JSON.stringify(row) + '. ' +
    'Consider using a custom email template, with placeholders that match the new ' + 
    'headers: Add-ons > Rose Task Manager > Settings')
  
}, // Utils_.throwNoColumnError()

/**
 * Send email
 *
 * @param {String} recipient
 * @param {String} subject
 * @param {String} plainBody
 * @param {Object} options
 * @param {String} logMessage
 *
 * @return {Boolean} email sent or not
 */
 
sendEmail: function(recipient, subject, plainBody, options, logMessage) {
  
  Log_.functionEntryPoint()
  
  Log_.fine('recipient: ' + recipient)
  Log_.fine('subject: ' + subject)
  Log_.fine('plainBody: ' + plainBody)
  
  if (options.hasOwnProperty('htmlBody')) {
    Log_.fine('htmlBody: ' + options.htmlBody)
  }
    
  if (options.hasOwnProperty('cc')) {
    Log_.fine('cc: ' + options.cc)
  }
    
  if (TEST_BLOCK_EMAILS) {
  
    Log_.warning('Email sending disabled')
    return false
    
  } else {
  
    var quota = MailApp.getRemainingDailyQuota()
    
    if (quota === 0) {
    
      throw new Error('You do not have any quota left for sending email')
    
    } else {
    
      Log_.fine('Email quota: ' + quota)
      MailApp.sendEmail(recipient, subject, plainBody, options)
      Log_.info(logMessage)
      return true
    }
  }

}, // Utils_.sendEmail()

/**
 * Set a setting to a new value
 *
 * @param {String} propertyName
 * @param {Object} newValue
 */
 
setSetting: function(propertyName, newValue) {
  
  Log_.functionEntryPoint()
  var properties = PropertiesService.getDocumentProperties()
  var oldValue = properties.getProperty(propertyName)
  properties.setProperty(propertyName, newValue)
  Log_.fine('Replaced ' + propertyName + ' value: "' + oldValue + '" with "' + newValue + '"')
 
}, // Utils_.setSetting()

/**
 * Get the email templates
 *
 * @param {String} propertyName
 * @param {String} defaultSubjectTemplate
 * @param {String} defaultBodyTemplate
 * 
 * @return {Object} templates
 */
 
getEmailTemplates: function(propertyName, defaultSubjectTemplate, defaultBodyTemplate) {
  
  Log_.functionEntryPoint()

  var draft = null
  var subjectTemplate
  var htmlBodyTemplate
  var plainBodyTemplate

  var templateSubject = PropertiesService
    .getDocumentProperties()
    .getProperty(propertyName)
   
  if (templateSubject === null || templateSubject === DEFAULT_DRAFT_TEXT) {
  
    Log_.info('Email template not available (' + propertyName + '), using default')
    subjectTemplate = defaultSubjectTemplate
    htmlBodyTemplate = ''
    plainBodyTemplate = defaultBodyTemplate
    
  } else {
  
    GmailApp.getDraftMessages().some(function(nextDraft) {
      if (nextDraft.getSubject() === templateSubject) {
        draft = nextDraft
        return true
      }
    })
  
    if (draft === null) {
    
      Log_.warning('Could not find the GMail draft (' + propertyName + ') so using default')
      subjectTemplate = defaultSubjectTemplate
      htmlBodyTemplate = ''
      plainBodyTemplate = defaultBodyTemplate  
      
    } else {
  
      subjectTemplate = draft.getSubject()
      htmlBodyTemplate = draft.getBody()
      plainBodyTemplate = draft.getPlainBody()
    }
  }
  
  return {
    subject: subjectTemplate,
    htmlBody: htmlBodyTemplate,
    plainBody: plainBodyTemplate
  }

}, // Utils_.getEmailTemplates()

/**
 * For each placeholder - "{{[placeholder]}}", strip out any HTML and replace 
 * it with the appropriate key value:
 * 
 * E.g. If the object were:
 *
 * {PlaceHolder1: newValue}
 *
 * In the template "{{PlaceHolder1}}" would be replaced with "newValue" even
 * if there were some HTML (<...>) inside the brackets.
 *
 * https://github.com/shantanu543/bulk-mail/blob/master/bulkmail.gs
 * 
 * @param {String} template HTML or plain text
 * @param {Object} rowObject replacement values: {[placeholder]: [new value]}
 *
 * @return {String} completed template
 */
  
fillInTemplate: function (template, rowObject) {
  
  Log_.functionEntryPoint()
  Log_.fine('template: ' + template)
  Log_.fine('rowObject: ' + JSON.stringify(rowObject))
  
  var completed = template.replace(/{{.*?}}/g, function(nextMatch) {
  
    Log_.fine('nextMatch: ' + nextMatch)
    
    // Remove any HTML inside the placeholder
    var placeholderValue = nextMatch.replace(/<.*?>/g, '')
    Log_.fine('placeholderValue (HTML stripped): ' + placeholderValue)
    
    // Strip the placeholder identifier
    placeholderValue = placeholderValue.substring(2, placeholderValue.length - 2)
    Log_.fine('placeholderValue (ids removed): ' + placeholderValue)
    
    var nextValue = rowObject[placeholderValue] || ''
    Log_.fine('nextValue: ' + nextValue)
    
    return nextValue
  })

  return completed

}, // Utils_.fillInTemplate()

} // Utils_
