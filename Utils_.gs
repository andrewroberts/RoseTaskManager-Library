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
 * Send a notification to my Universal Analytics account.
 */
 
postAnalytics: function(msg) {
  
  Log_.functionEntryPoint()

  if (!DISABLE_ANALYTICS) {
    if (UA) {
      UA.postAppView(msg)
    }
  }

}, // Utils_.postAnalytics()

/**
 * Get the email address of the user that authorised the script.
 */
 
getListAdminEmail: function() {
  
  Log_.functionEntryPoint()
  
  // Get the email of the user that authorised the script, probably the owner.
  return Session.getEffectiveUser().getEmail()
  
}, // Utils_.getListAdminEmail()

/** 
 * Find the column index of the field with this name.
 *
 * @param {Object} sheet (not required if headerValues defined)
 * @param {String} column name
 * @param {Array} headerValues (not required if sheet defined)
 *
 * @return {number} column 1-based index or -1
 */

getColumnPosition: function (sheet, columnName, headerValues) {

  Log_.functionEntryPoint()
    
  if (typeof headerValues === 'undefined') {
    headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  }
  
  for (var colIndex = 0; colIndex < headerValues.length; colIndex++) {
    
    var nextName = headerValues[colIndex]
    
    if (nextName === columnName) {
      
      Log_.fine(
        "Column '" + columnName + "' " + 
          'has position ' + (colIndex + 1) + ' ' +
          "in sheet '" + sheet.getName() + "'")
      
      // Return the column position 1-based as this is what's used by
      // SpreadsheetApp
      return colIndex + 1      
    }
  }
  
  Log_.warning(
    "Column '" + columnName + "' " + 
      "not found in sheet '" + sheet.getName() + "'")
  
  return -1
  
}, // Utils_.getColumnPosition()

/**
 * Set the value of a spreadsheet cell.
 *
 * @param {object} sheet
 * @param {number} rowIndex
 * @param {string} column name
 * @param {any} value
 */

setCellValue: function (sheet, rowIndex, columnName, value) {
  
  Log_.functionEntryPoint()
  
  var colIndex = this.getColumnPosition(sheet, columnName)
  
  if (colIndex === -1) {
    
    Log_.warning("Could not find column: " + columnName)
    
  } else {

    sheet.getRange(rowIndex, colIndex).setValue(value)
    
    Log_.fine(
      'Wrote value ' + value + ' ' +
        'to sheet ' + sheet.getName() + ' ' +
        'rowIndex: ' + rowIndex + ', ' + 
        'columnName: ' + columnName)
  }
    
}, // Utils_.setCellValue()

/**
 * Get the value of a spreadhsheet cell
 *
 * @param {object} sheet
 * @param {number} rowIndex
 * @param {string} column name
 * @return {any} value or null
 */

getCellValue: function(sheet, rowIndex, columnName) {
  
  Log_.functionEntryPoint()
  
  var columnIndex = this.getColumnPosition(sheet, columnName)
  var value = null
  
  if (columnIndex === -1) {
    
    Log_.warning("setCellValue_", "Could not find column: " + columnName)
   
  } else {
  
    value = sheet.getRange(rowIndex, columnIndex).getValue()
    
    Log_.fine(
      'Got ' + value + ', ' +
        'from sheet: ' + sheet.getName() + ' ' +
        'rowIndex: ' + rowIndex + ' ' +
          'columnName: ' + columnName)
  }
    
  return value    
  
}, // Utils_.getCellValue()

// Replaces markers in a template string with values define in a JavaScript data object.
//
// @param {string} template string containing markers, for instance ${"Column name"}
// @param {object} data: JavaScript object with values to that will replace markers. For instance
//                       data.columnName will replace marker ${"Column name"}
//
// @return {string} a string without markers. If no data is found to replace a marker, it is
//                  simply removed.
//

fillInTemplateFromObject: function(template, data) {
  
  Log_.finest("fillInTemplateFromObject", "data: " + data + " template: " + template)
  
  var complete = template
  
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g)

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string
  for (var i = 0; i < templateVars.length; ++i) {
    
    // normalizeHeader ignores ${"} so we can call it directly here
    var variableData = data[normalizeHeader(templateVars[i])]
    
    if (variableData instanceof Date) {
      
      // This is a Date (it has the method anyway!) so convert it
      // to a nicer string
      variableData = variableData.toDateString()
    }
    
    complete = complete.replace(templateVars[i], variableData || "")
  }

  return complete
  
  // Private Functions
  // -----------------
  
  // Normalizes a string, by removing all alphanumeric characters and using mixed case
  // to separate words. The output will always start with a lower case letter.
  // This function is designed to produce JavaScript object property names.
  //
  // @param {string} header String to normalize
  //
  // @return 
  
  function normalizeHeader(header) {
  
    var key = ''
    var upperCase = false
  
    for (var i = 0; i < header.length; ++i) {
  
      var letter = header[i]
  
      if (letter === " " && key.length > 0) {
        upperCase = true
        continue
      }
  
      if (!isAlnum(letter)) {
        continue
      }
  
      if (key.length == 0 && isDigit(letter)) {
        continue // first character must be a letter
      }
  
      if (upperCase) {
  
        upperCase = false
        key += letter.toUpperCase()
  
      } else {
  
        key += letter.toLowerCase()
      }
    }
  
    return key
    
    // Private Functions
    // -----------------
    
    // Returns true if the cell where cellData was read from is empty
    //
    // @param {string} cellData
    //
    // @return {boolean} result
    
    function isCellEmpty(cellData) {
    
      return typeof(cellData) === 'string' && cellData === ''
      
    } // fillInTemplateFromObject.normalizeHeader.isCellEmpty()
    
    // Returns true if the character char is alphabetical, false otherwise
    //
    // @param {string} char
    //
    // @return {boolean} result
    
    function isAlnum(char) {
    
      return char >= 'A' && char <= 'Z' ||
        char >= 'a' && char <= 'z' ||
        isDigit(char)
    
    } // fillInTemplateFromObject.normalizeHeader.isAlnum()
    
    // Returns true if the character char is a digit, false otherwise
    //
    // @param {string} char
    //
    // @return {boolean} result
    
    function isDigit(char) {
    
      return char >= '0' && char <= '9'
    
    } // fillInTemplateFromObject.normalizeHeader.isDigit()
    
  } // fillInTemplateFromObject.normalizeHeader()
  
}, // Utils_.fillInTemplateFromObject()

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
  var DISPLAY_FOREVER = -1
  
  if (spreadsheet !== null) {
    spreadsheet.toast(message, SCRIPT_NAME, DISPLAY_FOREVER)
  }

} // Utils_.toast()

} // Utils_
