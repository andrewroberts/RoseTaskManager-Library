// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 7 Oct 2014 19:40 GMT+1

/*
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

var Calendar_ = {

/**
 * Create a calendar and set a trigger to check it for events
 * to convert to tasks.
 */
  
create: function() {
  
  Log_.functionEntryPoint()
  
  // TODO - This takes 4s
  
  // Create the calendar for storing regular events/tasks.
  
  var calendars = CalendarApp.getCalendarsByName(REGULAR_TASK_CALENDAR_NAME)
  var numberOfCalendars = calendars.length
  
  if (numberOfCalendars === 0) {
    
    var calendar = CalendarApp.createCalendar(REGULAR_TASK_CALENDAR_NAME)
    
    if (calendar) {
      
      Log_.info('New calendar "' + REGULAR_TASK_CALENDAR_NAME + '" created')
      
    } else {
      
      throw new Error(
        'Failed to create calendar "' + 
         REGULAR_TASK_CALENDAR_NAME + '"')
    }
    
  } else if (numberOfCalendars === 1) {
    
    Log_.warning(
      'There is already one or more calendars called "' + 
      REGULAR_TASK_CALENDAR_NAME + '"')
      
  } else {
  
    throw new Error('There are already more than one calendars called "' + 
      REGULAR_TASK_CALENDAR_NAME + '". Please delete one and try reinstalling the add-on. ' + 
      REENABLE_CALENDAR_TEXT)
  }
    
  this.createTrigger()
    
}, // Calendar_.create()

/**
 * Convert calendar events to tasks.
 * 
 * Read in all of todays events and add a row to the task list 
 * spreadsheet for each.
 */
 
convertEventsToTasks: function() {

  Log_.functionEntryPoint()
  
  try {

    if (Utils_.checkIfAuthorizationRequired()) {
    
      // No point continuing, an email to re-auth has already been sent to 
      // the user
      throw new Error('Script needs to be re-authorized. ' + REENABLE_CALENDAR_TEXT)
    }
    
    var calendars = CalendarApp.getCalendarsByName(REGULAR_TASK_CALENDAR_NAME)
    
    if (calendars.length === 0) {
    
      throw new Error(
        'There is no calendar called "' + REGULAR_TASK_CALENDAR_NAME + '". ' + 
          'Re-create it or uninstall the Rose Task Manager Add-on for GSheets. ' +
          REENABLE_CALENDAR_TEXT)
    }
    
    if (calendars.length > 1) {
      
      throw new Error('Found ' + calendars.length + ' calendars ' + 
        'called ' + REGULAR_TASK_CALENDAR_NAME + 'when there should only be one. ' + 
        'Please delete or rename one. ' + REENABLE_CALENDAR_TEXT)
    }
    
    // Check today's events
    // --------------------
    //
    // The events are stored in UTC so we need to allow for that by 
    // defining todays date in terms of UTC to avoid picking up events 
    // either side of today.
    
    var today = new Date()
    
    var startOfDay = today
    
    startOfDay.setUTCHours(0)
    startOfDay.setMinutes(0)
    startOfDay.setSeconds(0)
    startOfDay.setMilliseconds(0)  
    
    var endOfDay = new Date(startOfDay.getTime() + 24 * 60 * 60 * 1000)
    
    var events = calendars[0].getEvents(startOfDay, endOfDay)
    
    if (events.length === 0) {
    
    // TODO (#107) - Reinstate Dialog when run from UI
/*    
      Dialog.show(
        REGULAR_TASK_CALENDAR_NAME, 
        'No events found in ' + REGULAR_TASK_CALENDAR_NAME + ' calendar today.',
        110)
*/        
      Log_.info('No events today')
      
      return
    } 
    
    Log_.info( 
      'Number of events today: ' + events.length + 
       ' start: ' + startOfDay + 
       ' end: ' + endOfDay)
/*
    Dialog.show(
      REGULAR_TASK_CALENDAR_NAME, 
      events.length + ' events found in ' + REGULAR_TASK_CALENDAR_NAME + ' calendar today.',
      110)
*/
    // Simulate a Form response
    // ------------------------
    
    var formId = Utils_.getFormId()
    
    if (formId === null) {
        
      throw new Error('There is no form attached to "' + TASK_LIST_WORK_SHEET_NAME + '". ' + 
        'You will need to re-install the add-on to recreate it.')
    }
    
    var form = FormApp.openById(formId)
    var items = form.getItems()
    
    for (var eventIndex = 0; eventIndex < events.length; eventIndex++) {
      
      Utils_.postAnalytics('CalendarEvent')
      
      // Get the event data.
      
      var title = events[eventIndex].getTitle()
      var description = events[eventIndex].getDescription()
      var location = events[eventIndex].getLocation()
      
      var response = form.createResponse()
      
      var textBoxItem = items[0].asTextItem().createResponse(title)
      response.withItemResponse(textBoxItem)
      
      textBoxItem = items[1].asTextItem().createResponse(location)
      response.withItemResponse(textBoxItem)
      
      var checkBoxItem = items[2].asCheckboxItem().createResponse([PRIORITY_NORMAL])
      response.withItemResponse(checkBoxItem)
      
      textBoxItem = items[3].asTextItem().createResponse(SCRIPT_NAME)
      response.withItemResponse(textBoxItem)
      
      textBoxItem = items[4].asTextItem().createResponse(Utils_.getListAdminEmail())
      response.withItemResponse(textBoxItem)
      
      textBoxItem = items[5].asParagraphTextItem().createResponse(description)
      response.withItemResponse(textBoxItem)
      
      response.submit()
      
      Log_.fine('Simulated a form response')
      
    } // for each event
    
  } catch(error) {
  
    // For any errors thrown in this daily check, disable the trigger until they've fixed it
    this.deleteTrigger(false)

    throw error
  }
  
}, // Calendar_.convertEventsToTasks()

/**
 * Create a trigger to check the calendar every night
 */
   
createTrigger: function() {
    
  Log_.functionEntryPoint()
    
  var properties = PropertiesService.getDocumentProperties()

  // For some reason there may be multiple triggers already - due to an 
  // earlier error maybe - so delete any existing ones

  ScriptApp.getProjectTriggers().forEach(function(trigger) {    
  
    if (trigger.getHandlerFunction() === 'onCalendarTrigger') {

      Log_.warning('deleting calendar trigger: ' + trigger.getUniqueId())
      ScriptApp.deleteTrigger(trigger)
      properties.deleteProperty(PROPERTY_CALENDAR_ID)
    }
  })
  
  // TODO - Change time according to calendar timezone, see getCalendarTime()
  
  if (CREATE_TRIGGERS) {
    
    var triggerId = ScriptApp.newTrigger('onCalendarTrigger')
      .timeBased()
        .everyDays(1)
        .atHour(4)
        .create()
      .getUniqueId()
    
    properties.setProperty(PROPERTY_CALENDAR_ID, triggerId)
    
    Log_.info('New calendar trigger created: ' + triggerId)
    
    // Refresh the sheet menu - it will be at least limited   
    onOpen_({authMode: ScriptApp.AuthMode.LIMITED})    
    
  } else {
    
    Log_.warning('Calendar trigger creation disabled')
  }
        
}, // Calendar_.createTrigger()

/**
 * Delete the trigger that checks the calendar every night
 *
 * @param {Boolean} calledWithUI 
 */
   
deleteTrigger: function(calledWithUI) {
    
  Log_.functionEntryPoint()
    
  var properties = PropertiesService.getDocumentProperties()

  ScriptApp.getProjectTriggers().forEach(function(trigger) {  
  
    if (trigger.getHandlerFunction() === 'onCalendarTrigger') {
    
      ScriptApp.deleteTrigger(trigger)
      properties.deleteProperty(PROPERTY_CALENDAR_ID)
      Log_.warning('deleted old calendar trigger')
    }
  })
  
  if (calledWithUI) {
  
    // Refresh the sheet menu - it will be at least limited   
    onOpen_({authMode: ScriptApp.AuthMode.LIMITED})    
  }
    
}, // Calendar_.deleteTrigger()

} // Calendar_
