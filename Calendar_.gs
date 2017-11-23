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
  
  // Create the calendar for storing regular events/tasks
  
  var properties = PropertiesService.getDocumentProperties()
  var calendarId = properties.getProperty(PROPERTY_CALENDAR_ID_NT)
  var calendar
  
  if (calendarId !== null) {
  
    calendar = CalendarApp.getCalendarById(calendarId)
    Log_.info('Using existing calendar "' + calendar.getName() + '" (' + calendarId + ')')
    
  } else { // calendarId === null
  
    var calendars = CalendarApp.getCalendarsByName(REGULAR_TASK_CALENDAR_NAME)
    var numberOfCalendars = calendars.length
    
    if (numberOfCalendars === 0) {
      
      calendar = CalendarApp.createCalendar(REGULAR_TASK_CALENDAR_NAME)
      
      if (calendar !== null) {      
        Log_.info('New calendar "' + REGULAR_TASK_CALENDAR_NAME + '" created')
      } else {
        throw new Error('Failed to create calendar "' + REGULAR_TASK_CALENDAR_NAME + '"')
      }
      
      calendarId = calendar.getId()
      
    } else if (numberOfCalendars === 1) {
    
      calendarId = calendars[0].getId()
      
      Log_.info(
        'Using existing calendar "' + REGULAR_TASK_CALENDAR_NAME + '"(' + calendarId + ')')
        
    } else {
    
      throw new Error('There are already more than one calendars called "' + 
        REGULAR_TASK_CALENDAR_NAME + '". Please delete one and try reinstalling the add-on. ' + 
        REENABLE_CALENDAR_TEXT)
    }
    
    properties.setProperty(PROPERTY_CALENDAR_ID_NT, calendarId)
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
  
  var ERROR_TEXT_NO_CALENDAR = 'There is no "' + REGULAR_TASK_CALENDAR_NAME + '" calendar. ' + 
    'Re-create it or uninstall the Rose Task Manager Add-on for GSheets. ' +
    REENABLE_CALENDAR_TEXT
      
  try {

    if (Utils_.checkIfAuthorizationRequired()) {
    
      // No point continuing, an email to re-auth has already been sent to 
      // the user
      throw new Error('Script needs to be re-authorized. ' + REENABLE_CALENDAR_TEXT)
    }
    
    var properties = PropertiesService.getDocumentProperties()
    var calendarId = properties.getProperty(PROPERTY_CALENDAR_ID_NT)
    var calendar
    
    if (calendarId === null) {
    
      calendar = getCalendarByName()
           
    } else {
    
      calendar = CalendarApp.getCalendarById(calendarId)
            
      if (calendar === null) {
        calendar = getCalendarByName()
      }
    }

    properties.setProperty(PROPERTY_CALENDAR_ID_NT, calendar.getId())

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
 
    var events = calendar.getEvents(startOfDay, endOfDay)   
    var eventCount = getEventCount()
    
    if (events.length === 0) {
    
    // TODO (#107) - Reinstate Dialog when run from UI
/*    
      Dialog.show(
        REGULAR_TASK_CALENDAR_NAME, 
        'No events found in ' + REGULAR_TASK_CALENDAR_NAME + ' calendar today.',
        110)
*/        
      Log_.fine('No events today') 
      eventCount.noEvents++
      setEventCount()
      return
    } 

    eventCount.someEvents++
    eventCount.totalEvents += events.length
    setEventCount()
  
    Log_.fine( 
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
    
    events.forEach(function(event) {
      
      var title = event.getTitle()
      var description = event.getDescription()
      var location = event.getLocation()
      
      var response = form.createResponse()
      
      var textBoxItem = items[0].asTextItem().createResponse(title)
      response.withItemResponse(textBoxItem)
      
      textBoxItem = items[1].asTextItem().createResponse(location)
      response.withItemResponse(textBoxItem)
      
      var checkBoxItem = items[2].asCheckboxItem().createResponse([PRIORITY_NORMAL])
      response.withItemResponse(checkBoxItem)
      
      textBoxItem = items[3].asTextItem().createResponse(SCRIPT_NAME)
      response.withItemResponse(textBoxItem)
      
      textBoxItem = items[5].asParagraphTextItem().createResponse(description)
      response.withItemResponse(textBoxItem)
      
      response.submit()
      
      Log_.fine('Simulated a form response')
      
    }) // for each event
    
  } catch(error) {
  
    // For any errors thrown in this daily check, disable the trigger until they've fixed it
    this.deleteTrigger(false)

    throw error
  }
  
  // Private Functions
  // -----------------

  /**
   * Get calendars
   * 
   * @return {Calendar} RTM calendar
   */
   
  function getCalendarByName() {
    
    Log_.functionEntryPoint()
    
    var calendars = CalendarApp.getCalendarsByName(REGULAR_TASK_CALENDAR_NAME)
    
    if (calendars.length === 0) {
      throw new Error(ERROR_TEXT_NO_CALENDAR)
    }
    
    if (calendars.length > 1) {
      
      throw new Error('Found ' + calendars.length + ' calendars ' + 
                      'called ' + REGULAR_TASK_CALENDAR_NAME + 'when there should only be one. ' + 
                      'Please delete or rename one. ' + REENABLE_CALENDAR_TEXT)
    }
    
    return calendars[0]
    
  } // Calendar_.convertEventsToTasks.getCalendarByName()


  /**
   * Get the present event count. Logging every calendar trigger fills up the 
   * log, so they are just counted instead
   * 
   * @return {Object} eventCount
   */
   
  function getEventCount() {
    
    Log_.functionEntryPoint()
    
    var eventCount = PropertiesService
      .getScriptProperties()
      .getProperty(PROPERTY_CALENDAR_TRIGGER_COUNT)
    
    if (eventCount === null) {
    
      eventCount = { 
        noEvents: 0,
        someEvents: 0, 
        totalEvents: 0 
      }
      
      Log_.fine('Initialised eventCount')
      
    } else {
  
      Log_.fine('eventCount: ' + eventCount)
      eventCount = JSON.parse(eventCount)
    }
    
    return eventCount 
      
  } // Calendar_.convertEventsToTasks.getEventCount()
  
  /**
   * Set the present event count. Logging every calendar trigger fills up the 
   * log, so they are just counted instead
   * 
   * @param {Object} eventCount
   */
   
  function setEventCount() {
    
    Log_.functionEntryPoint()
    Log_.fine('Setting event count: ' + JSON.stringify(eventCount))
    
    PropertiesService
      .getScriptProperties()
      .setProperty(PROPERTY_CALENDAR_TRIGGER_COUNT, JSON.stringify(eventCount))
      
  } // Calendar_.convertEventsToTasks.setEventCount()

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
      properties.deleteProperty(PROPERTY_CALENDAR_TRIGGER_ID)
    }
  })
  
  // TODO - Change time according to calendar timezone, see getCalendarTime()
  
  if (!TEST_BLOCK_TRIGGERS) {
    
    var triggerId = ScriptApp.newTrigger('onCalendarTrigger')
      .timeBased()
        .everyDays(1)
        .atHour(4)
        .create()
      .getUniqueId()
    
    properties.setProperty(PROPERTY_CALENDAR_TRIGGER_ID, triggerId)
    
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
      properties.deleteProperty(PROPERTY_CALENDAR_TRIGGER_ID)
      Log_.warning('deleted old calendar trigger')
    }
  })
  
  if (calledWithUI) {
  
    // Refresh the sheet menu - it will be at least limited   
    onOpen_({authMode: ScriptApp.AuthMode.LIMITED})    
  }
    
}, // Calendar_.deleteTrigger()

/**
 * Dump event count
 */
 
dumpEventCount: function () {
  
  Log_.functionEntryPoint()
  
  var properties = PropertiesService.getScriptProperties()
  
  var nullEventCount = JSON.stringify({
    noEvents: 0,
    someEvents: 0,
    totalEvents: 0
  })
  
  var eventCount = properties.getProperty(PROPERTY_CALENDAR_TRIGGER_COUNT) || nullEventCount
  Log_.info('Event Count: ' + eventCount)
  properties.deleteProperty(PROPERTY_CALENDAR_TRIGGER_COUNT)

} // Calendar_.dumpEventCount()

} // Calendar_
