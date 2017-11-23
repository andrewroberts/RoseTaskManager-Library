//234567890123456789012345678901234567890123456789012345678901234567890123456789

/*
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
 * README.gs
 * =========
 */

// Version 1.6
// -----------

// * Removed analytics (that wasn't used anymore)
// * Add sidebar
// * Use GMail draft for emails (in Sidebar > Settings)
// * Add "About" to sidebar
// * Email notifications "from" configurable from Settings
// * Replace calendar trigger log with event count
// * Use calendar ID rather than name (so the name can be changed and RTM 
//   will still find it

// Version 1.5.1
// -------------

// * Ignore missing Timestamp column
// * Add trace for where the "not from Task List" columns come from

// Version 1.5
// -----------

// * Use Firebase for production logging - BBLog

// Version 1.4.4
// -------------

// * Stop calendar trigger after re-auth email 

// Version 1.4.3
// -------------

// * Updated auth email text 

// Version 1.4.2
// -------------

// * BUGFIX: Re-arranged event handler init so user email available, and errors
//   emailed.

// Version 1.4.1
// -------------

// * Made sure menu refresh not called when trigger deleted in context of 
//   a trigger

// Version 1.4
// -----------

// * Updated Assert library to fix error email send
// * Updated "re-auth" email text
// * Removed myself from "re-auth" email

// Version 1.3
// -----------

// * Stop trigger on "daily calendar check" 
// * Added "Stop/start" trigger menu option
// * Removed delete of columns on install as causing error

// Version 1.2.3
// -------------

// * Better "no RTM calendar" error message.

// Version 1.2.2
// -------------

// * Don't check calendar if not authed

// Version 1.2.1
// -------------

// * Downgrade non task list form submission

// Version 1.2
// -----------

// * Send error emails to user

// Version 1.1.1 (Add-on v5) 
// -------------------------

// * Updated so errors displayed as alert not thrown

// Version 1.1 (Add-on v4)
// -----------------------

// * Updated code to latest libraries, style, etc.
// * Added status description to emails
// * Added re-auth code
// * Reorganised installation
// * Added manual Calendar trigger to production menu

// Version 1.0.0
// -------------

// * Automated installation (create form, format sheet, create calendar, etc)
//   so that RoseTM could be released as an add-on.
// * Introduced various libraries (or development versions of)
// * Started on re-formatting.

// Version 0.4.1
// -------------

// * Added started column and auto-complete on going to in-progress
// * Add UA tracking (after release) via cUAMeasure library.

// Version 0.4.0
// -------------

// * Hadn't checked that unit tests were running on 0.3.0, so small changes for that.
// * Changed CMMS_NAME to TM_NAME
// * Checked with new sheets - minor changes
// * Removed filter from sort menu and added new ones

// Version 0.3.0
// -------------

// * Added "wait" text to form
// * Not in the code - but realised I need to set the triggers at 4am to avoid
//   slipping into previous day as events are stored in UTC.
// * Renamed to RoseTM.

// Version 0.2.0
// -------------

// * Generally made more generic and easier to deploy; removed all reference to original 
//   deployment.
// * Removed need for fixed task list columns, as long as the required ones 
//   are in place the user can add and remove their own wherever they like. 
// * Removed the notifyAssignee() function as it seems to impersonal to ever be used.

// Version 0.1.1
// -------------
//
// * Added auto-completion of date when tasks are closed (CLOSED or DONE).

// Version 0.1.0
// -------------
//
// * First beta release.
