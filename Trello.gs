// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO
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

// TODO - Check if the card already exists, if so don't create a new one but 
// update existing
// TODO - Create list if one doesn't already exist

var Trello = {

  /**
   *
   */

  displayBoards: function () {
    
    Log_.functionEntryPoint()
    var trelloApp = new TrelloApp.App()
    var boards = trelloApp.getMyBoards()
    
    boards.forEach(function(board) {
      Log_.info(board.getId()+ ' ' + board.getName())
    })
    
  }, // Trello.displayBoards()
  
  /**
   *
   */

  uploadCards: function() {

    var trelloApp = new TrelloApp.App()
    
    // Go through each row in the spreadsheet
  
    var sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(TASK_LIST_WORK_SHEET_NAME)
      
    // TODO - make a function  
      
    var rows = sheet.getDataRange().getValues()
    var idColumnIndex = Utils_.getColumnPosition(sheet, SS_COL_ID, rows[0]) - 1
    var titleColumnIndex = Utils_.getColumnPosition(sheet, SS_COL_TITLE, rows[0]) - 1
    var descriptionColumnIndex = Utils_.getColumnPosition(sheet, SS_COL_NOTES, rows[0]) - 1
    var statusColumnIndex = Utils_.getColumnPosition(sheet, SS_COL_STATUS, rows[0]) - 1
        
    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      
      var currentRow = rows[rowIndex]
      
      if (currentRow[titleColumnIndex] === '') {
      
        // Do nothing if no card name
        
      } else {
      
        var status = currentRow[statusColumnIndex]
        var listId = getRtmListID()
        
        if (listId === '') {
          throw new Error('Failed to find list for ' + status)
        }
        
        Log_.fine('Found Trello list id: ' + listId)

        // TODO - Check if it already exists

        trelloApp.createCard({
          name: currentRow[idColumnIndex] + ' - ' + currentRow[titleColumnIndex], 
          desc: currentRow[descriptionColumnIndex],
          idList: listId,
        })       
      }      
      
    } // for each row
    
    return 'backlog items were uploaded successfully'
    
    // Private Functions
    // -----------------

    /**
     *
     */
    
    function getRtmListID() {
   
      var listName = ''
      
      // TODO - Put this into a table
      
      switch (status) {
      
      case STATUS_NEW:
        listName = 'New'
        break
        
      case STATUS_OPEN:
        listName = 'Open'
        break
        
      case STATUS_IN_PROGRESS:
        listName = 'In Progress'
        break
        
      case STATUS_IGNORED:
        listName = 'Ignored'
        break
        
      case STATUS_DONE:
        listName = 'Done'
        break
        
      default:
        throw new Error('Unexpected status type')        
      }
   
      Log_.fine('listName: ' + listName)
   
      var rtmListId = ''
      var foundList
    
      // Look for 'RTM List 1' in the 'Rose Task Manager' and 
      // add a new card to it
    
      trelloApp.getMyBoards().some(function(board) {
        
        if (board.getName() === TRELLO_BOARD_NAME) {
          
          Log_.fine('Found RTM board: ' + TRELLO_BOARD_NAME)
          
          var foundList = trelloApp.getBoardLists(board.getId()).some(function(list) {
    
            if (list.getName() === listName) {
      
              rtmListId = list.getId()
              Log_.fine('Found RTM list 1, id: ' + rtmListId)
              return true
            }
          })
          
          return foundList
        }
      })
      
      return rtmListId
      
    } // Trello.uploadCards.getRtmListID()    
  
  } // Trello.uploadCards()

} // Trello
