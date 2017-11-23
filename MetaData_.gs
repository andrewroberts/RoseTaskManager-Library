// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 20171110
/* jshint asi: true */

(function() {"use strict"})()

// GasTemplate.gs
// ==============

/*
 * Copyright (C) 2015-2018 Andrew Roberts
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

var MetaData_ = (function myFunction(ns) {
  
  /**
   * Add some meta data to a GSheet column
   *
   * @param {Sheet} sheet
   * @param {String} key The header name of the column (row 1)
   * @param {Number} startIndex 0-based index
   * 
   * @return {Object} result
   */
     
  ns.add = function(sheet, key, startIndex) {

    Log_.functionEntryPoint()
    
    Log_.fine('key: ' + key)
    Log_.fine('startIndex: ' + startIndex)

    var callingfunction = 'MetaData_.addMetaData()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    Assert.assertString(key, callingfunction, 'key not a string')
    Assert.assertNumber(startIndex, callingfunction, 'startIndex not a Number')
  
    var requests = [{
         
      // CreateDeveloperMetadataRequest
      createDeveloperMetadata: {
      
        // DeveloperMetaData
        developerMetadata: {
        
          // DeveloperMetaDataLocation with column scope  
          metadataKey: key,
          metadataValue: JSON.stringify({
              writtenBy:Session.getActiveUser().getEmail(),
              createdAt:new Date().getTime()
            }),
          location: {  
            dimensionRange: {
              sheetId:sheet.getSheetId(),
              dimension:"COLUMNS",
              startIndex:startIndex,    
              endIndex:startIndex + 1   
            }
          },
          visibility:"DOCUMENT"      
        }
      }
    }]
    
    var spreadsheetId = sheet.getParent().getId()
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId)
    return result
  
  } // MetaData_.addMetaData()
    
  /**
   * Get some meta data from a GSheet
   *
   * @param {Sheet} sheet
   * @param {String} key 
   * 
   * @return {Object} result
   */

  ns.get = function(sheet, key) {
  
    Log_.functionEntryPoint()
    Log_.fine('key: ' + key)

    var callingfunction = 'MetaData_.addMetaData()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    Assert.assertString(key, callingfunction, 'key not a string')

    var spreadsheetId = sheet.getParent().getId()
    var meta
    
    try {
    
      meta = cSAM.SAM.searchByKey (spreadsheetId , key)
    
    } catch (error) {
    
      if (error.name === 'GoogleJsonResponseException') {
        
        Log_.fine('The Sheets API cant be accessed')
        return []
      }   
    }

    var tidied = cSAM.SAM.tidyMatched(meta)
    
    if (tidied.length === 0) {
      Log_.fine('No meta data')
    } else {
      Log_.fine('Got meta data')
    }

    return tidied
    
  } // MetaData_.get()

  /**
   * Get a column's index 
   *
   * @param {Sheet} sheet
   * @param {String} key The header name of the column (row 1)
   * 
   * @return {Number} startIndex 0-based 
   */

  ns.getColumnIndex = function(sheet, key) {

    Log_.functionEntryPoint()

    var tidied = MetaData_.get(sheet, key)
    var startIndex = -1

    if (tidied.length !== 0) {
    
      if (tidied.length > 1) {
        Log_.warning('Using the first element, although there may be more')
      }
      
      startIndex = tidied[0].location.dimensionRange.startIndex 
      Log_.fine('startIndex: ' + startIndex)
      
    } else {
    
      Log_.fine('No meta data for start index')
    } 
      
    return startIndex

  } // MetaData_.getColumnIndex()

  /**
   * Remove some meta data from a GSheet
   *
   * @param {Sheet} sheet
   * @param {String} key 
   * 
   * @return {Object} result
   */
  
  ns.remove = function(sheet, key) {

    Log_.functionEntryPoint()

    var callingfunction = 'MetaData_.addMetaData()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    Assert.assertString(key, callingfunction, 'key not a string')

    // Get all the things and delete them in one go
    var requests = [{
      deleteDeveloperMetadata: {
        dataFilter:{
          developerMetadataLookup: {
            metadataKey: key
          }
        }
      }
    }]
  
    var spreadsheetId = sheet.getParent().getId()
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId);
    return result
    
  } // MetaData_.remove()  
  
  return ns
    
}) (MetaData_ || {})