// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 27 Feb 2016 21:19
/* jshint asi: true */

/*
 * Copyright (C) 2016 Andrew Roberts
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
 * Domain Monitoring.gs
 * ====================
 */

// Config
// ------

var DAYS_TO_EXPIRATION = 30
var EXPIRATION_DATE = 'A4'

// Constants
// ---------

var DOMAINS_SHEET_NAME = 'Domains'

var DOMAIN_COLUMN_INDEX          = 0
var OLD_DNS_COLUMN_INDEX         = 1 
var NEW_DNS_COLUMN_INDEX         = 2
var LAST_CHECKED_COLUMN_INDEX    = 3
var CHANGED_COLUMN_INDEX         = 4
var CHECK_URL_COLUMN_INDEX       = 5
var EXPIRED_COLUMN_INDEX         = 6
var EXPIRATION_DATE_COLUMN_INDEX = 7
var EMAIL_SENT_COLUMN_INDEX      = 8

// Log Library
// -----------

var LOG_LEVEL = Log.Level.INFO
var LOG_SHEET_ID = ''
var LOG_DISPLAY_FUNCTION_NAMES = Log.DisplayFunctionNames.NO

// Functions
// ---------

/**
 * 'Sheet open' event handler
 */
 
function onOpen() {

  SpreadsheetApp
    .getUi()
      .createMenu('Domain monitor')
        .addItem('Check domains', 'checkDomains')
        .addToUi()

} // onOpen()

/**
 * Check the name servers and access to a list of URLs
 */
 
function checkDomains() {

  Log.init({
    level: LOG_LEVEL, 
    sheetId: LOG_SHEET_ID,
    displayFunctionNames: LOG_DISPLAY_FUNCTION_NAMES})

  Log.functionEntryPoint()

  var EMAIL = PropertiesService.getScriptProperties().getProperty('EMAIL')
  var SHEET_URL = PropertiesService.getScriptProperties().getProperty('SHEET_URL')

  if (EMAIL === null || SHEET_URL === null) {
    Log.warning('No email address and/or sheet URL')
  }

  var dataRange = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(DOMAINS_SHEET_NAME)
    .getDataRange()
    
  var data = dataRange.getValues()
  var header = data.shift()
  var nsChanged = false
  var urlFailed = false
  var expirationDateSoon = false
  
  data.forEach(function(row) {

    var url = row[DOMAIN_COLUMN_INDEX]

    // Check the site is accessible
    // ----------------------------
   
    try {
   
      var response = UrlFetchApp.fetch('http://www.' + url, {muteHttpExceptions:true})
    
    } catch (error) {
    
      Log.warning(url + ' not accessible: ' + error.message)
      return
    }
    
    var responseCode = response.getResponseCode()
    
    if (responseCode === 200) {
    
      row[CHECK_URL_COLUMN_INDEX] = 'OK'
      
    } else {
    
      row[CHECK_URL_COLUMN_INDEX] = responseCode
      urlFailed = true
      Log.warning(url + ' returned code ' + responseCode)
    }
  
    // Check the name servers
    // ----------------------
  
    var oldDns = row[OLD_DNS_COLUMN_INDEX]
    var domain = whoisLookup_(url)
    
    if (domain.hasOwnProperty('dns1')) {
    
      var newDns = domain.dns1
      
      if (oldDns === domain.dns1) {
      
        row[CHANGED_COLUMN_INDEX] = 'No'
        
      } else {
      
        nsChanged = true
        row[NEW_DNS_COLUMN_INDEX] = newDns
        row[CHANGED_COLUMN_INDEX] = 'Yes'
        Log.warning(url + ' dns1 changed to ' + newDns + ' from ' + oldDns)      
      }
      
    } else {
    
      row[CHANGED_COLUMN_INDEX] = '?'      
    }
    
    // Check not due to expire soon
    // ----------------------------
    
    // Check that there is a "expiration" sheet for this domain

    // If not create one and update the formulas
    var expirationSheet = SpreadsheetApp.getActive().getSheetByName(url)
    
    if (expirationSheet === null) {
      
      var value = 'https://www.whois.com/whois/' + url
      var formula = '=IMPORTXML("' + value + '","//div[@class=' + "'df-value'" + ']")'
      
      expirationSheet = SpreadsheetApp.getActive().insertSheet(url)
      expirationSheet.getRange('A1').setFormula(formula)
      SpreadsheetApp.flush();
      Log.info('Created new raw sheet for ' + url)  
    }
        
    var today = (new Date()).getTime()
    
    var expiresOn = expirationSheet.getRange(EXPIRATION_DATE).getValue()

    if (expiresOn instanceof Date) {
    
      var expirationDate = expiresOn.getTime() - (DAYS_TO_EXPIRATION * 24 * 60 * 60 * 1000)
      
      if (today > expirationDate) {
  
        row[EXPIRED_COLUMN_INDEX] = 'Yes'
        
        if (row[EMAIL_SENT_COLUMN_INDEX] !== 'Yes') {
          expirationDateSoon = true
          Log.warning(url + ' expiration date soon ' + expiresOn)            
          sendEmail_('DOMAIN WARNING: Expiration date soon')
          row[EMAIL_SENT_COLUMN_INDEX] = 'Yes'
        }
          
      } else {
      
        row[EXPIRED_COLUMN_INDEX] = 'No'
      }
      
    } else {
    
      row[EXPIRED_COLUMN_INDEX] = '?'      
    }

    row[EXPIRATION_DATE_COLUMN_INDEX] = expiresOn
    row[LAST_CHECKED_COLUMN_INDEX] = new Date()

  })
  
  // Log results
  
  if (urlFailed) {
  
    sendEmail_('DOMAIN WARNING: URL response not 200')
    
  } else if (nsChanged) {
  
    sendEmail_('DOMAIN WARNING: Name Server Changed')
    
  } else if (expirationDateSoon) {
  
    // Checked in loop as need row number
        
  } else {

    sendEmail_('DOMAIN MONITOR: All domains OK')
    Log.info('URLs responding OK (200), name servers unchanged expiration date not for a while')
  }
  
  data.unshift(header)
  dataRange.setValues(data)
  
  domainSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(DOMAINS_SHEET_NAME)
  
  SpreadsheetApp.setActiveSheet(domainSheet)
  
} // checkDomains()

/**
 * Peform a Network Service Lookup 
 * 
 * @param {String} domainName A well-formed domain name to resolve
 *
 * @return {String} resolved DNS1
 */
 
function whoisLookup_(domainName) {

  Log.functionEntryPoint()

  var url = 'http://whoiz.herokuapp.com/lookup.json?url=' + domainName
  var result = UrlFetchApp.fetch(url, {muteHttpExceptions:true})
  var responseCode = result.getResponseCode()
  var response = result.getContentText()
  
  if (responseCode !== 200) {
    throw new Error(response.message)
  } 

  var data = JSON.parse(response)
  var domain = {}
  
  if (data.hasOwnProperty('nameservers')) {
    if (data.nameservers.length > 0) {
      domain.dns1 = data.nameservers[0].name 
    }
  }
  
  if (!domain.hasOwnProperty('dns1')) {
    sendEmail_('Could not get DNS for ' + domainName)
  }
  
  if (data.hasOwnProperty('expires_on') && data.expires_on !== null) { 
  
    var dateString = data.expires_on.slice(0,10)
    var year = dateString.slice(0,4)
    var month = dateString.slice(5,7) - 1
    var date = dateString.slice(8)
    domain.expires_on = new Date(year,month,date)
  } 
  
  return domain
  
} // whoisLookup_()

/**
  * Send a notification email
  */
  
function sendEmail_(message) {

  var EMAIL = PropertiesService.getScriptProperties().getProperty('EMAIL')
  var SHEET_URL = PropertiesService.getScriptProperties().getProperty('SHEET_URL')

  if (EMAIL === null || SHEET_URL === null) {
    throw new Error('No email address and/or sheet URL')
  }
  
  var html = HtmlService
    .createHtmlOutput('<a href="' + SHEET_URL + '">Results sheet</a>')
    .getContent()
  
  if (EMAIL !== '' && EMAIL !== null) {
    MailApp.sendEmail(EMAIL, message, SHEET_URL, {htmlBody: html})     
  } else {
    Log.warning('No email address stored in props')
  }
  
} // sendEmail_()
  
function test_whoisLookup() {
  var dns1 = whoisLookup_('andrewroberts.net')
  return
}