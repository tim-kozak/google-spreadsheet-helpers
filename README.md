# Google script spreadsheet helpers
Useful javascript functions for google spreadsheets

## 1. Tab Manipulations

### 1.1 Get existing tab(sheet)
>Returns tab instance if exists 
```javascript
function getTab(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  return sheet.getSheetByName(name);
}
```

### 1.2 Check if tab(sheet) exists
>Returns true or false if tab exists
```javascript
function isTabExist(name) {
  var tab = getTab(name);
  return tab ? true : false ;
}
```

### 1.3 Create new tab(sheet)
>Creates new tab if name is available
```javascript
function createTab(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!isTabExist(name)) {
    return sheet.insertSheet(name);
  }    
  return null;
}
```

### 1.4 Get or create tab(sheet)
>Returns existing tab or creates one 
```javascript
function getOrCreateTab(name) {
  return isTabExist(name) ? getTab(name) : createTab(name);
}
```

### 1.5 Copy tab(sheet)
>Duplicates tab with new name
```javascript
function copyTab(template,newName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  if (isTabExist(newName) || !isTabExist(template)) 
    return null;

  var newTab = getTab(template).copyTo(sheet);
  newTab.setName(newName);

  //Additional stuff
  //newTab.showSheet(); //unhide if template was hidden
  //sheet.setActiveSheet(newTab) //activate in browser
  //sheet.moveActiveSheet(1); //put infront

  return newTab;
}
```

## 2. Row manipulations

### 2.1 VLookup
>Search for a value in selected column - returns row number
```javascript
function vlookup(tab, columnNumber, searchValue) {
  var lastRow = sheet.getLastRow();
  
  var data=sheet.getRange(1,columnNumber,lastRow,columnNumber).getValues();
  var i;
  for(i=0;i<data.length;++i){
    if (data[i][0]==searchValue){
      return i+1;
    }
  }
  return lastRow;
}
```

## 3. Column manipulations

### 3.1 HLookup
>Search for a value in selected row - returns column address in A1 notation eq: AH
```javascript
function hlookup(sheet, rowNumber, searchValue) {
  var foundData = false;
  var foundIndex = 0;
  const lastColumn = sheet.getLastColumn();
  
  var range = sheet.getRange(rowNumber,1,1,lastColumn);
  var values = range.getValues();
  if (values && values[0]) {
    values = values[0];
    var i;
    for(i=0;i<values.length;++i){
      if (values[i]==searchValue){
        foundIndex = i;
        foundData = true;
        break;
      }
    }
  }
  if (!foundData) 
    return null;

  var columnNumber = foundIndex+1;
  var cellAddr = sheet.getRange(rowNumber, columnNumber).getA1Notation();
  var withNoDigits = cellAddr.replace(/[0-9]/g, '');
  return withNoDigits;
}
```

## 4. Range and Cell manipulations

### 4.1 Get range values
>Select range of cells in specific tab and return values
```javascript
function getRangeValues(tabName,range) {
  const tab = getTab(tabName);
  const rangeValuesMatrix = tab.getRange(range).getValues();
 
  return rangeValuesMatrix;
}
```

### 4.2 Get cell value
>Select a cell in specific tab and return its value
```javascript
function getRangeValue(tabName,range) {
  const tab = getTab(tabName);
  const value = tab.getRange(range).getValue();
  return value;
}
```

## 5. Thread manipulations

### 5.1 Wait/Sleep function
>Wait for specified amount of time to delay the execution. eq: helpful not to get "too many requests" from API  
```javascript
function sleep(milliseconds) {
    var start = new Date().getTime();
    while (new Date().getTime() < start + delay);
}
```

## 6. External requests

### 6.1 Simple GET request
>Perform GET request with some headers and query string
```javascript
function performGETRequest(resource,queryString) {
  const token = "some auth token";
  const host = "https://api.some.com/v1/";
  const url = host + resource + '?'+queryString+'&api_token='+token;
  const options = {
    'Accept': 'application/json',
    "followRedirects" : true,
    "muteHttpExceptions": true
  };
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response);
  
  //example status checks 
  if (!result["success"]) {
    return null;
  }
  
  //Additional stuff
  //sleep(100); //if you perform multiple requests in a loop
  return result["data"];
}
```

### 6.2 Simple PUT request
>Perform PUT request with some headers and data object
```javascript
function performPUTRequest(resource,data) {
  const token = "some auth token";
  const host = "https://api.some.com/v1/";
  const url = host + resource + '?'+queryString+'&api_token='+token;
  const options = {
    "contentType": "application/json",
    "method" : "PUT",
    "followRedirects" : true,
    "muteHttpExceptions": true,
    "payload" : JSON.stringify(data)
  };
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response);
  
  //example status checks 
  if (!result["success"]) {
    return null;
  }
  
  return result["data"];
}
```

## 7. Logging and Debugging

### 7.1 log() function
>Simple logging function with logging levels
```javascript
var LL_ALL = 'LOG_LEVEL_ALL';
var LL_DEBUG = 'LOG_LEVEL_DEBUG';
var LL_INFO = 'LOG_LEVEL_INFO';
var LL_ERROR = 'LOG_LEVEL_ERROR';

var LL = [LL_ERROR,LL_INFO,LL_DEBUG,LL_ALL]

var CURRENT_LL = LL_DEBUG; //LL_ALL, LL_DEBUG, LL_INFO, LL_ERROR 

function log(string,logLevel) {
  const logIndex = LL.indexOf(logLevel);
  const currentlogIndex = LL.indexOf(CURRENT_LL);

  if ( logIndex >=0 && logIndex <= currentlogIndex) {
    console.log(string)
  }
}
```
>Usage:
```javascript
    log('|--| range query - [done]',LL_INFO)
```

### 7.2 Output some debug data to a range
>Use some technical range to output debug data
```javascript
function getDebuggRange(length) {
  const start = 17;
  const end = length + length-1;
  var rangeString = "A"+length+":A"+end;
  return getTab("Config").getRange(rangeString);
}

function clearDebugRange() {
  var range = getDebuggRange(1000);
  range.clearContent();
}

function printSomethingToDebug(data) {
  clearDebugRange();

  if (data.length > 0) {
    var range = getDebuggRange(data.length); 
    range.setValues(data);
  }
}

if( CURRENT_LL == LL_DEBUG ) {
  printSomethingToDebug(data);
}
```
