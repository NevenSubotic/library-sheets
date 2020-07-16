// constants
const SHEET = SpreadsheetApp.getActiveSheet();
const LOG_IT = true;
const HEADER_ROW = 1;

function logIt(){
  if( !LOG_IT ){
    return
  }
  for (let i = 0; i < arguments.length; i++) {
    console.log(arguments[i]);
  }
}

function displayDraftsSelector(){
  const html = HtmlService.createTemplateFromFile("draftSelectionScreen").evaluate().setWidth(400).setTitle("Auswahl der Email");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getDraftsArr(){
  let output = [];
  const messages = GmailApp.getDraftMessages();
  
  messages.forEach( message => {  
    output.push({
      id: message.getId(),
      subject: message.getSubject()
    });
  });
  
  return JSON.stringify(output)  
}

function handleFormSubmit( draftId ){
  try {
    const email = getEmailInfo( draftId );
    if( !email ){
      throw "Could not get Draft";
    }
    const selectedRows = getSelectedRowsAsObjInArr( SpreadsheetApp.getActiveSheet(), 1 );
    if( !selectedRows.length ){
      throw "No rows selected";
    }      
    selectedRows.forEach( row => {
      const options = {
        htmlBody: fillPlaceholders(email.body, row),
        attachments: email.attachments        
      };
      GmailApp.sendEmail( row.Email, email.subject, options.htmlBody, options);      
    });    
    
  } catch( err ) {
    SpreadsheetApp.getUi().alert("ERROR: " + err)
  }
  
  function fillPlaceholders( body, valuesAsObj ){
    const headersArr = SHEET.getRange(HEADER_ROW, 1, 1, SHEET.getLastColumn()).getValues()[0];   
    const placeholdersInBody = getPlaceholdersInBody_( body );
    
    placeholdersInBody.forEach( placeholder => {
      const placeholderWithoutCurly = placeholder.substring(1, placeholder.length-1);
      const replaceText = valuesAsObj[placeholderWithoutCurly];
      if( headersArr.indexOf(placeholderWithoutCurly) === -1 || !replaceText ){
        return
      }    
      logIt(`Replacing ${placeholder} with value ${replaceText}`);
      body = body.replace(placeholder, replaceText);
    });
    
    return body
        
    function getPlaceholdersInBody_(body){
      const searchPattern = /{\w+}/g;
      const foundPlaceholdersInBody = body.match(searchPattern);
      logIt("Found Placeholders in Email: "+foundPlaceholdersInBody);
      return foundPlaceholdersInBody
    }
  }
  
  function getEmailInfo( draftId ){
    const draft = GmailApp.getMessageById( draftId );
    if( !draft ){
      return
    }
    const email = {
      subject: draft.getSubject(),
      body: draft.getBody(),
      attachments: draft.getAttachments()
    };
    logIt("Got email: "+email.subject);
    return email
  }
}


/**
* Checks to see if we are in the correct sheet
* @param {Object} Active Sheet - The curently active Sheet 
* @param {String} Sheet Name - The name of the sheet to test against
* @returns {Boolean} Is Active = Target
*/
function isSheetCorrect( currentSheet, targetSheetName ){
  return currentSheet.getSheetName() == targetSheetName
}


/**
* Extract G-Drive Id from Url
* 
* @param {string} A url to the G-Drive File
* @returns {string} The id
*/
function getFileIdFromFileUrl( url ) { 
  return url.match( /[-\w]{25,}/ )
}


/**
* Returns header as objects with colName : colNum
* 
* @param {Object} Sheet which is used
* @param {number} The header row
* @returns {Object} ColName: ColNum
*/
function getHeaderAsObjFromSheet( sheet, headerRow = 1 ){
  const headerArr = sheet.getRange( headerRow, 1, 1, sheet.getLastColumn() ).getValues()[0];  
  return convertArrToObj_( headerArr );
}

/**
* Returns an array of objects for each selected row headerName : rowValue
* Assumes a continues range is selected, ie no hidden, filtered or multiple ranges
* 
* @param {Object} Sheet - The currently active Sheet
* @param {number} Header Row - The row where the header is located
* @return {Array} Collection of rows as objects with rowNum
*/
function getSelectedRowsAsObjInArr( sheet, headerRow ){  
  const firstRow = sheet.getActiveRange().getRow();
  const lastRow  = sheet.getActiveRange().getLastRow();  
  const lastCol  = sheet.getLastColumn();
  const rowsAsArr = sheet.getRange( firstRow, 1, lastRow - firstRow + 1, lastCol ).getValues();
  const headerObj = getHeaderAsObjFromSheet(sheet, headerRow);
  
  const rowsAsObjInArr = [];
  rowsAsArr.forEach( (row, i) => rowsAsObjInArr.push( HELPER.convertRowArrToObj_(row, headerObj) ));
  rowsAsObjInArr.forEach( (rowObj, i) => rowObj["rowNum"] = firstRow + i );
  return rowsAsObjInArr
}

function headerColFun( header, range ){
  return range.indexOf(header) + 1     
}

function writeToSheet_( row, column, value ){
  SHEET.getRange( row, column ).setValue( value )
}

function writeToAnotherSheet_( sheet, row, col, value ){
  sheet.getRange(row, col).setValue(value)
}

function getHeader_(sheet, headerRow){
  var headerArr = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];  
  var headerObj= convertArrToObj_(headerArr);
  return headerObj
}

function headerColFun(header, range){
  var headerCol = range.indexOf(header);
  return headerCol+1;  
}

function getSelectedRowsAsObjInArr_( headerObj ){
  var firstRow = SHEET.getActiveRange().getRow();
  var lastRow  = SHEET.getActiveRange().getLastRow();  
  var rowsAsArr = SHEET.getRange(firstRow, 1, lastRow - firstRow + 1, SHEET.getLastColumn()).getValues();
  
  var rowsAsObjInArr = [];
  rowsAsArr.forEach(function( rowArr, index ){
    rowsAsObjInArr.push( convertRowArrToObj_(rowArr, headerObj) );
  });
  return rowsAsObjInArr
}

/**
+ Simple message alert to user
* @param {string} Message to display to user
*/
function alert( msg ){
  SpreadsheetApp.getUi().alert(msg)
}

const HELPER = (function(){
  
  function convertArrToObj_( anArray ){
    const asObj = {};
    anArray.forEach( (item, index) => {
                    asObj[item] = index + 1
                    });
    return asObj
  }
  
  function convertRowArrToObj_( rowAsArray, headerObj ){
    var rowAsObj = {};
    for(var header in headerObj){
      rowAsObj[header] = rowAsArray[ headerObj[header]-1 ]
    }
    return rowAsObj
  }
  
  function todayISO_(){
    return new Date().toISOString().substr(0,10)
  }
  
  function convertDateToISO_( aDate ){
    return aDate.toISOString().substr(0,10)
  }
  
  return {
    convertArrToObj_: convertArrToObj_,
    convertRowArrToObj_: convertRowArrToObj_,
    todayISO_: todayISO_,
    convertDateToISO_: convertDateToISO_
  }
  
})();
// moved to HELPER
function convertArrToObj_( anArray ){
  var asObj = {};
  anArray.forEach(function(item, index){
    asObj[item] = index + 1
  });
  return asObj
}

function convertRowArrToObj_( rowAsArray, headerObj ){
  var rowAsObj = {};
  for(var header in headerObj){
    rowAsObj[header] = rowAsArray[ headerObj[header]-1 ]
  }
  return rowAsObj
}

function todayISO(){
  return new Date().toISOString().substr(0,10)
}

/**
* Returns the name of the file based on the inputed url
* Use this within apps script, not a cell formula
* @param {string} url of the file
* @returns {string} name of the file
*/
function getFilenameFromUrl( url ){
  const ss    = SpreadsheetApp.getActive();
  
  const cleanUrl = url.replace("https://drive.google.com/open?id=","");
  const file = DriveApp.getFileById( cleanUrl );
  
  if( !file ){
    return "ERRPR: file not found";
  }
  
  return file.getName()  
}

Date.prototype.addHours = function(h) {
  this.setTime(this.getTime() + (h*60*60*1000));
  return this;
}
