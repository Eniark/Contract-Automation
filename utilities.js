function formatDate(date, format='DDMMYYYY', sep='.') {
  /*
    Formats the date according to <format> parameter
    Parameters:
    =========================================
      date: an object of class Date
      format: format into which to convert the string. Supports only 'DDMMMYYYY' and 'DDMMYYYY'
      sep: separator for <format> parameter. If null, returns array in [day, month, year] format
  */
  let month, day, year;  

  if (format=='DDMMYYYY'){
    day = String(date.getDate()).padStart(2, '0');
    month = String(date.getMonth() + 1).padStart(2, '0');
    year = String(date.getFullYear());

  }
  else if (format=='DDMMMMYYYY'){
    day = date.getDate();
    month = ukrainianMonths[date.getMonth() - 1];
    year = date.getFullYear();

  }
  
  return sep===null ? [day, month, year] : `${day}${sep}${month}${sep}${year}`;
}

function getCurrentTriggers () { 
  /*
    Gets triggers of this project
  */

  let triggers = ScriptApp.getProjectTriggers().filter((el) => el.getHandlerFunction()==TRIGGER_HANDLER_NAME)
  return triggers
} 

function deleteCurrentTriggers() {
  /*
    Deletes all triggers of this project
  */

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
   ScriptApp.deleteTrigger(triggers[i]);
 }
}

function _getLastRow(spreadSheet, startCol, endCol, TOLERANCE=10) { 
  /* 
    Function gets last row number in table
    Parameters:
    =========================================
      spreadSheet: object of spreadsheet
      startCol: left column boundary
      endCol: right column boundary
      TOLERANCE: how many empty rows to tolerate before declaring a row as end row. 
  */
  let lastRow = 1;
  let valuesBetween = null;
  let newLastRow = null;
  let isEmpty = false
  while (!isEmpty) {
    newLastRow = spreadSheet.getRange(lastRow, startCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();

    valuesBetween = spreadSheet.getRange(newLastRow + 1, startCol, TOLERANCE, endCol-startCol + 1).getValues()
    isEmpty = valuesBetween.every(row => row.every(el => el.toString().trim()===''))
    lastRow = newLastRow
  }

  return newLastRow
};
