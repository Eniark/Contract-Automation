function generateRandomInt(min, max) {
  /*
    Generates random integer between min and max
  */
    let difference = max - min;
    let rand = Math.random();
    rand = Math.floor( rand * difference);
    rand = rand + min;

    return rand;
}



function callChatGPT(prompt) {
  /*
    Calls ChatGPT api with a given prompt. Uses chat-gpt_3.5_turbo model.
    Parameters:
    =========================================
      prompt: a string that is passed to ChatGPT
  */
  if (OPENAI_API_TOKEN === null) {
    throw new Error('ChatGPT API Key script property is missing');
  }                  
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_TOKEN}`,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      messages: [{
        role: 'user',
        content: prompt
        }],
      model: 'gpt-3.5-turbo',
      temperature: 1,
      max_tokens: 2000,
    }),
  };
  const response = UrlFetchApp.fetch(apiUrl, options);
  const content = JSON.parse(response.getContentText());


  if (content.hasOwnProperty('error'))
  {
    throw new Error(ERROR_MSGS.GPT_RATE_LIMIT_EXCEEDED_ERROR)
  }
  console.log(content)

  return content;

}



function formatDate(date, format='DDMMYYYY', sep='.') {
  /*
    Formats the date according to <format> parameter
    Parameters:
    =========================================
      date: a Date object
      format: format into which to convert the string. Supports only 'DDMMMYYYY' and 'DDMMYYYY'
      sep: separator for <format> parameter. If null, returns array in [day, month, year] format
  */
  let month, day, year;  
  const [dd,mm,yyyy] = [date.getDate(),date.getMonth(), date.getFullYear()]

  if (format=='DDMMYYYY'){
    day = String(dd).padStart(2, '0');
    month = String(mm + 1).padStart(2, '0');
    year = String(yyyy);

  }
  else if (format=='DDMMMMYYYY'){
    day = dd;
    month = ukrainianMonths[mm];
    year = yyyy;
  }

  else if (format='DDMMYYYY HHMMSS') {
    const day = String(dd).padStart(2, '0');
    const month = String(mm + 1).padStart(2, '0'); // Months are zero-based
    const year = yyyy;
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');

    return `${day}${sep}${month}${sep}${year} ${hours}:${minutes}:${seconds}`;
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

function getLastFridayFromDate(date) {
  /* 
    Gets last Friday from a given date
    Parameters:
    =========================================
      date: an object of class Date
  */
  const daysUntilLastFriday = (date.getDay() + 2) % 7; 
  const lastFriday = new Date(date.getTime() - daysUntilLastFriday * 24 * 60 * 60 * 1000); 
  return lastFriday;
  }