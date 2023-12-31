
const letterDictionary = Object.fromEntries(
  Array.from('abcdefghijklmnopqrstuvwxyz').map((letter, index) => [letter.toString(), index + 1]));

class RangeSet { 
  /* 
  Class for mapping between String ranges to Numerical ranges
  =========================================
  */
  constructor(colRange) {
   
    this.range = colRange
    const [left, right] = colRange.split(':');
    this.letterColumnLeft = left[0];
    this.letterColumnRight = right[0];

    this.rowNumLeft = left.slice(1);
    this.rowNumRight = right.slice(1);
    this.columnNumLeft = letterDictionary[this.letterColumnLeft.toLowerCase()]
    this.columnNumRight = letterDictionary[this.letterColumnRight.toLowerCase()]


  }  
};

const columnIndexes = {
  checkBox      : 1,
  lname         : 2,
  fname         : 3,
  fatherhood    : 4,
  dateOfBirth   : 8,
  ipn           : 9,
  passportNo    : 10,
  address       : 12,
  campName      : 13,
  mobilePhone   : 21,
  email         : 22,
  contract      : 23  
};


const ukrainianMonths = [
  'січня','лютого','березня',
  'квітня','травня','червня',
  'липня','серпня','вересня',
  'жовтня','листопада','грудня'
  ];


const CURSES = [

  "Най тебе качка копне..",
  "А щоб той код луснув..",
  "A, cобача кров!",
  "Дідько б його вхопив..",
  "Такий код до діла, як свиня штани наділа..",
  "Дурна програма, як сало без хліба",
  "А бодай його чорти вхопили…",
]

const ERROR_MSGS = {
  GPT_ERROR : 'Помилка на стороні ChatGPT. Спробуйте ще раз.',
  GPT_RATE_LIMIT_EXCEEDED_ERROR: 'Помилка на стороні ChatGPT. Зачекайте 2 хвилини.',
  DEFAULT_ERROR : 'Сталась помилка. Зачекайте декілька хвилин.',
  NO_ROWS_SELECTED_INFO : 'Оберіть рядки, для яких потрібно згенерувати договори.',

}

  
// Global settings
const config = Object.values(getConfig())
let [ 
  PLACEHOLDER_TEMPLATE, 
  MAX_AMT_TRIGGERS,
  DOC_TEMPLATE_ID,
  NEW_DOCS_FOLDER_ID, 
  JOURNAL_ID, 
  JOURNAL_SHEET_NAME, 
  JOURNAL_TABLE_RANGE, 
  SOURCE_SHEET_NAME,
  SOURCE_TABLE_RANGE, 
  TRIGGER_HANDLER_NAME,
  EMAIL_RECIPIENT,
  OPENAI_API_TOKEN,
  AWAIT_LOCK_MS
  ] = config;

const DOC_TEMPLATE  = DriveApp.getFileById(DOC_TEMPLATE_ID);
const DEST_FOLDER   = DriveApp.getFolderById(NEW_DOCS_FOLDER_ID);
const SOURCE_SHEET  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET_NAME);
SOURCE_TABLE_RANGE  = new RangeSet(SOURCE_TABLE_RANGE)
JOURNAL_TABLE_RANGE = new RangeSet(JOURNAL_TABLE_RANGE)


class UserData {
  /* 
  Class to store data of users
  =========================================
  */

  constructor(rowData) {
    let [
      DocCreatedCheckBox,
      FirstName, 
      LastName, 
      Fatherhood, 
      Address, 
      IPN, 
      PassportNo, 
      Camp,
      DateOfBirth,
      MobilePhone,
      Email,
      rowIdx 
      , // skipping last element
    ] = rowData;


    this.FirstName = String(FirstName)
    this.LastName = String(LastName)
    this.Fatherhood = String(Fatherhood)
    this.Address = String(Address)
    this.IPN = String(IPN)
    this.PassportNo = String(PassportNo)
    this.MobilePhone = String(MobilePhone)
    this.PassportNo = String(PassportNo)
    this.DateOfBirth = String(DateOfBirth)
    this.DocCreatedCheckBox = DocCreatedCheckBox
    this.Email = String(Email)
    this.FullName = `${this.LastName} ${this.FirstName} ${this.Fatherhood}`;
    this.LastNameCapitalized = this.LastName.toUpperCase();


    this.FirstName = this.FirstName.trim()
    this.LastName = this.LastName.trim()
    this.Fatherhood = this.Fatherhood!==undefined ? this.Fatherhood.trim() : this.Fatherhood
    this.Address = this.Address!==undefined ? this.Address.trim() : this.Address
    this.IPN = this.IPN!==undefined ? this.IPN.trim() : this.IPN
    this.PassportNo = this.PassportNo!==undefined ? this.PassportNo.trim() : this.PassportNo
    this.MobilePhone = this.MobilePhone.length!==0 ? this.MobilePhone.trim().padStart(13, '+380'): this.MobilePhone
    this.RAW_Camp = Camp!==undefined ? Camp.trim() : Camp
    this.DateOfBirth = this.DateOfBirth!==undefined ? this.DateOfBirth.trim() : this.DateOfBirth
    this.DocCreatedCheckBox = this.DocCreatedCheckBox
    this.Email = this.Email!==undefined ? this.Email.trim() : this.Email
    this.rowIdx = rowIdx

    this.FullName = `${this.LastName} ${this.FirstName} ${this.Fatherhood}`;
    this.LastNameCapitalized = this.LastName.toUpperCase();
  }
}


function getMenu() {
  /* 
  Function to render menu for contracts for users to interract with the script
  =========================================
  */
  const UI = SpreadsheetApp.getUi();
  const menu = UI.createMenu('Договори');

  menu.addItem('Синхронізувати з журналом та згенерувати договори', 'main');
  menu.addToUi();
}

function main() {
  /* 
  Main entry to the script. Calls next function
  =========================================
  */
  var UI  = SpreadsheetApp.getUi();
    // var selectedData = SOURCE_SHEET.getActiveRangeList().getRanges()//.map(el=> el.getValues()[0]) // doesnt work with SHIFT + M1
  var lock = LockService.getScriptLock();


  // Check if script is already running
  try {
      lock.waitLock(AWAIT_LOCK_MS);
  } catch (e) {
      console.log(`INFO:Could not obtain lock after ${AWAIT_LOCK_MS/1000} seconds.`);
      UI.alert('От халепа...Скрипт уже запущений кимось.', 'Зачекайте декілька хвилин.',  UI.ButtonSet.OK)
      return
  }

  try
  {

    const maxColNumber = SOURCE_SHEET.getLastColumn()
      var selectedRange = SOURCE_SHEET.getActiveRange() // works with SHIFT + M1 and MOUSE, but not CTRL + M1
      var selectedData = selectedRange.getValues()
      
      // Check if user accidentally selected something and tries to create contracts for them
      if (maxColNumber!==selectedData[0].length || selectedData==null) {   
        return UI.alert('Увага!', ERROR_MSGS.NO_ROWS_SELECTED_INFO, UI.ButtonSet.OK)
      }
      
      
      // if (maxColNumber!==selectedData[0].length || selectedData==null) {   
      //   const lastRow =  SOURCE_SHEET.getLastRow()
      //   // _getLastRow(SOURCE_SHEET, SOURCE_TABLE_RANGE.columnNumLeft, SOURCE_TABLE_RANGE.columnNumRight)  
      //   selectedRange = SOURCE_SHEET.getRange(2, SOURCE_TABLE_RANGE.columnNumLeft, lastRow, SOURCE_TABLE_RANGE.columnNumRight-SOURCE_TABLE_RANGE.columnNumLeft + 1); 
      
      //   selectedData = selectedRange.getValues()//.slice(1);
      // } 
      let rowNumber = null;
      selectedData = selectedData.map((el,idx) => {
        rowNumber = selectedRange.getCell(idx+1, 1).getRow();
        return [...el, rowNumber]
      } )


      сopyToJournal(selectedData, JOURNAL_ID, JOURNAL_SHEET_NAME, UI)  
      lock.releaseLock();    
      



      
     
      


  }

  catch (error)
  {
    console.log(error)
    const today = formatDate(new Date(), format='DDMMYYYY HHMMSS', sep='/');
    
    GmailApp.sendEmail(EMAIL_RECIPIENT, 'Contracts & Journal Automation Failure', `A failed execute happened at: ${today}`)
    const curseIndex = generateRandomInt(min=0, max=CURSES.length)


    return UI.alert(CURSES[curseIndex], 
      Object.values(ERROR_MSGS).includes(error.message) ? error.message : ERROR_MSGS.DEFAULT_ERROR,  UI.ButtonSet.OK)
    // return UI.alert('Не вдалось запустити операцію', 'Зачекайте декілька хвилин',  UI.ButtonSet.OK)
  }

}

