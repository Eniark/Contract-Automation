

var nSkippedUsers = 0;

function сopyToJournal(data, journal_ID, journalSheetName, UI) {
  /* 
    Copies records without a contract to the journal -> obtains IDs for each contract -> calls function to generate new contract files
    Parameters:
    =========================================
      journal_ID: ID of target journal file
      journalSheetName: the name of sheet in the journal file
      UI: variable that holds the UI of the Spreadsheet
  */


    
     
    const journal = SpreadsheetApp.openById(journal_ID).getSheetByName(journalSheetName)
    const currentAmtTriggers = getCurrentTriggers().length // Check if script is already running
    if (currentAmtTriggers > 0) {
      return UI.alert('Не вдалось запустити операцію', 'Зачекайте декілька хвилин',  UI.ButtonSet.OK)
    }
      
    let rowsToCreateNewDocsFor = data.filter(row =>row[columnIndexes.contract] === '' && row[columnIndexes.fname]!=='' 
                                                    && row[columnIndexes.lname]!=='' 
                                                    && row[columnIndexes.dateOfBirth]!=='') // name, lname and date of birth must be present
      
      
      

    rowsToCreateNewDocsFor = rowsToCreateNewDocsFor.filter(el => el!==undefined).map(el => 
    new UserData(
      [
        el[columnIndexes.checkBox],
        el[columnIndexes.fname], 
        el[columnIndexes.lname], 
        el[columnIndexes.fatherhood], 
        el[columnIndexes.address], 
        el[columnIndexes.ipn], 
        el[columnIndexes.passportNo], 
        el[columnIndexes.campName], 
        formatDate(el[columnIndexes.dateOfBirth], format='DDMMYYYY'),
        el[columnIndexes.mobilePhone],
        el[columnIndexes.email],
        el.at(-1)
    ])) // el.at(-1) - index in excel spreadsheet



    const journalLastRow = _getLastRow(journal, JOURNAL_TABLE_RANGE.columnNumLeft, JOURNAL_TABLE_RANGE.columnNumRight)

    let lastContractID = journal.getRange(journalLastRow, JOURNAL_TABLE_RANGE.columnNumLeft).getValue()

    rowsToCreateNewDocsFor = fillNextIDs(rowsToCreateNewDocsFor, lastContractID)
    rowsToCreateNewDocsFor = rowsToCreateNewDocsFor.map(getFullCampAndDate).filter(el => el!==undefined) 

    if (rowsToCreateNewDocsFor.length>0)
      {
      const rangeForInsert = journal.getRange(journalLastRow + 1, 
                                              JOURNAL_TABLE_RANGE.columnNumLeft, 
                                              rowsToCreateNewDocsFor.length, 
                                              JOURNAL_TABLE_RANGE.columnNumRight - JOURNAL_TABLE_RANGE.columnNumLeft + 1 )

      journal.setActiveRange(rangeForInsert)
        .setValues(rowsToCreateNewDocsFor.map(obj => [
          obj.ContractID, 
          new Date(), 
          obj.FullName, 
          obj.CampDate, 
          obj.Camp]
        ))

      createNewGoogleDocs(rowsToCreateNewDocsFor, UI) // call next function
    }
}

function getFullCampAndDate(userObject) {
  /* 
  Preprocesses Camp input. Gets CampName and Camp dates
  =========================================
  Parameters:
    userObject: an object of UserData class
  */
  try 
  {
    if (userObject.Camp!==undefined) {

      const regex = new RegExp(`(?<Prefix>БУР-табір\\s*в)?\\s*(?<Location>[мс]\\s*\.\\s*[А-Яа-яіїє]+)\\s*(?<Region>[А-Яа-яіїє]+\\s*[А-Яа-яіїє]+)?\\s*\\((?<StartDate>\\s*\\d+\\s*[а-яА-Яіїє]*)\\s*-\\s*(?<EndDate>\\d+\\s*[а-яА-Яіїє]*)\\s*\\)`, "u");
      let [prefix, location, region, startDate, endDate] = Object.values(regex.exec(userObject.Camp).groups);

      let [endDayNo, endUkrainianMonth] = endDate.split(' ')
      let endMonthNo = ukrainianMonths.indexOf(endUkrainianMonth);
      let startDayNo = startDate;
      let startUkrainianMonth;
      let startMonthNo = endMonthNo;
      if (startDate.length > 2) {
        [startDayNo, startUkrainianMonth] = startDate.split(' ')
        startMonthNo = ukrainianMonths.indexOf(startUkrainianMonth);
      }
      const todaysYear = new Date().getFullYear(); 
      let endYear = todaysYear;
      if (endMonthNo==0)
        endYear++;
      
      startDate = new Date(todaysYear, startMonthNo,startDayNo )
      endDate = new Date(endYear, endMonthNo, endDayNo)
      startDateFormated = formatDate(startDate, format='DDMMYYYY', sep='.');
      endDateFormated = formatDate(endDate, format='DDMMYYYY', sep='.');

           
      const fullCampDate = startDateFormated + ' - ' + endDateFormated;
      userObject.CampStartDate = startDateFormated;
      userObject.CampEndDate = endDateFormated;
      userObject.CampDate = fullCampDate;
      userObject.Camp = (prefix===undefined ? '' : prefix) + ' ' +  
        (location===undefined ? '' : location) + ' ' + (region===undefined ? '' : region);
    }


    return userObject

    
  }
  catch (error) 
  {
    console.log(error)
    console.log(`Skipping ${++nSkippedUsers} record: \n\tFirstName ${userObject.FirstName}\n\tLastName ${userObject.LastName}`)
  }
}


function fillNextIDs(dataToInsert, startingID) {
  /* 
    Function increments last ID in journal and prepends result to each element in array
    Parameters:
    =========================================
      dataToInsert: array to prepend IDs
      startingID: number to start incrementing
  */
  let journal_ID = startingID ;
  dataToInsert.forEach(el => el.ContractID = ++journal_ID);
  return dataToInsert

}


