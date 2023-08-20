

var errorCount = 0;

function сopyToJournal(data, journal_ID, journalSheetName, UI) {
  /* 
    Copies records without a contract to the journal -> obtains IDs for each contract -> calls function to generate new contract files
    Parameters:
    =========================================
      journal_ID: ID of target journal file
      journalSheetName: the name of sheet in the journal file
      UI: variable that holds the UI of the Spreadsheet
  */


    
     
    const journal = SpreadsheetApp.openById(journal_ID).getSheetByName(journalSheetName);
    // const currentAmtTriggers = getCurrentTriggers().length // Check if script is already running
    // if (currentAmtTriggers > 0) {
    //   return UI.alert('Не вдалось запустити операцію', 'Зачекайте декілька хвилин',  UI.ButtonSet.OK);
    // }
      
    let rowsToCreateNewDocsFor = data.filter(row =>row[columnIndexes.contract] === '' && row[columnIndexes.fname]!=='' 
                                                    && row[columnIndexes.lname]!=='' 
                                                    && row[columnIndexes.dateOfBirth]!==''); // name, lname and date of birth must be present
      
      
      

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
    ])); // el.at(-1) - index in excel spreadsheet




    const journalLastRow = _getLastRow(journal, JOURNAL_TABLE_RANGE.columnNumLeft, JOURNAL_TABLE_RANGE.columnNumRight);

    let lastContractID = journal.getRange(journalLastRow, JOURNAL_TABLE_RANGE.columnNumLeft).getValue();

    rowsToCreateNewDocsFor = rowsToCreateNewDocsFor.map(getFullCampAndDate).filter(el => el!==undefined) ;
    
    rowsToCreateNewDocsFor = fillNextIDs(rowsToCreateNewDocsFor, lastContractID);
    const gptPrompt = 'Напиши ці ПІБ, якби вони були вписані у речення "Договір укладений між...":\n' + rowsToCreateNewDocsFor.map((el, idx)=>[idx+1+'. ', el.FullName+'\n']).flat().join('') + `${rowsToCreateNewDocsFor.length+1}. Гуліватий Юрій Дмитрович`;


    let gptResponse = callChatGPT(gptPrompt).choices[0].message.content;
    if (gptResponse.includes(':')) {

        gptResponse = gptResponse.split(':');
        gptResponse = gptResponse[gptResponse.length - 1].trim();

      }


      console.log(gptResponse)

      gptResponse= gptResponse.split('\n').map(el=> { 
          let content = el.split(' ').slice(-4).filter(el => el[0]===el[0].toUpperCase()).join(' ')
          const cleaned_OrudnyiFullname = content.replace(/^[,!\\?:;\\.]*/, '').replace(/[,!\\?:;\\.]*$/,'') // strip unneeded characters
          return cleaned_OrudnyiFullname
          })


      const testCase = gptResponse.slice(-1)[0].split(' ');
      if (!testCase.includes('Гуліватим'))
      {
        throw new Error(ERROR_MSGS.GPT_ERROR)
      }

      rowsToCreateNewDocsFor = rowsToCreateNewDocsFor.map((obj, idx) => {
        obj.Orudnyi_FullName = gptResponse[idx];
        return obj;
      });

      console.log(rowsToCreateNewDocsFor);


    if (rowsToCreateNewDocsFor.length>0)

      {
      rowsToCreateNewDocsFor = createNewGoogleDocs(rowsToCreateNewDocsFor, UI)
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

    
    }

    if (errorCount>0) {
      const curseIndex = generateRandomInt(min=0, max=CURSES.length)
      UI.alert(CURSES[curseIndex], 'Деяким волонтерам не вдалось згенерувати договори.',  UI.ButtonSet.OK)
    // ui.alert('Сумулька 😢', 'Деяким волонтерам не вдалось згенерувати договори.',  UI.ButtonSet.OK)
  }
}

function getFullCampAndDate(userObject, idx) {
  /* 
  Preprocesses Camp input. Gets CampName and Camp dates
  =========================================
  Parameters:
    userObject: an object of UserData class
  */
  try 
  {
    if (userObject.Camp!==undefined) {
      const ukrLetters = "[А-Яа-яіїєІЇЄ'`\-]"
      const regexPrefix = "(?<Prefix>БУР-табір\\s*в)"
      const regexLocation = `(?<Location>[смт{3}|с|м]+\\s*\.\\s*${ukrLetters}+)`
      const regexRegion = `(?<Region>${ukrLetters}+\\s*${ukrLetters}+)`
      const regexDate = `(?<StartDate>\\s*\\d+\\s*${ukrLetters}*)\\s*-\\s*(?<EndDate>\\d+\\s*${ukrLetters}*)`

      const regexCombinedString =  new RegExp(`${regexPrefix}?\\s*${regexLocation}\\s*${regexRegion}?\\s*\\(${regexDate}\\s*\\)`, "u");

      let [prefix, location, region, startDate, endDate] = Object.values(regexCombinedString.exec(userObject.Camp).groups);
      let [endDayNo, endUkrainianMonth] = endDate.split(' ')
      let endMonthNo = ukrainianMonths.indexOf(endUkrainianMonth);
      let startDayNo = startDate;
      let startUkrainianMonth;
      let startMonthNo = endMonthNo;

      if (startDate.length > 2) {
        [startDayNo, startUkrainianMonth] = startDate.split(' ')
        startMonthNo = ukrainianMonths.indexOf(startUkrainianMonth);
      }

      if (endMonthNo===-1 || startMonthNo===-1) {
        throw new Error(`No such months in UkrainianMonths`)
      }


      const todaysYear = new Date().getFullYear(); 
      let endYear = todaysYear;
      if (endMonthNo==0) {
        endYear++;
      }
      
      startDate = new Date(todaysYear, startMonthNo,startDayNo )
      endDate = new Date(endYear, endMonthNo, endDayNo)

      startDateFormated = formatDate(startDate, format='DDMMYYYY', sep='.');
      endDateFormated = formatDate(endDate, format='DDMMYYYY', sep='.');

           
      const fullCampDate = startDateFormated + ' - ' + endDateFormated;
      userObject.CampStartDate = startDateFormated;
      userObject.CampEndDate = endDateFormated;
      userObject.CampDate = fullCampDate;
      userObject.Camp = ((prefix===undefined ? '' : prefix) + ' ' +  
        (location===undefined ? '' : location) + ' ' + (region===undefined ? '' : region)).trim();
    }

    return userObject

    
  }
  catch (error) 
  {
    errorCount++;
    console.log(error)
    console.log(`Skipping record for Journal: \n\tFirstName ${userObject.FirstName}\n\tLastName ${userObject.LastName}
      \n\tRow Index: ${userObject.rowIdx}`)
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


