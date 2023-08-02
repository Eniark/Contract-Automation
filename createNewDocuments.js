

function replaceBodyPlaceholders(body, replacements) {
  /* 
  Replaces labels in documents with data
  Parameters:
  =========================================
    body: the document
    replacements: the key-value pairs to replace
  */
  for (const [key, value] of Object.entries(replacements)) {
    const preparedPlaceholder = PLACEHOLDER_TEMPLATE.replace("TEXT", key);
    body.replaceText(preparedPlaceholder, value);
  }

  return body;
}

function copyData(dataToCopy) { 
  /* 
  Creates new contracts in specified in config folder and fills them with data
  Parameters:
  =========================================
    dataToCopy: an array of records that need to be 
  */
  /* Pass only rows that need to be copied! */
  const len = dataToCopy.length - 1;
  const start = Date.now();
  let idx = 0;
  while (idx <= len) {
    const userData = dataToCopy[idx];
    try
    {
      Logger.log(`INFO:Processing #${idx + 1}... out of ${dataToCopy.length}`);
        
      const copiedDocument = DOC_TEMPLATE.makeCopy(`${userData.ContractID} ${userData.FirstName} ${userData.LastName}`, DEST_FOLDER);
      const doc = DocumentApp.openById(copiedDocument.getId());
      const body = doc.getBody();
      const today = new Date();
      const newMonth = today.getMonth() + 1
      today.setMonth(newMonth) // needed since formatDate subtracts 1 month
      const [day, month, year] = formatDate(today, format='DDMMMMYYYY', sep=null);
      

      const [startDD,startMM,startYYYY] = userData.CampStartDate.split('.')
      const campStartDate = new Date(startYYYY, startMM, startDD)
      const [endDD,endMM,endYYYY] = userData.CampEndDate.split('.')
      const campEndDate = new Date(endYYYY, endMM, endDD)

      let [campStartDD, campStartMMMM, campStartYYYY] = formatDate(campStartDate, format='DDMMMMYYYY', sep=null);
      let [campEndDD, campEndMMMM, campEndYYYY] = formatDate(campEndDate, format='DDMMMMYYYY', sep=null);
      
      // Replacing placeholders
      const replacements = {
        FullNameNominativ  : userData.FullName,  
        Address            : userData.Address,
        IPN                : userData.IPN,
        PassportNo         : userData.PassportNo,
        DateOfBirth        : userData.DateOfBirth,
        dd                 : day,
        mmmm               : month,
        yyyy               : year,
        ContractID         : userData.ContractID,
        LastNameCapitalized: userData.LastNameCapitalized,
        MobilePhone        : userData.MobilePhone,
        Email              : userData.Email,
        FirstName          : userData.FirstName,
        campStartDD        : campStartDD,
        campStartMMMM      : campStartMMMM,
        campStartYYYY      : campStartYYYY,
        campEndDD          : campEndDD,
        campEndMMMM        : campEndMMMM,
        campEndYYYY        : campEndYYYY,
        Camp               : userData.Camp
      };

      replaceBodyPlaceholders(body, replacements);
      doc.saveAndClose();

      const url = doc.getUrl();
      SOURCE_SHEET.getRange(userData.rowIdx, columnIndexes.contract + 1).setFormula(`=HYPERLINK("${url}";"Договір")`); // +2 to skip header
      // SOURCE_SHEET.getRange(userData.rowIdx, columnIndexes.checkBox + 1).setValue(userData.DocCreatedCheckBox)
    }
    catch (error)
    {
      console.log(error)
      console.log(`Caught error while creating document... \n\tFirstName: ${userData.FirstName}\n\tLastName: ${userData.LastName}`)
    }

    idx++;


  }
    
  let end = Date.now();
  Logger.log(`Finished in ${end - start}ms`);
  }

function createNewGoogleDocs(dataToCopy, UI) { 
  if (dataToCopy.length>30)
  {
    Logger.log(`INFO:Performing asynchronous calls.`);
    const batchAmount = Math.ceil(dataToCopy.length / MAX_AMT_TRIGGERS);  
    let amtOfTriggersToCreate = 0;
    // UI.alert('Операція успішна ✅', '',  UI.ButtonSet.OK)

    for (let i = 0; i < dataToCopy.length; i += batchAmount) {
      const batch = dataToCopy.slice(i, i + batchAmount);
      Async.call('copyData', batch);
      amtOfTriggersToCreate++;

    }
    Logger.log(`INFO:Created ${amtOfTriggersToCreate} triggers.`);

  }
  else
  {
    Logger.log(`INFO:Performing synchronous call with ${dataToCopy.length} records.`);
    Logger.log(`INFO:Created 0 triggers.`);
    // UI.alert('Операція успішна ✅', '',  UI.ButtonSet.OK);
    
    copyData(dataToCopy);


  }
}













