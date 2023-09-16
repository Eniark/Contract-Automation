

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



function getLastCampNumber(arrOfCamps) {
  /* 
  Get the number of last camp
  Parameters:
  =========================================
    arrOfCamps: a collection of camps 
  */
  const patternForNumber = '^(?<CampNumber>\\w*\\s*\\d+)';
  const patternForRemovingDates = '\\s*\\(.+\\)\\s*'
  const regexForNumber = new RegExp(patternForNumber);
  const regexForRemovingDates = new RegExp(patternForRemovingDates);
  const arrOfCampNumbers = arrOfCamps.map(el => {
    el = el.replace(regexForRemovingDates, '');
    const [campNumber] = Object.values(regexForNumber.exec(el).groups); // null here
    return parseInt(campNumber);
  });
  let lastCampNumber = Math.max(...arrOfCampNumbers);
  return lastCampNumber===-Infinity ? 0 : lastCampNumber
}

function getSubfolderIDs(destFolder, data) {
  /* 
  Gets a map of available <campName>: <forlderId>. Creates new subfolders for contracts in case such don't exist and add to map.
  Parameters:
  =========================================
    destFolder: folder where to create subfolders
    data: an array of records that need contracts 
  */

  let wrapperFolders = destFolder.getFolders();
  let nameIdMapping = {};
  let availableSubFolders = [];
  let availableSubFolders_noPrefix = [];
  const ukrLetters = "[А-Яа-яіїєІЇЄ'`\\-\\s]"
  const regexPatternForCamp = `^(?<Prefix>\\d*\\s*[А-Яа-яіїєІЇЄ'\`\-]*)(?:\\s+(?<Camp>${ukrLetters}+\\s+\\(\\d+\\.\\d+\\.\\d+\\s*-\\s*\\d+\\.\\d+\\.\\d+\\)))?$`;
  const regexCamp = new RegExp(regexPatternForCamp, 'u');
  locations = [...new Set(data.map(el => el.ShortLocation))];
  camps = [...new Set(data.map(el => [`${el.ShortLocation} (${el.CampStartDate} - ${el.CampEndDate})`, new Date(el.CampStartDate)]))];
  camps.sort()
  camps = camps.map(el => el[0]);

  let availableFolders = [];
  while (wrapperFolders.hasNext()) {
    let wrapperFolder = wrapperFolders.next();
    availableFolders.push(wrapperFolder.getName())
  }
  
  console.log(`Available folders: ${availableFolders}`)

  locations.forEach(el => { 
    if (availableFolders.indexOf(el) < 0) {
      destFolder.createFolder(el)
      console.log(`INFO:Created new folder with name: ${el}`);
    }
  });
  
  wrapperFolders = destFolder.getFolders();
  while (wrapperFolders.hasNext()) {
    let wrapperFolder = wrapperFolders.next();
    
    if (locations.includes(wrapperFolder.getName())) {
      const subFolders = wrapperFolder.getFolders();
      availableSubFolders = [];
      availableSubFolders_noPrefix = [];
      
      while (subFolders.hasNext()) {
        let subFolder = subFolders.next();

        // Extract camp name to use later in the Map
        const camp = subFolder.getName();
        const [prefix, noPrefixCamp] = Object.values(regexCamp.exec(subFolder.getName()).groups);
        availableSubFolders.push(camp);
        availableSubFolders_noPrefix.push(noPrefixCamp);
        nameIdMapping[noPrefixCamp===undefined ? prefix : noPrefixCamp] = subFolder.getId();
      }

      const subDestFolder = DriveApp.getFolderById(wrapperFolder.getId());
      let lastCampNumber = getLastCampNumber(availableSubFolders);
      camps.forEach(el => {
        if (availableSubFolders_noPrefix.indexOf(el) < 0 && el.includes(wrapperFolder.getName())) {
          const f = subDestFolder.createFolder(`${++lastCampNumber} заїзд ${el}`);
          console.log(`INFO:Created new subfolder with name: ${lastCampNumber} заїзд ${el}`);
          nameIdMapping[el] = f.getId();
        }
      });
    }

  }
    

  console.log(`INFO:Available Folders:\n${Object.keys(nameIdMapping)}`);
  console.log(nameIdMapping)
  return nameIdMapping;
  }


function copyData(dataToCopy) { 
  /* 
  Creates new contracts in specified in config folder and fills them with data
  Parameters:
  =========================================
    dataToCopy: an array of records that need contracts
  */

  const len = dataToCopy.length - 1;
  const start = Date.now();
  let idx = 0;
  let erronous_idxs = [];

  nameIdMapping = getSubfolderIDs(DEST_FOLDER, dataToCopy);
  while (idx <= len) {
    const userData = dataToCopy[idx];
    try
    {
      console.log(`INFO:Processing #${idx + 1}... out of ${dataToCopy.length}`);
      let subDestinationFolderID = nameIdMapping[`${userData.ShortLocation} (${userData.CampStartDate} - ${userData.CampEndDate})`]
      const subDestinationFolder = DriveApp.getFolderById(subDestinationFolderID)

      const copiedDocument = DOC_TEMPLATE.makeCopy(`${userData.ContractID} ${userData.FirstName} ${userData.LastName}`, subDestinationFolder);
      const doc = DocumentApp.openById(copiedDocument.getId());
      const body = doc.getBody();

      const [startDD,startMM,startYYYY] = userData.CampStartDate.split('.')
      const campStartDate = new Date(startYYYY, startMM-1, startDD)
      const [endDD,endMM,endYYYY] = userData.CampEndDate.split('.')
      const campEndDate = new Date(endYYYY, endMM-1, endDD)
      let contractGeneratedDate = new Date(campStartDate); // Create a new Date object to avoid modifying the original date
      contractGeneratedDate.setDate(contractGeneratedDate.getDate() - 1);
      if ([0,6].includes(contractGeneratedDate.getDay())) // Check if Saturday or Sunday
      {
        contractGeneratedDate = getLastFridayFromDate(contractGeneratedDate);
      }
      const [day, month, year] = formatDate(contractGeneratedDate, format='DDMMMMYYYY', sep=null);
      
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
        Camp               : userData.Camp,
        Orudnyi_FullName   : userData.Orudnyi_FullName
      };
      replaceBodyPlaceholders(body, replacements);
      doc.saveAndClose();

      const url = doc.getUrl();
      SOURCE_SHEET.getRange(userData.rowIdx, columnIndexes.contract + 1).setFormula(`=HYPERLINK("${url}";"Договір")`); 
      // SOURCE_SHEET.getRange(userData.rowIdx, columnIndexes.checkBox + 1).setValue(userData.DocCreatedCheckBox)
    }
    catch (error)
    {
      console.log(`ERROR:Caught error while creating document... \n\tFirstName: ${userData.FirstName}\n\tLastName: ${userData.LastName}`)
      console.log(`ERROR: ${error}`)
      erronous_idxs.push(idx);
    }

    idx++;



  }


    
  let end = Date.now();
  console.log(`Finished in ${end - start}ms`);

  return dataToCopy.filter((_, idx) => !erronous_idxs.includes(idx))

  }

function createNewGoogleDocs(dataToCopy, UI) { 
  // if (dataToCopy.length>30)
  // {
  //   console.log(`INFO:Performing asynchronous calls.`);
  //   const batchAmount = Math.ceil(dataToCopy.length / MAX_AMT_TRIGGERS);  
  //   let amtOfTriggersToCreate = 0;
  //   // UI.alert('Операція успішна ✅', '',  UI.ButtonSet.OK)

  //   for (let i = 0; i < dataToCopy.length; i += batchAmount) {
  //     const batch = dataToCopy.slice(i, i + batchAmount);
  //     Async.call('copyData', batch);
  //     amtOfTriggersToCreate++;

  //   }
  //   console.log(`INFO:Created ${amtOfTriggersToCreate} triggers.`);

  // }
  // else
  // {
  //   console.log(`INFO:Performing synchronous call with ${dataToCopy.length} records.`);
  //   console.log(`INFO:Created 0 triggers.`);
  //   // UI.alert('Операція успішна ✅', '',  UI.ButtonSet.OK);
    
  //   copyData(dataToCopy);


  // }

  console.log(`INFO:Performing synchronous call with ${dataToCopy.length} records.`);
  // console.log(`INFO:Created 0 triggers.`);
  // UI.alert('Операція успішна ✅', '',  UI.ButtonSet.OK);
  
  return copyData(dataToCopy);
}













