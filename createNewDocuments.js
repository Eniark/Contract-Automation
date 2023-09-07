

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


function getSubfolderIDs(destFolder, data) {
  /* 
  Creates new subfolders for contracts in case such don't exist.
  Parameters:
  =========================================
    destFolder: folder where to create subfolders
    data: an array of records that need contracts 
  */

  let subFolders = destFolder.getFolders();
  let nameIdMapping = {};
  camps = [...new Set(data.map(el => `${el.Camp} (${el.CampStartDate} - ${el.CampEndDate})`))];
  while (subFolders.hasNext()) {
    let folder = subFolders.next();
    nameIdMapping[folder.getName()] = folder.getId(); 
  }

  console.log(`INFO:Available Folders:\n${Object.keys(nameIdMapping)}`);
  camps.map(el => {
    if (!Object.keys(nameIdMapping).includes(el)) {
      const f = destFolder.createFolder(el);
      console.log(`INFO:Created new subfolder with name: ${el}`);
      nameIdMapping[el] = f.getId();
    }

  })

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
      let subDestinationFolderID = nameIdMapping[`${userData.Camp} (${userData.CampStartDate} - ${userData.CampEndDate})`]
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
  console.log(`INFO:Created 0 triggers.`);
  // UI.alert('Операція успішна ✅', '',  UI.ButtonSet.OK);
  
  return copyData(dataToCopy);
}













