/**
 * Homepage builder function
 */
function onDriveHomePageOpen() {
  return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle('Welcome to TITUS tags convertor')
    .setImageStyle(CardService.ImageStyle.SQUARE).setImageUrl(
      'https://res.cloudinary.com/crunchbase-production/image/upload/c_lpad,f_auto,q_auto:eco,dpr_1/v1464254968/hag3n5nntrfuhjbfoz89.png'
    )).setHeader(CardService.newCardHeader().setTitle('Select the file(s) you want\n to convert and tag')
      .setImageStyle(CardService.ImageStyle.SQUARE).setImageUrl(
        'https://icones.pro/wp-content/uploads/2021/06/icone-fleche-gauche-noir.png')).build();
}

/**
 * Function triggered when an item is selected
 * Display the card with selected items and button to convert
 * @param {array} e
 */
function onDriveItemsSelected(e) {
  let selectedItems = e.drive.selectedItems;
  let filesToConvertObj = {};
  let textToDisplay = "";
  // Reading every selected file and filling filesToConvertObj
  // this obj is then sent as a parameter when the button Apply labels is pressed
  for (let i = 0; i < selectedItems.length; i++) {
    let item = selectedItems[i];
    let title = item.title;
    filesToConvertObj[String(item.id)] = String(title) + "||" + String(item.mimeType);
    textToDisplay = textToDisplay + "\n" + title;
  }
  let card = CardService.newCardBuilder().addSection(CardService.newCardSection().addWidget(CardService
    .newTextParagraph().setText('File(s) selected:' + textToDisplay)))
  let switchConvert = CardService.newSwitch().setFieldName("convertToGoogle").setValue("true")
  card.addSection(CardService.newCardSection().addWidget(CardService.newDecoratedText().setText("Convert to Google")
    .setWrapText(true).setSwitchControl(switchConvert)))
  card.addSection(CardService.newCardSection().addWidget(CardService.newTextButton().setText("Apply labels")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(CardService.newAction()
      .setFunctionName('convertTheseFiles').setLoadIndicator(CardService.LoadIndicator.SPINNER).setParameters(
        filesToConvertObj))))
  return card.build();
}

/**
 * Function launched when button clicked
 * Convert the file and apply Google tags
 * @param {array} filesToConvertObj
 * @return {Card}
 */
function convertTheseFiles(filesToConvertObj) {
  const resultTableSS = SpreadsheetApp.openById(MATCHINGTABLESSID);
  const scanSheet = resultTableSS.getSheetByName("Matching Table");
  const scanSheetData = scanSheet.getDataRange().getValues();
  const scanSheetDataHeaders = scanSheetData[0];
  const titusKeys = scanSheetData.map(x => x[scanSheetData[0].indexOf("TITUS Key")]);
  let taggedFiles = [];
  let unreadibleXmlFileNames = [];
  let nonRecognizedFileNames = "";
  let nonRecognizedFiles = {};
  for (let id in filesToConvertObj.parameters) {
    let file = DriveApp.getFileById(id);
    let blob = file.getBlob().setContentType("application/zip");
    if (Utilities.unzip(blob).length > 0) {
      Utilities.unzip(blob).filter(z => z.getName().indexOf('docProps/custom.xml') != -1).forEach(z => {
        let xmlText = z.getDataAsString()
        let xmlKey = readTagFromXML(xmlText);
        // Check if the XML key ends with /0
        // Word and excel documents with the same tag will have the same XML key but with an extra /0 at the end for Excel files
        console.log("xmlKey", xmlKey);
        if (titusKeys.indexOf(xmlKey) > -1 && scanSheetData[titusKeys.indexOf(xmlKey)][scanSheetDataHeaders.indexOf(
          "Corresponding labels key ")] != "") {
          let taggedFile = applyLabelsAndConvertIfNecessary(file, scanSheetData[titusKeys.indexOf(xmlKey)][
            scanSheetDataHeaders.indexOf("Corresponding labels key ")], filesToConvertObj.formInput.convertToGoogle);
          taggedFiles.push(taggedFile);
        }
        else {
          // XML read but key not recognised
          nonRecognizedFiles[file.getId()] = file.getName();
          nonRecognizedFileNames = nonRecognizedFileNames + "\n" + file.getName();
        }
      })
    }
    else {
      // Not able to read the XML
      unreadibleXmlFileNames.push(file.getName())
    }
  }
  // Case 1: files converted
  if (Object.keys(nonRecognizedFiles).length === 0 && unreadibleXmlFileNames.length == 0) {
    let card = CardService.newCardBuilder().addSection(CardService.newCardSection().setHeader(
      'Process done') // optional
      .addWidget(CardService.newTextParagraph().setText('Your files have been tagged!')))
    let navigation = CardService.newNavigation().pushCard(card.build());
    return CardService.newActionResponseBuilder().setNavigation(navigation).build();
  }
  // Case 2: only not recognized XMLs: assess button
  else if (Object.keys(nonRecognizedFiles).length > 0 && unreadibleXmlFileNames.length == 0) {
    if (filesToConvertObj.formInputs.convertToGoogle) {
      nonRecognizedFiles["convertToGoogle"] = filesToConvertObj.formInputs.convertToGoogle[0];
    } else {
      nonRecognizedFiles["convertToGoogle"] = "false";
    }
    let card = CardService.newCardBuilder().addSection(CardService.newCardSection().setHeader(
      'Process done') // optional
      .addWidget(CardService.newTextParagraph().setText('Some tags were not recognized: ' +
        nonRecognizedFileNames))).addSection(CardService.newCardSection().addWidget(CardService
          .newTextButton().setText("Assess").setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('generateCards').setParameters(
            nonRecognizedFiles))))
    let navigation = CardService.newNavigation().pushCard(card.build());
    return CardService.newActionResponseBuilder().setNavigation(navigation).build();
  }
  // Case 3: not recognized XMLs and not readable XMLs
  else if (Object.keys(nonRecognizedFiles).length > 0 && unreadibleXmlFileNames.length > 0) {
    if (filesToConvertObj.formInputs.convertToGoogle) {
      nonRecognizedFiles["convertToGoogle"] = filesToConvertObj.formInputs.convertToGoogle[0];
    } else {
      nonRecognizedFiles["convertToGoogle"] = "false";
    }
    let card = CardService.newCardBuilder().addSection(CardService.newCardSection().setHeader(
      'Process done') // optional
      .addWidget(CardService.newTextParagraph().setText('Some tags were not recognized: ' +
        nonRecognizedFileNames))).addSection(CardService.newCardSection().addWidget(CardService
          .newTextButton().setText("Assess").setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('generateCards').setParameters(
            nonRecognizedFiles)))).addSection(CardService.newCardSection().setHeader('Warning') // optional
              .addWidget(CardService.newTextParagraph().setText(
                'We were not able to read the tags for the following files: ' + unreadibleXmlFileNames)))
    let navigation = CardService.newNavigation().pushCard(card.build());
    return CardService.newActionResponseBuilder().setNavigation(navigation).build();
  }
  // Case 4: only not readible XMLs
  else if (Object.keys(nonRecognizedFiles).length === 0 && unreadibleXmlFileNames.length > 0) {
    let card = CardService.newCardBuilder().addSection(CardService.newCardSection().setHeader(
      'Warning') // optional
      .addWidget(CardService.newTextParagraph().setText(
        'We were not able to read the tags for the following files: ' + unreadibleXmlFileNames)))
    let navigation = CardService.newNavigation().pushCard(card.build());
    return
    CardService.newActionResponseBuilder().setNavigation(navigation).build();
  }
}
/**
 * Generate N cards for each file to assess
 * displaying the form to specify the corresponding Google labels
 * param {array} e
 * @return {CardService.action}
 */
function generateCards(e) {
  let cardSection1Divider1 = CardService.newDivider();
  let navigation = CardService.newNavigation();
  let parameters = {};
  let endOfProcessCard = CardService.newCardBuilder().addSection(CardService.newCardSection().setHeader(
    'Process done') // optional
    .addWidget(CardService.newTextParagraph().setText('The labels has been applied and assessed!')))
  navigation.pushCard(endOfProcessCard.build())
  for (let id in e.parameters) {
    if (id != "convertToGoogle") {
      parameters = {};
      parameters["convertToGoogle"] = e.parameters["convertToGoogle"];
      parameters["id"] = id;
      parameters["name"] = e.parameters[id];
      let card = CardService.newCardBuilder().setName("Assess tag for " + id).setHeader(CardService
        .newCardHeader().setTitle('Please select the labels').setSubtitle(e.parameters[id])).addSection(
          CardService.newCardSection()
            // -- PRIVACY LABEL -- 
            .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.DROPDOWN)
              .setTitle("Privacy").setFieldName("Privacy").addItem("Basic Personal Data",
                "Basic Personal Data", true).addItem("Personal Data", "Personal Data", false).addItem(
                  "Sensitive Personal Data", "Sensitive Personal Data", false)).addWidget(cardSection1Divider1)
            // -- Export Control --
            .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.DROPDOWN)
              .setTitle("Export Control").setFieldName("Export Control").addItem("Not Technical",
                "Not Technical", true).addItem("Not In The Export Control List",
                  "Not In The Export Control List", true).addItem("Export Control Dual Use",
                    "Export Control Dual Use", false)).addWidget(cardSection1Divider1)
            // -- National Security --
            .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.DROPDOWN)
              .setTitle("National Security").setFieldName("National Security").addItem(
                "Not National Security", "Not National Security", false)).addWidget(cardSection1Divider1)
            // -- Company classification --
            .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.DROPDOWN)
              .setTitle("Company classification").setFieldName("Company classification").addItem(
                "Not Applicable", "Not Applicable", true).addItem("CLIENT Amber", "CLIENT Amber", false)
              .addItem("CLIENT Red", "CLIENT Red", false)).addWidget(cardSection1Divider1)
            // -- Business or Private data --
            .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.DROPDOWN)
              .setTitle("Business or Private data").setFieldName("Business or Private data").addItem(
                "Private Data", "Private Data", true).addItem("Business Data", "Business Data", false)))
        .addSection(CardService.newCardSection().addWidget(CardService.newTextButton().setText("Assess")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(CardService.newAction()
            .setFunctionName('goToNextCard').setParameters(parameters))))
      navigation.pushCard(card.build());
    }
  }
  return CardService.newActionResponseBuilder().setNavigation(navigation).build();
}

/**
 * Function to navigate through cards and assess files
 * @param {array} e
 * @return {CardService.Card}
 */
function
  goToNextCard(e) {
  const correspondingLabelsKey = addThisTagToMatchingTable(e);
  applyLabelsAndConvertIfNecessary(DriveApp.getFileById(e.parameters.id), correspondingLabelsKey, e
    .parameters.convertToGoogle);
  let navigation = CardService.newNavigation().popCard();
  return CardService.newActionResponseBuilder().setNavigation(navigation).build();
}

/**
 * Get file, extract XML and write in matching table the corresponding labels
 * @param {array} e
 * @return {string} correspondingLabelsKey
 */
function
  addThisTagToMatchingTable(e) {
  file = e.parameters;
  const resultTableSS = SpreadsheetApp.openById(MATCHINGTABLESSID);
  const scanSheet = resultTableSS.getSheetByName("Matching Table");
  const newRow = [];
  newRow.push(file.name);
  let blob = DriveApp.getFileById(file.id).getBlob().setContentType("application/zip");
  Utilities.unzip(blob).filter(z => z.getName().indexOf('docProps/custom.xml') != -1).forEach(z => {
    newRow.push(z.getDataAsString());
    newRow.push(readTagFromXML(z.getDataAsString()));
  })
  let correspondingLabelsKey = "";
  let googleLabelIDsMap = getGoogleLabelIDsMap(resultTableSS);
  let formInputs = e.formInputs;
  for (let input in formInputs) {
    correspondingLabelsKey = correspondingLabelsKey + "//" + googleLabelIDsMap[input] + "/" +
      googleLabelIDsMap[formInputs[input]];
  }
  newRow.push(correspondingLabelsKey);
  scanSheet.appendRow(newRow);
  return correspondingLabelsKey;
}

/**
 * Extract key from XML file
 * Format: /X/X/X/X/X
 * @param {string} xmlText
 * @return {string} key
 */
function readTagFromXML(xmlText) {
  let pattern = new
    RegExp("<vt:lpwstr>(.*?)</vt:lpwstr>", "g")
  titusKeys = [];
  let exctract = xmlText.match(pattern);
  let key = "";
  for (let j = 1; j < exctract.length; j++) {
    if (exctract[j].replaceAll("<vt:lpwstr>", "").replaceAll("</vt:lpwstr>", "") != "") {
      key = key + "/" + exctract[j].replaceAll("<vt:lpwstr>", "").replaceAll("</vt:lpwstr>", "");
    }
  }
  return key;
}

/**
 * Gets data from sheet Google labels IDs and build a map out of it
 * @param {SpreadsheetApp.Spreadsheet} resultTableSS
 * @return {array} labelIDsMap
 */
function
  getGoogleLabelIDsMap(resultTableSS) {
  const labelSheet = resultTableSS.getSheetByName("Google labels IDs");
  const labelSheetData = labelSheet.getDataRange().getValues();
  let labelIDsMap = {};
  for (let i = 0; i < labelSheetData.length; i++) {
    labelIDsMap[labelSheetData[i][0]] = labelSheetData[i][1];
  }
  return labelIDsMap;
}

/**
 * Main function to apply labels, convert to Google format if radio button checked
 * @param {DriveApp.File} file
 * @param {string} googleTagKey
 * @param {string} convertToGoogle
 * @return {array} taggedFile
 */
function
  applyLabelsAndConvertIfNecessary(file, googleTagKey, convertToGoogle) {
  // If so, convert file to Google and apply labels
  if (convertToGoogle == "true" && String(file.getId().length < 44)) {
    let xBlob = file.getBlob();
    let name = file.getName();
    let parents = file.getParents();
    let folderParentId;
    while (parents.hasNext()) {
      var folder = parents.next();
      folderParentId = folder.getId();
    }
    let newFileInfo = {
      title: name,
      parents: [{
        id: folderParentId
      }]
    };
    newFile = Drive.Files.insert(newFileInfo, xBlob, {
      convert: true
    });
    fileToTagId = newFile.getId();
    // Remove original
    file.setTrashed(true)
  } else {
    fileToTagId = file.getId();
  }
  // Convert Google label key to field modification object
  let googleTagKeyLabelsTab = googleTagKey.split("//");
  let fieldModificationsTab = [];
  for (let i = 1; i < googleTagKeyLabelsTab.length; i++) {
    let labelObj = {};
    let labelTab = googleTagKeyLabelsTab[i].split("/");
    labelObj["fieldId"] = String(labelTab[0]);
    labelObj["setSelectionValues"] = String(labelTab[1]);
    fieldModificationsTab.push(labelObj);
  }
  // Apply labels
  Drive.Files.modifyLabels(JSON.stringify({
    "kind": "drive#modifyLabelsRequest",
    "labelModifications": [{
      "labelId": DATACLASSIFICATIONLABELID, // CLIENT 'Data Classification' label id
      "fieldModifications": fieldModificationsTab
    }]
  }), fileToTagId)
  let taggedFile = {};
  taggedFile["id"] = file.getId();
  taggedFile["name"] = file.getName();
  return taggedFile;
}

const MATCHINGTABLESSID = '';
const DATACLASSIFICATIONLABELID = '';
