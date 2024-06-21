function generateIGVContract() {
  // the tab of the contract
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IGV Contracts");
  const sheetData = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const rowIndex = sheet.getLastRow();

  const lcName = sheet.getRange(rowIndex, 5).getValue();
  const contractType = sheet.getRange(rowIndex, 6).getValue();
  if (contractType == "Reallocation Agreement") {
    var partner_name = sheet.getRange(rowIndex, 57).getValue();
  } else {
    var partner_name = sheet.getRange(rowIndex, 7).getValue();
  }

  // the drive folder that the script will save the contracts at
  const folder = DriveApp.getFolderById(`1mbXm8uBzCFtuWUPe_AILlRYmcGzjVFnT`);

  // creating the new file
  const contractIDIndex = referenceSheet
    .createTextFinder(contractType)
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getRow())[0];
  const contractID = referenceSheet.getRange(contractIDIndex, 2).getValue();
  const template = DriveApp.getFileById(contractID);
  const name = `${contractType} - ${partner_name}`;
  const newFile = template.makeCopy(name, folder);
  console.log(newFile.getUrl());

  // Google Docs
  const doc = DocumentApp.openById(newFile.getId());
  const docBody = doc.getBody();

  var sendCol = sheet
    .createTextFinder("Email Sent?")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  var aiesecerName = sheet
    .createTextFinder("AIESECer Name")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  var aiesecerName = sheet.getRange(rowIndex, aiesecerName).getValue();

  if (sheet.getRange(rowIndex, sendCol).getValue() == true) return;

  if (contractType == "Reallocation Agreement") {
    const pairs = referenceSheet
      .createTextFinder(`Reallocation`)
      .matchEntireCell(true)
      .findAll()
      .map((x) => [x.getRow(), x.getColumn()]);
    var emails = [];
    for (let i = 0; i < pairs.length; i++) {
      console.log(i);
      var colIndex = sheet
        .createTextFinder(referenceSheetData[pairs[i][0] - 1][pairs[i][1]])
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getColumn());
      var replaced = referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1];
      var value = sheetData[rowIndex - 1][colIndex - 1];

      if (replaced.toString().includes("sd}}")) {
        if (value == undefined || value == "") {
          value = "";
          docBody.replaceText(replaced, value);
          continue;
        }
        var value = Utilities.formatDate(
          new Date(value),
          "GMT+3",
          "dd/MM/yyyy"
        );
      }
      if (replaced.toString().includes("open}}")) {
        if (value == undefined || value == "") {
          value = "";
          docBody.replaceText(replaced, value);
          continue;
        }
      }
      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "Reference Code"
      ) {
        var date = Utilities.formatDate(new Date(), "GMT+3", dateFormat);
        var value =
          convertTexttoNumber(partner_name) +
          date +
          Math.floor(Math.random() * 100000 + 1);
        sheet.getRange(rowIndex, colIndex).setValue(value);
      }

      if (referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "LC Name") {
        var value = sheet.getRange(rowIndex, 5).getValue();
      }
      if (referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "GV") {
        var value = "GV";
      }

      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "AIESECer E-mail"
      ) {
        emails.push(sheet.getRange(rowIndex, 3).getValue());
        continue;
      }

      docBody.replaceText(replaced, value);
    }
    doc.saveAndClose();
    var docblob = doc.getAs("application/pdf");
    docblob.setName(doc.getName() + ".pdf");
    var file = DriveApp.createFile(docblob);
    var fileId = file.getId();
    const toFolder = DriveApp.getFolderById(`${lcsFolders[`${lcName}`]}`);

    moveFileId(fileId, toFolder.getId());

    MailApp.sendEmail({
      to: `${emails.join(",")}`,
      subject: `${name} Contract`,
      cc: mcvpIGV,
      body: `Hello ${aiesecerName},\nGreeting from AIESEC in Egypt.\n\nYou can find your copy of the contract that should be signed in the next few days with your partner.\n\nThank you.\n\nPlease, dont' reply. This is an automated mail.`,
      attachments: [newFile.getAs(MimeType.PDF)],
    });

    sheet
      .getRange(rowIndex, sendCol)
      .setValue(true)
      .setBackground("green")
      .setFontColor("white");
    return;
  }

  const pairs = referenceSheet
    .createTextFinder(`IGV`)
    .matchEntireCell(true)
    .findAll()
    .map((x) => [x.getRow(), x.getColumn()]);
  var emails = [];
  for (let i = 0; i < pairs.length; i++) {
    console.log(i);
    var colIndex = sheet
      .createTextFinder(referenceSheetData[pairs[i][0] - 1][pairs[i][1]])
      .matchEntireCell(true)
      .findAll()
      .map((x) => x.getColumn());
    var replaced = referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1];
    var value = sheetData[rowIndex - 1][colIndex - 1];

    if (replaced.toString().includes("sd}}")) {
      if (value == undefined || value == "") {
        value = "";
        docBody.replaceText(replaced, value);
        continue;
      }
      var value = Utilities.formatDate(new Date(value), "GMT+3", "dd/MM/yyyy");
    }
    if (referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "Reference Code") {
      var indices = referenceSheet
        .createTextFinder(`${lc}`)
        .matchEntireCell(true)
        .findAll()
        .map((x) => [x.getRow(), x.getColumn() + 1]);
      var lcCode = lcMap[`${lc}`];
      var date = Utilities.formatDate(new Date(), "GMT+3", dateFormat);
      var value = lcCode + convertTexttoNumber(partner_name) + date;
      sheet.getRange(rowIndex, colIndex).setValue(value);
    }
    if (referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "LC Name") {
      var lc = sheetData[rowIndex - 1][colIndex - 1];
      continue;
    }

    if (referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "AIESECer E-mail") {
      emails.push(
        sheetData[rowIndex - 1][
          parseInt(referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1]) - 1
        ]
      );
      continue;
    }
    if (value == undefined) value = "";
    if (value == "Other") {
      let colIndex = sheet
        .createTextFinder(
          `If "other" is chosen above, write the name of the industry.`
        )
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getColumn());
      value = sheetData[rowIndex - 1][colIndex - 1];
    }
    docBody.replaceText(replaced, value);
  }
  doc.saveAndClose();
  var docblob = doc.getAs("application/pdf");
  docblob.setName(doc.getName() + ".pdf");
  var file = DriveApp.createFile(docblob);
  var fileId = file.getId();
  const toFolder = DriveApp.getFolderById(`${lcsFolders[`${lcName}`]}`);

  moveFileId(fileId, toFolder.getId());

  MailApp.sendEmail({
    to: `${emails.join(",")}`,
    subject: `${name} Contract`,
    cc: mcvpIGV,
    body: `Hello ${aiesecerName},\nGreeting from AIESEC in Egypt.\n\nYou can find your copy of the contract that should be signed in the next few days with your partner.\n\nThank you.\n\nPlease, dont' reply. This is an automated mail.`,
    attachments: [newFile.getAs(MimeType.PDF)],
  });

  sheet
    .getRange(rowIndex, sendCol)
    .setValue(true)
    .setBackground("green")
    .setFontColor("white");
}

function convertTexttoNumber(partnerName) {
  var column = "",
    length = partnerName.length;
  for (var i = 0; i < length; i++) {
    column += partnerName.codePointAt(i);
  }

  return column.substring(0, 10);
}

function moveFileId(fileId, toFolderId) {
  var file = DriveApp.getFileById(fileId);
  var source_folder = DriveApp.getFileById(fileId).getParents().next();
  var folder = DriveApp.getFolderById(toFolderId);
  folder.addFile(file);
  source_folder.removeFile(file);
}
