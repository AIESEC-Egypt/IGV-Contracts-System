const referenceSheet =
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference");
const referenceSheetData = referenceSheet
  .getRange(1, 1, referenceSheet.getLastRow(), referenceSheet.getLastColumn())
  .getValues();

// LC Codes from Expa padded with 0 on the left
const lcMap = {
  "Ain Shams University": "1789",
  "AAST in Cairo": "1322",
  GUC: "2570",
  Damietta: "0109",
  "AAST in Alexandria": "1788",
  Suez: "0015",
  Mansoura: "0171",
  Helwan: "2124",
  MIU: "2125",
  Menofia: "1727",
  "Kafr Sheikh": "2524",
  "6th October University": "2820",
  AUC: "1489",
  "Cairo University": "1064",
  Tanta: "1725",
  Zagazig: "1114",
  MUST: "2818",
  MSA: "2817",
  Alexandria: "899",
  "Beni Suef": "2126",
  "Luxor & Aswan": "2114",
};

const ecbSheetsMap = {};

const lcsFolders = {};

const mcvpIGV = "m.alswaf@aiesec.org.eg";

const dateFormat = "yyyyddMM";
