// --------------- Indepent Variables ---------------
const LOG = {FAT:"[0] FATAL",ERR:"[1] ERROR",WAR:"[2] WARNING",INF:"[3] INFO"};
const libCodes = ["BRDX","BRET","CLRF","GRNO","LYON","MRSL","MOPL","NNCY","NANT","NRMD","PBLV","MLVL","PVDS","PVSM","STET","STRB","TOUL","VRSL","LILL","PAYV","PAYM","IUAR","MALQ","IMVT","PLVT"];
const headers = {
  "branchcode":{index:0,forms:["branchcode","code_ecole"]},
  "biblionumber":{index:1,forms:["biblionumber"]},
  "no_abonnement_koha":{index:2,forms:["no_abonnement_koha","no_abonnement_koha"]},
  "numero":{index:3,forms:["numero"]},
  "date_parution":{index:4,forms:["date_parution"]},
  "date_reception":{index:5,forms:["date_reception"]},
  "statut_arrive_manquant":{index:6,forms:["statut_arrive_manquant"]},
  "Localisation":{index:7,forms:["localisation"]},
  "Cote":{index:8,forms:["cote"]},
}
const mandatoryCols = ["branchcode","biblionumber","numero","date_parution","date_reception","statut_arrive_manquant"];
const reportData = [];

// --------------- Indepent functions ---------------
function normalizeOutput(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim()
}

// --------------- Variable dependant functions ---------------
function appendToReport(gravity,sheetName,index,type,message) {
  reportData.push([gravity,sheetName,index,type,message]);
}

// ------------------------------------------------------------
//                      MAIN 
// ------------------------------------------------------------
// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
  // --------------- Create all worksheets ---------------
  // This is mostly for tests but I guess it's also a sefaty ?
  ["Erreurs","Fusionnee"].forEach((name) => {
    if ((workbook.getWorksheet(name))) {
      workbook.getWorksheet(name).delete();
    }
  })
  const SHEETS = {
    REPORT:{sheet:workbook.addWorksheet("Erreurs"),rowIndex:0},
    MERGED:{sheet:workbook.addWorksheet("Fusionnee"),rowIndex:0},
  }

  // --------------- Define workbook functions ---------------
  function addLines(sheet, data) {
    // sheet is the object, not the actual Worksheet element, data is an array of array
    sheet.sheet.getRangeByIndexes(sheet.rowIndex, 0, data.length, data[0].length).setValues(data);
    sheet.rowIndex += data.length;
  }
  
  // --------------- Set up worksheets ---------------
  // Add headers on relevent sheets
  // don't add to merged yet, it's less of a pain to add it at the end
  // addLines(SHEETS.MERGED, [["branchcode","biblionumber","no_abonnement_koha","numero","date_parution","date_reception","statut_arrive_manquant","Localisation","Cote","sheetName"]]);
  appendToReport("Gravité","Feuille","Numero de ligne","Type","Message");

  // --------------- Prepare loops ---------------
  const data = []; // Will hold all the data until added to the sheet
  const originalSheetCount = workbook.getWorksheets().length;
  let processedSheets = 0; // Tracks number of sheets actually processed

  // Loop through each sheet
  workbook.getWorksheets().forEach((sheet) => {
    let sheetName = sheet.getName();
    // If sheet name is not in the expected list, ignore it
    if (!(libCodes.includes(sheetName))) {
      return;
    }
    processedSheets++;

    let values = sheet.getUsedRange().getTexts();

    // Get column index in this sheet
    let thisSheetColIndex = {};
    let thisSheetHeaders = values[0].map((txt) => normalizeOutput(String(txt)));
    let missingCols = {mandatory:[],other:[]};
    Object.keys(headers).forEach((key) => {
      thisSheetColIndex[key] = -1;
      thisSheetHeaders.forEach((txt, index) => {
        if (headers[key].forms.includes(txt)) {
          thisSheetColIndex[key] = index;
          return
        }
      })
      // Check if the column is missing
      if (thisSheetColIndex[key] === -1) {
        if (mandatoryCols.includes(key)) {
          missingCols.mandatory.push(key);
        } else {
          missingCols.other.push(key);
        }
      }
    })
    // Output fatal errors (warning if not mandatory) for missing columns
    if (missingCols.mandatory.length > 0) {
      appendToReport(LOG.FAT,sheetName,-1,"Colonnes vitales manquantes","Colonnes : " + missingCols.mandatory.join(", "));
    }
    if (missingCols.other.length > 0) {
      appendToReport(LOG.WAR,sheetName,-1,"Colonnes facultatives manquantes","Colonnes : " + missingCols.other.join(", "));
    }

    // Loop through data excluding headers
    values.slice(1).forEach((row) => {
      rowData = {
        "branchcode":"",
        "biblionumber":"",
        "no_abonnement_koha":"",
        "numero":"",
        "date_parution":"",
        "date_reception":"",
        "statut_arrive_manquant":"",
        "Localisation":"",
        "Cote":"",
        "sheetName":sheetName
      }
      Object.keys(headers).forEach((key) => {
        if (thisSheetColIndex[key] !== -1) {
          rowData[key] = String(row[thisSheetColIndex[key]]);
        }
      })

      // -------- Push the row to the data container --------
      data.push(
        [
          rowData["branchcode"],
          rowData["biblionumber"],
          rowData["no_abonnement_koha"],
          rowData["numero"],
          rowData["date_parution"],
          rowData["date_reception"],
          rowData["statut_arrive_manquant"],
          rowData["Localisation"],
          rowData["Cote"],
          rowData["sheetName"]
        ]
      )
    })
  });

  // --------------- Report sheet processing ---------------
  // Check if sheets logic is fine. Should be originl count - 2 sheets because of "Modèle" & "Etat dépouillement"
  if ((originalSheetCount-4) == processedSheets) {
    appendToReport(LOG.INF, "Fusionnee",-9,"Nombre d'écoles traitées","Total : " + processedSheets)
  } else {
    appendToReport(LOG.ERR, "Fusionnee",-9,"Nombre d'écoles traitées","Total : " + processedSheets + "(nombre attendu : " + (originalSheetCount-4) + ")")
  }

  // --------------- Add data to the worksheets ---------------
  addLines(SHEETS.REPORT, reportData);
  SHEETS.MERGED.sheet.getRange("A:I").setNumberFormatLocal("@");
  addLines(SHEETS.MERGED, data);
}
