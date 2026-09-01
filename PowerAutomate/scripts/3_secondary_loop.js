// --------------- Independant Variables ---------------
const LOG = {FAT:"[0] FATAL",ERR:"[1] ERROR",WAR:"[2] WARNING",INF:"[3] INFO"};
const colIndex = {
  "branchcode":0,
  "biblionumber":1,
  "no_abonnement_koha":2,
  "numero":3,
  "date_parution":4,
  "date_reception":5,
  "statut_arrive_manquant":6,
  "Localisation":7,
  "Cote":8,
  "sheetName":9
}

// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
  // --------------- Check if this step should execute ---------------
  const mainBibnb = workbook.getWorksheet("PowerAutomateHidden").getRange("A1").getText();
  if (mainBibnb == "") {
    return
  }
  // --------------- Variables ---------------
  const sheetMerged = workbook.getWorksheet("Fusionnee");
  const data = sheetMerged.getUsedRange().getTexts(); // Will hold all the data until added to the sheet
  const reportData = [];

  // --------------- Define workbook functions ---------------
  function addDataToReport() {
    // data is an array of array
    let sheet = workbook.getWorksheet("Erreurs");
    sheet.getRangeByIndexes(sheet.getUsedRange().getRowCount(), 0, reportData.length, reportData[0].length).setValues(reportData);
  }

  data.forEach((row, index) => {
    function getData(name) {
      // Return the value of a column for this row
      return row[colIndex[name]]
    }
    let sheetName = getData("sheetName");
    function reportThis(gravity,type,message) {
      reportData.push([gravity,sheetName,index+1,type,message]);
    }
    let currentColName = null; // Temp just to avoid writing a billion times the same string
    let currentColVal = null; // Temp just to avoid writing a billion times the same string

    // -------- Force biblionumber if 80% has the same --------
    currentColName = "biblionumber";
    currentColVal = getData(currentColName);
    if (currentColVal != mainBibnb) {
      reportThis(LOG.WAR,"Biblionumber différent du biblionumber principal","Ancienne valeur : " + currentColVal);
      row[colIndex[currentColName]] = mainBibnb;
    }
  })

  // --------------- Add data to the worksheets ---------------
  addDataToReport();
  sheetMerged.getUsedRange().clear();
  sheetMerged.getRange("A:I").setNumberFormatLocal("@");
  data.unshift(["branchcode","biblionumber","no_abonnement_koha","numero","date_parution","date_reception","statut_arrive_manquant","Localisation","Cote","sheetName"]);
  sheetMerged.getRangeByIndexes(0,0,data.length,data[0].length).setValues(data);
}