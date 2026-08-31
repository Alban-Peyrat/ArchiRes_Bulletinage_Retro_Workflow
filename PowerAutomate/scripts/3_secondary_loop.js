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
const mandatoryCols = ["branchcode","biblionumber","numero","date_parution","date_reception","statut_arrive_manquant"];

// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
  // --------------- Variables ---------------
  const sheetMerged = workbook.getWorksheet("Fusionnee");
  const sheetReport = workbook.getWorksheet("Erreurs");
  const reportData = [];

// --------------- Variable dependant functions ---------------
  function addLines(sheet, data) {
    // Data is an array of array
    sheet.getRangeByIndexes(sheet.getUsedRange().getValues().length+1, 0, data.length, data[0].length).setValues(data);
  }

  sheetMerged.getUsedRange().getValues().forEach((row, index) => {
    function getData(name) {
      // Return the value of a column for this row
      return row[colIndex[name]]
    }
    function reportThis(gravity,type,message) {
      reportData.push([gravity,getData("sheetName"),index+1,type,message]);
    }

    // -------- Force biblionumber if 80% has the same --------
    let mainBibnb = workbook.getWorksheet("PowerAutomateHidden").getRange("A1").getValue();
    let currentColName = "biblionumber";
    if ((getData(currentColName) != mainBibnb) && (mainBibnb != "")) {
      reportThis(LOG.WAR,"Biblionumber différent du biblionumber principal","Ancienne valeur : " + getData(currentColName));
      row[colIndex[currentColName]] = mainBibnb;
    }

    // -------- Check missing vital data --------
    // We do that in 2nd loop so we can use bibnb edition
    let missingData = [];
    mandatoryCols.forEach((colName) => {
      if (getData(colName) == "") {
        missingData.push(colName);
      }
    })
    if (missingData.length > 0) {
      reportThis(LOG.ERR,"Informations vitales absentes","Colonnes : " + missingData.join(", "));
    }
  })

  // --------------- Add data to the worksheets ---------------
  addLines(sheetReport, reportData);
  addLines(sheetMerged, data);
}