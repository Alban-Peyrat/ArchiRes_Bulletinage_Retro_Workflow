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
const reversedColIndex = Object.fromEntries(Object.entries(colIndex).map(([key, value]) => [value, key]));
const mandatoryCols = ["branchcode","biblionumber","numero","date_parution","date_reception","statut_arrive_manquant"];
const mandatoryColsIndex = [];
mandatoryCols.forEach((colName) => {
  mandatoryColsIndex.push(colIndex[colName]);
})


// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
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
    // -------- Check missing vital data --------
    let missingData = [];
    mandatoryColsIndex.forEach((colIndex) => {
      if (row[colIndex] == "") {
        missingData.push(reversedColIndex[colIndex]);
      }
    })
    if (missingData.length > 0) {
      reportThis(LOG.ERR,"Informations vitales absentes","Colonnes : " + missingData.join(", "));
    }
  })

  // --------------- Add data to the worksheets ---------------
  addDataToReport();
}