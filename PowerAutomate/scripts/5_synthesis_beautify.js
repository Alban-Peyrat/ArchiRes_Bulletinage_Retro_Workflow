// --------------- Independant Variables ---------------
const LOG = {FAT:"[0] FATAL",ERR:"[1] ERROR",WAR:"[2] WARNING",INF:"[3] INFO"};
// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
  const sheetMerged = workbook.getWorksheet("Fusionnee");
  const sheetReport = workbook.getWorksheet("Erreurs");
  const sheetSynthesis = workbook.addWorksheet("Synthese");

  // --------------- Set up worksheets position ---------------
  sheetSynthesis.setPosition(1);
  sheetReport.setPosition(2);
  sheetMerged.setPosition(3);

  // --------------- Remove sheet name col ---------------
  sheetMerged.getRangeByIndexes(0, 9, sheetMerged.getUsedRange().getRowCount(), 1).delete(ExcelScript.DeleteShiftDirection.left);

  // --------------- Report data looks ---------------
  // -------- Sort data --------
  let reportUsedRange = sheetReport.getUsedRange();
  // Ty Copilot
  reportUsedRange.getSort().apply([
      { key: 0, ascending: true }, // Gravité
      { key: 2, ascending: true }  // Index
    ],
    true, // match case 
    true // has headers
  );
  // -------- Fix col length --------
  reportUsedRange.getFormat().autofitColumns();
  // -------- Color gravity --------
  // Still Copilot stuff cleaned up
  [[LOG.ERR,"#FFC7CE"], [LOG.WAR,"#FFEB9C"]].forEach((tuple) => {
    let rule = reportUsedRange.addConditionalFormat(
      ExcelScript.ConditionalFormatType.containsText
    );
    rule.getTextComparison().setRule({
      operator: ExcelScript.ConditionalTextOperator.contains,
      text: tuple[0]
    });
    rule.getTextComparison().getFormat().getFill().setColor(tuple[1]);
  })

  // --------------- Synthesis with pivot table ---------------
  // Still Copilot with clean up
  const pivotTable = sheetSynthesis.addPivotTable(
    "Synthese_Rapport_Erreur",
    reportUsedRange,
    sheetSynthesis.getRange("A1")
  );
  // Rows
  pivotTable.addRowHierarchy(pivotTable.getHierarchy("Gravité"));
  pivotTable.addRowHierarchy(pivotTable.getHierarchy("Type"));
  // Columns
  pivotTable.addColumnHierarchy(pivotTable.getHierarchy("Feuille"));
  // Count rows
  let dataField = pivotTable.addDataHierarchy(
    pivotTable.getHierarchy("Message")
  );
  dataField.setSummarizeBy(ExcelScript.AggregationFunction.count);
  dataField.setName("Nombre de lignes");
}
