function addLines(sheet, currentRow, data) {
  // data is an array of array
  // return the new current row
  sheet.getRangeByIndexes(currentRow, 0, data.length, data[0].length).setValues(data);
  return currentRow + data.length;
}

// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
  let libCodes = ["BRDX","BRET","CLRF","GRNO","LYON","MRSL","MOPL","NNCY","NANT","NRMD","PBLV","MLVL","PVDS","PVSM","STET","STRB","TOUL","VRSL","LILL","PAYV","PAYM","IUAR","MALQ","IMVT","PLVT"]
  let mergedHeader = ["branchcode","biblionumber","no_abonnement_koha","numero","date_parution","date_reception","statut_arrive_manquant","Localisation","Cote","SheetName"];
  let originalSheetCount = workbook.getWorksheets().length; // Store now the original sheet count
  // Create a new worksheet for the report & add the header
  let reportSheet = workbook.addWorksheet("Report");
  let reportCurrentRow = 0; // This will track the line in report sheet
  function appendToReport(gravity,sheetName,index,type,message) {
    reportCurrentRow = addLines(reportSheet, reportCurrentRow, [[gravity,sheetName,index+1,type,message]]);
  }
  appendToReport("Gravité","Feuille","Numero de ligne","Type","Message");

  // --------------- Merge worksheets ---------------
  // Original marging worksheets script created by B.Perez 20241023, wiht clean up and arragement for this project
  // https://github.com/bastienperez/office-scripts-excel/blob/e7ee8fbe82f5d9b03ce8b83a9930b2785631e712/concatenate-worksheets-into-one/concatenate-worksheets-into-one.osts
  
  // Create a new worksheet for the combined data & add the header
  let mergedSheet = workbook.addWorksheet("Fusionnee");
  let mergedCurrentRow = 0; // This will track the line in combined sheet
  let processedSheets = 0; // Tracks number of sheets actually processed
  mergedCurrentRow = addLines(mergedSheet, mergedCurrentRow, [mergedHeader]);

  // Loop through each sheet
  workbook.getWorksheets().forEach((sheet) => {
    // If sheet name is not in the expected list, leave
    if (!(libCodes.includes(sheet.getName()))) {
      return;
    }
    processedSheets++;

    // Get the data from the used range, excluding the header
    let usedRange = sheet.getUsedRange();
    let data = usedRange.getValues().slice(1);
    // Add sheet name as the last column for each row
    let dataWithSheetName = data.map(row => {
      row.push(sheet.getName());
      return row;
    });
    // Add data to the marged sheet
    mergedCurrentRow = addLines(mergedSheet, mergedCurrentRow, dataWithSheetName);
  });
  mergedSheet.getUsedRange().setNumberFormatLocal("@");
  mergedSheet.getUsedRange().setNumberFormat([["@"]]);

  // --------------- Report sheet processing ---------------
  // Check if sheets logic is fine. Should be originl count - 2 sheets because of "Modèle" & "Etat dépouillement"
  if ((originalSheetCount-2) == processedSheets) {
    appendToReport("INFO", "Fusionnee",-2,"Nombre d'écoles traitées",processedSheets)
  } else {
    appendToReport("ERROR", "Fusionnee",-2,"Nombre d'écoles traitées",processedSheets + "(nombre attendu : " + (originalSheetCount-2) + ")")
  }

  // --------------- Check data ---------------
  let mergedData = mergedSheet.getUsedRange().getValues();
  function getColIndex(name){
    // Return the column index of a property
    return mergedHeader.indexOf(name)
  }

  // Check if there's a main biblionumber
  let mainBibnb = null;
  let bibnbCount = {};
  for (let row = 1; row < mergedData.length; row++) {
    let key = String(mergedData[row][getColIndex("biblionumber")]).trim();
    // If does not exist yet, add to the object
    if (!(key in bibnbCount)) {
      bibnbCount[key] = 0;
    }
    bibnbCount[key]++;
  }
  // For each bibnb, check if one is in more than 80% of lines
  for (key in bibnbCount) {
    if (bibnbCount[key]>mergedData.length*0.8) {
      mainBibnb = key;
      break;
    }
  }

  // Start at 1 to skip the headers
  for (let row = 1; row < mergedData.length; row++) {
    let currentRow = mergedData[row];
    let currentColName = null; // Temp just to avoid writing a billion times the same string
    // -------- Funciton definition --------
    function getData(name) {
      // Return the value of a column for this row
      return currentRow[getColIndex(name)]
    }
    let sheetName = getData("SheetName");
    function fix(colName, value) {
      // Changes the cell inside the merged data
      mergedSheet.getCell(row, getColIndex(colName)).setValue(value); // Column C
    }
    function reportThis(gravity,type,message) {
      appendToReport(gravity,sheetName,row,type,message)
    }

    // -------- Force biblionumber if 80% has the same --------
    currentColName = "biblionumber";
    if ((getData(currentColName) != mainBibnb) && (mainBibnb !== null)) {
      reportThis("WARNING","Biblionumber différent du biblionumber principal","Ancienne valeur : " + getData(currentColName));
      fix(currentColName, mainBibnb);
    }

    // -------- Force branchcode to sheetname if different --------
    currentColName = "branchcode";
    if (getData(currentColName) !== sheetName) {
      reportThis("WARNING","Branchcode différent du nom de la feuille","Branchcode : " + getData(currentColName));
      fix(currentColName, sheetName);
    }

    // -------- Status & locaiton --------
    // Normalize (remove diacritic, trim, lowercase)
    currentColName = "statut_arrive_manquant";
    let newStatus = getData(currentColName).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    // Check if value is wrong
    if (!(["manquant","arrive","arrivee"].includes(newStatus))) {
      reportThis("ERROR","Statut arrivé manquant erronné",getData(currentColName))
    }
    // IF "manquant", force location to be empty
    if ((newStatus === "manquant") && (getData("Localisation") == "")) {
      reportThis("INFO","Manquant, suppression de sa localisation","Ancienne valeur : " + getData("Localisation"));
      fix("Localisation", "");
    }
    // Push the normlization after all checks so we keep original value
    // Actually I the data I'm reading might not be the one in cells so this might chage absolutely nothing lol hihi
    fix(currentColName, newStatus);


    // TO DO
    // Both dates : if date can be interpreted, set it to string
    // --Ask first-- force no_abonnement if different in a school
    // Both date, can't be interpreted
    // Missing data in numero, branchcode, biblionumber, date_parution, date_reception, stautt_arrive_anquant
  }
}