const kohaLocations = ["1er étage","1er étage réserve","2e étage","2e étage bureau doc","Accueil","Architectes","Architecture","Archives","Archives nationales","Arpège","Arts","Atelier documentaire","Audiovisuel","Bibliothèque","Bureau bas","Bureau doc","Bureau haut","Bureau interne","Cartothèque","Centre d'art","Centre de documentation","Passages","Construction","DSA","Espace Métier","Fonds ancien","Fonds courant","Fonds diapos","Fonds photo aérienne","Fonds régional","Fonds TPFE","Fonds travaux étudiants","GRECAU","Hors format","IPRAUS","Labo ARIA","Labo ARTOPOS","LIFAM","Laboratoires","Libre accès","Magasin","Magasin 1","Matériauthèque","Monographies","Niveau haut","PAVE","Paysage","Placard 1","RDC Réserve","Recherche","Réserve","Réserve 1","Réserve 2","Réserve 3","Réserve Mûrier","Revues","Rez de chaussée","Salle 1","Salle de lecture","Salle des archives/ouvrages doubles","Sciences humaines","Service informatique","Services administratifs","Territoire","Urbanisme","Usuels","Vidéothèque","Vitrine Prof","VRD","Inconnu","Atelier maquette","Fonds revues","Laboratoire de recherche en architecture (LRA)","Archives départementales","Fonds Auzelle","Fonds Huet","Fonds Huet ancien","Labo LAURE","Serveur ENSA","En ligne","Fonds BD","DPEA","Placard","Réserve de cours","Fonds Jean Aubert","Atelier Bois","Fonds ancien réserve","Revues réserve","Revues vitrine","Fonds travaux d'atelier","Laboratoire de recherche","Fonds Guerrand","Meuble à plans 1","Meuble à plans 2","Meuble à plans 3","Meuble à plans 4","Meuble à plans 5","Quarantaine","Littérature - BD","Espace Pédagogie","Master RBW","Escape game","Fonds Hervé Dupont","Écologie","Potager","Fonds Pinon","Fonds Pinon ancien","Mezzanine Vercors","Cohen","Mezzanine Chartreuse","Mezzanine Belledonne","Salle Ailefroide","Salle détente","Littérature grise"];
const libCodes = ["BRDX","BRET","CLRF","GRNO","LYON","MRSL","MOPL","NNCY","NANT","NRMD","PBLV","MLVL","PVDS","PVSM","STET","STRB","TOUL","VRSL","LILL","PAYV","PAYM","IUAR","MALQ","IMVT","PLVT"]
const mergedHeader = ["branchcode","biblionumber","no_abonnement_koha","numero","date_parution","date_reception","statut_arrive_manquant","Localisation","Cote","SheetName"];

function normalizeOutput(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim()
}

// Add an mapping normalized / original version of the locations to try and fix them
const kohaLocationsNormalized = {};
let kohaLocationsIgnored = [];
kohaLocations.forEach((location) => {
  let normalized = normalizeOutput(location);
  // Only add normalized version if it is unique
  if ((!(normalized in kohaLocationsNormalized)) && (!(kohaLocationsIgnored.includes(normalized)))) {
    kohaLocationsNormalized[normalized] = location;
  } else {
    delete kohaLocationsNormalized[normalized];
    kohaLocationsIgnored.push(normalized); // yes it there are at laest 3 forms they're duplciated here, idc
  }
})

function addLines(sheet, currentRow, data) {
  // data is an array of array
  // return the new current row
  sheet.getRangeByIndexes(currentRow, 0, data.length, data[0].length).setValues(data);
  return currentRow + data.length;
}

function toYYYYMMDD(value) {
  value = String(value).trim();
  let date = null;
  // Excel date
  if (/^\d+$/.test(value)) {
    // Copilot did the number conversion and it seems to work
    date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return date.toISOString().split("T")[0];
  // String : check valdity
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    date = new Date(value);
    // Check if it's a real calendar date
    if (!isNaN(date.getTime())) {
      return date.toISOString().split("T")[0];
    }
  }
  return null;
}

function locationIsValid(location) {
  // Returns if the locations is valid
  return kohaLocations.includes(location)
}

function getSimilarLocation(location) {
  // Returns a location if one exists with the same nromalized version
  let normalized = normalizeOutput(location);
  if (normalized in kohaLocationsNormalized) {
    return kohaLocationsNormalized[normalized];
  }
  return null;
}


// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
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
      mergedSheet.getCell(row, getColIndex(colName)).setValue(value);
    }
    function reportThis(gravity,type,message) {
      appendToReport(gravity,sheetName,row,type,message);
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

    // -------- Status & location --------
    // Normalize (remove diacritic, trim, lowercase)
    currentColName = "statut_arrive_manquant";
    let newStatus = normalizeOutput(getData(currentColName));
    // Check if value is wrong
    if (!(["manquant","arrive","arrivee"].includes(newStatus))) {
      reportThis("ERROR","Statut arrivé manquant erronné",getData(currentColName))
    }
    // IF "manquant", force location to be empty
    let location = getData("Localisation");
    if ((newStatus === "manquant") && (location != "")) {
      reportThis("INFO","Manquant, suppression de sa localisation","Ancienne valeur : " + location);
      fix("Localisation", "");
    // Else, check if location is a legal value
    } else if ((!(locationIsValid(location))) && (location != "")) {
      // If not valid, try to find a valid alternative
      let altLocation = getSimilarLocation(location);
      if (altLocation === null) {
        reportThis("ERROR","Localisation inexistante",location);
      } else {
        reportThis("WARNING","Correction de la localisation","Nouvelle valeur : " + altLocation + " (ancienne :" + location + ")");
        fix("Localisation",altLocation);
      }
    }
    // Push the normlization after all checks so we keep original value
    // Actually the data I'm reading is not the one in cells so this changes absolutely nothing lol hihi
    fix(currentColName, newStatus);

    // -------- Dates --------
    // Convert to string and check if they're real (if interpreted from a string)
    ["date_parution","date_reception"].forEach((colName) => {
      let newDate = toYYYYMMDD(getData(colName));
      if (newDate === null) {
        reportThis("ERROR","Impossible d'interpréter " + colName, "Ancienne valeur : " + getData(colName));
      // Don't psh an empty value to avoid triggering two errors on the same problem
      } else {
        fix(colName, newDate);
      }
    })
  }
  // --------------- Check missing vital data ---------------
  // We do that after transformation so we can see their results
  // We could do this in the second loop, but doing it this way makes sure we do have the finalized data
  // And tbh it's less of a apin to do hihi
  mergedData = mergedSheet.getUsedRange().getValues();
  for (let row = 1; row < mergedData.length; row++) {
    let missingData = [];
    let currentRow = mergedData[row];
    ["branchcode","biblionumber","numero","date_parution","date_reception","statut_arrive_manquant"].forEach((colName) => {
      if (currentRow[getColIndex(colName)] == "") {
        missingData.push(colName);
      }
    })
    if (missingData.length > 0) {
      appendToReport("ERROR",currentRow[getColIndex("SheetName")],row,"Informations vitales absentes","Colonnes : " + missingData.join(", "));
    }
  }
}

// TO DO
// Sort report lines by error -> warning -> INFO
// Make display a bit easier (col lengths, color stuff)
// TCD du nombres d'erreurs