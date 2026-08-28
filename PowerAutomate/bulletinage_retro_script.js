// ------------------------------------------------------------
//              Workbook free declarations
// ------------------------------------------------------------
// --------------- Indepent functions ---------------
function normalizeOutput(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim()
}

function toYYYYMMDD(value) {
  value = String(value).trim();
  let date = null;
  // Excel date
  if (/^\d+$/.test(value)) {
    // If between 1800 & 2050, assume it's a year only
    // This means that anything bewteen 1904-12-04 & 1905-08-11 is ignored
    if ((parseInt(value) > 1800) && (parseInt(value) < 2050)) {
      return value + "-01-01";
    }
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
  } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) {
    return [value.substring(6,10),value.substring(0,2),value.substring(3,5)].join("-")
  }
  return null;
}

// --------------- Variables ---------------
const kohaLocations = ["1er étage","1er étage réserve","2e étage","2e étage bureau doc","Accueil","Architectes","Architecture","Archives","Archives nationales","Arpège","Arts","Atelier documentaire","Audiovisuel","Bibliothèque","Bureau bas","Bureau doc","Bureau haut","Bureau interne","Cartothèque","Centre d'art","Centre de documentation","Passages","Construction","DSA","Espace Métier","Fonds ancien","Fonds courant","Fonds diapos","Fonds photo aérienne","Fonds régional","Fonds TPFE","Fonds travaux étudiants","GRECAU","Hors format","IPRAUS","Labo ARIA","Labo ARTOPOS","LIFAM","Laboratoires","Libre accès","Magasin","Magasin 1","Matériauthèque","Monographies","Niveau haut","PAVE","Paysage","Placard 1","RDC Réserve","Recherche","Réserve","Réserve 1","Réserve 2","Réserve 3","Réserve Mûrier","Revues","Rez de chaussée","Salle 1","Salle de lecture","Salle des archives/ouvrages doubles","Sciences humaines","Service informatique","Services administratifs","Territoire","Urbanisme","Usuels","Vidéothèque","Vitrine Prof","VRD","Inconnu","Atelier maquette","Fonds revues","Laboratoire de recherche en architecture (LRA)","Archives départementales","Fonds Auzelle","Fonds Huet","Fonds Huet ancien","Labo LAURE","Serveur ENSA","En ligne","Fonds BD","DPEA","Placard","Réserve de cours","Fonds Jean Aubert","Atelier Bois","Fonds ancien réserve","Revues réserve","Revues vitrine","Fonds travaux d'atelier","Laboratoire de recherche","Fonds Guerrand","Meuble à plans 1","Meuble à plans 2","Meuble à plans 3","Meuble à plans 4","Meuble à plans 5","Quarantaine","Littérature - BD","Espace Pédagogie","Master RBW","Escape game","Fonds Hervé Dupont","Écologie","Potager","Fonds Pinon","Fonds Pinon ancien","Mezzanine Vercors","Cohen","Mezzanine Chartreuse","Mezzanine Belledonne","Salle Ailefroide","Salle détente","Littérature grise"];
const kohaLocationsNormalized = {};
const libCodes = ["BRDX","BRET","CLRF","GRNO","LYON","MRSL","MOPL","NNCY","NANT","NRMD","PBLV","MLVL","PVDS","PVSM","STET","STRB","TOUL","VRSL","LILL","PAYV","PAYM","IUAR","MALQ","IMVT","PLVT"];
const mergedHeader = ["branchcode","biblionumber","no_abonnement_koha","numero","date_parution","date_reception","statut_arrive_manquant","Localisation","Cote"];
const colIndex = Object.fromEntries(mergedHeader.map((name, index) => [name, index])); // Ty copilot
const LOG = {ERR:"[1] ERROR",WAR:"[2] WARNING",INF:"[3] INFO"};
const reportData = [];
// Add an mapping normalized / original version of the locations to try and fix them
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

// --------------- Variable dependant functions ---------------
function appendToReport(gravity,sheetName,index,type,message) {
  reportData.push([gravity,sheetName,index,type,message]);
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

// ------------------------------------------------------------
//                      MAIN 
// ------------------------------------------------------------
// Do not ever remove ": ExcelScript.Workbook" or the script will fail
function main(workbook: ExcelScript.Workbook) {
  const originalSheetCount = workbook.getWorksheets().length; // Store now the original sheet count
  // --------------- Create all worksheets ---------------
  const SHEETS = {
    REPORT:{sheet:workbook.addWorksheet("Erreurs"),rowIndex:0},
    MERGED:{sheet:workbook.addWorksheet("Fusionnee"),rowIndex:0},
    SYNTHESIS:{sheet:workbook.addWorksheet("Synthese"),rowIndex:0}
  }

  // --------------- Define workbook functions ---------------
  function addLines(sheet, data) {
    // sheet is the object, not the actual Worksheet element, data is an array of array
    sheet.sheet.getRangeByIndexes(sheet.rowIndex, 0, data.length, data[0].length).setValues(data);
    sheet.rowIndex += data.length;
  }
  
  // --------------- Set up worksheets ---------------
  // Add headers on relevent sheets
  addLines(SHEETS.MERGED, [mergedHeader]);
  appendToReport("Gravité","Feuille","Numero de ligne","Type","Message"); // Yes this outputs numéro de ligne1 I'm aware
  SHEETS.SYNTHESIS.sheet.setPosition(1);
  SHEETS.REPORT.sheet.setPosition(2);
  SHEETS.MERGED.sheet.setPosition(3);

  // --------------- Prepare loops ---------------
  const data = []; // Will hold all the data until added to the sheet
  let processedSheets = 0; // Tracks number of sheets actually processed
  const bibnbCount = {}; // Tracks each bibnb apparition and how much of them exists

  // ---------------------------------------------
  //                  FIRST LOOP
  // ---------------------------------------------
  // Loop through each sheet
  workbook.getWorksheets().forEach((sheet) => {
    // If sheet name is not in the expected list, ignore it
    if (!(libCodes.includes(sheet.getName()))) {
      return;
    }
    processedSheets++;

    // Loop through data excluding headers
    sheet.getUsedRange().getValues().slice(1).forEach((row) => {
      // -------- Set up utils for the loop --------
      // ---- Function definition ----
      function getData(name) {
        // Return the value of a column for this row
        return row[colIndex[name]]
      }
      function fix(colName, value) {
        // Changes the cell inside the data container
        row[colIndex[colName]] = value;
      }
      function reportThis(gravity,type,message) {
        appendToReport(gravity,sheetName,data.length+1,type,message);
      }
      // ---- Name definition ----
      let sheetName = sheet.getName();
      let currentColName = null; // Temp just to avoid writing a billion times the same string

      // -------- Force branchcode to sheetname if different --------
      currentColName = "branchcode";
      if (getData(currentColName) !== sheetName) {
        reportThis(LOG.WAR,"Branchcode différent du nom de la feuille","Branchcode : " + getData(currentColName));
        fix(currentColName, sheetName);
      }

      // -------- Store bibnb for main bibnb detection later --------
      let bibnb = String(getData("biblionumber")).trim();
      if (!(bibnb in bibnbCount)) {
        bibnbCount[bibnb] = 0;
      }
      bibnbCount[bibnb]++;

      // -------- Status & location --------
      // Normalize (remove diacritic, trim, lowercase)
      currentColName = "statut_arrive_manquant";
      let newStatus = normalizeOutput(getData(currentColName));
      // Check if value is wrong
      if (!(["manquant","arrive","arrivee"].includes(newStatus)) && (getData(currentColName) != "")) {
        reportThis(LOG.ERR,"Statut arrivé manquant erronné",getData(currentColName))
      }
      // IF "manquant", force location to be empty
      let location = getData("Localisation");
      if ((newStatus === "manquant") && (location != "")) {
        reportThis(LOG.INF,"Manquant, suppression de sa localisation","Ancienne valeur : " + location);
        fix("Localisation", "");
      // Else, check if location is a legal value
      } else if ((!(locationIsValid(location))) && (location != "")) {
        // If not valid, try to find a valid alternative
        let altLocation = getSimilarLocation(location);
        if (altLocation === null) {
          reportThis(LOG.ERR,"Localisation inexistante",location);
        } else {
          reportThis(LOG.WAR,"Correction de la localisation","Nouvelle valeur : " + altLocation + " (ancienne : " + location + ")");
          fix("Localisation",altLocation);
        }
      }
      // Push the normlization after all checks so we keep original value
      fix(currentColName, newStatus);

      // -------- Dates --------
      // Convert to string and check if they're real (if interpreted from a string)
      ["date_parution","date_reception"].forEach((colName) => {
        let newDate = toYYYYMMDD(getData(colName));
        if (newDate === null) {
          reportThis(LOG.ERR,"Impossible d'interpréter " + colName, "Ancienne valeur : " + getData(colName));
        // Don't psh an empty value to avoid triggering two errors on the same problem
        } else {
          fix(colName, newDate);
        }
      })

      // -------- Push the row to the data container -------- 
      data.push(row)
    })
  });

  // ---------------------------------------------
  //            BETWEEN LOOP CALCULATIONS
  // ---------------------------------------------
  // For each bibnb, check if one is in more than 80% of lines
  let mainBibnb = null; // 
  for (key in bibnbCount) {
    if (bibnbCount[key]>data.length*0.8) {
      mainBibnb = key;
      break;
    }
  }
  // ---------------------------------------------
  //                  SECOND LOOP
  // ---------------------------------------------
  data.forEach((row, index) => {
    function getData(name) {
      // Return the value of a column for this row
      return row[colIndex[name]]
    }
    function reportThis(gravity,type,message) {
      // This is going to be awkward if there is no branchcode, but anyway this should not hapen. And we have the index
      appendToReport(gravity,getData("branchcode"),index+1,type,message);
    }

    // -------- Force biblionumber if 80% has the same --------
    let currentColName = "biblionumber";
    if ((getData(currentColName) != mainBibnb) && (mainBibnb !== null)) {
      reportThis(LOG.WAR,"Biblionumber différent du biblionumber principal","Ancienne valeur : " + getData(currentColName));
      row[colIndex[currentColName]] = mainBibnb;
    }

    // -------- Check missing vital data --------
    // We do that in 2nd loop so we can use bibnb edition
    let missingData = [];
    ["branchcode","biblionumber","numero","date_parution","date_reception","statut_arrive_manquant"].forEach((colName) => {
      if (getData(colName) == "") {
        missingData.push(colName);
      }
    })
    if (missingData.length > 0) {
      reportThis(LOG.ERR,"Informations vitales absentes","Colonnes : " + missingData.join(", "));
    }
  })

  // --------------- Report sheet processing ---------------
  // Check if sheets logic is fine. Should be originl count - 2 sheets because of "Modèle" & "Etat dépouillement"
  if ((originalSheetCount-2) == processedSheets) {
    appendToReport(LOG.INF, "Fusionnee",-1,"Nombre d'écoles traitées","Total : " + processedSheets)
  } else {
    appendToReport(LOG.ERR, "Fusionnee",-1,"Nombre d'écoles traitées","Total : " + processedSheets + "(nombre attendu : " + (originalSheetCount-2) + ")")
  }

  // --------------- Add data to the worksheets ---------------
  addLines(SHEETS.REPORT, reportData);
  addLines(SHEETS.MERGED, data);
  SHEETS.MERGED.sheet.getUsedRange().setNumberFormatLocal("@");
  SHEETS.MERGED.sheet.getUsedRange().setNumberFormat([["@"]]);

  // --------------- Report data looks ---------------
  // -------- Sort data --------
  let reportUsedRange = SHEETS.REPORT.sheet.getUsedRange();
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
  const pivotTable = SHEETS.SYNTHESIS.sheet.addPivotTable(
    "Synthese_Rapport_Erreur",
    reportUsedRange,
    SHEETS.SYNTHESIS.sheet.getRange("A1")
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
