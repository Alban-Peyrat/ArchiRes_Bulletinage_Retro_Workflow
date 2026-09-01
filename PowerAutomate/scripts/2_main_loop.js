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
  // --------------- Variables ---------------
  const sheetMerged = workbook.getWorksheet("Fusionnee");
  const reportData = [];
  const data = sheetMerged.getUsedRange().getValues(); // Will hold all the data until added to the sheet
  const bibnbCount = {}; // Tracks each bibnb apparition and how much of them exists

  // --------------- Define workbook functions ---------------
  function addDataToReport() {
    // data is an array of array
    let sheet = workbook.getWorksheet("Erreurs");
    sheet.getRangeByIndexes(sheet.getUsedRange().getRowCount(), 0, reportData.length, reportData[0].length).setValues(reportData);
  }

  data.forEach((row, index) => {
    // -------- Set up utils for the loop --------
    // ---- Function definition ----
    function getData(name) {
      // Return the value of a column for this row
      return row[colIndex[name]]
    }
    let sheetName = getData("sheetName");
    function fix(colName, value) {
      // Changes the cell inside the data container
      row[colIndex[colName]] = value;
    }
    function reportThis(gravity,type,message) {
      reportData.push([gravity,sheetName,index+1,type,message]);
    }
    // ---- Name definition ----
    let currentColName = null; // Temp just to avoid writing a billion times the same string
    let currentColVal = null; // Temp just to avoid writing a billion times the same string

    // -------- Force branchcode to sheetname if different --------
    currentColName = "branchcode";
    currentColVal = getData(currentColName);
    if (currentColVal !== sheetName) {
      reportThis(LOG.WAR,"Branchcode différent du nom de la feuille","Branchcode : " + currentColVal);
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
    currentColVal = getData(currentColName);
    let newStatus = normalizeOutput(getData(currentColName));
    // Check if value is wrong
    if (!(["manquant","arrive","arrivee"].includes(newStatus)) && (currentColVal != "")) {
      reportThis(LOG.ERR,"Statut arrivé manquant erronné",currentColVal)
    }
    // IF "manquant", force location to be empty
    currentColName = "Localisation";
    currentColVal = getData(currentColName);
    if ((newStatus === "manquant") && (currentColVal != "")) {
      reportThis(LOG.INF,"Manquant, suppression de sa localisation","Ancienne valeur : " + currentColVal);
      fix(currentColName, "");
    // Else, check if location is a legal value
    } else if ((!(locationIsValid(currentColVal))) && (currentColVal != "")) {
      // If not valid, try to find a valid alternative
      let altLocation = getSimilarLocation(currentColVal);
      if (altLocation === null) {
        reportThis(LOG.ERR,"Localisation inexistante",currentColVal);
      } else {
        reportThis(LOG.WAR,"Correction de la localisation","Nouvelle valeur : " + altLocation + " (ancienne : " + currentColVal + ")");
        fix(currentColName,altLocation);
      }
    }
    // Push the status normlization after all checks so we keep original value
    fix("statut_arrive_manquant", newStatus);

    // -------- Dates --------
    // Convert to string and check if they're real (if interpreted from a string)
    ["date_parution","date_reception"].forEach((colName) => {
      currentColVal = getData(colName);
      let newDate = toYYYYMMDD(currentColVal);
      if (newDate === null) {
        reportThis(LOG.ERR,"Impossible d'interpréter " + colName, "Ancienne valeur : " + currentColVal);
      // Don't psh an empty value to avoid triggering two errors on the same problem
      } else {
        fix(colName, newDate);
      }
    })
  })
  // --------------- Add data to the worksheets ---------------
  addDataToReport();
  sheetMerged.getUsedRange().clear();
  sheetMerged.getRange("A:I").setNumberFormatLocal("@");
  sheetMerged.getRangeByIndexes(0,0,data.length,data[0].length).setValues(data);

  // ---------------------------------------------
  //            BETWEEN LOOP CALCULATIONS
  // ---------------------------------------------
  // Add an hidden worksheet with the main bibnb
  if (!(workbook.getWorksheet("PowerAutomateHidden"))) {
    workbook.addWorksheet("PowerAutomateHidden");
    workbook.getWorksheet("PowerAutomateHidden").setVisibility(ExcelScript.SheetVisibility.hidden);
  } else {
    workbook.getWorksheet("PowerAutomateHidden").getRange("A1").clear();
  }
  // For each bibnb, check if one is in more than 80% of lines
  for (key in bibnbCount) {
    if (bibnbCount[key]>data.length*0.8) {
      workbook.getWorksheet("PowerAutomateHidden").getRange("A1").setValue(key);
      break;
    }
  }
}