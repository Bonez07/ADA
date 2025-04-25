//Config
const mode = "real"; //"test" or "real"
const termAndWeek = ""; //Write in format "T_W_"
const houseOnDuty = ""; //House on duty (3 letters). For SLT weeks write "SLT", for probation duty weeks "Probation Nominees"

//Sheets & docs IDs
const rosterTemplateID = ""; //Duty roster template
//Real
const realPBAttendanceSheetID = ""; //PB attendance sheet
const realHouseAttendanceSheetID = ""; //House/nominees attendance sheet
const realDatabaseSheetID = ""; //Duty allocation database sheet
const realRosterFolderID = ""; //Folder where duty roster is created
//Test
const testPBAttendanceSheetID = "";
const testHouseAttendanceSheetID = "";
const testDatabaseSheetID = "";
const testRosterFolderID = "";

const copyRosterTemplate = true; //If true duty roster template will be copied, if false names will be directly put on the template
const schoolPrefectColor = "#000000"; //color of school prefect names on roster
const housePrefectColor = "#ff0000"; //color of house prefect/nominee names on roster

//Allocation settings (all between 0 and 1)
const subcommProbabilityScalingFactor = 0.7; //how many exco/emcee duties subcomm excos get compared to regular excos (between 0 and 1)
const yearRepScoreOffset = 5; //The higher this value, the less duties year reps will get compared to prefects

const dutiesThisWeekWeightingPB = 0.6; //The higher this value, the more ADA will avoid giving prefects multiple duties in the same week
const dutiesThisWeekWeightingHouse = 0.9; //Same as the previous value but for house prefects
const generalScoreWeighting = 0.85; //The higher this value, the more duties will be given to those who have less total duties
const timeSinceSpecificDutyWeighting = 0.6; //The higher this value, the higher the weighting of the time since specific duty compared to the number of specific duties
const worseScoreWeighting = 0.23; //The higher this value, the more weighting will be on the worse of the general and specific score

//Database format (1-indexed) (col 1 = col A)
const fullNameCol = 1;
const nameCol = 2;
const dataCol = 3;
const pastDutiesCol = 4;
const unavailableDaysCol = 5;
const firstPrefectRow = 3;

//Globals
var prefectorialBoard;
var housePrefects;
var rosterDoc;
var dutyCount = 0;
var numberOfDuties = 0;
function main() {
  console.log("Initializing");
  initialize();
  console.log("Generating attendance sheets");
  prefectorialBoard.generateAttendanceSheet(housePrefects);
  if (houseOnDuty != "SLT") {
    housePrefects.generateAttendanceSheet();
  }
  console.log("Processing duties");
  let duties = getDutiesList(rosterDoc);
  duties = processDuties(duties);
  duties = linkDuties(duties, prefectorialBoard, housePrefects);
  numberOfDuties = duties.length; 
  //Sort duties into forced, exco, PB, and HP
  let forcedDuties = [], excoDuties = [], pbDuties = [], hpDuties = [];
  let dutyDays = []
  for (const duty of duties) {
    dutyDays.push(duty.day);
    if (duty.requirements.length == 1) {
      forcedDuties.push(duty);
      continue;
    }
    if (duty.requirements[3].includes("E") && duty.prefects == "PB") {
      excoDuties.push(duty);
      continue;
    }
    if (duty.prefects == "PB") {
      pbDuties.push(duty);
      continue;
    }
    if (duty.prefects == "HP") {
      hpDuties.push(duty);
      continue;
    }
  }
  console.log("Allocating duties");
  //Allocate house cap duty
  if (!(houseOnDuty == "SLT" || houseOnDuty == "Probation Nominees")) {
    prefectorialBoard.allocateHouseCap(houseOnDuty, [...new Set(dutyDays)]);
  }
  //Allocate forced duties
  for (const duty of forcedDuties) {
    if (duty.prefects == "PB") {
      prefectorialBoard.forceDuty(duty);
    }
    else if (duty.prefects == "HP") {
      housePrefects.forceDuty(duty);
    }
  }
  //Allocate Exco duties
  while (excoDuties.length > 0) {
    let index = findNextDutyIndex(excoDuties, prefectorialBoard);
    prefectorialBoard.allocateDuty(excoDuties[index])
    excoDuties.splice(index, 1)
  }
  //Allocate PB duties
  while (pbDuties.length > 0) {
    let index = findNextDutyIndex(pbDuties, prefectorialBoard);
    prefectorialBoard.allocateDuty(pbDuties[index])
    pbDuties.splice(index, 1)
  }
  //Allocate HP duties
  while (hpDuties.length > 0) {
    let index = findNextDutyIndex(hpDuties, housePrefects);
    housePrefects.allocateDuty(hpDuties[index])
    hpDuties.splice(index, 1)
  }
  //update database & attendance sheets
  console.log("Updating database, attendance sheets, roster")
  prefectorialBoard.updateDatabase();
  prefectorialBoard.placeCheckboxes();
  if (houseOnDuty != "SLT") {
    housePrefects.updateDatabase();
    housePrefects.placeCheckboxes();
  }
  //update roster
  findAndReplace("<<T_W_>>", termAndWeek);
  findAndReplace("<<HOUSE_ON_DUTY>>", houseOnDuty);
  console.log(`Done! - Allocated ${dutyCount} out of ${numberOfDuties} duties`);
}
function syncTestSpreadsheets() {
  console.log("Updating Database");
  replaceSheetContent(realDatabaseSheetID, testDatabaseSheetID);
  console.log("Updating PB Attendance sheet");
  replaceSheetContent(realPBAttendanceSheetID, testPBAttendanceSheetID);
  console.log("Updating House Attendance sheet");
  replaceSheetContent(realHouseAttendanceSheetID, testHouseAttendanceSheetID);
  console.log("Done!")
}
function replaceSheetContent(sourceID, targetID) {
  // Open the source and target spreadsheets.
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceID);
  var targetSpreadsheet = SpreadsheetApp.openById(targetID);
  
  // Create a temporary sheet in the target to avoid having no sheets.
  var tempSheet = targetSpreadsheet.insertSheet('TempSheetForDeletion');
  
  // Delete all existing sheets in the target spreadsheet.
  var targetSheets = targetSpreadsheet.getSheets();
  for (var i = 0; i < targetSheets.length; i++) {
    // Skip the temporary sheet
    if (targetSheets[i].getName() !== 'TempSheetForDeletion') {
      targetSpreadsheet.deleteSheet(targetSheets[i]);
    }
  }
  // Copy each sheet from the source spreadsheet to the target spreadsheet.
  var sourceSheets = sourceSpreadsheet.getSheets();
  for (var j = 0; j < sourceSheets.length; j++) {
    var copiedSheet = sourceSheets[j].copyTo(targetSpreadsheet);
    // Set the sheet name to match the source sheet name.
    copiedSheet.setName(sourceSheets[j].getName());
  }
  // Remove the temporary sheet
  targetSpreadsheet.deleteSheet(tempSheet);
}
function initialize() {
  //init: create doc, open sheets and check mode
  //Init sheets
  var pbAttendanceSheet;
  var houseAttendanceSheet;
  var databaseSheet;
  var rosterFolder = undefined;
  if (mode.toLowerCase().trim() == "test") {
    pbAttendanceSheet = SpreadsheetApp.openById(testPBAttendanceSheetID);
    houseAttendanceSheet = SpreadsheetApp.openById(testHouseAttendanceSheetID);
    databaseSheet = SpreadsheetApp.openById(testDatabaseSheetID);
    if (testRosterFolderID) {
      rosterFolder = DriveApp.getFolderById(testRosterFolderID);
    }
  }
  else if (mode.toLowerCase().trim() == "real") {
    pbAttendanceSheet = SpreadsheetApp.openById(realPBAttendanceSheetID);
    houseAttendanceSheet = SpreadsheetApp.openById(realHouseAttendanceSheetID);
    databaseSheet = SpreadsheetApp.openById(realDatabaseSheetID);
    if (realRosterFolderID) {
      rosterFolder = DriveApp.getFolderById(realRosterFolderID);
    }
  }
  else {
    throw new Error(`Invalid mode: "${mode}" - please set to either "real" or "test"`);
  }
  //Create prefect classes 
  prefectorialBoard = new Prefects("PB", "#ffffff", pbAttendanceSheet, databaseSheet, rosterDoc);
  //Replace house captains data for SLT weeks
  if (houseOnDuty == "SLT") {
    for (let i = 0; i < prefectorialBoard.prefects.length; i++) {
      if (prefectorialBoard.prefects[i].data[3] == "H") {
        prefectorialBoard.prefects[i].data = prefectorialBoard.prefects[i].data.substring(0, 3) + "N";
      }
    }
  }
  else {
    switch (houseOnDuty.toUpperCase().trim()) {
      case "PROBATION NOMINEES":
        housePrefects = new Prefects("Probation Nominees", "#ffffff", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "CKS":
        housePrefects = new Prefects("CKS", "#00ffff", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "GHK":
        housePrefects = new Prefects("GHK", "#ffff00", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "LSG":
        housePrefects = new Prefects("LSG", "#b7b7b7", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "OLD":
        housePrefects = new Prefects("OLD", "#ff0000", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "SVM":
        housePrefects = new Prefects("SVM", "#9900ff", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "TCT":
        housePrefects = new Prefects("TCT", "#ff9900", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "THO":
        housePrefects = new Prefects("THO", "#00ff00", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      case "TKK":
        housePrefects = new Prefects("TKK", "#0000ff", houseAttendanceSheet, databaseSheet, rosterDoc);
        break;
      default:
        throw new Error(`Invalid house: ${houseOnDuty}`);
    }
  }
  //Init roster doc
  if (copyRosterTemplate) {
    console.log("Creating duty roster document");
    //copy roster template
    var sourceDoc = DriveApp.getFileById(rosterTemplateID);
    //Make a copy of the document
    if (rosterFolder != undefined) {
      var copiedDoc = sourceDoc.makeCopy(houseOnDuty + " " + termAndWeek, rosterFolder);
    }
    else {
      var copiedDoc = sourceDoc.makeCopy(houseOnDuty + " " + termAndWeek);
    }
    //Get the ID of the copied document
    rosterDoc = DocumentApp.openById(copiedDoc.getId());
    //TODO check if doc already exists
  }
  else {
    rosterDoc = DocumentApp.openById(rosterTemplateID);
  }
}
function findAndReplace(find, replace) {
  let targetColor = undefined;
  if (find.substring(2, 4) == "PB") {
    find = find.slice(0, 2) + find.slice(5);
    targetColor = schoolPrefectColor;
  }
  else if (find.substring(2, 4) == "HP") {
    find = find.slice(0, 2) + find.slice(5);
    targetColor = housePrefectColor;
  }
  var body = rosterDoc.getBody();
  // Escape regex special characters in the find string
  var escapedStr = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // Find the first occurrence of the text
  var foundElement = body.findText(escapedStr);
  
  while (foundElement) {
    var textElement = foundElement.getElement().editAsText();
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    var fgColor = textElement.getForegroundColor(start);
    // Check if the text's foreground color matches the target color.
    if (targetColor == undefined || fgColor === targetColor || (fgColor === null && targetColor === schoolPrefectColor)) {
      // Capture the attributes and explicitly capture the foreground color.
      var attributes = textElement.getAttributes(start);
      
      // Delete the matched text range and insert the replacement.
      textElement.deleteText(start, end);
      textElement.insertText(start, replace);
      
      // Reapply the captured attributes to the new text.
      textElement.setAttributes(start, start + replace.length - 1, attributes);
      // Explicitly set the foreground color if it exists.
      if (fgColor) {
        textElement.setForegroundColor(start, start + replace.length - 1, fgColor);
      }
      break; // Only replace the first occurrence with the correct color.
    } else {
      // If not matching, look for the next occurrence.
      foundElement = body.findText(escapedStr, foundElement);
    }
  }
}
function getDutiesList() {
  //duty format {id: "", day: , prefects: "", requirements: [], raw: ""}
  let duties = [];
  let rawDuties = extractRawDuties();
  for (const rawDuty of rawDuties) {
    let splitDuty = rawDuty.split("_");
    //check if duty length is correct
    if (splitDuty.length != 4) {
      console.error(`Invalid duty: "${rawDuty}" - the duty does not contain 4 parts separated by underscores as expected`)
      continue;
    }
    let duty = {prefects: splitDuty[0], id: splitDuty[1], day: Number(splitDuty[2]), requirements: splitDuty[3], raw: rawDuty}
    duty.requirements = duty.requirements.match(/(\([^)]*\)|\[[^\]]*\])|./g).map(c => c.startsWith('(') ? c.slice(1, -1) : c);
    //error checking
    if (!(duty.day > 0 && duty.day < 6)) {
      console.error(`Invalid duty: "${rawDuty}" - "${splitDuty[2]}" is not a valid day. Ensure day is between 1 and 5`);
      continue;
    }
    if (duty.id.length != 2) {
      console.warn(`Invalid duty: "${rawDuty}" - "${duty.id}" is not a valid ID. Ensure duty ID is 2 characters long (seen ${duty.id.length})`);
    }
    if (duty.requirements.length != 1 && duty.requirements.length != 4) {
      console.error(`Invalid duty: "${rawDuty}" - "${splitDuty[3]}" is not a valid set of requirements. Ensure it contains 4 elements (multiple letters in brackets are counted one element with multiple possible requirements) or 1 element in brackets if it is a forced duty (seen ${duty.requirements.length} elements)`);
      continue;
    }
    duties.push(duty);
  }
  return duties;
}
function extractRawDuties() {
  let duties = [];
  const body = rosterDoc.getBody();
  // Use findText to locate the first occurrence
  let searchResult = body.findText('{{(.*?)}}');
  
  while (searchResult !== null) {
    const element = searchResult.getElement().asText();
    const start = searchResult.getStartOffset();
    const end = searchResult.getEndOffsetInclusive();
    
    // Extract the matching text (including the curly braces)
    const fullMatch = element.getText().substring(start, end + 1);
    // Remove the curly braces and trim whitespace
    const innerText = fullMatch.slice(2, -2).trim();
    
    // Get the foreground color for the range of the matched text.
    // This returns the CSS-style color code (e.g., "#000000")
    const color = element.getForegroundColor(start);
    
    if (color == schoolPrefectColor || color == null) {
      duties.push("PB_" + innerText);
    }
    else if (color == housePrefectColor) {
      duties.push("HP_" + innerText);
    }
    else {
      console.warn(`Unrecognised color "${color}" of duty "${innerText}" - please make the text ${schoolPrefectColor} for PB and ${housePrefectColor} for house prefects/nominees`);
    }
    // Look for the next match after the current one.
    searchResult = body.findText('{{(.*?)}}', searchResult);
  }
  return duties;
}
function linkDuties(rawDuties, pb, hp) {
  let notLinkedDuties = rawDuties.filter(duty => !(/\[.+?]/.test(duty.raw)));
  let linkedDuties = rawDuties.filter(duty => /\[.+?]/.test(duty.raw));
  //for each linked duty
  for (let dutyIndex = 0; dutyIndex < linkedDuties.length; dutyIndex++) {
    //for each requirement
    for (let requirementIndex = 0; requirementIndex < linkedDuties[dutyIndex].requirements.length; requirementIndex++) {
      //if requirement contains []
      if (/\[.+?]/.test(linkedDuties[dutyIndex].requirements[requirementIndex])) {
        //loop all other duties paired and test
        let matches = []; //list of second duty indices corresponding to links
        for (let secondDutyIndex = 0; secondDutyIndex < linkedDuties.length; secondDutyIndex++) {
          //if it is on the same day
          if (linkedDuties[dutyIndex].day == linkedDuties[secondDutyIndex].day) {
            if (linkedDuties[dutyIndex].requirements[requirementIndex] == linkedDuties[secondDutyIndex].requirements[requirementIndex]) {
              //Match
              matches.push(secondDutyIndex);
            }
          }
        }
        //Choose which duty to put requirement on
        //Loop through the matches, modify them and check number of elligible
        let noElligible = [];
        let otherSymbol = "*"
        for (const match of matches) {
          if (linkedDuties[match].requirements[requirementIndex].match(/\[([^-\]]+)(?:-[^\]]*)?\]/)[1].includes("/")) {
            otherSymbol = linkedDuties[match].requirements[requirementIndex].match(/\[[^/\]]*\/([^-\]]+)(?:-[^\]]*)?\]/)[1];
          }
          linkedDuties[match].requirements[requirementIndex] = linkedDuties[match].requirements[requirementIndex].match(/\[([^\-\/\]]+)(?:[-\/][^\]]*)?\]/)[1];
          if (linkedDuties[match].prefects == "PB") {
            noElligible.push(pb.findGoodElligible(linkedDuties[match]).length);
          }
          if (linkedDuties[match].prefects == "HP") {
            noElligible.push(hp.findGoodElligible(linkedDuties[match]).length);
          }
        }
        let possible = [];
        let sumOfNoElligible = 0;
        for (let i = 0; i < matches.length; i++) {
          possible.push(i);
          sumOfNoElligible += noElligible[i];
        }
        //radomly pick from possible, find corresponding match
        const randomPick = (arr, probs, r = Math.random()) => arr[probs.findIndex(p => (r -= p) < 0)];
        const probabilities = possible.map(possible => (noElligible[possible]/sumOfNoElligible));
        let chosenIndex = randomPick(possible, probabilities); //matches index that is chosen
        //replace requirement for all the other matches as otherSymbol (all of them were replaced earlier)
        for (let i = 0; i < matches.length; i++) {
          if (i != chosenIndex) {
            linkedDuties[matches[i]].requirements[requirementIndex] = otherSymbol;
          }
        }
      }
    }
  }
  return [...notLinkedDuties, ...linkedDuties];
}
function processDuties(rawDuties, pb) {
  //Shuffle duties
  duties = [...rawDuties]
  for (let i = duties.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [duties[i], duties[j]] = [duties[j], duties[i]];
  }
  duties.sort((a, b) => [2, 4].includes(a.day) - [2, 4].includes(b.day));
  duties.forEach(duty => {
    if (duty.prefects == "PB") {
      if (duty.id == "EX") {
        if (duty.requirements[3] == "ES") {
          const noSubcomms = prefectorialBoard.prefects.filter(prefect => prefect.data[3] == "S" && !(prefect.unavailableDays.includes(duty.day))).length;
          const noExcos = prefectorialBoard.prefects.filter(prefect => prefect.data[3] == "E" && !(prefect.unavailableDays.includes(duty.day))).length;
          const subcommProbability = Math.min((2*subcommProbabilityScalingFactor*noSubcomms)/(noExcos + subcommProbabilityScalingFactor*noSubcomms), 0.9);
          duty.requirements[3] = Math.random() < subcommProbability ? "S" : "E";
        }
      }
      //special case for emcee duties
      else if (duty.id == "MC") {
        if (duty.requirements[3] == "ES") {
          const noSubcomms = prefectorialBoard.prefects.filter(prefect => prefect.data[3] == "S" && !(prefect.unavailableDays.includes(duty.day))).length;
          const noExcos = prefectorialBoard.prefects.filter(prefect => prefect.data[3] == "E" && !(prefect.unavailableDays.includes(duty.day))).length;
          const subcommProbability = (subcommProbabilityScalingFactor*noSubcomms)/(noExcos+noSubcomms);
          duty.requirements[3] = Math.random() < subcommProbability ? "S" : "E";
        }
      }
    }
  });
  return duties;
}
function findNextDutyIndex(duties, prefects) {
    //find duty with least elligible
    let noElligible = duties.map(duty => prefects.findGoodElligible(duty).length + 0.3*prefects.findElligible(duty).length);
    let minimum = Math.min(...noElligible);
    noElligible = duties.map(duty => {
      value = prefects.findGoodElligible(duty).length + 0.3*prefects.findElligible(duty).length;
      if (minimum > 3) {
        if ([1, 3, 5].includes(duty.day)) {
          value -= minimum/4;
        }
      }
      if (minimum > 6) {
        //for rare duties make them allocate first
        if (duty.id == "FA") {
          value -= minimum/4;
        }
        else if (duty.id == "DO" || duty.id == "CA" || duty.id == "MC") {
          value -= minimum/6;
        }
        else if (duty.id == "PA" || duty.id == "BG") {
          value -= minimum/9;
        }
      }
      return value;
      });
    minimum = Math.min(...noElligible);
    if (minimum <= 0) {
      //if no good elligible find regular elligible
      noElligible = duties.map(duty => prefects.findElligible(duty).length);
      minimum = Math.min(...noElligible);
    }
    //find index
    let index = noElligible.indexOf(minimum);
    return index;
}
function standardize(value, list) {
  if (!Array.isArray(list) || list.length === 0) {
    return 0;
  }
  let mean = list.reduce((sum, num) => sum + num, 0) / list.length;
  let stddev = Math.sqrt(list.reduce((sum, num) => sum + Math.pow(num - mean, 2), 0) / list.length);
  if (stddev === 0) {
    return 0;
  }
  return (value - mean) / stddev;
}
class Prefects {
  constructor(name, color, attendanceSheet, databaseSheet, rosterDoc) {
    this.name = name;
    this.color = color; //in hex
    //generate and open attendance sheet
    this.attendanceSheet = attendanceSheet;
    this.attendanceWorksheet = undefined;
    //open database and get data
    this.databaseWorksheet = databaseSheet.getSheetByName(this.name);
    this.fullData = this.databaseWorksheet.getDataRange().getValues();
    this.rosterDoc = rosterDoc;
    //create prefects list based on data
    this.prefects = this.processData();
    this.prefectsAttendanceLength = this.databaseWorksheet.getLastRow()-(firstPrefectRow-1);
    //format {number: "", name: "", data: "", pastDuties: "", unavailableDays (this needs to be processed): "", adjacent: "", dutiesThisWeek: "" totalDuties: , dutyDays: []}
  }
  forceDuty(duty) {
    dutyCount++;
    let prefect;
    if (!isNaN(duty.requirements[0])) {
      prefect = this.prefects.find(person => person.number == duty.requirements[0]);
    }
    else {
      prefect = this.prefects.find(person => person.name == duty.requirements[0]);
      if (prefect == undefined) {
        prefect = this.prefects.find(person => person.fullName == duty.requirements[0]);
      }
    }
    if (prefect == undefined) {
      console.error(`Prefect "${duty.requirements[0]}" not found for forced duty "${duty.raw}"`);
    }
    this.incrementDuties(prefect, duty);
    return prefect
  }
  allocateDuty(duty) {
    dutyCount++;
    let elligible = this.findElligible(duty);
    let initialElligible = elligible;
    if (elligible.length == 0) {
      console.error(`No one available for "${duty.raw}" - ensure there are not more duties than available prefects, that prefects can match the duty requirements, and that the requirements are keyed in correctly`);
      return undefined;
    }
    //check adjacent
    elligible = elligible.filter(prefect => !(prefect.adjacentDutyDays.includes(duty.day)));
    if (elligible.length == 0) {
      elligible = initialElligible;
      console.warn(`No non-adjacent day matches for "${duty.raw}" - prefect may get duty 2 days in a row`);
    }
    else {
      initialElligible = elligible;
    }
    if (duty.id != "EX") {
      //prevent giving 3 duties in 1 week
      elligible = elligible.filter(prefect => prefect.dutiesThisWeek < 2);
      if (elligible.length == 0) {
        elligible = initialElligible;
        console.warn(`No 2 duties in week matchess for "${duty.raw}" - prefect may get 3 duties in the week`);
      }
      else {
        initialElligible = elligible;
      }
      //check previous duty
      elligible = elligible.filter(prefect => !(prefect.pastDuties.substring(prefect.pastDuties.length - 3).includes(duty.id)));
      if (elligible.length == 0) {
        elligible = initialElligible;
        console.warn(`No different previous duty matches for "${duty.raw}" - prefect may get the same type of duty twice in a row"`);
      }
    }
    //scoring remaining prefects
    elligible.forEach(prefect => {
      if (prefect.data[3] == "S" && duty.id != "EX") {
        prefect.timeSinceSpecificDuty = (prefect.pastDuties.split('').reverse().join('').indexOf(" " + duty.id.split('').reverse().join(''))+prefect.pastDuties.replaceAll("EX ", "").split('').reverse().join('').indexOf(" " + duty.id.split('').reverse().join('')))/6;
        if ((prefect.pastDuties.replaceAll("EX ", "").split(" ").length - 1) != 0) {
          prefect.totalSpecificDuties = (prefect.pastDuties.split(duty.id).length - 1)*(prefect.pastDuties.split(" ").length - 1)/(prefect.pastDuties.replaceAll("EX ", "").split(" ").length - 1);
        }
        else {
          prefect.totalSpecificDuties = prefect.pastDuties.split(duty.id).length - 1;
        }
      }
      else {
        prefect.timeSinceSpecificDuty = prefect.pastDuties.split('').reverse().join('').indexOf(" " + duty.id.split('').reverse().join(''))/3;
        prefect.totalSpecificDuties = prefect.pastDuties.split(duty.id).length - 1;
      }
    });
    //create total list variables for standardization
    let totalDuties = elligible.map(prefect => prefect.totalDuties);
    let timesSinceSpecificDuty = elligible.map(prefect => prefect.timeSinceSpecificDuty);
    let totalSpecificDuties = elligible.map(prefect => prefect.totalSpecificDuties);
    //standardize and score
    elligible.forEach(prefect => {
      if (prefect.timeSinceSpecificDuty < 0 || prefect.timeSinceSpecificDuty === NaN) {
        prefect.timeSinceSpecificDuty = Math.max(...timesSinceSpecificDuty) + 1;
      }
      if (this.name == "PB" || this.name == "Probation Nominees") {
        prefect.generalScore = (1-dutiesThisWeekWeightingPB)*(prefect.totalDuties - Math.min(...totalDuties)) + (dutiesThisWeekWeightingPB)*(prefect.dutiesThisWeek);
        prefect.specificScore = timeSinceSpecificDutyWeighting*(-standardize(prefect.timeSinceSpecificDuty, timesSinceSpecificDuty)) + (1-timeSinceSpecificDutyWeighting)*(standardize(prefect.totalSpecificDuties, totalSpecificDuties));
      }
      else {
        prefect.generalScore = (1-dutiesThisWeekWeightingHouse)*(prefect.totalDuties - Math.min(...totalDuties)) + (dutiesThisWeekWeightingHouse)*(prefect.dutiesThisWeek);
        prefect.specificScore = timeSinceSpecificDutyWeighting*(-standardize(prefect.timeSinceSpecificDuty, timesSinceSpecificDuty)) + (1-timeSinceSpecificDutyWeighting)*(standardize(prefect.totalSpecificDuties, totalSpecificDuties));
      }
    });
    let generalScores = elligible.map(prefect => prefect.generalScore);
    elligible.forEach(prefect => {
      prefect.generalScore = standardize(prefect.generalScore, generalScores);
      //total score is average of average of scores and maximum of scores (average scewed towards higher value)
      if (duty.id == "EX") {
        //for exco specific score is only used as a discriminator when tied
        prefect.totalScore = (95/100)*prefect.generalScore + (5/100)*prefect.specificScore;
      }
      else {
        prefect.totalScore = (worseScoreWeighting)*Math.max(prefect.generalScore, prefect.specificScore) + (1-worseScoreWeighting)*((generalScoreWeighting)*prefect.generalScore + (1-generalScoreWeighting)*prefect.specificScore);
      } 
      if (prefect.data[3] == "Y") {
        prefect.totalScore += yearRepScoreOffset;
      }  
      //add bias for non-singaporeans non-christians so that SC can be saved for duties that require them
      prefect.totalScore += (Number(prefect.data[0] == "S") + Number(prefect.data[1] == "C"))/8;
    });
    let finalOptions = elligible.filter(prefect => Math.min(...elligible.map(person => person.totalScore)) == prefect.totalScore);
    //randomly select from lowest scores
    let prefect = finalOptions[Math.floor(Math.random() * finalOptions.length)];
    //increment duties
    this.incrementDuties(prefect, duty);
    return prefect;
  }
  allocateHouseCap(houseOnDuty, days) {
    let houseCaptain = this.prefects.find(prefect => prefect.data.substring(5) == houseOnDuty);
    if (houseCaptain == undefined) {
      console.error(`No House Captain found for ${houseOnDuty} - ensure house is marked on the House Captain's data in the database (e.g. SNGH-${houseOnDuty})`);
      return;
    }
    for (const day of days) {
      this.incrementDuties(houseCaptain, {id: "HC", day: day, prefects: "PB", requirements: [], raw: "PB_HC_" + String(day) });
    }
  }
  incrementDuties(prefect, duty) {
    //increase total duties
    this.prefects[this.prefects.indexOf(prefect)].totalDuties++;
    //increase total duties that week
    this.prefects[this.prefects.indexOf(prefect)].dutiesThisWeek++;
    //add to past duties
    if (this.prefects[this.prefects.indexOf(prefect)].pastDuties.length > 0 && this.prefects[this.prefects.indexOf(prefect)].pastDuties[this.prefects[this.prefects.indexOf(prefect)].pastDuties.length - 1] != " ") {
      this.prefects[this.prefects.indexOf(prefect)].pastDuties += (" ");
    }
    if (Math.max(...this.prefects[this.prefects.indexOf(prefect)].dutyDays, duty.day) == duty.day) {
      //if last duty of the week
      this.prefects[this.prefects.indexOf(prefect)].pastDuties += (duty.id + " ");
    }
    else {
      //put in the correct position otherwise
      const sortedArr = [...this.prefects[this.prefects.indexOf(prefect)].dutyDays, duty.day].sort((a, b) => b - a);
      const index = this.prefects[this.prefects.indexOf(prefect)].pastDuties.length - (3 * sortedArr.indexOf(duty.day));
      this.prefects[this.prefects.indexOf(prefect)].pastDuties = this.prefects[this.prefects.indexOf(prefect)].pastDuties.slice(0, index) + (duty.id + " ") + this.prefects[this.prefects.indexOf(prefect)].pastDuties.slice(index);
    }
    //remove from available
    this.prefects[this.prefects.indexOf(prefect)].unavailableDays.push(duty.day);
    this.prefects[this.prefects.indexOf(prefect)].dutyDays.push(duty.day);
    //update adjacent duties
    this.prefects[this.prefects.indexOf(prefect)].adjacentDutyDays.push(duty.day, duty.day + 1, duty.day - 1);
    //add checkbox to list
    //this.attendanceWorksheet.getRange(this.databaseWorksheet.getRange(prefect.number, duty.day + 1).getA1Notation()).insertCheckboxes();
    this.checkboxes[prefect.number-3][duty.day-1] = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    //find and replace on document
    if (duty.id != "HC") {
      findAndReplace("{{" + duty.raw + "}}", prefect.name);
    }
  }
  generateAttendanceSheet(house=this) { 
    //pass in house for PB
    var templateSheet = this.attendanceSheet.getSheetByName("Template");
    //Generate sheet
    if (!this.attendanceSheet.getSheetByName(termAndWeek)) {
      this.attendanceWorksheet = templateSheet.copyTo(this.attendanceSheet).setName(termAndWeek);
    }
    else {
      //If sheet already exists load the existing worksheet
      this.attendanceWorksheet = this.attendanceSheet.getSheetByName(termAndWeek);
      console.warn("Attendance sheet with the same name already exists for " + this.name + "\nCheckboxes will be placed on the existing worksheet but will not remove existing checkboxes")
    }
    //Set color and house
    this.attendanceWorksheet.setTabColor(house.color);
    this.attendanceWorksheet.getRange("A1").setValue("On Duty: " + house.name);
    this.attendanceWorksheet.getRange("A1").setBackground(house.color);
    //For retrieving total number of duties
    const attendanceWorksheets = this.attendanceSheet.getSheets();
    //For PB and nominees
    if (this.name == "PB" || this.name == "Probation Nominees") {
      let names = this.databaseWorksheet.getRange(firstPrefectRow, 1, this.prefectsAttendanceLength, 2).getValues();
      let values = Array.from({ length:  names.length}, () => ["", ""]);
      for (var row = firstPrefectRow; row < values.length+firstPrefectRow; row++) {
        //if there is a name in the database
        if (names[row-firstPrefectRow][0].length > 0 && names[row-firstPrefectRow][1].length > 0) {
          if (attendanceWorksheets.length > 1) {
            //Takes values from the second last sheet
            values[row-firstPrefectRow][1] = [`='${attendanceWorksheets[attendanceWorksheets.length - 2].getName()}'!H${String(row)}+G${String(row)}`];
          }
          else {
            //Othewise take just from this weeks duties
            values[row-firstPrefectRow][1] = [`=G${String(row)}`];
          }
          values[row-firstPrefectRow][0] = `=COUNTIF(B${String(row)}:F${String(row)},FALSE)+COUNTIF(B${String(row)}:F${String(row)},TRUE)`;
        }
      }
      //Upadte values on the sheet
      this.attendanceWorksheet.getRange(firstPrefectRow, 7, this.prefectsAttendanceLength, 2).setValues(values);
    }
    //for house prefects
    else {
      //copy names from database and paste them in the first column
      let fullNames = this.databaseWorksheet.getRange(firstPrefectRow, fullNameCol, this.prefectsAttendanceLength, 1).getValues();
      let names = this.databaseWorksheet.getRange(firstPrefectRow, nameCol, this.prefectsAttendanceLength, 1).getValues();
      let values = Array.from({ length:  fullNames.length}, () => ["", ""]);
      this.attendanceWorksheet.getRange(firstPrefectRow, 1, this.prefectsAttendanceLength, 1).setValues(fullNames);
      //for each worksheet filter if both name and color match
      let previousAttendanceWorksheets = attendanceWorksheets.filter(sheet => {
        if ((sheet.getTabColor() == this.color) && (sheet.getRange("A1").getValue().includes(this.name))) {
          return true;
        }
        if (sheet.getTabColor() == this.color) {
          console.warn(`Sheet "${sheet.getName()}" matches the house color but not the house name \nEnsure attendance sheets for the same house all have the same tab color and contain the house name in cell A1`);
          return true;
        }
        if (sheet.getRange("A1").getValue().includes(this.name)) {
          console.warn(`Sheet "${sheet.getName()}" matches the house name but not the house color \nEnsure attendance sheets for the same house all have the same tab color and contain the house name in cell A1`);
          return true;
        }
        return false;});
      //get total values from 2nd last worksheet that matches (to exclude the current)
      for (var row = firstPrefectRow; row < values.length+firstPrefectRow; row++) {
        if (fullNames[row-firstPrefectRow][0].length > 0 && names[row-firstPrefectRow][0].length > 0) {
          if (previousAttendanceWorksheets.length > 1) {
            values[row-firstPrefectRow][1] = `='${previousAttendanceWorksheets[previousAttendanceWorksheets.length-2].getName()}'!H${String(row)}+G${String(row)}`;
          }
          else {
            values[row-firstPrefectRow][1] = `=G${String(row)}`;
          }
          values[row-firstPrefectRow][0] = `=COUNTIF(B${String(row)}:F${String(row)},FALSE)+COUNTIF(B${String(row)}:F${String(row)},TRUE)`;
        }
      }
      //update values on sheet
      this.attendanceWorksheet.getRange(firstPrefectRow, 7, this.prefectsAttendanceLength, 2).setValues(values);
    }
    let finalTotalDuties = this.attendanceWorksheet.getRange(1, 8, this.databaseWorksheet.getLastRow(), 1).getValues();
    for (let prefectIndex = 0; prefectIndex < this.prefects.length; prefectIndex++) {
      this.prefects[prefectIndex].totalDuties = finalTotalDuties[this.prefects[prefectIndex].number - 1][0];
    }
    this.checkboxes = this.attendanceWorksheet.getRange(firstPrefectRow, 2, this.prefectsAttendanceLength, 5).getDataValidations();
  }
  findElligible(duty) {
    //filter and return prefects who meet requirements and are available
    return this.prefects.filter(prefect => {
      for (let i = 0; i < 4; i++) {
        //if requirements are not met, return false
        if (!(duty.requirements[i].includes(prefect.data[i]) || duty.requirements[i] == "*")) {
          return false;
        }
      }
      return !(prefect.unavailableDays.includes(duty.day));
      });
  }
  findGoodElligible(duty) {
    return this.prefects.filter(prefect => {
      for (let i = 0; i < 4; i++) {
        //if requirements are not met, return false
        if (!(duty.requirements[i].includes(prefect.data[i]) || duty.requirements[i] == "*")) {
          return false;
        }
      }
      if (duty.id != "EX") {
        return !(prefect.unavailableDays.includes(duty.day) || prefect.adjacentDutyDays.includes(duty.day) || (prefect.pastDuties.substring(prefect.pastDuties.length - 3).includes(duty.id)) || prefect.dutiesThisWeek > 1);
      }
      else {
        return !(prefect.unavailableDays.includes(duty.day) || prefect.adjacentDutyDays.includes(duty.day) || prefect.dutiesThisWeek > 1);
      }});
  }
  placeCheckboxes() {
    this.attendanceWorksheet.getRange(firstPrefectRow, 2, this.prefectsAttendanceLength, 5).setDataValidations(this.checkboxes);
  }
  updateDatabase() {
    let newPastDutiesCol = this.fullData.map(row => [row[pastDutiesCol-1]]);
    for (const prefect of this.prefects) {
      newPastDutiesCol[prefect.number - 1][0] = prefect.pastDuties;
    }
    this.databaseWorksheet.getRange(1, pastDutiesCol, newPastDutiesCol.length, 1).setValues(newPastDutiesCol);
  }
  processData() {
    let fullData = this.fullData;
    //given raw data from database, convert to list of objects (prefects) 
    //loop through rows of raw data
    let prefects = []
    for (let row = 1; row < fullData.length; row++) {
      let prefect = {number: row + 1, name: fullData[row][nameCol-1].trim(), fullName: fullData[row][fullNameCol-1].trim(), data: fullData[row][dataCol-1].trim(), pastDuties: fullData[row][pastDutiesCol-1], unavailableDays: fullData[row][unavailableDaysCol-1].toString().trim().split("").map(Number), adjacentDutyDays: [], dutiesThisWeek: 0, totalDuties: 0, dutyDays: []};
      if (prefect.name != "") {
        if (prefect.data.length < 4) {
          console.warn(`Invalid data for prefect "${prefect.fullName}": data must be at least 4 characters (seen ${prefect.data})`);
        }
        prefects.push(prefect);
      }
    }
    return prefects;
  }
} 
