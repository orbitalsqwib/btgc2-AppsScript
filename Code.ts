import "google-apps-script";

// ===================================================================
// +=-=+=-=+ Bid-to-Timing Generator for Cookhouse 2 (BTGC2) +=-=+=-=+
// ===================================================================
// Source code can be found at:
// https://github.com/orbitalsqwib/btgc2-AppsScript
// ===================================================================

// -----(1)------ setup enums and interfaces

// 1-indexed
enum TimeslotRows {
  "0530 Collection Time" = 4,
  "0750 Collection Time" = 5,
  "0530 - 0610" = 8,
  "0610 - 0650" = 9,
  "0650 - 0730" = 10,
  "1130 Collection Time" = 12,
  "1350 Collection Time" = 13,
  "1130 - 1210" = 16,
  "1210 - 1250" = 17,
  "1250 - 1330" = 18,
  "1730 Collection Time" = 20,
  "1950 Collection Time" = 21,
  "1730 - 1810" = 24,
  "1810 - 1850" = 25,
  "1850 - 1930" = 26,
  "2030 - 2050" = 28,
  "2050 - 2110" = 29,
  "2110 - 2130" = 30,
}

// 1-indexed
enum SheetColumn {
  "Timestamp" = 1,
  "ContactNo",
  "Company",
  "EatingStrength",
  "Remarks",
  "Week",
  "TimeslotStart",
}

interface Bid {
  info: string;
  timeslot: string;
}

interface Day {
  bids: Bid[];
}

type Week = Day[];
const Week = (len: number) =>
  Array.from({ length: len }, () => ({
    bids: [],
  })) as Week;

// -----(2)------ define constants

// apps script constants
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const ui = SpreadsheetApp.getUi();

// defines the number of timeslot columns there are in a day
const dayTimeslotCols = 7;
const templateCols = 35 + 2;

// -----(3)------ define utility functions

// merges a source week into a target week
const mergeWeeks = (source: Week, target: Week) => {
  // add bids for existing days in target from source
  for (let day = 0; day < target.length; day++) {
    target[day].bids = target[day].bids.concat(source[day].bids);
  }
};

// parses bids for a given week from a bidding entry
const parseWeekBidsFromEntry: (entry: any[]) => Week = (entry) => {
  // initialise working vars
  let week = Week(7);
  const companyName = entry[SheetColumn["Company"] - 1];
  const eatingStrength = entry[SheetColumn["EatingStrength"] - 1];
  const scanStartIndex = SheetColumn["TimeslotStart"] - 1;

  // loop through entry records beginning from first timeslot
  for (let i = scanStartIndex; i < entry.length; i++) {
    // get array index for current day based on position in bidding entry
    const dayIndex = Math.floor((i - scanStartIndex) / dayTimeslotCols);

    // if an entry exists for the timeslot
    if (entry[i] != "")
      // add bid to week for the appropriate day
      week[dayIndex].bids.push({
        info: companyName + " (" + eatingStrength + ")",
        timeslot: entry[i],
      });
  }

  // return generated bids
  return week;
};

// parses a given DDMMYY string to a Date object
const parseDDMMYYToDate = (text: string) => {
  const day = parseInt(text.slice(0, 2));
  const mth = parseInt(text.slice(2, 4));
  const yr = parseInt(text.slice(4)) + 2000;
  return new Date(yr, mth - 1, day);
};

// -----(4)------ define menu functions

// generates a timing sheet based on bidding entries in "Form Responses"
const generateTimingSheetFromBids = () => {
  // initialize sheets
  const inputSheet = spreadsheet.getSheetByName("Form Responses");

  // guard inputSheet must exist
  if (!inputSheet) {
    spreadsheet.insertSheet("Form Responses");
    ui.alert(
      '[ERROR]: No "Form Responses" sheet found! Generating input sheet.'
    );
  }

  // delete and re-clone cookhouse timings output sheet
  if (spreadsheet.getSheetByName("Cookhouse Timings"))
    spreadsheet.deleteSheet(spreadsheet.getSheetByName("Cookhouse Timings")!);

  const outputSheet = spreadsheet
    .getSheetByName("Cookhouse Timings Template")!
    .copyTo(spreadsheet)
    .setName("Cookhouse Timings")
    .showSheet();

  // generate new week dictionary
  let weeksMap = new Map<string, Week>();

  // read all bidded timings from Form Responses
  const inputValues = inputSheet?.getDataRange().getValues().slice(1);
  if (inputValues == undefined) {
    SpreadsheetApp.getUi().alert("No timings in Form Responses!");
    return;
  }

  // for all bid entries,
  for (let entryRow of inputValues) {
    // generate bids from each entry
    let weekBids = parseWeekBidsFromEntry(entryRow);

    // if week exists in week dictionary, merge new bids into existing week
    let weekname = entryRow[SheetColumn["Week"] - 1];
    if (weeksMap.has(weekname)) mergeWeeks(weekBids, weeksMap.get(weekname)!);
    // else, add new week to week dictionary
    else weeksMap.set(weekname, weekBids);
  }

  // sort by weekname and split into days
  const sortedWeeks = Array.from(weeksMap.entries()).sort((a, b) =>
    a[0] > b[0] ? 1 : -1
  );
  const sortedDays = sortedWeeks.map((week) => week[1]).flat();

  // init day counterfor timings sheet
  const currentDay = parseDDMMYYToDate(sortedWeeks[0][0].slice(0, 6));

  // create a mapping of a day's timeslots to the companies inhabiting the timeslots
  let coyTimeslotMap = new Map<string, string>();

  // for all sorted days
  for (let i = 0; i < sortedDays.length; i++) {
    // calculate output column for day
    const col = 3 + i * 2;

    // re-initialize the coy-timeslot map
    coyTimeslotMap.clear();

    // for all bids of the day
    for (let bid of sortedDays[i].bids) {
      // compile all coy bids per timeslot
      coyTimeslotMap.set(
        bid.timeslot,
        (coyTimeslotMap.get(bid.timeslot)
          ? coyTimeslotMap.get(bid.timeslot) + "\n"
          : "") + bid.info
      );
    }

    // transcribe all compiled timeslot data for the day onto the output sheet
    Array.from(coyTimeslotMap.entries()).forEach(([timeslot, info]) => {
      outputSheet.getRange(TimeslotRows[timeslot], col).setValue(info);
    });

    // update day header for column
    outputSheet
      .getRange(1, col)
      .setValue(Utilities.formatDate(currentDay, "UTC+8", "EEEE (ddMMyy)"));

    // increment day counter for next day
    currentDay.setDate(currentDay.getDate() + 1);
  }

  // hide trailing unused columns on template
  const firstEmptyCol = 3 + sortedDays.length * 2;
  for (let col = firstEmptyCol; col <= outputSheet.getLastColumn(); col++)
    outputSheet.hideColumns(col);

  // finally, set output sheet as active sheet
  spreadsheet.setActiveSheet(outputSheet);
};

// -----(5)------ define spreadsheet ui functions

// onOpen trigger for spreadsheet apps
const onOpen = () => {
  ui.createMenu("BTGC2 Functions")
    .addItem("Generate Timing Sheet From Bids", "generateTimingSheetFromBids")
    .addToUi();
};
