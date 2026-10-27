// script.js

/**
 * The xl2cal system converts an event schedule (typically, the Whitstable
 * Yacht Club) racing schedule from an Excel spreadsheet to .ics format 
 * that can be loaded into a calendar.
 *
 * Details of usage and the requirements of the spreadsheet can be found in
 * instructions.html.
 * In general, xl2cal attempts to be as lenient as possible with spreadsheet
 * data such as dates and times, and permits a variety of formats for these.
 * See comments below about peculiarities of Excel formats, etc.
 *
 * @author Richard Jones
 * @copyright Richard Jones, 2023
 * https://github.com/rejones/wyc/web
 */

/*
Copyright 2023-present Richard Jones

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

'use strict';

const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Weds'];
const monthNames = ['January','February','March','April','May','June','July','August',
          'September','October','November','December'];

/** 
 * Month prefixes to numbers
 * Month numbers are strings to help with padding
 */
const months = new Map([
  ['JAN', '1'],
  ['FEB', '2'],
  ['MAR', '3'],
  ['APR', '4'],
  ['MAY', '5'],
  ['JUN', '6'],
  ['JUL', '7'],
  ['AUG', '8'],
  ['SEP', '9'],
  ['OCT', '10'],
  ['NOV', '11'],
  ['DEC', '12']
]);

/** 
 * Month prefixes to full names
 */
const fullMonths = new Map([
  ['JAN', 'JANUARY'],
  ['FEB', 'FEBRUARY'],
  ['MAR', 'MARCH'],
  ['APR', 'APRIL'],
  ['MAY', 'MAY'],
  ['JUN', 'JUNE'],
  ['JUL', 'JULY'],
  ['AUG', 'AUGUST'],
  ['SEP', 'SEPTEMBER'],
  ['OCT', 'OCTOBER'],
  ['NOV', 'NOVEMBER'],
  ['DEC', 'DECEMBER']
]);

const MONTHS_PREFIX_LEN = 3;

/**
 * Hash of column names to numbers.
 *  Invariants: mappings are unique, i.e. no two keys may have the same value
 *              cannot have both CDAY or CMONTH and CDATE
 */
const columns = new Map();
const CDAY = 'Day no.';
const CMONTH = 'Month';
const CDATE = 'Date';
const CHW = 'HW';
const CSTART = 'Start';
const CEND = 'End';
const CDURATION = 'Duration';
const CEVENT = 'Event';
const CCALENDAR = 'Calendar';
const columnValues = [CDAY, CMONTH, CDATE, CEVENT, CSTART, CEND, CDURATION, CHW, CCALENDAR].sort();
const columnTips = new Map([
    [CDAY,      "Number of the day of the event (e.g. '1' or '1st').\n" +
                `MUST specify \'${CDAY}\' or \'${CDATE}\'.`],
    [CMONTH,    "Month: either number, name or prefix of name.\n" +
                `MUST specify \'${CMONTH}\' or \'${CDATE}\'.`],
    [CDATE,     "Date, e.g 'Sat 2 September 2023'.\n" +
                "Excel converts other dates to this format." +
                `MUST specify \'${CDATE}\' or \'${CDAY}\' and \'${CMONTH}\'.`],
    [CEVENT,    "MUST specify the name of the event" +
                `If \'${CHW}\' isspecified, that field is appended to the \'${CEVENT}\' field`],
    [CSTART,    "MUST specify start time of the event (e.g. '11:00', '11.00' or '1000')"],
    [CEND,      "OPTION: end time of the event (e.g. '11:00', '11.00' or '1000')\n" +
                "If not specified, the duration is used"],
    [CHW,       "OPTION: High water time, appended to the event\n" +
                "If not be specified, or event does not use, nothing is added."],
    [CDURATION, "OPTION: duration of the event (e.g. '3' or '2:30')\n" +
                "If not specified, the default duration is used"],
    [CCALENDAR, "OPTION: this column allows different calendars\n" +
                "You can then select which calendar(s) to export"],
  ]);


// Constants
const VERSION = '2.0'; // for iCalendar
const H = 'H';         // needed for the DEFAULT_DURATION value
const ALERT = 2;     // 2 hours warning

// To Be Announced times
const TBAhour = 0;             // so set the period to be 00.00 - 23.00
const TBAend_hour = 23;
const TBAstring = ' (times TBC)'; // string to add to TBA events
// Not Applicable times
//const NAhour = 0;              // so set the period to be 00.00 - 23.00
//const NAend_hour = 23;

const T = 'T';
const Z = 'Z';      // 'Z' means UTC rather than local time

// Default durations
const DEFAULT_DURATION = '2:00';    // hours for a single race 
const DEFAULT_DEFAULT_DURATION_HOUR = 2;
const DEFAULT_DEFAULT_EXTENSION_HOUR = 1;
const DEFAULT_DEFAULT_DURATION_MIN = 0;
let DEFAULT_DURATION_HOUR = DEFAULT_DEFAULT_DURATION_HOUR;
let DEFAULT_EXTENSION_HOUR = DEFAULT_DEFAULT_EXTENSION_HOUR;
let DEFAULT_DURATION_MIN = DEFAULT_DEFAULT_DURATION_MIN;

// Regular expressions
const TBA_PATT = /^(TBA|TBC|N\/?A|-)$/i;                   // TBA, tbc, N/A etc 
//const NA_PATT = /^(N\/?A|-$)/i;                      // N/A or NA Could strengthen with ^$
const TIME_PATT = /^(\d\d?)[:.h]?(\d\d)(?:[:.]?\d\d)?\s*(am|pm)?\.?$/i;                   // A time such as 9:30, 12.15, 2h00 or 1100 and optional seconds
const DURATION_PATT = /^(\d\d?)(?:h|[h:.](\d\d)(?:[:.]?\d\d)?)?$/;        // A duration as 1, 2:30, 3.15, 4h or 2h00 but not 1100
const ANY_TIME_PATT = /^(\d\d?[:.h]?\d\d(?:[:.]?\d\d)?\s*(?:am|pm)?|TB[AC]|N\/?A|-)$/i;  // A time or TBA or N/A etc 
const DATE_PATT = /^(\d\d?)[-\/. ](\d\d?)[-\/. ](\d{2,4})$/;   // A date such as 17/10/2026 or 1.6.26 
const REV_DATE_PATT = /^(\d\d\d\d)[-\/. ](\d\d?)[-\/. ](\d\d)$/;   // A date such as 2026/10/19 or 26.6.19 
const ORDINAL_DAY_PATT = /^(\d\d?)(st|nd|rd|th)?$/i; // Ordinal day eg 12th
const INTEGER_PATT = /^\d+$/;                        // An integer


// Spreadsheet data
const MAX_SPREADSHEET_SIZE = 1024 * 1024;// Maximum spreadsheet size
const MAX_COLUMNS = 60;                  // Maximum number of columns for decent performance
const MAX_ROWS = 1000;                   // Maximum number of rows for decent performance
const JANUARY_1_2020 = 43480;            // Excel value for 1/1/2020
let theSheet = 0;                        // The sheet to read
const theData = [];                      // The data from the spreadsheet
let calendarsFound;                      // Calendars found in the spreadsheet
let theLineNumber;                       // The number of the current line in the spreadsheet.

/** Export button for spreadsheet should only be created once */
let hasExportBtn = false;

/** Modal dialog */
let modal;


// If user cancels the run..
class AbortError extends Error {
  constructor(message) {
    super(message);
    this.name = "AbortError";
  }
}

// Capture all otherwise uncaught errors
window.onerror = function (evt, source, lineno, colno, error) {
  if (error instanceof TypeError) {
      console.trace(evt, source, lineno, colno, error);
      alert(`Line  ${1+lineno}: UNEXPECTED TypeError.`);
  } 
  if (error instanceof ReferenceError) {
      console.trace(evt, source, lineno, colno, error);
      alert(`Line  ${1+lineno}: ReferenceError. Please report.`);
  }
  else if (!(error instanceof AbortError)) {
    console.warn(`Line ${1+lineno}: UNEXPECTED ERROR\n`, evt, error);
    alert(`Line ${1+lineno}: UNEXPECTED ERROR. Please report.`);
  }
}


/**
 * Report errors
 * @param e the error
 * @throws e if not otherwise caught
 */
function reportError(e) {
  if (e instanceof AbortError) {
    console.warn(e.message);
    alert(e.message);          // a deliberate, user-readable stop
  } 
  else {
    throw e;                   // real bugs still reach window.onerror
  }
}


const OKCancel = '\n(Press OK to continue or Cancel to abort)';
/**
 * Warn user and log
 * @param msg The alert message
 * @return true if user wants to continue, else false
 */
function warnUserOrCancel(msg) {
  console.warn(msg);
  //alert(msg);
  return confirm(msg + OKCancel);
}


/**
 * Warn about bad record in the spreadsheet
 * @param msg The error message
 * @param lno The number of the line on which the error was found
 * @param evt The name of the event
 */
function bad(msg, lno, evt) {
  const message = evt ? 
      `Line ${lno}: ${msg} so ignoring this line.\n\t${evt}` :
      `Line ${lno}: ${msg} so ignoring this line.` ;
  console.warn(message);
  console.trace(message);
  if (confirm(message)) {
    return;
  } else {
    alert("You have cancelled the run");
    throw new AbortError('Run aborted');
  }
}


crypto.randomUUID(); // Check early for https: (or file:) not http:

/** 
 * Add listener on events to 
 * (1) choose the spreadsheet
 * (2) generate a dropdown list of years, centred on current year
 * (3) get the name or number of the sheet to use
 * (4) get a prefix to use for events
 */
window.addEventListener('DOMContentLoaded', () => {
  const fileChooser = document.getElementById('file-chooser');
  const yearDropdown = document.getElementById('year-dropdown');
  const dz = document.getElementById('drop-zone');
  
  fileChooser.addEventListener('change', loadFileEvent);
  dz.addEventListener('drop', dropHandler);
  dz.addEventListener('dragover', dragOverHandler);

  // Generate the range of years
  const currentYear = new Date().getFullYear();
  const startYear = currentYear - 5;
  const endYear = currentYear + 10;

  // Set these years in a dropdown list, with currentYear selected
  for (let year = startYear; year <= endYear; year++) {
    const option = document.createElement('option');
    option.value = year;
    option.textContent = year;
    option.selected = year == currentYear+1;
    yearDropdown.appendChild(option);
  }

  // Allow user to choose a default duration for an event
  const durationBox = document.getElementById('duration');
  durationBox.addEventListener('change', () => {
    const duration = durationBox.value;
    let matchTime;
    if (duration) {
      matchTime = duration.match(DURATION_PATT); 
    }
    if (matchTime) {
      DEFAULT_DURATION_HOUR = +matchTime[1];
      DEFAULT_DURATION_MIN =  matchTime[2] ? +(matchTime[2].slice(1)) : 0; // remove [:h]
    } 
    else {
      durationBox.value = DEFAULT_DURATION;
    }
    if ((DEFAULT_DURATION_HOUR > 23) || (DEFAULT_DURATION_MIN > 59)) {
      durationBox.value = DEFAULT_DURATION;
      DEFAULT_DURATION_HOUR = DEFAULT_DEFAULT_DURATION_HOUR;
      DEFAULT_DURATION_MIN = DEFAULT_DEFAULT_DURATION_MIN;
    }
  });
});


/**
 * Handle file dropped on landing site
 * @param {Event} ev - The event
 */
function dropHandler(ev) {
  // Prevent default behaviour (prevent file from being opened)
  ev.preventDefault();
  const item = ev.dataTransfer.items ?
    ev.dataTransfer.items[0] :
    ev.dataTransfer.files[0];
  if (item.kind === "file") {
    const file = item.getAsFile();
    if (!file)
      return;
    if (!/\.xlsx$/i.test(file.name)) { 
      alert('Please choose an .xlsx spreadsheet.'); 
      return;
    }  
    // set file.name in the fileChooser
    document.getElementById('file-chooser').files = ev.dataTransfer.files;
    loadFile(file);
  }
}

/**
 * Handle file dragged onto landing site
 * @param {Event} ev - The event
 */
function dragOverHandler(ev) {
  // Prevent default behaviour (prevent file from being opened)
  ev.preventDefault();
}


/**
 * Read the spreadsheet file
 * @param {Event} event - The event
 */
function loadFileEvent(event) {
  const file = event.target.files[0];
  if (!file)
    return;
  loadFile(file);
}


/**
 * Check that sheet dimensions are reasonable
 * @param sheet the sheet
 * @param name the name of the sheet
 * @throws Error if too many rows or columns
 */
function checkSheetSize(sheet, name) {
  if (!sheet || !sheet['!ref'])        // empty sheet
    return;

  // range is an object with start (s) and end (e) properties for rows (r) and columns (c),
  // all zero-indexed
  const range = XLSX.utils.decode_range(sheet['!ref']);
  if (range.e.c + 1> MAX_COLUMNS) 
    throw new AbortError(`Sheet "${name}" has too many columns (${range.e.c + 1}).`);
  if (range.e.r + 1> MAX_ROWS) 
    throw new AbortError(`Sheet "${name}" has too many rows (${range.e.r + 1}).`);
}


/**
 * Read the spreadsheet and render it in the document
 * @param {File} file - The file loaded
 */
function loadFile(file) {
  //console.log(file.name);
  if (file.size >  MAX_SPREADSHEET_SIZE) {
    alert(`${file.name} is too large.`);
    return;
  }

  const reader = new FileReader();
  reader.onerror = () => alert('Could not read the file.');
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
  
      let workbook;
      try {
        workbook = XLSX.read(data, { cellStyles: true, type: 'array', dateNF: 'd mmmm yyyy' });
      }
      catch (e) {
        throw new AbortError(`${file.name} could not be read as an .xlsx spreadsheet.`);
      }
  
      let sheetName;
      let jsonData;
      const options = { skipHidden: true, raw: true, header: 1 };  // was raw:false
  
      if (workbook.SheetNames.length == 1) {
        // only one sheet
        sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
  
        checkSheetSize(sheet, sheetName); // Check that this sheet is a reasonable size
        jsonData = XLSX.utils.sheet_to_json(sheet, options);
        replaceSheetGroup();
        addSpreadsheetHeading();
        const sheetChooser = document.getElementById("sheet-names");
        sheetChooser.replaceChildren();
        renderTable(jsonData);
      } 
      else { 
        // choose sheet from drop-down list
        replaceSheetGroup();
        addSpreadsheetHeading();
        const select = document.createElement("select");
        select.id = "select-sheet";
  
        workbook.SheetNames.forEach(function(sh) {
          const option = document.createElement("option");
          option.value = sh;
          option.text = sh;
          select.appendChild(option);       
        });
  
        select.addEventListener('change', (event) => {
          try {
            sheetName = event.target.value;
            const sheet = workbook.Sheets[sheetName];
            checkSheetSize(sheet, sheetName); // Check that this sheet is a reasonable size
            jsonData = XLSX.utils.sheet_to_json(sheet, options);
            renderTable(jsonData);
          }
          catch (e) {
            reportError(e);
          }
        });
  
        const div = document.createElement("div");
        div.style.display = "flex";
        const label = document.createElement("label");
        label.innerHTML = 'Select sheet to use:';
        label.htmlFor = "select-sheet";
        div.appendChild(label);
        div.appendChild(select);
        const sheetChooser = document.getElementById("sheet-names");
        sheetChooser.replaceChildren(div);   // Replace any previous div
  
        // Assume the first sheet initially
        sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        checkSheetSize(sheet, sheetName); // Check that this sheet is a reasonable size
        jsonData = XLSX.utils.sheet_to_json(sheet, options);
        // Clear any previous column selections
        renderTable(jsonData);
      }
    }
    catch (e) {
      reportError(e);
    }
  };
  
  reader.readAsArrayBuffer(file);
  return true;
}


/**
 * Clear any previous children of the sheet-group
 */
function replaceSheetGroup() {
  const sheetGroup = document.getElementById("sheet-group");
  const sheetHdr = document.createElement("h2");
  sheetHdr.id = "sheet-header";
  const sheetNamesDiv = document.createElement("div");
  sheetNamesDiv.id = "sheet-names";
  const colInstrs = document.createElement("p");
  colInstrs.id = "columns-instructions";
  const sheetTbl = document.createElement("table");
  sheetTbl.id = "sheet-table";
  sheetGroup.replaceChildren(sheetHdr, sheetNamesDiv, colInstrs, sheetTbl);
}


/**
 * Add spreadsheet section title and instructions
 */
function addSpreadsheetHeading() {
  const sheetHeader = document.getElementById('sheet-header');
  sheetHeader.innerHTML = "2. The spreadsheet";
  const instructions = document.getElementById('columns-instructions');
  instructions.innerHTML = 
      "<em>Select the columns to use from the dropdown lists in the first row</em>";
}


/**
 * Parse and render the spreadsheet on screen as a table
 * Note: use DOM calls rather than string to guard against HTML injection
 * @param {Array} data - The sheet to parse and display
 */
function renderTable(data) {

  // display number with 2 digits, prepending `0` if necessary 
  const f = (n) => n.toString().padStart(2, "0");

  const sheetTable = document.getElementById('sheet-table');
  const numColumns = data.map(row => row.length)
                     .reduce( (a,b) => Math.max(a,b), 0);
  
  const tbl = document.createDocumentFragment();
  const headRow = document.createElement('tr');
  for (let i = 0; i < numColumns; i++) {
    const td = document.createElement('td');
    const sel = document.createElement('select');
    sel.className = 'column-select';
    sel.add(new Option('...', 'empty'));
    for (const option of columnValues) {
      const o = new Option(option, option);
      o.title = columnTips.get(option);
      sel.add(o);
    }
    td.appendChild(sel);
    headRow.appendChild(td);
  }
  tbl.appendChild(headRow); 
  
  // Process the data, also saving it as strings
  theData.length = 0; // clear in case user has added the spreadsheet more than once
  for (const row of data) {
    const tr = document.createElement('tr');
    let newRow = [];
    
    //console.log(row);
    for (const cell of row) {
      //console.log(cell, cell%1, typeof cell);
      // Convert cell value to string
      let cellValue = cell == undefined ? '' : String(cell);
  
      // Unless 'cellStyles: true' is given as an option to XLSX.read(),
      // Excel serial dates and times are represented as a real number.
      // The integral part represents the date by the number of days since 1 Jan 1900;
      // the fractional part represents the time as a fraction of 24 hours.
      // With 'cellStyles: true', 
      // dates and times are exported as Date objects.
      // However, this seems very fragile as it is dependant of very precise
      // formatting in Excel.

      // not used anymore
      /*
      if (cell instanceof Date) {
        console.log(`Found an excelDate on row ${row}: ${cell}, ${cell%1}, ${typeof cell}`);
        // Try to tidy incomplete Excel values
        if (cell.getFullYear() >= 1900 && cell.getHours()==0 && cell.getMinutes()==0) { // Looks like a date
          cellValue = `${dayNames[cell.getDay()]} ${cell.getDate()} ${monthNames[cell.getMonth()]}`;
          const year = cell.getFullYear();
          if (year > 1904) { // Otherwise, user probably didn't specify year so likely to be 1899 or 1903
            cellValue = cellValue + ` ${year}`;
          }
        }
        else { // Probably a time
          cellValue = `${f(cell.getHours())}:${f(cell.getMinutes())}`;
        }
      }
      */
      if (cell instanceof Date) {
        //console.log(`Found an excelDate on row ${row}: ${cell}, ${cell%1}, ${typeof cell}`);
        // SheetJS can be a fraction of a millisecond short (10:00 comes out as 09:59:59.999),
        // so round to the nearest minute before reading the hours and minutes
        const d = new Date(Math.round(cell.getTime() / 60000) * 60000);
        if (d.getFullYear() >= 1900) { // Has a date part (any time part is dropped)
          cellValue = `${dayNames[d.getDay()]} ${d.getDate()} ${monthNames[d.getMonth()]}`;
          if (d.getFullYear() > 1904) { // Otherwise, user probably didn't specify year
            cellValue = cellValue + ` ${d.getFullYear()}`;
          }
        }
        else { // A time on its own, which SheetJS dates in 1899
          cellValue = `${f(d.getHours())}:${f(d.getMinutes())}`;
        }
      }

      // Check if the cell format is "hh:mm"
      else if (typeof cell === 'number') {
        //if (cell % 1 !== 0) { // A non-integer so assume this is a time
        if (cell % 1 !== 0 && cell < JANUARY_1_2020) { // A non-integer so assume this is a time
          
          // Steve Gray spreadsheet has times as decimal numbers, which is weird.
          // With the Claude changes, a duration cell of 2.5 is treated as 2h50, not 2h30
          if ((1 <= cell) && (cell < 24.00)) {
            const hours = Math.floor(cell);
            const minutes = Math.round((cell - hours) * 100);
            if (minutes < 60) { // looks like a time
              //cellValue = `${hours}.${f(minutes)}`;
              cellValue = `${hours}.${minutes}`;
            }
          }
          else {
            // Assume it's an Excel time
            /*
            const excelTime = cell % 1; 
            const hours = Math.floor(excelTime * 24);
            const minutes = Math.round(((excelTime * 24) % 1) * 60);
            */
            const totalMinutes = Math.round((cell % 1) * 1440);
            const hours = Math.floor(totalMinutes / 60) % 24;
            const minutes = totalMinutes % 60;
            cellValue = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
          }
        }
        else if (cell > JANUARY_1_2020) { // Assume it's a date after 1 Jan 2020
          //const date = new Date(Date.UTC(0, 0, cell-1));
          const date = new Date(Date.UTC(0, 0, Math.floor(cell-1)));
          if (!isNaN(date.getFullYear())) {
            cellValue = date.toDateString();
          }
        }
      }

      /*
      else {
        let matchDate;
        // Defend against regex DOS attacks, otherwise trailing spaces is polynomial
        //if (cell && (matchDate = cell.trim().match(/^(\d\d?)\/(\d\d?)\/(\d{2,4})$/))) {
        if (cell && (matchDate = cell.trim().match(DATE_PATT))) { 
          // Looks like a date.
          // Excel default format seems to be m/d/yy but we want to display and use d/m/yy
          cellValue = `${matchDate[2]}/${matchDate[1]}/${matchDate[3]}`;
        }
      }
      */
          
      const td = document.createElement('td');
      td.textContent = cellValue;        // never parsed as HTML
      tr.appendChild(td);
      newRow.push(cellValue);
    }

    tbl.appendChild(tr);
    theData.push(newRow);
  }

  sheetTable.replaceChildren(tbl);

  // Add horizontal and vertical sliders to the table
  const container = document.createElement('div');
  container.className = 'table-container';
  const wrapper = document.createElement('div');
  wrapper.className = 'table-wrapper';
  sheetTable.parentNode.insertBefore(container, sheetTable);
  container.appendChild(wrapper);
  wrapper.appendChild(sheetTable);

  // Add border style to the table
  sheetTable.style.border = '1px solid black';

  // Add event listener to each dropdown list
  const dropdowns = document.querySelectorAll('.column-select');
  dropdowns.forEach((dropdown, index) => {
    dropdown.addEventListener('change', (event) => {
      // clear any entry with this value in order to maintain unique mapping
      let key = getKeyByValue(columns, index);
      if (key) {
        columns.delete(key);
      }
      const selectedOption = event.target.value;
      if (selectedOption != 'empty') { 
        columns.set(selectedOption, index);
        // Disallow both Day+Month and Date
        if (selectedOption === CDATE) {
          columns.delete(CDAY);
          columns.delete(CMONTH);
        }
        else if ((selectedOption === CDAY) || (selectedOption === CMONTH)) {
          columns.delete(CDATE);
        }
      }
      
      // Clear any other dropdowns with the same value
      // Disallow both Day+Month and Date columns
      dropdowns.forEach((dd, i) => {
        if (i != index) {
          let val = dd.value;
          if (columns.get(val) != i) {
            dd.value = 'empty';
          }
          if ((selectedOption === CDATE) &&
              ((dd.value === CDAY) || (dd.value === CMONTH))) {
            dd.value = 'empty';
          }
          else if (((selectedOption === CDAY) || (selectedOption === CMONTH)) &&
                   (dd.value === CDATE)) {
            dd.value = 'empty';
          }
        }
      });

      //Enable export button once required columns are chosen
      const exportBtn = document.getElementById('export-button');
      exportBtn.disabled = !( 
                   ((columns.has(CDAY) && columns.has(CMONTH)) || columns.has(CDATE))
                    && columns.has(CSTART)
                    && columns.has(CEVENT)
                            );
    });
  });

  // Create export button unless already present
  if (!hasExportBtn) {
    const exportBtn = document.createElement('button');
    exportBtn.id = 'export-button';
    exportBtn.innerHTML = 'Export';
    exportBtn.disabled = true;
    exportBtn.addEventListener('click', exportCalendar);
    document.body.appendChild(exportBtn);
    hasExportBtn = true;
  } else {
    document.getElementById('export-button').disabled = true;
  }
}


/**
 * Find the (first) key with a specified value in a hashmap 
 * @param {Map} map 
 * @param value
 * @return The key or null if value is not present
 */
function getKeyByValue(map, value) {
  for (let key of map.keys()) {
    if (map.get(key) === value) {
      return key;
    }
  }
  return null;
}


/**
 * Search the spreadsheet for calendars, if any
 * Want only Calendar entries for event rows (not headers, etc)
 @ return {Set} The set of calendar names (possibly empty)
 */
function findCalendars() {
  const calendarsFound = new Set();
  if (columns.has(CCALENDAR)) {
    const startCol = columns.get(CSTART);
    const calCol = columns.get(CCALENDAR);
    for (let row of theData) {
      const start = row[startCol];
      if (!start)
        continue;
      // Replicate patterns for Start time in generateICal
      if ( start.match(TIME_PATT) ||  
           start.match(TBA_PATT) ||  
           start.match(NA_PATT) ) {  
        //console.log(row[columns.get(CEVENT)], row[calCol]);
        if (row[calCol]) {
          calendarsFound.add(row[calCol]);
        }
      }
    }
  }
  return calendarsFound;
}


/** Close modal dialog */
function cancelSelect() {
  //console.log('called cancel');
  modal.close();
}


/**
 * Export rows for selected calendars
 * 1. Popup modal with calendars to select
 * 2. Generate the iCalendar entries
 * 3. Open a window with a link for the export.
 *    This is a window rather than a modal to allow multiple iCalendar
 *    to be generated
 */
function exportSelect() {
  try {
    const calendarsToExport = new Set();
    if (columns.has(CCALENDAR) && (calendarsFound.size > 0)) {
      const checkedCalendars = document.querySelectorAll('input[name=calendar-checkbox]:checked');
      checkedCalendars.forEach(function(cal) {
        calendarsToExport.add(cal.value);
      });
      //setCalendarsToExport(calendarsToExport);
      modal.close();
    }
  
    // Generate iCal
    const iCal = generateICal(theData, calendarsToExport);
    if (iCal) {
      //console.log(iCal);
    
      //Open new window with iCal data
      openICalWindow(calendarsToExport, iCal);
    }
  } catch (e) {
    if (!(e instanceof AbortError)) 
      throw e; 
  }
}


/**
 * Export the calendar to export
 * NB. The required fields have been chosen
 */
function exportCalendar() {
  // First, search the spreadsheet for calendars, if any
  calendarsFound = findCalendars();
  console.log(calendarsFound, calendarsFound.size);
  if (calendarsFound.size > 0) {
    // If any found, ask which calendars to use
    selectCalendars(calendarsFound);
  }
  else {
    // else go straight to export all
    exportSelect();
  }
}


/**
 * Escape special characters to prevent break outs
 * @param s a string that may contain special characters
 * @return the string with characters replace by HTML codes
 */
function escapeHTML(s) {
  return String(s).replace(/[&<>"']/g,
    c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}


/**
 * Popup window to select calendars to export
 * @param calendarsFound - The calendars found in the CCALENDAR column
 */
function selectCalendars(calendarsFound) {
  if (!modal) modal = new SV.Modal('select-calendars-modal');
  /*
  modal.getModalElement().addEventListener('sv.modal.close', function (ev) {
    console.log('Closed the modal');
  });
  */
   
  // inject some content
  let html = `
    <p>Calendars found:</p>
  `;
  let n = 0;
  calendarsFound.forEach(function(cal) {
    const id = `cal-${n++}`;
    html += `
      <input type="checkbox" id="${id}" name="calendar-checkbox" value="${escapeHTML(cal)}" />
      <label for="${id}">${escapeHTML(cal)}</label><br>`;
  });
  html += `
    <p>If no calendars are selected, all will be exported.</p>
  `;
  html += `
      <button id="cancel-calendar-choice-button">Cancel</button>
      <button id="export-calendar-choice-button">Export</button>
  `;

  modal.inject(html, 'Select calendars to export');
  modal.getContentElement().querySelector('#cancel-calendar-choice-button')
    .addEventListener('click', cancelSelect);
  modal.getContentElement().querySelector('#export-calendar-choice-button')
    .addEventListener('click', exportSelect);
  modal.show();
}


/** 
 * Generate the DTSTAMP 
 * @return A DTSTAMP for now (the time this ical was created)
 */
function createDTSTAMP() {
  const now = new Date();
  const year = now.getUTCFullYear();
  const month = String(now.getUTCMonth() + 1).padStart(2, '0');
  const day = String(now.getUTCDate()).padStart(2, '0');
  const hours = String(now.getUTCHours()).padStart(2, '0');
  const minutes = String(now.getUTCMinutes()).padStart(2, '0');
  const seconds = String(now.getUTCSeconds()).padStart(2, '0');
  return `${year}${month}${day}${T}${hours}${minutes}${seconds}`;
}


/**
 * Is a string a month (or prefix)?
 * @param month The possible month string
 * @return True if so
 */
function isMonth(month) {
  month = month.replace(/[\.,\s]+$/, '').toUpperCase();
  return monthNames.map(m => m.toUpperCase()).find((m) => {
      return m.startsWith(month);
    }) != undefined;
}


/**
 * Is a string a weekday (or prefix)?
 * @param  day The possible weekday string
 * @return True if so
 */
function isWeekDay(day) {
  day = day.replace(/[\.,\s]+$/, '').toUpperCase();
  return dayNames.map(m => m.toUpperCase()).find((m) => {
      return m.startsWith(day);
    }) != undefined;
}


/**
 * Parse a day in a format like 'Sunday 12th March' or '12/3/'
 * @param rowDate The date
 * @return [day number, month, year], e.g. [12, 'March', 2023] or
 *         [day number, month] if no year given
 */
function parseDate(rowDate) {
  //console.log('parseDate', rowDate, typeof rowDate);
  if (!rowDate)
    return null;

  // First, try Excel Date format which is read as a number
  if (!isNaN(rowDate)) {
    // Try converting from Excel date to JS date
    console.log(`rowDate (${rowDate}) is a number (line ${theLineNumber}`);
    if (rowDate < JANUARY_1_2020) {
      if (!warnUserOrCancel(`Line ${theLineNumber}: "${rowDate}" does not look like a date.`)) {
        // TODO If user accepts, then they also get a year differs warning with a year interpreted
        // as an Excel year, e.g. "12" shown as "1900". This is misleading. 
        // For the moment, I have changed the year differs message.
        return null;
      }
    }
      
    const date = new Date(Date.UTC(0, 0, rowDate -1));
    console.log(date);
    if (isNaN(date.getFullYear())) {
      //console.warn(`parseDate returns null! for ${rowDate}`);
      return null;
    }
    return [date.getDate().toString(), date.getMonth() + 1, date.getFullYear()].map((e) => {return e.toString()});
  }
  
  // Next, try dd/mm/yy and yyyy/mm/dd
  let matchDate;
  if (matchDate = rowDate.match(DATE_PATT)) {
    const day = matchDate[1];
    const month = matchDate[2];
    let year = +matchDate[3];
    if ((0 <= year) && (year < 100)) {
      year = year + 2000;
    }
    if (isNaN(new Date(year, month, day).valueOf())) {
      console.warn('parseDate returns null!');
      return null;
    }
    return [day, month, year.toString()];
  }
  else if (matchDate = rowDate.match(REV_DATE_PATT)) {
    const day = matchDate[3];
    const month = matchDate[2];
    let year = +matchDate[1];
    if ((0 <= year) && (year < 100)) {
      year = year + 2000;
    }
    if (isNaN(new Date(year, month, day).valueOf())) {
      console.warn('parseDate returns null!');
      return null;
    }
    return [day, month, year.toString()];
  }


  //Next, try numbers and words

  // Define the regular expression for the delimiters (spaces and commas)
  const delimiterRE = /[-,\s]+/;
  const yearRE = /^\d{4}$/;

  // New version
  //  all of this to allow dates like 2 Jan
  const ordinalRE =  /\b\d\d?\s*(st|nd|rd|th)/ig;
  const ordinals = rowDate.match(ordinalRE) ?? [];      // get and patterns like "2 nd" into a list
  let remainder = rowDate.replace(ordinalRE, '')        // remove any ordinal string
                         .split(delimiterRE)            // split into a list of words
                         .filter(it => it !== '')       // remove any empty elements
                         .filter(x => !isWeekDay(x))    // remove any weekday names
                         .map(it => it.toUpperCase());  // make all upper case
                         // remainders contains month-names, cardinal numbers, junk
  console.log('ordinals ', ordinals);
  console.log('remainder ', remainder);

  // gather ordinal days
  let days = ordinals.map(t => t.match(/^\d\d?/)[0]);   // guaranteed that all elements match
  console.log('days', days);
  if (days.length > 1)
    return null;                                        // Too many days found 
  let theDay = (days.length == 1) ? days[0]             // we have the DAY
                                  : null;
  // if days.length == 0, No ordinal found but there may be a valid day number

  // look for a single month
  const months = remainder.filter(isMonth);              // get any months
  console.log('months', months);
  if (months.length != 1)                                // Either no month or more than 1!
    return null;
  const theMonth = months[0];                            // we have the MONTH

  remainder = remainder.filter(x => !isMonth(x));        // remove any months
  console.log('remainder without months', remainder);
  // remainders contains cardinal numbers, junk

  // look for a year as yyyy
  let theYear = null;
  let years = remainder.filter(y => y.match(/\b\d\d\d\d\b/));
  switch (years.length) {
    case 0: break;                                      // might be OK if we have eg 12 Jan 26
    case 1: theYear = years[0];                         // we have the YEAR
            break;
    default: return null;                               // too many years
  }
  remainder = remainder.filter(y => !/\b\d\d\d\d\b/.test(y)); // Remove 4-digit years 
  // remainders contains non 4-digit cardinal numbers, junk

  // might have 2 1- or 2-digit numbers, e.g. from an input 10 Jan 26
  let numbers = remainder.filter(y => y.match(/\b\d\d?\b/));       // 1- or 2-digit numbers
  console.log('2-digit numbers', numbers);
  remainder = remainder.filter(y => !/\b\d\d?\b/.test(y));         // Remove 4-digit years
  console.log('remainder should be empty', remainder);
  if (remainder.length != 0)                                       // this must be junk
    return null;

  if (!theDay) {
    if (numbers.length != 0) 
      theDay = numbers.shift();              // if not already found, the first element left must be the day dd
    else
      return null;                           // no day found
  }
  if (!theYear) { 
    switch (numbers.length) {
      case 0: theYear = null;                // No year found - fix this later
      case 1: theYear = numbers.shift();     // if not already found, the second element must be the year yy
              theYear = +theYear + 2000;     // 20yy
              break;
      default: return null;                  // too many 2-digit numbers
    }
  }

  if (numbers.length != 0)                   // Should have used all the numbers (day, year if any)
    return null;

  return theYear ? [ theDay, theMonth, theYear ] 
                 : [ theDay, theMonth ];
  

  /*
  let dayMonthYear = rowDate.split(delimiterRE)
                     .filter(it => it !== '')              // remove any empty elements
                     .map(it => it.toUpperCase());         // make all upper case
                     //.filter(it => !isDay(it));          // remove any day names 
  console.log('dayMonthYear', dayMonthYear);
  const years = dayMonthYear
                .filter (it => yearRE.test(it));         // get any year
  console.log('years', years);
  const dayMonth = dayMonthYear
                   .filter(it => !(yearRE.test(it)));   // remove any year
  console.log('dayMonth', dayMonth);
  const days = dayMonth
               .filter(it => /^\d+(ST|ND|RD|TH)?$/.test(it))    // only day strings 
               .map(it => it.replace(/(ST|ND|RD|TH)$/,''));    // remove any ordinal characters
  console.log('days', days);
  const months = dayMonth.filter(isMonth);              // get any month(s)
  console.log('months',months);
  console.log(days[0], months[0], years[0]);
  if (days.length != 1 || months.length != 1) {                  // must have day and month numbers
    return null;  
  }
  // TODO Could check that any weekday given is correct.
  //console.log(years.length);
  return years.length ? 
    [days[0], months[0], years[0]] :
    [days[0], months[0]];
  */
}


/**
 * Get the time
 * @line The line
 * @time The time column value
 * @return The hour, minute and true if TBA/ NA or -,  or null if not a valid time
 */
function getTimes(line, time) {
  let matchTime;
  if (matchTime = time.match(TIME_PATT)) {
    //const theHour = matchTime[1].padStart(2, '0');
    //const theMin = matchTime[2].padStart(2, '0');;
    let theHour = matchTime[1];
    const theMin = matchTime[2] || 0;
    const am_pm = matchTime.slice(-1)[0];                      // the last element
    console.log('matchTime', matchTime, 'am_pm', am_pm);
    if (/^PM$/i.test(am_pm) && +theHour != 12) {               // convert to 24 hour clock
      theHour = +theHour + 12;
      if (theHour > 24) {                                      // was a bad time like 16.30pm
        return null;
      }
    }
    else if (/^AM$/i.test(am_pm)) {
      if (+theHour == 12) {                                    // 12.15am is 00:15
        theHour = 0;
      }
      else if (+theHour > 12) {                                // a bad time like 15.00am
        return null;
      }
    }
    return {H: +theHour, M: +theMin, TBA: false};          // +null returns 0
  }
  else if (time.match(TBA_PATT) || time === '-' || time === '') {
    //return {H: TBAhour, M: 0, TBA: true};           
    return {TBA: true};           
  }
  /*
  else if (time.match(NA_PATT)) { 
    return {H: NAhour, M: 0, TBA: true};
  } 
  */
  else return null;
}


/**
 * Get a duration
 * @line The line
 * @time The duration column value
 * @return The hour, minute and true if TBA/ NA or -,  or null if not a valid duration
 */
function getDuration(line, time) {
  let matchTime;
  if (matchTime = time.match(DURATION_PATT)) {
    const theHour = matchTime[1];
    const theMin = matchTime[2] || DEFAULT_DURATION_MIN;  
    console.log(`getDuration (line ${theLineNumber}): time=${time} H=${theHour} M=${theMin}`);
    return {H: +theHour, M: +theMin, TBA: false};
  }
  else if (time.match(TBA_PATT) || time === '-' || time === '') {
    return {H: DEFAULT_DURATION_HOUR, M: DEFAULT_DURATION_MIN, TBA: true};
  }
  /*
  else if (time.match(NA_PATT)) { 
    return {H: DEFAULT_DURATION_HOUR, M: DEFAULT_DURATION_MIN, TBA: true};
  } 
  */
  else return null;
}


/**
 * Get the date without warnings
 * @line The line
 * @theDefaultYear The default year chosen by the user
 * @return true if the line includes a date
 */
function hasDate(line, theDefaultYear) {
  let theDay, theMonth, theYear;

  if (columns.has(CDAY) && columns.has(CMONTH)) {
    theYear = theDefaultYear;
    const monthCol = columns.get(CMONTH); 
    const lineMonth = line[monthCol] || '';         // Guard against empty cell
    // allow either month number or string
    if (/^\d\d?$/.test(lineMonth)) {
      theMonth =  line[monthCol];
    }
    else {
      theMonth = months.get(lineMonth.toUpperCase().substr(0, MONTHS_PREFIX_LEN));
      if (!theMonth) {
        return false;
      }
      if (theMonth < 1 || theMonth > 12) {
        return false;
      }
    }

    // get the day number
    let matchDay;
    if (matchDay = (line[columns.get(CDAY)] || '').match(ORDINAL_DAY_PATT)) {  // Guard against empty cell
      theDay = matchDay[1];
    } else {
      return false; 
    }

  } else if (columns.has(CDATE)) {
    // parse a date
    const dateCol = columns.get(CDATE);
    if (!line[dateCol]) {
      return null;
    }
    const dayMonthYear = parseDate(line[dateCol]);
    //console.log(dayMonthYear);
    if (!dayMonthYear) {
      return false;
    }

    if (dayMonthYear.length === 2) {
      dayMonthYear.push(theDefaultYear);
    }

    theDay = dayMonthYear[0];
    //theMonth = dayMonthYear[1].match(/^\d+$/) ? 
    theMonth = dayMonthYear[1].match(INTEGER_PATT) ?
      dayMonthYear[1] :
      months.get(dayMonthYear[1].substr(0, 3));
    //console.log('theDay - theMonth', theDay, theMonth);
    theYear = dayMonthYear[2];
  }

  // Check that the date is valid
  if (new Date(theYear, theMonth-1, theDay).getMonth() != theMonth-1) {
    return false;
  }

  return true;
}


/**
 * Generate the iCalendar data
 * @param data The data
 * @param calendarsToExport The calendars selected for export
 * @return The iCalendar as text
 */
function generateICal(data, calendarsToExport) {
  const DTSTAMP = createDTSTAMP();
  const startCol = columns.get(CSTART);
  const eventCol = columns.get(CEVENT);
  const startPattern = /^\d{4}$|^\d{2}:\d{2}$/;

  const theDefaultYear = document.getElementById('year-dropdown').value;
  const DEFAULT_DURATION = +document.getElementById('duration').value;
  //console.log('DEFAULT_DURATION', DEFAULT_DURATION);
  let thePrefix = document.getElementById('event-prefix').value;

  // Print header for calendars
  let text = printCALhdr();
  if (thePrefix.length) {
    thePrefix += ' ';
  }

  // Process the source file
  for (let lineNum in data) {
    let line = data[lineNum];
    if (!line || !line.length) {
      continue;
    }

    theLineNumber = parseInt(lineNum) + 1; //The Excel line number, for error reporting
  
    // Defend against regex DOS attacks, otherwise trailing spaces is polynomial
    line = line.map(it => it.trim()); 

    // Warn use if there's no valid start
    const start = line[startCol];
    if ( !start || !start.match(ANY_TIME_PATT) ) {
      // Ignore lines like header lines where date and start are both invalid
      if (!hasDate(line, theDefaultYear)) {
        console.warn(`Ignoring line ${theLineNumber}: Valid date and/or start not specified\n\t${line}`);
        continue;
      }
      else {
        if (warnUserOrCancel(`Ignoring line ${theLineNumber}: Valid start time "${start}" not specified\n\t${line}`)) 
          continue;
        else
          return null;
      }
    }

    // Skip rows for unwanted calendars
    if (columns.has(CCALENDAR) && (calendarsToExport.size > 0)) {
      const theCal = line[columns.get(CCALENDAR)];
      if (!calendarsToExport.has(theCal)) {
        continue;
      }
    }

    // GET THE DAY AND MONTH

    let theDay, theMonth, theYear;

    if (columns.has(CDAY) && columns.has(CMONTH)) {
      theYear = theDefaultYear;
      const monthCol = columns.get(CMONTH); 
      const lineMonth = line[monthCol];
      if (!lineMonth) {         // Guard against empty cell
        bad(`EMPTY MONTH`, theLineNumber, line[eventCol]);
        continue;
      }
      // allow either month number or string
      if (/^\d\d?$/.test(lineMonth)) {
        theMonth =  line[monthCol];
        if (theMonth < 1 || theMonth > 12) {
          //bad(`Invalid month ${theMonth} ${lineMonth} ${line[monthCol]}`, theLineNumber, line[eventCol]);
          bad(`INVALID MONTH "${line[monthCol]}"`, theLineNumber, line[eventCol]);
          continue;
        }
      }
      else {
        if (!isMonth(lineMonth)) {
          bad(`BAD MONTH NAME "${lineMonth}"`, theLineNumber, line[eventCol]);
          continue;
        }
        theMonth = months.get(lineMonth.toUpperCase().substr(0, MONTHS_PREFIX_LEN));
        if (!theMonth) {
          bad(`CANNOT UNDERSTAND MONTH "${lineMonth}"`, theLineNumber, line[eventCol]);
          continue;
        }
        if (theMonth < 1 || theMonth > 12) {
          //bad(`Invalid month ${theMonth} ${lineMonth} ${line[monthCol]}`, theLineNumber, line[eventCol]);
          bad(`INVALID MONTH "${line[monthCol]}"`, theLineNumber, line[eventCol]);
          continue;
        }
      }

      // GET THE DAY NUMBER
      //theDay = line[columns.get(CDAY)];
      let matchDay;
      if (matchDay = (line[columns.get(CDAY)] || '').match(ORDINAL_DAY_PATT)) {  // Guard against empty cell
        theDay = matchDay[1];
      } else {
        if (theDay)
          bad(`CANNOT UNDERSTAND DAY NUMBER "${theDay || ''}"`, theLineNumber, line[eventCol]); 
        else
          bad(`CANNOT UNDERSTAND DAY NUMBER`, theLineNumber, line[eventCol]); 
        continue; 
      }

    } 

    else if (columns.has(CDATE)) {
      // parse a date
      const dateCol = columns.get(CDATE);
      if (!line[dateCol]) {
        bad(`EMPTY DATE "${line[dateCol]}"`, theLineNumber, line[eventCol]);
        continue;
      }
      const dayMonthYear = parseDate(line[dateCol]);
      if (!dayMonthYear) {
        bad(`BAD DATE "${line[dateCol]}"`, theLineNumber, line[eventCol]);
        continue;
      }
      //console.log(dayMonthYear);
  
      // check that any year value matches the year chosen 
      if ((dayMonthYear.length === 3) &&
          (dayMonthYear[2] != theDefaultYear)) {
        if (!warnUserOrCancel(`Line ${theLineNumber}: Year "${line[dateCol]}" differs from ${theDefaultYear}\nIs this what you meant?`)) {
            //text = null;
            return null;
        }
      }
      else if (dayMonthYear.length === 2) {
        dayMonthYear.push(theDefaultYear);
      }
  
      theDay = dayMonthYear[0];
      theMonth = dayMonthYear[1].match(INTEGER_PATT) ?  
        dayMonthYear[1] :
        months.get(dayMonthYear[1].substr(0, 3));
      if (theMonth < 1 || theMonth > 12) {
        bad(`INVALID MONTH "${theMonth}"`, theLineNumber, line[eventCol]);
        continue;
      }
      //console.log('theDay - theMonth', theDay, theMonth);
      theYear = dayMonthYear[2];
    }

    // Check that the date is valid 
    if (new Date(theYear, theMonth-1, theDay).getMonth() != theMonth-1) {
      console.log(new Date(theYear, theMonth-1, theDay).getMonth());
      console.log(`theYear ${theYear}, theMonth-1 ${theMonth-1}, theDay ${theDay}`);
      bad(`NOT A VALID DATE`, theLineNumber, line[eventCol], line[eventCol]);
      continue;
    }
    
    // GET THE EVENT. 

    // Discard any lines with no event.
    let theEvent = line[columns.get(CEVENT)];
    if ( !theEvent ) {
      console.warn(`Ignoring line ${theLineNumber} ${line}: no event specified`);
      if(!warnUserOrCancel(`Ignoring line ${theLineNumber}: no event specified\n\t${line}`)) {
        return null;
      }
      continue;
    }

    // GET THE START AND END TIMES

    // times before 1000 are sometimes recorded with only 3 digits
    let theHour;
    let theMin;
    let theEndHour;
    let theEndMin;

    // Must have a Start time
    const theTimes = getTimes(line, start);
    if (!theTimes) {
      bad(`CANNOT UNDERSTAND START TIME "${line[columns.get(CSTART)]}"`, theLineNumber, line[eventCol]); 
      continue; 
    }

    /*
    switch (theHour) {
      case TBAhour: theEndHour = TBAend_hour;     
                    theEndMin = 0;
                    theEvent = theEvent + TBAstring;
                    break;
      case NAhour: theEndHour = NAend_hour;
                   theEndMin = 0;
                   break;
      default: theEndHour = +theHour + +DEFAULT_DURATION_HOUR;
               theEndMin = +DEFAULT_DURATION_MIN;
    }
      */

    if (theTimes.TBA) {
      theHour = TBAhour;
      theMin = 0;
      theEndHour = TBAend_hour;
      theEndMin = 0;
      theEvent = theEvent + TBAstring;
    }
    else {
      theHour = theTimes.H;
      theMin = theTimes.M;
      theEndHour = +theHour + +DEFAULT_DURATION_HOUR;
      theEndMin = +theMin + +DEFAULT_DURATION_MIN;
    }
    if (theHour > 23 || theMin >59) {
     bad(`INVALID START TIME "${line[columns.get(CSTART)]}"`, theLineNumber, line[eventCol]); 
      continue; 
    }


    // May have an End time
    let hasEnd = false;
    if (columns.has(CEND) && line[columns.get(CEND)]) {
      hasEnd = true;
      const endCol = columns.get(CEND);
      const theEndTimes = getTimes(line, line[endCol]);
      if (!theEndTimes) {
        bad(`CANNOT UNDERSTAND END TIME "${line[columns.get(CEND)]}"`, theLineNumber, line[eventCol]); 
        continue; 
      }
      //console.log(`Line ${theLineNumber}: theHour: ${theHour}:${theMin}:${theTimes.TBA}, theEndHour: ${theEndHour}:${theEndMin}:${theEndTimes.TBA}`);
      if (theEndTimes.TBA) {
        theEndHour = TBAend_hour;
        theEndMin = 0;
        if (!theEvent.match(TBAstring))
          theEvent = theEvent + TBAstring;
      }
      else {
        theEndHour = theEndTimes.H;
        theEndMin = theEndTimes.M;
      }
      if (+theEndHour > 23 || +theEndMin >59) {
        bad(`INVALID END TIME "${line[columns.get(CEND)]}"`, theLineNumber, line[eventCol]); 
        continue; 
      }
      //console.log(`Line ${theLineNumber}: theEndHour: ${theEndHour}:${theEndMin}:${theEndTimes.TBA}`);


      /* BETTER TO CHECK AFTER WE HAVE A DURATION, IF ANY
      let matchTime;
      if (matchTime = line[endCol].match(TIME_PATT)) {
        if ((+theHour > + matchTime[1]) ||
            ((+theHour == +matchTime[1]) && (theMin >= +matchTime[2]))) {
          if (!warnUserOrCancel(`Event on line ${theLineNumber} ${line[eventCol]? line[eventCol] : ''} has mismatched start and end times\n` +
                `${line[startCol]} and ${line[endCol]}`)) {
            text = null;
            return null;
          }
        }
        theEndHour = matchTime[1].padStart(2, '0');
        theEndMin = matchTime[2];
        if (theEndHour > 23 || theEndMin >59) {
          bad(`END TIME ${line[endCol]} IS INVALID`, theLineNumber, line[eventCol]); 
          continue; 
        }
      }
      else {
        bad(`CANNOT UNDERSTAND End TIME ${line[endCol]}`, theLineNumber, line[eventCol]); 
        continue; 
      }
      */
    }

    // Or may have a Duration
    let hasDuration = false;
    if (columns.has(CDURATION) && line[columns.get(CDURATION)]) {
      hasDuration = true;

      if (hasEnd) {
        if (!warnUserOrCancel(`Line ${theLineNumber}: Event end overrides duration\n\t${line}`))
          return null;
      }

      else {
        //console.log(line[columns.get(CEVENT)], line[columns.get(CSTART)], line[columns.get(CDURATION)]);
        const theDurationTimes = getDuration(line, line[columns.get(CDURATION)]);
        if (!theDurationTimes) {
          bad(`CANNOT UNDERSTAND DURATION TIME "${line[columns.get(CDURATION)]}"`, theLineNumber, line[eventCol]); 
          continue; 
        }
        //if ( (theDurationTimes.H != TBAhour) && (theDurationTimes != NAhour) ) {
        if (!theDurationTimes.TBA) {
          if (+theDurationTimes.H === 0 && +theDurationTimes.M === 0) {
            bad(`DURATION "${line[columns.get(CDURATION)]}" MUST BE GREATER THAN 0`, theLineNumber, line[eventCol]); 
            continue; 
          }
          if (+theDurationTimes.M >= 60) {
            bad(`DURATION MINUTES "${line[columns.get(CDURATION)]}" MUST BE LESS THAN 60`, theLineNumber, line[eventCol]); 
            continue; 
          }
          theEndHour = +theHour + +theDurationTimes.H;
          theEndMin = +theMin + +theDurationTimes.M;
        }
        console.log(`theEnd ${theEndHour}:${theEndMin}, duration ${theDurationTimes.H}:${theDurationTimes.M}`);
        /*
        let matchTime;
        if (matchTime = line[columns.get(CDURATION)].match(DURATION_PATT)) {
          theEndHour = +theHour + +matchTime[1];
          theEndMin = matchTime[2] ? +theMin + +matchTime[2] : 0;
        }
        else {
          bad(`CANNOT UNDERSTAND THE DURATION ${line[columns.get(CDURATION)]}`, theLineNumber, line[columns.get(CEVENT)]); 
          continue; 
        }
        */
      }
    }

    // Event with multiple races
    if (!hasEnd && !hasDuration) {
      // assume more than one race if event string includes 
      // more than one separate number, and allow an extra hour 
      theEndHour = +theHour + ( /\d\D+\d/.test(theEvent)
                                  ? +DEFAULT_DURATION_HOUR + +DEFAULT_EXTENSION_HOUR
                                  : +DEFAULT_DURATION_HOUR
                               );
      theEndMin = +theMin + +DEFAULT_DURATION_MIN;
    }

    // Fix / warn about out of range times
    while (+theEndMin >= 60) {
      theEndHour = +theEndHour + 1;
      theEndMin -= 60;
    }
    if (+theEndHour >= 24) {
      if (!warnUserOrCancel(`Line ${theLineNumber}: EVENT SPANS MIDNIGHT!\n\t${theEvent}\n\tEnd changed to 23:59.`)) {
        text = null;
        return null;
      }
      theEndHour = '23';
      theEndMin = '59';
    }
    if ( (+theEndHour < +theHour) || ((+theEndHour == +theHour) && (+theEndMin <= +theMin)) ) {
      bad(`Line ${theLineNumber}: EVENT ENDS BEFORE IT STARTS!\n\t${theEvent} from ${theHour}:${theMin} to ${theEndHour}:${theEndMin}`, theLineNumber, line[eventCol]);
      continue;
    }
      
    // Get highwater time
    const highwater = columns.has(CHW) ? line[columns.get(CHW)] || '' : ''; // Guard against empty cell
   
    // Print the record
    text += printICAL(DTSTAMP, 
                      String(theDay), String(theMonth), String(theYear),
                      String(theHour).padStart(2, '0'), String(theMin).padStart(2, '0'),
                      String(theEndHour).padStart(2, '0'), String(theEndMin).padStart(2, '0'),
                      thePrefix, theEvent, String(highwater));
    // For testing
    console.log(`+++EVENT ${theEvent} ${theDay}/${theMonth}/${theYear} ${String(theHour).padStart(2, '0')}:${String(theMin).padStart(2, '0')}-${String(theEndHour).padStart(2, '0')}:${String(theEndMin).padStart(2, '0')}`);
  }

  text += printEOCAL();
  //console.log(text);
  return text;
}


/**
 * Convert hour to UTC.
 * Always output UTC times. If we were to write TZ=timezone, then 
 * strict conformance with the  RFC 5545 specification) requires that 
 * this timezone be fully specified (including BST and GMT start dates and
 * offsets, etc).  Of course, these dates are the last Sundays in March/October,
 * so they change... :-(
 * @param year The year
 * @param month The month
 * @param day The day
 * @param hour The hour
 * @return the hour as UTC, padded to 2 digits
 */
function convHourUTC(year, month, day, hour) {
  const dt = new Date(year, month-1, day, hour);
  return String(dt.getUTCHours()).padStart(2, '0');
}


/**
 * @return iCalendar (and vCalendar) header
 */
function printCALhdr() {
  // TODO nice to include X-WR-CALNAME: property
  const hdr = 
`BEGIN:VCALENDAR
VERSION:${VERSION}
PRODID:-//Richard Jones//xl2cal//EN
CALSCALE:GREGORIAN
`;
  return hdr;
}

/**
 * Clean ics text by preventing injections and escaping , ; \
 * as per RFC 5545
 * @param s a string
 * @return cleaned string
 */
function icsText(s) {
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\r\n|\r|\n/g, '\\n');
}


const MAX_LINE_OCTETS = 75;       // RFC 5545 limit, excluding the line break
const utf8 = new TextEncoder();

/**
 * Fold a content line to at most 75 octets per line (RFC 5545), breaking
 * between words where possible. Counts UTF-8 octets, never splits a
 * multi-byte character, and only breaks mid-word if a single word is
 * longer than a whole line.
 * @param {string} line The unfolded content line (no line break)
 * @param {string} eol  The line break to use (default '\n')
 * @return {string} The folded line
 */
function foldLine(line, eol = '\n') {
  const chars = Array.from(line);                       // by code point, so surrogate pairs stay whole
  const sizes = chars.map(c => utf8.encode(c).length);  // 1 to 4 octets each
  const pieces = [];
  let start = 0;

  while (start < chars.length) {
    // The first line holds 75 octets; a continuation line has a leading
    // fold space that uses one of its 75, leaving 74 for its content.
    const capacity = pieces.length === 0 ? MAX_LINE_OCTETS : MAX_LINE_OCTETS - 1;

    // Greedily take as many characters as fit
    let end = start;
    let octets = 0;
    while (end < chars.length && octets + sizes[end] <= capacity) {
      octets += sizes[end];
      end++;
    }

    if (end < chars.length) {                           // the rest won't fit: we must break
      // Prefer to break just before a space, so words stay whole
      let brk = end;                                    // a break at 'end' is before chars[end]
      while (brk > start && chars[brk] !== ' ') brk--;
      if (brk > start) {
        while (brk > start + 1 && chars[brk - 1] === ' ') brk--;  // no trailing spaces on the line
        end = brk;
      }
      // else: one word longer than a line; break at 'end' (mid-word)
    }

    pieces.push(chars.slice(start, end).join(''));
    start = end;
  }
  return pieces.join(eol + ' ');                        // continuation lines begin with one space
}


/**
 * Print iCalendar entry
 * @param DTSTAMP   The DTSTAMP
 * @param thePrefix The prefix for events
 * @param theDay    The day number
 * @param theMonth  The month number
 * @param theYear   The year
 * @param theStart  The start hour
 * @param theMin    The start minute 
 * @param theEnd    The end hour
 * @param theEndMin The end minute
 * @param theEvent  The event name
 * @param theHighwater Other text to add to event (e.g. high water time)
 * @return iCalendar entry
 */
function printICAL (DTSTAMP, theDay, theMonth, theYear, theStart, theMin, 
                    theEnd, theEndMin, thePrefix, theEvent, theHighwater) {
  //console.log(DTSTAMP, theDay, theMonth, theYear, theStart, theMin, theEnd, theEndMin, thePrefix, theEvent, theHighwater);
  const hw = theHighwater === '' ? '' : `, HW=${theHighwater}`;
  const summary = theEvent + hw;
  // $summary =~ s/.{63}\K/\n /sg; # fold lines longer than 75 octets
  const day = theYear.padStart(4, '0') + theMonth.padStart(2, '0') + theDay.padStart(2, '0');
  const start = convHourUTC(theYear, theMonth, theDay, theStart) + theMin + '00';
  const end = convHourUTC(theYear, theMonth, theDay, theEnd) + theEndMin + '00';
  /*
  let alarm = start - ALERT;
  if (alarm < 0) { 
    warnUserOrCancel(`Alarm set for previous day: ${theEvent}`);
    alarm = "000000";
  } else { 
    alarm = String(alarm).padStart(4, '0');
  }
  */
  const uid = crypto.randomUUID(); // Requires https: (or file:) not http:
  const eventText = icsText(thePrefix + summary);
  const entry = 
`BEGIN:VEVENT
CREATED:${DTSTAMP}${Z}
UID:${uid}
DTSTAMP:${DTSTAMP}${Z}
DTSTART:${day}${T}${start}${Z}
DTEND:${day}${T}${end}${Z}
${foldLine('SUMMARY:' + eventText)}
BEGIN:VALARM
TRIGGER:-PT${ALERT}H
ACTION:DISPLAY
${foldLine('DESCRIPTION:' + eventText)}
END:VALARM
END:VEVENT
`;
  return(entry);
}


/**
 * @return Calendar (and vCalendar) trailer
 */
function printEOCAL() {
  const eoc = 
`END:VCALENDAR
`;
  return eoc;
}


/**
 * Sanitise the name of the download file
 * @param name the raw name
 * @return the sanitised name
 */
function safeFilename(name) {
  return name.replace(/[^\p{L}\p{N}_-]+/gu, '_').replace(/^_+|_+$/g, '').slice(0, 60) || 'myCalendar';
}


/**
 * Open a new window with a download link
 * @param calendarsToExport The names of the calendars chosen for export
 * @param iCal The iCalendar data
 */
function openICalWindow(calendarsToExport, iCal) {
  const newWindow = window.open('', '_blank', 'width=400,height=400');
  if (!newWindow) {
    alert('Please allow pop-ups for this page to see the calendar.'); 
    return; 
  }
  newWindow.opener = null;
  newWindow.document.write('<html><head><title>Export iCalendar</title><link rel="stylesheet" type="text/css" href="styles.css"></head><body>');

  const heading = newWindow.document.createElement("h2");
  let title = 'iCalendar generated';
  let filename;
  if (calendarsToExport.size > 0) {
    title += ' for ' + Array.from(calendarsToExport).join(', ');
    filename = safeFilename(Array.from(calendarsToExport).join('')) +'.ics';
  } else {
    filename = 'myCalendar.ics';
  }
  console.log('filename', filename);
  heading.textContent = title;   // textContent rather than innerHTML to prevent injection
  newWindow.document.body.appendChild(heading);

  const blob = new Blob([iCal], {type:'text/calendar;charset=utf-8'});
  const link = newWindow.document.createElement("a");
  link.id = 'download-link';
  link.download = filename;
  link.textContent = "Download iCalendar";
  link.href = window.URL.createObjectURL(blob);
  newWindow.document.body.appendChild(link);

  const pre = newWindow.document.createElement("pre");
  pre.textContent = iCal; 
  newWindow.document.body.appendChild(pre);
  newWindow.document.write('</body></html>');
  newWindow.addEventListener('pagehide', () => URL.revokeObjectURL(link.href));
}


/*
TODO Bugs and possible improvements.

1. Sanity check on columns chosen.
   When the user selects a column, sniff entries in this column to see if most
   of them look plausible. Use 'most' not 'all' to allow for lines to be ignored
   or typos in entries.
   Checks would include
   . isDay() - cardinal or ordinal number
   . isMonth() - cardinal number, or month name or abbreviation
   . isTime() - \d\d\d\d, \d\d:\d\d, \d\d.\d\d, TBA, TBC, NA, N/A
   Probably, don't bother as we now let user bail out early.

2. Could check that any weekday given is correct.

From Robert 10/9/23
It would be good to have a description on what the dropdowns are looking for, and what happens if they're not selected. For example:If a field in the dropdowns isn't used, what happens? For example: End; Duration and Calendar. Does it assume something for these?
  ADDED tooltips with brief explanation, including whether MUST or OPTION and what happens if option not sepecified.
  IMPROVED descriptions of columns, and what happens if optional column is not chosen.
Day, Month and Date dropdowns. Not clear what Day needs (date number or day of week), and whether all are needed. 
  FIXED: Changed 'Day' to 'Day no.'. Prevented selection of Date and Day+Month.
Other
  ADDED a Default Duration box.
  CHANGED Year box to Default Year.
  ADDED tooltips to all initial components, i.e. the drop zone, the Choose File button, Default Year, Default Duration, Events Prefix

From Steve 2/10/23
Hidden rows ADDED skipHidden (which requires cellStyles:true)
Header row is not necessarily the first row, which may be too short. FIXED to use max width select boxes
Uses non-integer as a time, e.g. 12.3 means "12:30". CHANGED to guess that a non-integer in range 0..24 is a time.

October 2026
Added a cancel button to the warning dialog to allow the user to abort generating a calendar rather than having to click 'OK' to every error.
Sorted column header options.
Improved error message when years mismatch
Added a VALARM for ALERT hours before
*/
