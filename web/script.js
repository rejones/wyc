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

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

/*
*/

const dayNames = ['Sun', 'Mon', 'Tues', 'Weds', 'Thurs', 'Fri', 'Sat'];
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
const columnValues = [CDAY, CMONTH, CDATE, CEVENT, CSTART, CEND, CDURATION, CHW, CCALENDAR];
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
const ADVANCE = 2;     // 2 hours warning

// To Be Announced times
const TBAhour = '09';             // so set the period to be 9.00 - 17.00
const TBAend_hour = '17';
const TBAstring = ' (times TBC)'; // string to add to TBA events
// Not Applicable times
const NAhour = '09';              // so set the period to be 9.00 - 17.00
const NAend_hour = '17';

const T = 'T';
const Z = 'Z';      // 'Z' means UTC rather than local time

// Default durations
const DEFAULT_DURATION = '2:00';    // hours for a single race 
const DEFAULT_DEFAULT_DURATION_HOUR = 2;
const DEFAULT_DEFAULT_DURATION_MIN = 0;
let DEFAULT_DURATION_HOUR = DEFAULT_DEFAULT_DURATION_HOUR;
let DEFAULT_DURATION_MIN = DEFAULT_DEFAULT_DURATION_MIN;

// Spreadsheet data
let theSheet = 0;                        // The sheet to read
const theData = [];                      // The data from the spreadsheet
let allCalendars;                        // Calendars to use

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
  if (!(error instanceof AbortError)) {
    warn(`Unexpected error at line ${lineno}!\n`, evt, error);
    alert(`Unexpected error at line ${lineno}!\n` +
        error +
        "\nPlease report");
  }
}

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
  
  fileChooser.addEventListener('change', loadFileEvent);

  // Generate the range of years
  const currentYear = new Date().getFullYear();
  const startYear = currentYear - 1;
  const endYear = currentYear + 2;

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
    duration = durationBox.value;
    let matchTime;
    if (duration) {
      matchTime = duration.match(/^(\d\d?)([:\.]\d\d)?/);
    }
    if (matchTime) {
      DEFAULT_DURATION_HOUR = +matchTime[1];
      if (matchTime[2]) {
        matchTime[2] = matchTime[2].match(/\d\d/);
      }
      DEFAULT_DURATION_MIN =  matchTime[2] ? +matchTime[2] : 0;
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
  loadFile(file);
}


/**
 * Read the spreadsheet and render it in the document
 * @param {File} file - The file loaded
 */
function loadFile(file) {
  //console.log(file.name);

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { cellStyles: true, type: 'array' });
    let sheetName;
    const options = { skipHidden: true, raw: false, header: 1 };

    if (workbook.SheetNames.length == 1) {
      // only one sheet
      sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
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
        sheetName = event.target.value;
        const sheet = workbook.Sheets[sheetName];
        jsonData = XLSX.utils.sheet_to_json(sheet, options);
        renderTable(jsonData);
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
      jsonData = XLSX.utils.sheet_to_json(sheet, options);
      // Clear any previous column selections
      renderTable(jsonData);
    }
  };

  reader.readAsArrayBuffer(file);
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
 * @param {Array} data - The sheet to parse and display
 */
function renderTable(data) {

  // display number with 2 digits, prepending `0` if necessary 
  const f = (n) => n.toString().padStart(2, "0");

  const sheetTable = document.getElementById('sheet-table');
  const numColumns = data.map(row => row.length)
                     .reduce( (a,b) => Math.max(a,b), 0);
  columns.clear();
  let html = '';
  let newRow;

  // Insert new row of options at index 0
  for (let i = 0; i < numColumns; i++) {
    html += '<td><select class="column-select">';
    html += `<option value=empty>...</option>`;
    let title;
    for (let option of columnValues) {
      html += `<option value="${option}" title="${columnTips.get(option)}">${option}</option>`;
    }
    html += '</select></td>';
  }
  
  // Process the data, also saving it as strings
  theData.length = 0; // clear in case user has added the spreadsheet more than once
  for (let row of data) {
    html += '<tr>';
    newRow = [];
    
    //console.log(row);
    for (let cell of row) {
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
      if (cell instanceof Date) {
        // Try to tidy incomplete Excel values
        if (cell.getHours()==0 && cell.getMinutes()==0) { // Looks like a date
          cellValue = `${dayNames[cell.getDay()]} ${cell.getDate()} ${monthNames[cell.getMonth()]}`;
          // TODO Don't really want to put the day into theData
          const year = cell.getFullYear();
          if (year > 1904) { // Otherwise, user probably didn't specify year so likely to be 1899 or 1903
            cellValue = cellValue + ` ${year}`;
          }
        }
        else { // Probably a time
          cellValue = `${f(cell.getHours())}:${f(cell.getMinutes())}`;
        }
      }

      // Check if the cell format is "hh:mm"
      else if (typeof cell === 'number') {
        if (cell % 1 !== 0) { // A non-integer so assume this is a time
          
          // Steve Gray spreadsheet has times as decimal numbers, which is weird.
          if ((1 <= cell) && (cell < 24.00)) {
            const hours = Math.floor(cell);
            const minutes = Math.round((cell - hours) * 100);
            if (minutes < 60) { // looks like a time
              cellValue = `${hours}.${f(minutes)}`;
            }
          }
          else {
            // Assume it's an Excel time
            const excelTime = cell % 1; 
            const hours = Math.floor(excelTime * 24);
            const minutes = Math.round(((excelTime * 24) % 1) * 60);
            cellValue = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
          }
        }
        else if (cell > 43480) { // Assume it's a date after 1 Jan 2020
          const date = new Date(Date.UTC(0, 0, cell-1));
          if (!isNaN(date.getFullYear())) {
            cellValue = date.toDateString();
          }
        }
      }

      else {
        let matchDate;
        // Defend against regex DOS attacks, otherwise trailing spaces is polynomial
        if (cell && (matchDate = cell.trim().match(/^(\d\d?)\/(\d\d?)\/(\d{2,4})$/))) {
          // Looks like a date.
          // Excel default format seems to be m/d/yy but we want to display and use d/m/yy
          cellValue = `${matchDate[2]}/${matchDate[1]}/${matchDate[3]}`;
        }
      }

          
      //console.log(cellValue);
      html += `<td>${cellValue}</td>`;
      newRow.push(cellValue);
    }

    html += '</tr>';
    theData.push(newRow);
  }

  sheetTable.innerHTML = html;

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
      if ( start.match(/^(\d\d?)[:\.]?(\d\d)/) ||
           start.match(/TBA|TBC|-/i) ||
           start.match(/N\/?A/i) ) {
        calendarsFound.add(row[calCol]) 
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
  const calendarsToExport = new Set();
  if (columns.has(CCALENDAR)) {
    const checkedCalendars = document.querySelectorAll('input[name=calendar-checkbox]:checked');
    checkedCalendars.forEach(function(cal) {
      calendarsToExport.add(cal.value);
    });
    //setCalendarsToExport(calendarsToExport);
    modal.close();
  }

  // Generate iCal
  const iCal = generateICal(theData, calendarsToExport);
  //console.log(iCal);

  //Open new window with iCal data
  openICalWindow(calendarsToExport, iCal);
}


/**
 * Export the calendar to export
 * NB. The required fields have been chosen
 */
function exportCalendar() {
  // First, search the spreadsheet for calendars, if any
  allCalendars = calendarsFound = findCalendars();
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
 * Popup window to select calendars to export
 * @param calendarsFound - The calendars found in the CCALENDAR column
 */
function selectCalendars(calendarsFound) {
  if (calendarsFound.size > 0) {

    modal = new SV.Modal('select-calendars-modal');
    /*
    modal.getModalElement().addEventListener('sv.modal.close', function (ev) {
	console.log('Closed the modal');
    });
    */
   
    // inject some content
    let html = `
      <p>Calendars found:</p>
    `;
    calendarsFound.forEach(function(cal) {
      html += `
        <input type="checkbox" id="${cal}" name="calendar-checkbox" value="${cal}" />
        <label for="${cal}">${cal}</label><br>
      `;
    });
    html += `
      <p>If no calendars are selected, all will be exported.</p>
    `;
    html += `
        <button id="cancel-calendar-choice-button" onclick="cancelSelect()">Cancel</button>
        <button id="export-calendar-choice-button" onclick="exportSelect()">Export</button>
    `;

    modal.inject(html, 'Select calendars to export');
    modal.show();
  }
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
 * @param {String} day The possible month string
 * @return True if so
 */
function isMonth(month) {
  month = month.toUpperCase();
  return monthNames.map(m => m.toUpperCase()).find((m) => {
      return m.startsWith(month);
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
    //console.log(`rowDate (${rowDate}) is a number`);
    const date = new Date(Date.UTC(0, 0, rowDate -1));
    //console.log(date);
    if (isNaN(date.getFullYear())) {
      //console.warn(`parseDate returns null! for ${rowDate}`);
      return null;
    }
    //console.log(rowDate, date.getDate(), date.getMonth() + 1, date.getFullYear());
    return [date.getDate().toString(), date.getMonth() + 1, date.getFullYear()].map((e) => {return e.toString()});
  }
  
  // Next, try dd/mm/yy
  if (matchDate = rowDate.match(/^(\d\d?)\/(\d\d?)\/(\d{2,4})$/)) {
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

  //Next, try numbers and words

  // Define the regular expression for the delimiters (spaces and commas)
  const delimiterRE = /[,\s]+/;
  const yearRE = /^\d{4}$/;
  let dayMonthYear = rowDate.split(delimiterRE)
                     .filter(it => it !== '')              // remove any empty elements
                     .map(it => it.toUpperCase());         // make all upper case
                     //.filter(it => !isDay(it));          // remove any day names TODO Needs to accept all forms of days
  //console.log('dayMonthYear', dayMonthYear);
  const years = dayMonthYear
                .filter (it => yearRE.test(it));         // get any year
  //console.log('years', years);
  const dayMonth = dayMonthYear
                   .filter(it => !(yearRE.test(it)));   // remove any year
  //console.log('dayMonth', dayMonth);
  const days = dayMonth
               .filter(it => /^\d+[A-Z]*$/.test(it))    // only day strings 
               .map(it => it.replace(/[A-Z]*$/,''));    // remove any ordinal characters
  //console.log('days', days);
  //const months = dayMonth
  //               .filter(it => /^[A-Z]+$/.test(it)); 
  const months = dayMonth.filter(isMonth);              // get any month(s)
  //console.log('months',months);
  //console.log(days[0], months[0], years[0]);
  if (!days.length || !months.length)                   // must have day and month numbers
    return null;  
  return years.length ? 
    [days[0], months[0], years[0]] :
    [days[0], months[0]];
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
    line = data[lineNum];
    if (!line || !line.length) {
      continue;
    }
  
    // Defend against regex DOS attacks, otherwise trailing spaces is polynomial
    line = line.map(it => it.trim()); 

    // Discard header and any other unlikely lines
    const start = line[startCol];
    if ( !start || !start.match(/\d+[:\.]?\d+|TB[AC]|N\/?A/i) ) {
      console.warn(`Ignoring line ${lineNum} ${line}: no start time specified`);
      continue;
    }

    // Skip rows for unwanted calendars
    if (columns.has(CCALENDAR)) {
       const theCal = line[columns.get(CCALENDAR)];
       if (!calendarsToExport.has(theCal)) {
         continue;
       }
    }

    // Get the day and month
    let theDay, theMonth, theYear;

    if (columns.has(CDAY) && columns.has(CMONTH)) {
      theYear = theDefaultYear;
      const monthCol = columns.get(CMONTH); 
      const lineMonth = line[monthCol];
      // allow either month number or string
      if (/^\d\d?$/.test(lineMonth)) {
        theMonth =  line[monthCol];
      }
      else {
        const mth = months.get(lineMonth.toUpperCase().substr(0, MONTHS_PREFIX_LEN));
        if (!mth) {
          bad(`Cannot understand month \"${lineMonth}\"`, lineNum, line[eventCol]);
          continue;
        }
        theMonth = 0 + mth;
        if (theMonth < 1 || theMonth > 12) {
          bad(`Invalid month ${month} ${lineMonth} ${line[monthCol]}`, lineNum, line[eventCol]);
          continue;
        }
      }

      // get the day number
      //theDay = line[columns.get(CDAY)];
      let matchDay;
      if (matchDay = line[columns.get(CDAY)].match(/^(\d\d?)(st|nd|rd|th)?/i)) {
        theDay = matchDay[1];
      } else {
        //if (!(/^\d\d?$/.test(theDay))) {
        bad(`Cannot understand day number \"${theDay}\"`, lineNum, line[eventCol]); 
        continue; 
      }

    } else if (columns.has(CDATE)) {
      // parse a date
      const dateCol = columns.get(CDATE);
      if (!line[dateCol]) {
        bad(`Empty date ${line[dateCol]}`, lineNum, line[eventCol]);
        continue;
      }
      const dayMonthYear = parseDate(line[dateCol]);
      if (!dayMonthYear) {
        bad(`Bad date ${line[dateCol]}`, lineNum, line[eventCol]);
        continue;
      }
      //console.log(dayMonthYear);
  
      // check that any year value matches the year chosen 
      if ((dayMonthYear.length === 3) &&
          (dayMonthYear[2] != theDefaultYear)) {
        warnUser(`Different year (${dayMonthYear[2]})on line ${lineNum}\nIs this what you meant?`);
      }
      else if (dayMonthYear.length === 2) {
        dayMonthYear.push(theDefaultYear);
      }
  
      theDay = dayMonthYear[0];
      theMonth = dayMonthYear[1].match(/^\d+$/) ?
        dayMonthYear[1] :
        months.get(dayMonthYear[1].substr(0, 3));
      //console.log('theDay - theMonth', theDay, theMonth);
      theYear = dayMonthYear[2];
    }

    // Check that the date is valid
    if (new Date(theYear, theMonth-1, theDay).getMonth() != theMonth-1) {
      bad(`${theDay}/${theMonth}/${theYear} is not a valid date`, lineNum, line[eventCol]);
      continue;
    }

    // Get the event. 
    // Discard any lines with no event.
    let theEvent = line[columns.get(CEVENT)];
    if ( !theEvent ) {
      console.warn(`Ignoring line ${lineNum} ${line}: no event specified`);
      continue;
    }

    // Get the start and end times
    // times before 1000 are sometimes recorded with only 3 digits
    let theHour;
    let theMin;
    let theEndHour;
    let theEndMin;

    // Must have a Start time
    let matchTime;
    if (matchTime = line[startCol].match(/^(\d\d?)[:\.]?(\d\d)/)) {
      theHour = matchTime[1].padStart(2, '0');
      theMin = matchTime[2];
    }
    else if (line[startCol].match(/TBA|TBC|-/i) || 
              line[startCol] === '') {
      theHour = TBAhour;
      theMin = '00';
      theEndHour = TBAend_hour;
      theEndMin = '00';
      theEvent = theEvent + TBAstring;
    }
    else if (line[startCol].match(/N\/?A/i)) {
      theHour = NAhour;
      theMin = '00';
      theEndHour = NAend_hour;
      theEndMin = '00';
    }   
    else {
      bad(`Cannot understand Start time ${line[columns.get(CSTART)]}`, lineNum, line[eventCol]); 
      continue; 
    }

    // May have an End time
    if (columns.has(CEND) && line[columns.get(CEND)]) {
      const endCol = columns.get(CEND);
      if (matchTime = line[endCol].match(/^(\d\d?)[:\.]?(\d\d)/)) {
        if ((+theHour > + matchTime[1]) ||
            ((+theHour == +matchTime[1]) && (theMin >= +matchTime[2]))) {
          warnUser(`Event on line ${lineNum} ${line[eventCol]? line[eventCol] : ''} has mismatched start and end times\n` +
                `${line[startCol]} and ${line[endCol]}`);
        }
        theEndHour = matchTime[1];
        theEndMin = matchTime[2];
      }
      else {
        bad(`Cannot understand End time ${line[endCol]}`, lineNum, line[eventCol]); 
        continue; 
      }
    }

    // Or may have a Duration
    else if (columns.has(CDURATION) && line[columns.get(CDURATION)]) {
      //console.log(line[columns.get(CEVENT)], line[columns.get(CSTART)], line[columns.get(CDURATION)]);
      if (matchTime = line[columns.get(CDURATION)].match(/^(\d\d?)([:\.]\d\d)?/)) {
        theEndHour = +theHour + +matchTime[1];
        if (matchTime[2]) {
          matchTime[2] = matchTime[2].match(/\d\d/);
        }
        theEndMin = matchTime[2] ? +theMin + +matchTime[2] : 0;
      }
      else {
        bad(`Cannot understand the Duration ${line[columns.get(CDURATION)]}`, lineNum, line[columns.get(CEVENT)]); 
        continue; 
      }
    }

    else {
      // assume more than one race if event string includes 
      // more than one separate number, and allow an extra hour 
      theEndHour = +theHour + (/\d\D+\d/.test(theEvent) ? DEFAULT_DURATION_HOUR + 1 : DEFAULT_DURATION_HOUR);
      theEndMin = +theMin + DEFAULT_DURATION_MIN;
    }

    // Fix / warn about out of range times
    while (theEndMin >= 60) {
      theEndHour = +theEndHour + 1;
      theEndMin -= 60;
    }
    if (theEndHour >= 24) {
      warnUser(`Event on line ${lineNum} cannot span midnight! ${theEvent}, ${theHour}`);
      theEndHour = '23';
      theEndMin = '59';
    }
    theEndHour = String(theEndHour).padStart(2, '0');
    theEndMin = String(theEndMin).padStart(2, '0');


    // Get highwater time
    const highwater = columns.has(CHW) ? line[columns.get(CHW)] : '';
   
    // Print the record
    text += printICAL(DTSTAMP, theDay, theMonth, theYear, theHour, theMin, theEndHour, theEndMin, thePrefix, theEvent, highwater);
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
PRODID:Generated by xl2cal.html (Richard Jones, 2023)
CALSCALE:GREGORIAN
`;
  return hdr;
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
  let alarm = start - ADVANCE;
  if (alarm < 0) { 
    warnUser(`Alarm set for previous day: ${theEvent}`);
    alarm = "000000";
  } else { 
    alarm = String(alarm).padStart(4, '0');
  }
  const uid = crypto.randomUUID();
  const entry = 
`BEGIN:VEVENT
CREATED:${DTSTAMP}${Z}
UID:${uid}
DTSTAMP:${DTSTAMP}${Z}
DTSTART:${day}${T}${start}${Z}
DTEND:${day}${T}${end}${Z}
SUMMARY:${thePrefix}${summary}
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
 * Open a new window with a download link
 * @param calendarsToExport The names of the calendars chosen for export
 * @param iCal The iCalendar data
 */
function openICalWindow(calendarsToExport, iCal) {
  const newWindow = window.open('', '_blank', 'width=400,height=400');
  newWindow.document.write('<html><head><title>Export iCalendar</title><link rel="stylesheet" type="text/css" href="styles.css"></head><body>');

  const heading = newWindow.document.createElement("h2");
  let title = 'iCalendar generated';
  let filename;
  if (calendarsToExport.size > 0) {
    title += ' for ' + Array.from(calendarsToExport).join(', ');
    filename = Array.from(calendarsToExport).join('')+'.ics';
  } else {
    filename = 'myCalendar.ics';
  }
  heading.innerHTML = title;
  newWindow.document.body.appendChild(heading);

  const blob = new Blob([iCal], {type:'text/plain'});
  const link = newWindow.document.createElement("a");
  link.id = 'download-link';
  link.download = filename;
  link.innerHTML = "Download iCalendar";
  link.href = window.URL.createObjectURL(blob);
  newWindow.document.body.appendChild(link);

  const pre = newWindow.document.createElement("pre");
  pre.innerHTML = iCal;
  newWindow.document.body.appendChild(pre);
  newWindow.document.write('</body></html>');
}


/**
 * Warn user and log
 * @param msg The alert message
 */
function warnUser(msg) {
  console.warn(msg);
  alert(msg);
}


/**
 * Warn about bad record in the spreadsheet
 * @param msg The error message
 * @param lno The number of the line on which the error was found
 * @param evt The name of the event
 */
function bad(msg, lno, evt) {
  const message = evt ? 
      `BAD RECORD: ${msg} on line ${lno} for event ${evt} so ignoring this line.` :
      `BAD RECORD: ${msg} on line ${lno} so ignoring this line.` ;
  console.warn(message);
  if (confirm(message)) {
    return;
  } else {
    alert("You've have cancelled the run");
    throw new AbortError('Run aborted');
  }
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

*/
