// script.js

/**
 * Convert WYC racing schedule into .ics format that can be loaded into a calendar.
 * @author Richard Jones
 * @copyright Richard Jones, 2023
 * https://github.com/rejones/wyc
 */

/** 
 *Month prefixes to numbers
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
const daysInMonth = [31,28,31,30,31,30,31,31,39,31,30,31];

/**
 * Hash of column names to numbers.
 *  Invariant: mappings are unique, i.e. no two keys may have the same value
 */
const columns = new Map();
const columnValues = ['Day', 'Month', 'Date', 'HW', 'Start', 'End', 'Duration', 'Event', 'Calendar'];


// Constants
const VERSION = '2.0'; // for iCalendar
const DURATION = 2;    // hours for a single race 
const H = 'H';         // needed for the DURATION value
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

// Spreadsheet data
let theSheet = 0;                        // The sheet to read
const theData = [];
let allCalendars;                                               // Calendars to use

/** Export button for spreadsheet should only be created once */
let hasExportBtn = false;

/** Modal dialog */
let modal;

/* Deal with errors
window.onerror = function (error, source, lineno, colno, error) {
  alert("Unexpected error!\n" +
        error +
        "\nPlease report to Richard Jones");
}
*/

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
    option.selected = year == currentYear;
    yearDropdown.appendChild(option);
  }
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
    //console.log('DataTransferItemList/DataTransfer', `… file.name = ${file.name}`);
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
    const workbook = XLSX.read(data, { type: 'array' });
    let sheetName;

    if (workbook.SheetNames.length == 1) {
      // only one sheet
      sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      replaceSheetGroup();
      addSpreadsheetHeading();
      const sheetChooser = document.getElementById("sheet-names");
      sheetChooser.replaceChildren();
      renderTable(jsonData);
    } 
    else { 
      replaceSheetGroup();
      addSpreadsheetHeading();
      // choose sheet from drop-down list
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
        jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
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
      jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
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

  const sheetTable = document.getElementById('sheet-table');
  const numColumns = data.length > 0 ? data[0].length : 0;
  columns.clear();
  let html = '';
  let newRow;

  // Insert new row of options at index 0
  for (let i = 0; i < numColumns; i++) {
    html += '<td><select class="column-select">';
    html += `<option value=empty>...</option>`;
    for (let option of columnValues) {
      html += `<option value="${option}">${option}</option>`;
    }
    html += '</select></td>';
  }
  
  // Process the data, also saving it as strings
  theData.length = 0; // clear in case user has added the spreadsheet more than once
  for (let row of data) {
    html += '<tr>';
    newRow = [];
    
    console.log(row);
    for (let cell of row) {
      console.log(cell, cell%1, typeof cell);
      // Convert cell value to string
      let cellValue = cell == undefined ? '' : String(cell);
  
      // Check if the cell format is "hh:mm"
      if (typeof cell === 'number') {
        if (cell % 1 !== 0) { 
          // Assume this is a time
          console.log('here');
          const hours = Math.floor(cell * 24);
          const minutes = Math.round(((cell * 24 ) % 1) * 60);
          cellValue = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        }
        else if (cell > 43481) { // days between 1 Jan 1900 and 1 Jan 2020
          // Assume that this is a date
          const date = new Date(Date.UTC(0, 0, cell-1));
          if (!isNaN(date.getFullYear())) {
            cellValue = date.toDateString();
          }
        }
      }
          
      console.log(cellValue);

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
      }
      
      // Clear any other dropdowns with the same value
      dropdowns.forEach((dd, i) => {
        if (i != index) {
          let val = dd.value;
          if (columns.get(val) != i) {
            dd.value = 'empty';
          }
        }
      });

      //Enable export button once required columns are chosen
      const exportBtn = document.getElementById('export-button');
      exportBtn.disabled = !( 
                   ((columns.has('Day') && columns.has('Month')) || columns.has('Date'))
                    && columns.has('Start')
                    && columns.has('Event')
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
 * @param map 
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
  if (columns.has('Calendar')) {
    const startCol = columns.get('Start');
    const calCol = columns.get('Calendar');
    const pattern = /^\d{4}$|^\d{2}:\d{2}$/; // looks like a Start time
    for (let row of theData) {
      const start = row[startCol];
      if (pattern.test(start)) {
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
  if (columns.has('Calendar')) {
    const checkedCalendars = document.querySelectorAll('input[name=calendar-checkbox]:checked');
    checkedCalendars.forEach(function(cal) {
      calendarsToExport.add(cal.value);
    });
    //setCalendarsToExport(calendarsToExport);
    modal.close();
  }

  // Generate iCal
  const iCal = generateICal(theData, calendarsToExport);
  console.log(iCal);

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
 * @param calendarsFound - The calendars found in the 'Calendar' column
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
 * @return A DTSTAMP for now
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
 * @param day The possible month string
 * @return True if so
 */
function isMonth(month) {
  month = month.toUpperCase();
  return ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST',
          'SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'].find((m) => {
      return m.startsWith(month);
    }) != undefined;
}


/**
 * Parse a day in a format like 'Sunday 12th March' or '12/3/'
 * @param rowDate The date string
 * @return [day number, month], e.g. [12, 'March']
 */
function parseDate(rowDate) {
  console.log('parseDate', rowDate, typeof rowDate);
  if (!rowDate)
    return null;

  // First, try Excel Date format which is read as a number
  // TODO not needed now that renderTable converts Excel serial values to date strings 
  if (!isNaN(rowDate)) {
    // Try converting from Excel date to JS date
    alert(`rowDate (${rowDate}) is a number`);
    const date = new Date(Date.UTC(0, 0, rowDate -1));
    console.log(date);
    if (isNaN(date.getFullYear())) {
      console.log('parseDate returns null!');
      return null;
    }
    console.log(rowDate, date.getDate(), date.getMonth() + 1, date.getFullYear());
    console.log([date.getDate(), date.getMonth() + 1, date.getFullYear()]);
    return [date.getDate().toString(), date.getMonth() + 1, date.getFullYear()].map((e) => {return e.toString()});
  }
  
  //Next, try numbers and words

  // Define the regular expression for the delimiters (spaces and commas)
  const delimiterRE = /[,\s]+/;
  const yearRE = /^\d{4}$/;
  let dayMonthYear = rowDate.split(delimiterRE)
                     .filter(it => it !== '')              // remove any empty elements
                     .map(it => it.toUpperCase());         // make all upper case
                     //.filter(it => !isDay(it));             // remove any day names TODO Needs to accept all forms of days
  //console.log('dayMonthYear', dayMonthYear);
  const years = dayMonthYear
                .filter (it => yearRE.test(it));         // get any year
  //console.log('years', years);
  const dayMonth = dayMonthYear
                   .filter(it => !(yearRE.test(it)));   // remove any year
  //console.log('dayMonth', dayMonth);
  const days = dayMonth
               .filter(it => /^\d+[A-Z]*$/.test(it))   // only day strings 
               .map(it => it.replace(/[A-Z]*$/,''));   // remove any ordinal characters
  //console.log('days', days);
  //const months = dayMonth
  //               .filter(it => /^[A-Z]+$/.test(it)); 
  const months = dayMonth.filter(isMonth);             // get any month(s)
  //console.log('months',months);
  console.log(days[0], months[0], years[0]);
  if (!days.length || !months.length)          // must have day and month numbers
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
  const startCol = columns.get('Start');
  const startPattern = /^\d{4}$|^\d{2}:\d{2}$/;

  const theDefaultYear = document.getElementById('year-dropdown').value;
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
    if ( !start.match(/\d+[:\.]?\d+|TB[AC]|N\/?A/i) ) {
      console.log(`Ignoring line ${lineNum}: ${line}`);
      continue;
    }

    // Skip rows for unwanted calendars
    if (columns.has('Calendar')) {
       const theCal = line[columns.get('Calendar')];
       if (!calendarsToExport.has(theCal)) {
         continue;
       }
    }

    // Get the day and month
    let theNum, theMonth, theYear;

    if (columns.has('Day') && columns.has('Month')) {
      theYear = theDefaultYear;
      const lineMonth = line[columns.get('Month')];
      // allow either month number or string
      theMonth = (/^\d\d?$/.test(lineMonth)) ?
                  line[columns.get('Month')] :
                  0 + months.get(lineMonth.toUpperCase().substr(0, MONTHS_PREFIX_LEN));
      if (theMonth < 1 || theMonth > 12) {
        bad(`Invalid month ${month} ${lineMonth} ${line[columns.get('Month')]}`, lineNum);
        continue;
      }

      // get the day number
      //theNum = line[columns.get('Day')];
      let matchDay;
      if (matchDay = line[columns.get('Day')].match(/^(\d\d?)(st|nd|rd|th)?/i)) {
        theNum = matchDay[1];
        if ((theNum < 1) || (theNum > new Date(theDefaultYear, theMonth, 0).getDate())) {
          bad (`${theNum} is out of range for a day in month ${theMonth}`, lineNum); 
          continue; 
        }
      } else {
        //if (!(/^\d\d?$/.test(theNum))) {
        bad(`Cannot understand day number \"${theNum}\"`, lineNum); 
        continue; 
      }

    } else if (columns.has('Date')) {
      // parse a date
      const dayMonthYear = parseDate(line[columns.get('Date')]);
      if (!dayMonthYear) {
        bad(`Bad date ${line[columns.get('Date')]}`, lineNum);
        continue;
      }
      console.log(dayMonthYear);
  
      // check that any year value matches the year chosen 
      if ((dayMonthYear.length === 3) &&
          (dayMonthYear[2] != theDefaultYear)) {
        alert(`Different year (${dayMonthYear[2]})on line ${lineNum}\nIs this what you meant?`);
      }
      else if (dayMonthYear.length === 2) {
        dayMonthYear.push(theDefaultYear);
      }
  
      theNum = dayMonthYear[0];
      theMonth = dayMonthYear[1].match(/^\d+$/) ?
        dayMonthYear[1] :
        months.get(dayMonthYear[1].substr(0, 3));
      //console.log('theNum - theMonth', theNum, theMonth);
      theYear = dayMonthYear[2];
    }

    // Get the event
    let theEvent = line[columns.get('Event')];

    // Get the start and end times
    // times before 1000 are sometimes recorded with only 3 digits
    let theHour;
    let theMin;
    let theEndHour;
    let theEndMin;

    // Must have a Start time
    let matchTime;
    if (matchTime = line[columns.get('Start')].match(/^(\d\d?)[:\.]?(\d\d)/)) {
      theHour = matchTime[1];
      theMin = matchTime[2];

      // May have an End time
      if (columns.has('End') && 
          (matchTime = line[columns.get('End')].match(/^(\d\d?)[:\.]?(\d\d)/))) {
          theEndHour = matchTime[1].padStart(2, '0');
          theEndMin = matchTime[2];
      }

      // Or may have a Duration
      else if (columns.has('Duration') &&
                 (matchTime = line[columns.get('Duration')].match(/^(\d\d?)[:\.]?(\d\d)/))) {
          theEndHour = +theHour + +matchTime[1];
          theEndMin = +theMin + +matchTime[2];
          while (theEndMin > 60) {
            theEndHour = +theEndHour + 1;
            theEndMin -= 60;
          }
          theEndHour = String(theEndHour).padStart(2, '0');
          theEndMin = String(theEndMin).padStart(2, '0');
      }

      else {
        // assume more than one race if event string includes 
        // more than one separate number, and allow an extra hour 
        theEndHour = +theHour + (/\d\D+\d/.test(theEvent) ? DURATION + 1 : DURATION);
        theEndMin = theMin;
        //console.log('times', theHour, theMin, theEndHour);
        if (theEndHour > 24) {
          alert(`Event on line ${lineNum} cannot span midnight! ${theEvent}, ${theHour}`);
          theEndHour = '23';
          theEndMin = '59';
        }
        theEndHour = String(theEndHour).padStart(2, '0');
        theEndMin = String(theEndMin).padStart(2, '0');
      }
      theHour = theHour.padStart(2, '0'); 
    }

    else if (line[columns.get('Start')].match(/TBA|TBC|-/i) || 
              line[columns.get('Start')] === '') {
      theHour = TBAhour;
      theMin = '00';
      theEndHour = TBAend_hour;
      theEndMin = '00';
      theEvent = theEvent + TBAstring;
    }

    else if (line[columns.get('Start')].match(/N\/?A/i)) {
      theHour = NAhour;
      theMin = '00';
      theEndHour = NAend_hour;
      theEndMin = '00';
    }
   
    else {
      bad(`Cannot understand Start time ${line[columns.get('Start')]}`, lineNum); 
      continue; 
    }

    // Get highwater time
    const highwater = columns.has('HW') ? line[columns.get('HW')] : '';
   
    // Print the record
    text += printICAL(DTSTAMP, theNum, theMonth, theYear, theHour, theMin, theEndHour, theEndMin, thePrefix, theEvent, highwater);
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
 * @param DTSTAMP The DTSTAMP
 * @param thePrefix The prefix for events
 * @param theNum    The day number
 * @param theMonth  The month number
 * @param theYear  The year
 * @param theStart  The start hour
 * @param theMin    The start minute 
 * @param theEnd    The end hour
 * @param theEndMin The end minute
 * @param theEvent The event name
 * @param theHighwater Other text to add to event (e.g. hig hwater time)
 * @return iCalendar entry
 */
function printICAL (DTSTAMP, theNum, theMonth, theYear, theStart, theMin, 
                    theEnd, theEndMin, thePrefix, theEvent, theHighwater) {
  console.log(DTSTAMP, theNum, theMonth, theYear, theStart, theMin, theEnd, theEndMin, thePrefix, theEvent, theHighwater);
  const hw = theHighwater === '' ? '' : `, HW=${theHighwater}`;
  const summary = theEvent + hw;
  // $summary =~ s/.{63}\K/\n /sg; # fold lines longer than 75 octets
  const day = theYear.padStart(4, '0') + theMonth.padStart(2, '0') + theNum.padStart(2, '0');
  const start = convHourUTC(theYear, theMonth, theNum, theStart) + theMin + '00';
  const end = convHourUTC(theYear, theMonth, theNum, theEnd) + theEndMin + '00';
  let alarm = start - ADVANCE;
  if (alarm < 0) { 
    alert(`Alarm set for previous day: ${theEvent}`);
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
 * Warn about bad record in the spreadsheet
 * @param msg The error message
 * @param lno The number of the line on which the error was found
 */
function bad(msg, lno) {
  if (confirm(`BAD RECORD: ${msg} on line ${lno} so ignoring this line.`)) {
    return;
  } else {
    alert("You've have cancelled the run");
    throw new Error('Run aborted');
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
2. TODO Improve placement of select-box components.
3. Default and actually year
4. Days in month
*/
