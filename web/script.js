// script.js
// Convert WYC racing schedule into .ics format that can be loaded into a calendar.

//TODO
//import { v4 as uuidv4 } from 'uuid';


// Months 
const months = new Map([
  ['JAN', 1],
  ['FEB', 2],
  ['MAR', 3],
  ['APR', 4],
  ['MAY', 5],
  ['JUN', 6],
  ['JUL', 7],
  ['AUG', 8],
  ['SEP', 9],
  ['OCT', 10],
  ['NOV', 11],
  ['DEC', 12]
]);

// Hash of column names to numbers.
// Invariant: mappings are unique, i.e. no two keys may have the same value
const columns = new Map();
const columnValues = ['Day', 'Month', 'Date', 'HW', 'Start', 'End', 'Duration', 'Event', 'Calendar'];


// Constants
const VERSION = '2.0'; // for iCalendar
const TZ = 'Europe/London';
const DURATION = 2;   // hours for a single race 
const H = 'H';        // needed for the DURATION value
const ADVANCE = 2;    // 2 hours warning

// To Be Announced
const TBAhour = '09';             // so set the period to be 9.00 - 17.00
const TBAend_hour = '17';
const TBAstring = ' (times TBC)'; // string to add to TBA events
// Not Applicable
const NAhour = '09';              // so set the period to be 9.00 - 17.00
const NAend_hour = '17';

const T = 'T';
const Z = 'Z';      // 'Z' means UTC rather than local time

let theYear;  // The year to generate the calendar for

//Calendars to use
let allCalendars;
let calendarsToExport = new Set();

function setCalendarsToExport(cals) {
  console.log('calendarsToExport', cals);
  calendarsToExport = cals;
}

// The spreadsheet data
let jsonData;

// Drag and drop handlers
function dropHandler(ev) {
  console.log("File(s) dropped");

  // Prevent default behavior (Prevent file from being opened)
  ev.preventDefault();

  if (ev.dataTransfer.items) {
    // Use DataTransferItemList interface to access the file(s)
    const item = ev.dataTransfer.items[0];
    // If dropped items aren't files, reject them
    if (item.kind === "file") {
      const file = item.getAsFile();
      console.log('DataTransferItemList', `… file.name = ${file.name}`);
      handleFile(file);
    }
  } else {
    // Use DataTransfer interface to access the file(s)
    const file = ev.dataTransfer.files[0];
    console.log('DataTransfer', `… file.name = ${file.name}`);
    handleFile(file);
  }
}

function dragOverHandler(ev) {
  console.log("File(s) in drop zone");

  // Prevent default behavior (Prevent file from being opened)
  ev.preventDefault();
}


// Choose the spreadsheet
window.addEventListener('DOMContentLoaded', () => {
  const fileChooser = document.getElementById('file-chooser');
  const yearDropdown = document.getElementById('year-dropdown');
  
  fileChooser.addEventListener('change', handleFileEvent);

  // Generate the range of years
  const currentYear = new Date().getFullYear();
  const startYear = currentYear - 1;
  const endYear = currentYear + 2;

  for (let year = startYear; year <= endYear; year++) {
    const option = document.createElement('option');
    option.value = year;
    option.textContent = year;
    option.selected = year == currentYear;
    yearDropdown.appendChild(option);
  }
  
  // Access the selected year from the dropdown
  theYear = yearDropdown.value;
  console.log('Selected year:', theYear);
});


// Read the spreadsheet file
function handleFileEvent(event) {
  const file = event.target.files[0];
  handleFile(file);
}


function handleFile(file) {
  console.log(file.name);

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Assuming the first sheet in the workbook
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Add title and instructions here
    const sheetHeader = document.getElementById('sheet-header');
    sheetHeader.innerHTML = "2. The spreadsheet";
    const instructions = document.getElementById('columns-instructions');
    instructions.innerHTML = 
        "<em>Select the columns to use from the dropdown lists in the first row</em>";

    renderTable(jsonData);
  };

  reader.readAsArrayBuffer(file);
}


// Render the spreadsheet on screen
function renderTable(data) {
  const sheetTable = document.getElementById('sheet-table');
  const numColumns = data.length > 0 ? data[0].length : 0;
  let html = '';

  // Insert new row of options at index 0
  for (let i = 0; i < numColumns; i++) {
    html += '<td><select class="column-select">';
    html += `<option value=empty>...</option>`;
    for (let option of columnValues) {
      html += `<option value="${option}">${option}</option>`;
    }
    html += '</select></td>';
  }

  for (let row of data) {
    html += '<tr>';
    
    for (let cell of row) {
      let cellValue = cell == undefined ? '' : String(cell); // Convert cell value to string
  
      // Check if the cell format is "hh:mm"
      if (typeof cell === 'number' && cell % 1 !== 0) {
        const hours = Math.floor(cell * 24);
        const minutes = Math.round(((cell * 24 ) % 1) * 60);
        cellValue = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
      }

      html += `<td>${cellValue}</td>`;
    }

    html += '</tr>';
  }

  sheetTable.innerHTML = html;

  // Add horizontal and vertical sliders
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
      
      // Reset other dropdowns
      dropdowns.forEach((dd, i) => {
        if (i != index) {
          let val = dd.value;
          if (columns.get(val) != i) {
            dd.value = 'empty';
          }
        }
      });

      //Enable export button?
      const exportBtn = document.getElementById('export-button');
      exportBtn.disabled = !( 
                   (columns.get('Day') && columns.get('Month') || columns.get('Date'))
                    && columns.get('Start')
                    && columns.get('Event')
                            );
    });
  });

  // Export button
  const exportBtn = document.createElement('button');
  exportBtn.id = 'export-button';
  exportBtn.innerHTML = 'Export';
  exportBtn.disabled = true;
  exportBtn.addEventListener('click', exportCalendar);
  document.body.appendChild(exportBtn);
}


// Find the key in a hash table 'object' with value 'value'
function getKeyByValue(map, value) {
  for (let key of map.keys()) {
    if (map.get(key) === value) {
        return key;
    }
  }
  return null;
}

// Search the spreadsheet for calendars, if any
// Want only Calendar entries for event rows (not headers, etc)
function findCalendars() {
  const calendarsFound = new Set();
  if (columns.get('Calendar')) {
    const startCol = columns.get('Start');
    const calCol = columns.get('Calendar');
    const pattern = /^\d{4}$|^\d{2}:\d{2}$/;
    for (let row of jsonData) {
      const start = row[startCol];
      if (typeof start === 'number' || pattern.test(start)) {
        calendarsFound.add(row[calCol]) 
      }
    }
  }
  return calendarsFound;
}


// Modal dialog and listeners
let modal;


function cancelSelect() {
  console.log('called cancel');
}


function exportSelect() {
  const calendarsToExport = new Set();
  const checkedCalendars = document.querySelectorAll('input[name=calendar-checkbox]:checked');
  checkedCalendars.forEach(function(cal) {
    calendarsToExport.add(cal.value);
  });
  setCalendarsToExport(calendarsToExport);
  modal.close();

  // Generate iCal
  const iCal = generateICal(jsonData);

  //Open new window with iCal data
  openICalWindow(calendarsToExport, iCal);

}


// Export the calendar to export
// NB. The required fields have been chosen
function exportCalendar() {
  // First, search the spreadsheet for calendars, if any
  allCalendars = calendarsFound = findCalendars();
  selectCalendars(calendarsFound);
  console.log(calendarsFound);
}


// Popup window to select calendars to export
function selectCalendars(calendarsFound) {
  // New window to select calendars (if any) and iCalendar file
  const numCalendarsFound = calendarsFound.size;
  if (numCalendarsFound > 0) {

    modal = new SV.Modal('select-calendars-modal');
    modal.getModalElement().addEventListener('sv.modal.close', function (ev) {
	console.log('Closed the modal');
    });
   
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
        <button id="cancel-calendar-choice-button" onclick="cancelSelect()">Cancel</button>
        <button id="export-calendar-choice-button" onclick="exportSelect()">Export</button>
    `;

    modal.inject(html, 'Select calendars to export');
    modal.show();
  }
}


// Generate DTSTAMP
function makeDTSTAMP() {
  const now = new Date();
  const year = now.getUTCFullYear();
  const month = String(now.getUTCMonth() + 1).padStart(2, '0');
  const day = String(now.getUTCDate()).padStart(2, '0');
  const hours = String(now.getUTCHours()).padStart(2, '0');
  const minutes = String(now.getUTCMinutes()).padStart(2, '0');
  const seconds = String(now.getUTCSeconds()).padStart(2, '0');
  return '${year}${month}${day}${T}{hours}${minutes}${seconds}';
}

// Parse a day in a format like 'Sunday 12th March'
// Return [day number, month], e.g. [12, 'March']
function parseDate(row) {
  const rowDate = row[columns.get('Date')];
  // Define the regular expression for the delimiters (spaces and commas)
  const delimiterRE = /[,\s]+/;
  const yearRE = /^\d{4}$/;
  let dayMonthYear = rowDate.split(delimiterRE)
                     .filter(it => it !== '')              // remove any empty elements
                     .map(it => it.toUpperCase())          // make all upper case
                     .filter(it => !it.endsWith('DAY'));   // remove any day names
  const years = dayMonthYear
                .filter (it => yearRE.text(it));
  const dayMonth = dayMonthYear
                   .filter(it => !(/^\d{4}$/.text(it))); // remove any year
  const days = dayMonth
               .filter(it => /^\d+[A-Z]*$/.test(it))   // only day strings 
               .map(it => it.replace(/[A-Z]*$/,''));    // remove any ordinal characters
  const months = dayMonth
                 .filter(it => /^[A-Z]+$/.test(it)); 
  let rv = (!days.length || !months.length) ?
    null:  // something went wrong
    [days[0], months[0]];
  if (rv && year.length) {
    rv =  [days[0], months[0], years[0]];
  }
  return rv;
}

// Generate the iCalendar data
// TODO
function generateICal(data) {
  const DTSTAMP = makeDTSTAMP();
  const startCol = columns.get('Start');
  const startPattern = /^\d{4}$|^\d{2}:\d{2}$/;

  // Print header for calendars
  let text = printCALhdr();

  // Process the source file
  for (let lineNum in data) {
    line = data[lineNum];
  
    // Defend against regex DOS attacks, otherwise trailing spaces is polynomial
    line = line.map(it => typeof it === 'string' ? it.trim() : it); 

    // Discard header and any other unlikely lines
    const start = line[startCol];
    if ( (typeof start != 'number') || !start.match(/\d+[:\.]?\d+|TB[AC]|N\/?A/i) ) {
      console.log(`Ignoring line ${lineNum}: ${line}`);
      continue;
    }

    // Get the day and month
    let theNum, theMonth;

    if (columns.get('Day') && columns.get('Month')) {
      const lineMonth = line.get('Month');
      theMonth = (/^\d\d?$/.test(lineMonth)) ?
                  line.get('Month') :
                  0 + months.get(lineMonth.toUpperCase().substr(0, MONTHS_PREFIX_LEN));
      if (theMonth < 1 || theMonth > 12) {
        bad(`Invalid month ${month} ${lineMonth} ${line[columns.get('Month')]}`, lineNum);
        continue;
      }

      // get the day number
      theNum = line[columns.get('Day')];
      if (!(/^\d\d?$/.test(theNum))) {
        bad(`Cannot understand day number \"${theNum}\"`, lineNum); 
        continue; 
      }
      if ((theNum < 1) || (theNum > new Date(year, theMonth, 0).getDate())) {
        bad (`${theNum} is out of range for a day in month ${theMonth}`, lineNum); 
        continue; 
      }

    } else if (columns.get('Date')) {
      // parse a date
      const dayMonthYear = parseDate(line);
      if (!dayMonthYear) {
        bad(`Bad date ${line[columns.get('Date')]}`, lineNum);
        continue; 
      }
  
      if ((dayMonthYear.length === 3) &&
          (dayMonthYear[2] != theYear)) {
        alert(`Year doesn't match on line ${lineNum}`);
      }
  
      theNum = months.get(dayMonthYear[0].substr(0, 3));
      theMonth = dayMonthYear[1];
    }

    // Get the event
    const theEvent = line[columns.get('Event')];

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
          theEndHour = theEndHour.padStart(2, '0');
      }

      else {
        // assume more than one race if event string includes 
        // more than one separate number, and allow an extra hour 
        theEndHour = theHour + (/\d\D+\d/.test(theEvent) ? DURATION + 1 : DURATION);
        theEndMin = theMin;
        if (theEndHour > 24) {
          alert(`Event on line ${lineNum} cannot span midnight! ${theEvent}, ${theHour}`);
          theEndHour = '23';
          theEndMin = '59';
        }
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
    const highwater = line[columns.get('HW')];
   
    // Print the record
    text += printICAL(theNum, theMonth, theHour, theMin, theEndHour, theEndMin, theEvent, highwater);
  }

  text += printEOCAL();
  //console.log(text);
  return text;
}


// Convert hour to UTC
// Always output UTC times. If we were to write TZ=timezone, then 
// strict conformance with the  RFC 5545 specification) requires that 
// this timezone be fully specified (including BST and GMT start dates and
// offsets, etc).  Of course, these dates are the last Sundays in March/October,
// so they change... :-(
function convUTC(year, month, day, hour) {
  const dt = new Date(year, month-1, day, hour);
  return dt.getUTCHours().padStart(2, '0');
}


// iCalendar (and vCalendar) header
// TODO nice to include X-WR-CALNAME: property
function printCALhdr() {
  return 
`BEGIN:VCALENDAR
VERSION:${VERSION}
PRODID:Richard Jones xl2cal.html generated
CALSCALE:GREGORIAN
`;
}


// Print iCalendar entry
function printICAL (theNum, theMonth, theStart, theMin, 
                    theEnd, theEndMin, theEvent, theHighwater) {
  const hw = theHighwater === '' ? '' : `, HW=${theHighwater}`;
  const summary = theEvent + hw;
  // $summary =~ s/.{63}\K/\n /sg; # fold lines longer than 75 octets
  const day = theYear.padStart(4, '0') + theMonth.padStart(2, '0') + theNum.padStart(2, '0');
  const start = convUTC(theYear, theMonth, theNum, theStart);
  const end = convUTC(theYear, theMonth, theNum, theEnd);
  let alarm = start - ADVANCE;
  if (alarm < 0) { 
    alert(`Alarm set for previous day: ${theEvent}`);
    alarm = "000000";
  } else { 
    alarm = alarm.padStart(4, '0');
  }
  start = start + theMin + '00';
  end = end + theEndMin + '00';
  const uid = '9e7de45c-51d4-43b4-8895-477b45926c3c'; //FIXME uuidv4();
  return 
`BEGIN:VEVENT
CREATED:${DTSTAMP}${Z}
UID:${uid}
DTSTAMP:${DTSTAMP}${Z}
DTSTART:${day}${T}${start}${Z}
DTEND:${day}${T}${end}${Z}
SUMMARY:WYC ${summary}
END:VEVENT
EOI
`;
}

// Calendar (and vCalendar) trailer
function printEOCAL() {
  return 
`END:VCALENDAR
`;
}

// Open a new window with a download link
// calendarsToExport: the names of the calendars chosen for export
// iCal: the iCalendar data
function openICalWindow(calendarsToExport, iCal) {
  const newWindow = window.open('', '_blank', 'width=400,height=400');
  newWindow.document.write('<html><head><title>Export iCalendar</title><link rel="stylesheet" type="text/css" href="styles.css"></head><body>');

  const heading = newWindow.document.createElement("h2");
  let title = 'iCalendar generated';
  let filename;
  if (calendarsToExport) {
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

function bad(msg, lno) {
  alert(`BAD RECORD: ${msg} on line ${lno} so ignoring this line.`);
}
