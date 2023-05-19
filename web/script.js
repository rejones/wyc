// script.js

// Hash of column numbers (either as A,B,C... or 1,2,3... of necessary fields.
// Invariant: mappings are unique, i.e. no two keys may have the same value
const columns = new Map();
const columnValues = ['Day', 'Month', 'Date', 'HW', 'Start', 'End', 'Duration', 'Event', 'Calendar'];

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
  const selectedYear = yearDropdown.value;
  console.log('Selected year:', selectedYear);
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


// Export the calendar to export
// NB. The required fields have been chosen
function exportCalendar() {
  // First, search the spreadsheet for calendars, if any
  allCalendars = calendarsFound = findCalendars();
  selectCalendars(calendarsFound);
}


// Popup window to select calendars to export
function selectCalendars(calendarsFound) {
  // New window to select calendars (if any) and iCalendar file
  const numCalendarsFound = calendarsFound.size;
  if (numCalendarsFound > 0) {
    
    const calendarsWindow = window.open('', 'selectCalendars', 'width=400,height=300, resizable=0');
    let html = `
      <html>
        <head>
          <title>Choose calendars</title>
          <link rel="stylesheet" type="text/css" href="styles.css">
          <style>
            position: fixed; /* Stay in place */
            z-index: 1; /* Sit on top */
          </style>
        </head>
        <body>
          <h2>Select calendars to export</h2>
    `;
    calendarsFound.forEach(function(cal) {
      html += `
        <input type="checkbox" id="${cal}" value="${cal}" />
        <label for="${cal}">${cal}</label><br>
      `;
    });
    html += `
        <button id="cancel-calendar-choice-button" onclick="cancel()">Cancel</button>
        <button id="export-calendar-choice-button" onclick="exportCalendars()">Export</button>
        <script>
          function cancel() {
            console.log('called cancel');
            window.close(); // Close the new window
          }
          
          function exportCalendars() {
            console.log('called exportCalendars');
            const calendarsToExport = new Set();
    `;
    calendarsFound.forEach(function(cal) {
    html += `
            if (document.getElementById("${cal}").checked) {
              calendarsToExport.add(document.getElementById("${cal}").value);
            }
      `;
    });
    html += `
            const parentWindow = window.opener;
            parentWindow.setCalendarsToExport(calendarsToExport);
            window.close(); // Close the new window
          }
        </script>
      </body>
    </html>
  `;

  calendarsWindow.document.write(html);
  }
}
