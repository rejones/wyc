# Generate a calendar file from WYC sailing events Excel file

[Whitstable Yacht Club] (https://wyc.org.uk) (WYC) publishes its sailing events calendar 
as an Excel spreadsheet (and a webpage).
wyc.pl is a tool for convert this spreadsheet (saved as a `.csv` file) 
to a file suitable for importing into calendars like Apple or Google Calendar etc.

## Usage

`wyc.pl [-h] [-v] [-d] [-z] calendar [year] < in.csv > outfile`

The default output format is iCalendar (`.ics`) as per 
[RFC 5545] (https://www.rfc-editor.org/rfc/rfc5545). 
Other supported but obsolete formats are:
-  vCalendar (`.vcs`)
-  DateBook for Palm mobile devices (`.dba`)

The `calendar` argument is the type of events to generate,
i.e. for WYC, "Main" or "Cadet".

Options:
```
  -h 		Print this help.
  -v		Use .vcs format rather than iCalendar.
  -d		Use .dba format rather than iCalendar.
  -z		Use UTC rather than local time.
```
## History

The tool has undergone a number of changes over the years, 
from the first version supporting Palm Pilot DateBook calendars
to more modern calendars that import iCalendar version 2 files.
Currently, the Perl tool is pretty hard-wired to the column
structure of the WYC spreadsheet, but I plan to improve this.

Richard Jones
17.1.2023
