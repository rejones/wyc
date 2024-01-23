# Generate a calendar file from an Excel file

Organisations 
(such as [Whitstable Yacht Club](https://wyc.org.uk))
often construct their events calendar in an Excel spreadsheet 
before publishing on a web page, etc.
The tools provided here convert spreadsheet data into to a format suitable 
for importing into calendars like Apple or Google Calendar etc.

## Tools

Two tools are provided, one as a web application and one as a command-line tool.

### Web version

_xl2cal_ is a web application. 
Its use is simple. The user drags an `.xlsx` spreadsheet onto a web page,
and then selects which sheet, and which columns on that sheet, to use.
The output is an iCalendar (`.ics`) file as per 
[RFC 5545] (https://www.rfc-editor.org/rfc/rfc5545).
Written in plain JavaScript, it relies on the very excellent Excel
parser [SheetJS](https://sheetjs.com).

### Command-line version

_wyc.pl_ is a command-line tool written in Perl.
Its use is more complicated than the web application.
The input is a `.csv` file.
The output may be `.ics`, `.vcs` (vCalendar), or even `.dba` (DateBook for 
Palm mobile devices)!
Command-line arguments specify the columns of the `.csv` file to use.

This version is not being actively maintained.

