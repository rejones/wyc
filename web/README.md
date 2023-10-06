 How to use xl2cal
-----------------

 x2cal is a simple tool to convert a schedule, such as a sailing club racing calendar,
 from a spreadsheet to an iCalendar that can be imported into Apple Calendar, Google Calendar, etc.

### 1. Creating a spreadsheet for your schedule

 *xl2cal* assumes that you have created your schedule
 in an Excel spreadsheet, with filename extension `.xlsx`.   
 *More formats may be added in the future.*

 The layout of the spreadsheet is flexible,
 but *xl2cal* expects to find:

- one event per row, and
- columns specifying the day of the event 
  (events spanning days must be added as separate events), 
  the start time of the event, 
  and a name for the event. 
  See section [Selecting columns for export](#3-selecting-columns-for-export) for more detail.
 
 All event rows must provide this information. 
 Any row that does not look like an event is ignored. 
 Examples of ones that will be ignored include heading rows, blank rows,
 any rows without at least a date, start time and event name,
 etc.

 ### 2. Loading your spreadsheet into xl2cal

 Load your spreadsheet into *xl2cal* by either:

- dragging the spreadsheet file onto the landing site labelled
  `Drag and drop a spreadsheet file here` or
- using the `Choose file` button.
 
 By default, *xl2cal* assumes that the sheet to use is the first
 in the spreadsheet.
 However, you can change this with the
 `Sheet name or number` box.
 Note that sheets are numbered from 0.

 You may also wish to select the calendar year for your calendar
 from the `Select Year` drop-down list on the left.
 The default is the current year.

 Set default duration of events with the `Duration` box.
 Times may be hours (e.g. "2") or hours and minutes (e.g. "0:45" or "0.45").

 If you want all calendar entries to use a common prefix (such as "WYC"),
 enter it in the `Events Prefix` box on the right.

### 3. Selecting columns for export

 The cells of your spreadsheet will then be loaded into a table.
 Cells in the first row of the table are column heading buttons.
 Use these to select the columns used by your spreadsheet.
 Some columns are mandatory; others are optional.

#### Mandatory columns

1. The date of the events must be specified by selecting  
    
    - EITHER a **Day no.** (number) column *and*
      a **Month** column,
    - OR a **Date** column.
     
     *xl2cal* will not allow a **Date** column to be selected
     if a **Day no.** or a **Month** column
     are, and vice-versa.  

     If days are specified, the **Day no.** column may contain *either*
     cardinal (e.g. "6") *or* ordinal (e.g. "6th") numbers.

     If months are specified, the **Month** column may contain either cardinal numbers
     (e.g. "6), or month names or prefixes (e.g. "June" or "Jun").  

     If a **Date** is specified, most Excel formats should work
     (e.g. "Saturday 1 July", "1/6/2023", etc).

2. A **Start** column. 
     Entries in the start column should be times using a 24-hour clock.  

     Acceptable entries include "1400" and "14:00".
     Excel Time format can be used here (see under [Errors](#errors)
     below).
     However, "TBA", "TBC", "NA" and "N/A" are also accepted.

     Lines without valid Start times (such as headers, blank lines) are
     ignored.

3. An **Event** column, describing the event.

#### Optional columns

These columns are optional and do not need to be selected.
- A high-water (**HW**) column.  
     If this column is specified,
     any entry in this column (e.g. "1500") will be appended to the
     entry in the **Event** column (e.g. "Race 1" becomes "Race 1, HW=1500").

- An **End** column.  
     If this column is specified, any entry in this column will be used for the
     end time of an event.
     Otherwise, the end time will be the start time plus the duration.

     Entries in the start column should be times using a 24-hour clock.

     Acceptable entries include "1400" and "14:00",
     and Excel Time format (see [Errors](#errors) below).

- A **Duration** column.  
     If this column is specified, any entry in this column will be used for the
     duration of an event.
     Otherwise, the event will be given the default duration from the
     `Duration` box.
     Acceptable entries include "2", "2:30" and Excel Time format
     (see [Errors](#errors) below).

- A **Calendar** column.  
     This is intended to allow separate calendars to be generated from a
     single spreadsheet,
     by tagging event rows in the spreadsheet with the name of the
     appropriate calendar.
     Before exporting an iCalendar file, the tool will ask which calendars it
     should generate entries for.
     If no Calendar column is specified, or if no calendars are selected for
     export, the iCalendar file output will contain entries for all events.

#### Errors

 Take care that dates and times are properly entered in your spreadsheet.
 In particular,

- If you specify a **Date** column, ensure that there are no missing
spaces between words, no misspelt words.

- Be careful with Excel Date and Time formats.  
     Excel usually represents these as
     the number of days since 1 January 1900 (the integral part)
     / the fraction of 24 hours (the fractional part).
     However, one Excel preference uses 1904 rather than 1900;
     avoid the 1904 format.

     If you use a Date or Time format, ensure that it is for the year you want:
     Excel is quite good at using a year in the early 1900s instead!
 
 *xl2cal* will check each row in your spreadsheet.
 Any row that does not provide a valid date (however specified)
 or without a valid Start time will be ignored,
 with a warning.
 This allows you to have blank lines or other non-event rows in your
 spreadsheet.
 You have two choices if a warning pops up for a row with invalid data.

- To continue generating your iCalendar without this row,
  click `OK`.
- to cancel generating your iCalendar, 
  click `Cancel`.
 
### 4. Selecting the calendars to export

 Click the `Export` button to start to generate your calendar in
 iCalendar format.
 This button will only be enabled (with a blue background) if you have selected sufficient
 columns.

 If you have selected one of the columns of your spreadsheet as a "Calendars"
 column,
 then a pop-up window will appear,
 asking you to choose which calendar(s) to export.
 *xl2cal* will generate entries for rows with "Calendars" cells matching
 the ones you have chosen.
 If you do not choose any,
 or if no "Calendar" column was selected,
 the iCalendar file generated will contain entries for all events.
 Click the `Export` button to complete
 generating your calendar.

### 5. Download the schedule as an iCalendar file

 When the iCalendar has been generated,
 a window will appear inviting you to download the calendar.
 This window also displays the text of the iCalendar generated.
 Click the `Download iCalendar` button at the top
 of the window to download the iCalendar.
 You should choose to save the iCalendar in a file with extension
 `.ics`.

 You can then import the `.ics` file into the calendar app of your choice,
 e.g. Apple or Google Calendar.
 To check that the calendar generated is correct, import it into a temporary
 calendar in your app first. This is what I do on MacOS.

1. In the calendar app, create a new calendar.
2. Import the `.ics` file into this calendar.
3. Check the dates and times.
4. Delete this calendar...
5. ...but, if the entries are all correct, merge them into 
   another calendar (e.g. I have a separate 'Sailing' calendar).
 
 
