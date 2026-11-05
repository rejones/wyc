# xl2cal date/time test workbooks

Tests `script_b6749e.js` (the version uploaded on 20 Oct 2026). Every row has an **Expected** cell (or, for heading rows, a cell comment).

## How to run

1. Open the page, choose **Default Year = 2027** (the page pre-selects next year, so this is already the default in 2026). Leave Default duration at 2 and Events Prefix at WYC.
2. Load a workbook and set the column dropdowns as below. Other columns (No., Expected, Notes) stay unselected.
3. Click Export. **Press OK on every dialog**; each one is part of the test. Dialogs quote the Excel row number.
4. In the exported text, events are named `WYC Case X`. Times in the `.ics` are UTC, so in summer months (BST) they are one hour earlier than the local times in the Expected column.

| Workbook | Dropdown mapping |
|---|---|
| A date column | Date→Date, Start→Start, Event→Event |
| B day and month | Day no.→Day, Month→Month, Start→Start, Event→Event |
| C times | Date, Start, End, Duration, Event → the columns of the same name |
| D native cells | same as C (cells are real Excel dates/times with number formats) |
| E short rows | five sheets, pick each sheet in turn; map every column named in its sheet. The unlabelled Expected/Notes columns stay unselected |

### Expected classes
- **ACCEPT**: event produced, no dialog, at the stated time. **WARN, then ACCEPT**: a dialog, then the event.
- **REJECT**: a "BAD RECORD" or "Ignoring line" dialog and no event.
- **INVALID**: must be rejected or give a warning; silently producing an event is a failure.
- **EITHER**: accept correctly, or reject/warn, but never mis-read.
- **IGNORE**: heading row; no event, no error. **AMBIGUOUS**: informational.

### Results

`xl2cal_test_predicted_results.csv` has every row's predicted outcome (as determined by Clause). Text cells are exact. For workbook D (real Excel dates/times) Claude assumed SheetJS formats cells as Excel displays them, e.g. built-in date format 14 as `m/d/yy`. Compare the on-screen table with the "Predicted text received" in Notes.

My responses after fixing any bugs are in the are in the Responses column of each spreadsheet.

---


### What Claude found (as of 12 October 2026)

**As uploaded: 143 pass, 226 fail, 193 crash.** Five bugs cause nearly all of it. Fixing them gives **313 pass, 56 fail**, and the 56 are parsing leniency or design questions (below).

### Bugs (verified by applying the fix)
1. **`theDefaultYear` is not defined in `hasDate()`** (lines 939, 981). It is declared only inside `generateICal()`. With Day + Month columns, the header row triggers a ReferenceError, so every such export fails. Fix: add `const theDefaultYear = document.getElementById('year-dropdown').value;` at the top of `hasDate()`.
2. **`matchTime` is undeclared in `generateICal()`** (lines 1174–1202). Under `'use strict'`, any row with an End or Duration value throws. Fix: add `let matchTime;` beside `let theEndMin;`.
3. **Months October–December from a month *name* give a corrupt DTSTART** such as `202701002T…`. `theMonth = 0 + mth` (lines 951, 1079) concatenates strings: `0 + '10'` is `"010"`. Fix: `theMonth = mth;`.
4. **Short rows crash** (`line[monthCol]` or the Day cell is `undefined` when trailing cells are blank, so `.toUpperCase()` / `.match()` throws). This hits title and heading rows too, and the user sees the misleading "use https" message. Fix: `line[monthCol] || ''` and `(line[columns.get(CDAY)] || '').match(...)` in both `hasDate()` and `generateICal()`.
5. **Blank trailing HW cell gives `, HW=undefined`** in the event name. Fix: `(line[columns.get(CHW)] || '')`.

### Remaining 56 failures: your decisions
- **Text or custom-format slash dates are read month-first.** `dd/mm/yyyy` custom formats (D rows 4–6; common in LibreOffice files) become 1 Feb instead of 2 Jan, silently. `13/1/2027` is rejected.
- **Plain whole-hour durations are rejected** (`3`, `2`, `1`, `5`), although the tooltip says "e.g. 3 or 2:30". `TIME_PATT` needs two minute digits. The Default duration box has the same problem: typing `3` resets it to 2:00.
- **TBA/TBC/N/A events end at start + 2 h** (09:00–11:00), not 17:00 as `TBAend_hour` suggests, because the default-duration branch overwrites it.
- **Minutes are not range-checked:** start `10:60`/`10:75`/`99:99` produce a corrupt DTSTART; end `12:60`, duration `2:60`/`2:75` are silently carried over; zero duration `0:00` gives a zero-length event.
- **Lenient parsing accepts invalid text silently:** day `2.5`, `123`, `2st`, `2xyz`; months `Janvier`, `Junk`, `Marathon`, `Jan-Feb`, `Mayhem` (only the first 3 letters count); date ranges and `2 and 3 January`; wrong weekdays (`Mon 2 January 2027`); End `12:00-13:00`, `1200abc`.
- **False rejects of plausible forms:** ISO `2027-01-02`, `2-Jan-27` (Excel's built-in d-mmm-yy), `02-01-2027`, `02.01.2027`, `Jan.` with a full stop, `Sat 02/01/2027`, and a date plus time in one cell. These are safe failures.
- **Excel times in `h:mm:ss`, `h:mm AM/PM` and General** produce an "Ignoring line" warning (safe). But a General fraction duration `0.1041…` (2:30) is silently read as 0h10, and a time shown as `0.42` as 00:42.
