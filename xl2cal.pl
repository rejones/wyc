#!/usr/bin/perl

# Convert WYC racing schedule into format that can be loaded into a calendar.
# Three formats are provided. The default is iCalendar (.ics) suitable for
# importing into e.g. Google Calender or Apple Calendar. Apple Calendar will
# also accept vCalendar format (.vcs). Finally, the tool also supports the
# obsolete DataBook format used by Palm Pilots (which may need to be converted 
# to .dba with convdb on Windows).

use strict;
use Getopt::Long;
#use String::Approx qw(amatch)
use DateTime;
use UUID 'uuid';
use Date::Manip qw(ParseDate Date_Init Date_DaysInMonth);
&Date_Init("DateFormat=non-US");
use Data::Dumper;
use vars qw($opt_h $opt_v $opt_d $opt_tz $opt_n);

sub help($);
sub bad($$);
sub AA_int($);
sub printDBAhdr();
sub printCALhdr();
sub printDBA($$$$$$$);
sub printVCAL($$$$$$$$);
sub printICAL($$$$$$$$);
sub printEOCAL();
sub sanity($$$$);

my $USAGE = 'Usage: wyc.pl [-h] [-v] [--datebook] [--tz timezone] --col field=column_number... [calendars] [year] < in.csv > out.ics';

# Print help message
sub help($) {
  my $usage = shift;
  print <<"EOH"
$usage
Convert CSV output grabbed by Excel from the WYC racing schedule
to a file suitable for a calendar app. The default output format is
iCalendar (.ics) as per RFC 5545. Other supported but old/obsolete 
formats are:
  vCalendar (.vcs)
  DateBook for Palm mobile devices (.dba)

Note: Do not put commas in any cell in the spreadsheet as this will
confuse the CSV file.

The required calendar arguments are the column numbers of the key fields
and, optionally, the types of event to include, and/or the year 
if the calendar is being generated for a year other than the current one.

Columns are specified by '--col field=column_number' where
field           Day, Month, Date, HW, Start, End, Duration, Event, Calendar.
column_number   Specified either alphabetically or numerically 
                (counting from 1 or A).
Field names are case-insensitive.

Mandatory fields that MUST be specified by --cols are:
(1) The date column(s) of the event must be specified by
    EITHER --col Day=num and --col Month=num OR --col Date=num. 
    -c can be used instead of --col.
    If Day is specified, an entry in the Day column may be  
    a cardinal (e.g. 6) or an ordinal (e.g. 6th) number.
    If Month is specified, the entries should be either a cardinal number
    (e.g. 6) or a month name or prefix (e.g. June or Jun).
    If a Date is specified, almost any format is accepted
(2) The Start column. Entries in the start column should usually be times 
    using a 24-hour clock. Acceptable entries include 1400, 14:00 and 14.00.
    However, TBA, TBC, NA and N/A are also accepted.
    Lines without valid Start entries (such as headers, blank lines) 
    are ignored.
(3) The Event column. 

Optional fields are:
(1) The high-water (HW) columns. An entry in this column (e.g. 1500), 
    if any, is appended to the entry in the Event column (e.g. Race 1)
    to give e.g. (Race 1, HW=1500).
(2) An End or Duration column. If neither is provided, events are given
    a default duration.
(3) A Calendar column. This is intended to allow separate calendars 
    to be generated from a single spreadsheet. If no Calendar column
    is specified, all events are generated. Otherwise, you must also
    provide a list of the event types to include in the calendar.
    For example, if you have events for a Main calendar, a Cadets calendar
    and for both, the command-line for the Cadets calendar might look 
    something like this.
      xl2cal.pl --col Date=3 --col HW=11 --col Start=13 --col Event=14 --col Calendar=20 Cadet,Both
    Note that there are no spaces around the comma.
(4) A time zone (default: Europe/London)
    
Other options:
  -h 		Print this help.
  -v		Use .vcs format rather than iCalendar.
  --datebook    Use .dba format rather than iCalendar.
  -n note       Used only for DateBook to provide a colour and icon
EOH
} 


my %months = (
  'JAN' => 1,
  'FEB' => 2,
  'MAR' => 3,
  'APR' => 4,
  'MAY' => 5,
  'JUN' => 6,
  'JUL' => 7,
  'AUG' => 8,
  'SEP' => 9,
  'OCT' => 10,
  'NOV' => 11,
  'DEC' => 12 );
my $MONTHS_PREFIX_LEN = 3;

# Hash of column numbers (either as A,B,C... or 1,2,3... of necessary fields.
# To be consistent with Excel etc, count from A/1 on the CLI but from 0
# internally.
my %columns = (
#  'DAY' => 5, 
#  'MONTH' => 8,
#  'HW' => 11,
#  'START' => 13,
#  'EVENT' => 14, 
#  'CALENDAR' => 20,
##  'RACES' => 7,    # number of races
##  'RACE_NO' => 13, # e.g. 1 or 5&6
##  'CB' => 14,      # (CB) or blank
);

my $VERSION = '2.0'; # for iCalendar
my $TZ = 'Europe/London';
my $DURATION = 2;   # hours for a single race <-----------------------------------------------
my $H = 'H';        # needed for the DURATION value
my $ADVANCE = 2;    # 2 hours warning

# To Be Announced
my $TBAhour = '09';             # so set the period to be 9.00 - 17.00
my $TBAend_hour = '17';
my $TBAstring = ' (times TBC)'; # string to add to TBA events
# Not Applicable
my $NAhour = '09';              # so set the period to be 9.00 - 17.00
my $NAend_hour = '17';

# This note provided a colour and icon for the DateBk5 calendar
my $note = '##@@PC@@@A@@@@@@@@p=0D=0A';

# Hash of events to generate calendar for
my %calendar = ();

GetOptions('h'   => \$opt_h,                    # help
           'col=s' => sub {                     # Specify columns
                        my @kv = split /=/, @_[1];
                        if ($kv[0] =~ /HIGH\s*WATER/i) {
                          $kv[0] = 'HW';
                        }
                        $columns{uc $kv[0]} = AA_int($kv[1]);
                      }, 
           'v'   => \$opt_v,                    # vCalendar
           'datebook'   => \$opt_d,             # Palm DateBook
           'tz=s'   => \$TZ,                    # Time zone
           'n=s' => \$opt_n                     # Add note for DateBook
	   );

if ($opt_h) {
  help($USAGE);
  exit 0;
}

if (defined $opt_n) {
  warn "-n option only used by DateBook\n" unless defined $opt_d;
  $note = $opt_n;
}

$VERSION = '1.0' if (defined $opt_v); # for vCalendar

# Check that the required columns are present
die "You must specify either --col DAY=num and --col MONTH=num, or --col DATE=num\n"
  unless (exists($columns{'DAY'}) && exists($columns{'MONTH'})) || exists($columns{'DATE'});
die "You must specify --col START=num\n" unless exists($columns{'START'});
die "You must specify --col EVENT=num\n" unless exists($columns{'EVENT'});

# Deal with remaining optional arguments: the Year and/or a list of Calendars
my $year = 1900 + (localtime)[5];
die 'Too many arguments' if $#ARGV > 1;
while (my $a = shift(@ARGV)) {
  if ($a =~ /^\d{4}$/) { # looks like Year
    die "Bad year: \"$a\"" if ($a < 2006 || $a > 2100);
    $year = $a;
  } 
  else { # should be a list of calendars
    foreach my $c (split(/,/, $a)) {
      if ($c =~ /^\w+$/) { # looks like a calendar
        $calendar{$c} = 1;
      } else {
        die "\"$c\" is not a valid calendar name (only letters and numbers allowed)";
      }
    }
  }
}
     
my $T = 'T';
my $Z = 'Z';      # 'Z' means UTC rather than local time
my $DTSTAMP;


# Print header for calendars
if ($opt_d) {
  printDBAhdr();
} else {
  my $dt = DateTime->now;
  $dt->set_time_zone('UTC');
  $DTSTAMP = $dt->ymd('').$T.$dt->hms('');
  printCALhdr();
}

# Process the source file
while (<>) {
  chomp;
  $_ =~ s/\r//g;        # if there is a stray carriage return
  $_ =~ s/"//g;         # excel sometimes puts ".." around entries
  my @line = split /,/;

  # Defend against regex DOS attacks
  foreach my $i (0 .. $#line) {
    $line[$i] =~ s/^\s+|(?<=!\s)\s+$//; # otherwise trailing spaces is polynomial
  }

  # Discard header and any other unlikely lines
  die "Start column is too large\n" if $columns{'START'} > $#line;
  unless ($line[$columns{'START'}] =~ /\d+[:\.]?\d+|TB[AC]|N\/?A/i) { # doesn't look like a time
    warn "Ignoring line $.: $_\n";
    next;
  }

  # Discard this line unless the entry is for this calendar
  if (exists($columns{'CALENDAR'})) {
    die 'Calendar column-specifier is too large' if $columns{'CALENDAR'} > $#line;
    next unless exists($calendar{$line[$columns{'CALENDAR'}]});
  }

  # Get the day and month
  my $day;
  my $num;
  my $month;
  if (exists($columns{'DAY'}) && exists($columns{'MONTH'})) {
    # get the month number
    $month = ($line[$columns{'MONTH'}] =~ /^\d\d?$/) ?
      $line[$columns{'MONTH'}] :
      0 + $months{substr(uc $3,0, $MONTHS_PREFIX_LEN)};
    unless (($month >= 1 && $month <= 12)) {
      bad("Invalid month $month $line[$columns{'MONTH'}]", $.); 
      next; 
    }

    # get the day number
    $num = $line[$columns{'DAY'}];
    unless ($num =~ /^\d\d?$/) {
      bad("Cannot understand day number \"$num\"", $.); 
      next; 
    }
    unless (($num >= 1) && ($num <= Date_DaysInMonth($month, $year))) {
      bad ("$num is out of range for a day in month $month", $.); 
      next; 
    }
  
  } elsif (exists($columns{'DATE'})) {
    # parse a date
    my $date = ParseDate($line[$columns{'DATE'}]);
    if (! $date) {
      bad("Bad date $line[$columns{'DATE'}]", $.);
      next;
    }
    if ($date =~ /^(\d{4})(\d{2})(\d{2})/) {
      warn("Year doesn't match on line $.\n") unless ($1 == $year);
      $month = $2;
      $num = $3;
    }
    else {
      bad('Cannot find the date', $.);
      next;
    }
  } else {
    die "PANIC on line $.";
  }

  # Get the event
  my $event = $line[$columns{'EVENT'}];

  # Get the start and end times
  # times before 1000 are sometimes recorded with only 3 digits
  my $hour;
  my $min;
  my $end_hour;
  my $end_min;

  if ($line[$columns{'START'}] =~ /^(\d\d?)[:\.]?(\d\d)/) {
    $hour = $1;
    $min = $2;
    if (exists($columns{'END'}) && ($line[$columns{'END'}] =~ /^(\d\d?)[:\.]?(\d\d)/)) {
      $end_hour = sprintf("%02d", $1);
      $end_min = $2;
    } elsif (exists($columns{'DURATION'}) && ($line[$columns{'DURATION'}] =~ /^(\d\d?)[:\.]?(\d\d)/)) {
      $end_hour = $hour + sprintf("%02d", $1);
      $end_min = $min + $2;
      while ($end_min > 60) {
        $end_hour = $end_hour + 1;
        $end_min = $end_min - 60;
      }
    } else {
      # assume more than one race if event string includes 
      # more than one separate number, and allow an extra hour 
      $end_hour = $hour + (($event =~ /\d\D+\d/) ? $DURATION + 1 : $DURATION);
      $end_min = $min;
      if ($end_hour > 24) {
        warn "Event on line $. cannot span midnight! $event, $hour\n";
        $end_hour = '23';
        $end_min = '59';
      }
    }
    $hour = sprintf("%02d", $hour);
  } 
  elsif ($line[$columns{'START'}] =~ /TBA|TBC|-/i || $line[$columns{'START'}] eq '') {
    $hour = $TBAhour;
    $min = '00';
    $end_hour = $TBAend_hour;
    $end_min = '00';
    $event = $event . $TBAstring;
  }
  elsif ($line[$columns{'START'}] =~ /N\/?A/i) {
    $hour = $NAhour;
    $min = '00';
    $end_hour = $NAend_hour;
    $end_min = '00';
  }
  else {
    bad("Cannot understand Start time $line[$columns{'START'}]", $.); 
    next; 
  }


  # Get highwater time
  my $highwater = $line[$columns{'HW'}];
  
  #print the record
  if ($opt_d) {
    my $duration = $end_hour - $hour; # FIXME or omit DateBook
    printDBA($num, $month, $hour, $min, $duration, $event, $highwater);
  } elsif ($opt_v) {
    printVCAL($num, $month, $hour, $min, $end_hour, $end_min, $event, $highwater);
  } else {
    printICAL($num, $month, $hour, $min, $end_hour, $end_min, $event, $highwater);
  }
}

#print trailer
printEOCAL() unless $opt_d;


# SUBROUTINES -----------------------------------------

# Bad input warning
sub bad($$) { 
  my ($msg, $lno) = @_;
  warn ("BAD RECORD: $msg on line $lno so ignoring this line.\n");
}


# Convert hour to UTC
# Always output UTC times. If we were to write TZ=timezone, then 
# strict conformance with the  RFC 5545 specification) requires that 
# this timezone be fully specified (including BST and GMT start dates and
# offsets, etc).  Of course, these dates are the last Sundays in March/October,
# so they change... :-(
sub convUTC($$$$) {
  my ($year, $month, $day, $hour) = @_;
  my $dt = DateTime->new(
    year      => $year,
    month     => $month,
    day       => $day,
    hour      => $hour,
    time_zone => $TZ,
  );
  $dt->set_time_zone('UTC');
  return sprintf("%02d", $dt->hour);
}


# Convert column number (whether integer or AA-style) to zero-based integer 
# Allow liberal interpretations of integer
sub AA_int($) {
  my $aa = shift(@_);
  if ($aa =~ /(\d+)/) {
    return $1 - 1;
  } else {
    my $val = 0;
    while ($aa =~ /([a-z])/ig) {
      my $ch = uc $1;
      print "$ch\n";
      $val = $val * 26 + (ord($ch) - ord('A')) + 1;
    }
    die "Bad column number: $aa" unless $val > 0;
    return $val - 1;
  }
}


# Write header

# DBA format is:
# day/month/year	hour:minute	duration	details (HW=highwater)
sub printDBAhdr () {
  print "#WYC racing schedule $year\n";
  printf "%s\t%s\t%s\t%s", '%d/%m/%y', '%h:%i', '%t', '%v';
  print "\t%n" if defined $opt_n;
  print "\n";
}


# iCalendar and vCalendar header
# TODO nice to include X-WR-CALNAME: property
sub printCALhdr() {
  print <<"EOH"
BEGIN:VCALENDAR
VERSION:$VERSION
PRODID:Richard Jones xl2cal.pl generated
CALSCALE:GREGORIAN
EOH
}


# Print .dba entry
sub printDBA ($$$$$$$){
  my($num, $month, $hour, $min, $duration, $event, $highwater) = @_;
  my $hour = $1;
  my $min = $2;
  print "$num/$month/$year\t";
  print "$hour:$min\t$duration\t";
  print  "WYC $event";
  printf ", HW=%s", $highwater unless $highwater eq '';
  print "\n";
  print "$note\n.\n" if defined $opt_n;
}  

# Print vCalendar entry
sub printVCAL ($$$$$$$$){
  my($num, $month, $start, $min, $end, $end_min, $event, $highwater) = @_;
  my $hw = $highwater eq '' ? '' : ", HW=$highwater";
  my $day = sprintf "%4d%02d%02d", $year, $month, $num;
  $start = convUTC($year, $month, $num, $start);
  $end = convUTC($year, $month, $num, $end);
  my $alarm = $start - $ADVANCE;
  if ($alarm < 0) { 
    warn "Alarm set for previous day: $event\n";
    $alarm = "000000";
  } else { 
    $alarm = sprintf "%04d00", $alarm;
  }
  # sanity check
  #sanity($day, $start, $end, $alarm);
  $start = $start.$min.'00';
  $end = $end.$end_min.'00';
  print <<"EOV"
BEGIN:VEVENT
SUMMARY:WYC $event$hw
DTSTAMP:$DTSTAMP$Z
DTSTART:$day$T$start$Z
DTEND:$day$T$end$Z
END:VEVENT
EOV
#DALARM:$day$T$alarm$Z
}


# Print iCalendar entry
sub printICAL ($$$$$$$$){
  my($num, $month, $start, $min, $end, $end_min, $event, $highwater) = @_;
  my $hw = $highwater eq '' ? '' : ", HW=$highwater";
  my $summary = $event.$hw;
  $summary =~ s/.{63}\K/\n /sg; # fold lines longer than 75 octets
  my $day = sprintf "%4d%02d%02d", $year, $month, $num;
  $start = convUTC($year, $month, $num, $start);
  $end = convUTC($year, $month, $num, $end);
  my $alarm = $start - $ADVANCE;
  if ($alarm < 0) { 
    warn "Alarm set for previous day: $event\n";
    $alarm = "000000";
  } else { 
    $alarm = sprintf "%04d00", $alarm;
  }
  # sanity check
  #sanity($day, $start, $end, $alarm);
  $start = $start.$min.'00';
  $end = $end.$end_min.'00';
  my $uid = uuid();
  print <<"EOI"
BEGIN:VEVENT
CREATED:$DTSTAMP$Z
UID:$uid
DTSTAMP:$DTSTAMP$Z
DTSTART:$day$T$start$Z
DTEND:$day$T$end$Z
SUMMARY:WYC $summary
END:VEVENT
EOI
#DALARM:$day$T$alarm$Z
}


# Calendar and vCalendar trailer
sub printEOCAL() {
  print "END:VCALENDAR\n";
}

# Sanity check on lengths
# TODO do we need this?
sub sanity($$$$) {
  my ($day, $start, $end, $alarm) = @_;
  die "Bad length for day $day\n"     unless length($day) == 8;   # yyyymmdd 
  die "Bad length for start $start\n" unless length($start) == 6; # hhmm00
  die "Bad length for end $end\n"     unless length($end) == 6;   # hhmm00
  die "Bad length for alarm $alarm\n" unless length($alarm) == 6; # hhmm00
}

=begin comment
Improvements
/. Add to the spreadsheet a CALENDAR column with values e.g. Main, Cadets, 
   Both. Use this either to filter in Excel or add it to the user interface
   and filter in the tool.
/. Ensure that times in an format (hhmm, hh:mm, any more?) can be parsed.
/. Choose date from either a pair of DAY and MONTH columns or a single DATE
   column (e.g. Mon 12 Aug) which would need to be parsed and allow variations.
/. Rather than hard-wired column numbers, get these from the user (but beware
   that Excel counts from A/1) but perl counts from 0.
4. Provide a GUI rather than using the command line. This would overcome the 
   problems in (3) as user could select the spreadsheet, the tab to use, the
   columns and the output file name.
   Window 1: choose spreadsheet and output file name.
   Window 2: select tab then columns.
   Errors window: need a list of problems encountered.
5. Rewriting in Python might help both (3) and (4). 
   E.g. could specify key-value pairs on the commandline for columns.
   https://www.geeksforgeeks.org/python-key-value-pair-using-argparse/amp/
   suggests:

        import argparse 
          
        class keyvalue(argparse.Action): 
            # Constructor calling 
            def __call__(self, parser, namespace, values, option_string = None): 
                setattr(namespace, self.dest, dict()) 

                for value in values: 
                    key, value = value.split('=') 
                    # assign into dictionary 
                    getattr(namespace, self.dest)[key] = value 

        parser = argparse.ArgumentParser() 
        parser.add_argument('--cols', nargs='*', action = keyvalue) 
        args = parser.parse_args() 
        print(args.cols)

   Python also provides a GUI.

6. Alternatively, provide a web service, e.g. using JavaScript.

7. Operate on the .xlsx directly rather than on a .csv. Security issues?
=end comment
