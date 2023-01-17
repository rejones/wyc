#!/usr/bin/perl

# Convert WYC racing schedule into format that can be loaded into a calendar.
# Three formats are provided. The default is iCalendar (.ics) suitable for
# importing into e.g. Google Calender or Apple Calendar. Apple Calendar will
# also accept vCalendar format (.vcs). Finally, the tool also supports the
# obsolete DataBook format used by Palm Pilots (which may need to be converted 
# to .dba with convdb on Windows.

use strict;
use Getopt::Long;
#use String::Approx qw(amatch)
use DateTime;
use UUID 'uuid';
use vars qw($opt_h $opt_v $opt_d $opt_z $opt_n);

sub help($);
sub printDBAhdr();
sub printCALhdr();
sub printDBA($$$$$$$);
sub printVCAL($$$$$$$);
sub printEOCAL();
sub sanity($$$$);

my $USAGE = "Usage: wyc.pl [-h] [-v] [-d] [-z] calendar [year] < in.csv > out.vcs";

# Print help message
sub help($) {
  my $usage = shift;
  print <<"EOH"
$usage
Convert CSV output grabbed by Excel from the WYC racing schedule
to a file suitable for a calendar app. The default output format is
iCalendar (.ics) as per RFC 5545. Other supported but obsolete 
formats are:
  vCalendar (.vcs)
  DateBook for Palm mobile devices (.dba)

The calendar argument is the type of events to generate,
i.e. "Main" or "Cadet".

Options:
  -h 		Print this help.
  -v		Use .vcs format rather than iCalendar.
  -d		Use .dba format rather than iCalendar.
  -z		Use UTC rather than local time.
EOH
} 
  

my %months = (
  "JAN" => 1,
  "FEB" => 2,
  "MAR" => 3,
  "APR" => 4,
  "MAY" => 5,
  "JUN" => 6,
  "JUL" => 7,
  "AUG" => 8,
  "SEP" => 9,
  "OCT" => 10,
  "NOV" => 11,
  "DEC" => 12 );
my $MONTHS_PREFIX_LEN = 3;

# 2023 format
# column numbers 'Programme' sheet of the spreadsheet provided (starting from 0)
my $DAY = 5; 
my $MONTH = 8;
my $HW = 11;
my $START = 13;
my $EVENT = 14; 
my $CALENDAR = 20;
#my $RACES = 7;    # number of races
#my $RACE_NO = 13; # e.g. 1 or 5&6
#my $CB = 14;      # (CB) or blank

my $VERSION = '2.0';
my $TZ = 'TZID=Europe/London';
my $DURATION = 2; # hours for a single race
my $H = 'H';      # needed for the DURATION value
my $ADVANCE = 2;  # 2 hours warning

# not announced yet
my $TBAhour = '10';             # so set the period to be 10.00 - 16.00
my $TBAmin = '00'; 
my $TBAduration = 6;
my $TBAstring = ' (times TBC)'; # string to add to TBA events
# not applicable
my $NAhour = '10';              # so set the period to be 10.00 - 16.00
my $NAmin = '00'; 
my $NAduration = 6;

# This note provided a colour and icon for REJ's DateBk5 calendar
#my $NOTE = '##@@PC@@@A@@@@@@@@p=0D=0A';
my $note = '';


GetOptions("h"   => \$opt_h,
           "v"   => \$opt_v,
           "d"   => \$opt_d,
           "z"   => \$opt_z,
           "n:s" => \$opt_n
	   );
if ($opt_h) {
  help($USAGE);
  exit 0;
}

$note = $opt_n if defined $opt_n;

# Set the type of events, Main or Cadet
my $calendar = shift @ARGV;
unless (($calendar eq 'Main') || ($calendar eq 'Cadet')) {
  help($USAGE);
  exit 1;
}

# Optional year, otherwise use current year
my $YEAR = (@ARGV == 1) ?
             shift @ARGV :
             1900 + (localtime)[5];
die "Bad year: \"$YEAR\"" if ($YEAR < 2006 || $YEAR > 2100);

my $T = 'T';
my $Z = defined $opt_z ? 'Z' : '';      # 'Z' means UTC rather than local time
my $DTSTAMP;

# From 2013, WYC no longer has the month on a line on its own;
# instead, each entry is ddd [n]n month, e.g. Mon 1 April
#my $month = 1;
#
#skip source file headers
#while (<>) {
#  my @line = split /,/;
#  my $date = $line[$DAY];
#  my $dateUC = uc $date;
#  if (exists $months{$dateUC}) {
#    $month = $months{$dateUC};
#    last;
#  }
#}

# Print header
if ($opt_d) {
  printDBAhdr();
} else {
  my $dt = DateTime->now;
  $DTSTAMP = $dt->ymd('').$T.$dt->hms('');
  printCALhdr();
}

# Process the source file
while (<>) {
  next if $. < 2;       # discard header
  chomp;
  $_ =~ s/"//g;         # excel sometimes puts ".." around entries
  my @line = split /,/;

  # Is the entry for this calendar?
  next unless ($line[$CALENDAR] =~ /Both|\Q$calendar/);

#  my $date = $line[$DAY];

#  # update the month
#  my $dateUC = uc $date;
#  if (exists $months{$dateUC}) {
#    $month = $months{$dateUC};
#    next;
#  }

#  # convert the day
#  # unfortunately the WYC data is unreliably formatted here
#  my $day;
#  my $num;
#  my $month;
#  if ($date =~ /^([A-Za-z]+)\s*([0-9]+)[snrt][tdh]\s*([A-Za-z]+)/) { #eg Sun 6th Mar
#  	$day = $1;
#  	$num = $2;
#       $month = $months{substr(uc $3,0, $MONTHS_PREFIX_LEN)};
#  }
#  else {warn "header or BAD RECORD \"$_\" at line $.\ndate=\"$date\"\nCannot parse date.\n"; next; }
  
  # get the day number
  my $num = $line[$DAY];
  unless ($num =~ /^\d\d?$/) {
    warn "BAD RECORD \"$_\" at line $.\nCannot parse day number \"$num\".\n"; 
    next; 
  }
  unless (($num >= 1) && ($num <= 31)) {
    warn "BAD RECORD \"$_\" at line $.\nCannot parse day number.\n"; 
    next; 
  }

  # get the month number
  my $month = $line[$MONTH];
  unless ($month =~ /^\d+\d?$/) {
    warn "BAD RECORD \"$_\" at line $.\nCannot parse month number.\n"; 
    next; 
  }
  unless (($month >= 1 && $month <= 12)) {
    warn "BAD RECORD \"$_\" at line $.\nCannot parse month number.\n"; 
    next; 
  }

  # get the event
  my $event = $line[$EVENT];
  #$event = "$event $line[$RACE_NO]" unless $line[$RACE_NO] eq '';
  #$event = "$event $line[$CB]" unless $line[$CB] eq '';
  $event =~ s/\s+$//g;
  $event =~ s/\s+,/,/g;

#  # sanity check number of races
#  if ($line[$RACES] > 1) {
#    unless ($event =~ /\d\&\d/) {
#      warn "BAD RECORD \"$_\" at line $.\nWrong number of races\n"; 
#      next;
#    }
#  }

  # get the start time
  # there has been no WYC consistency here :-(
  my $hour;
  my $min;
  if ($line[$START] =~ /^(\d\d):?(\d\d)/) {
    $hour = $1;
    $min = $2;
#    $duration = $DURATION + ($line[$RACES] - 1); # Allow extra hour for each additional race
  } 
  # times before 1000 are sometimes recorded with only 3 digits
  elsif ($line[$START] =~ /^(\d):?(\d\d)/) {
    $hour = 0 . $1;
    $min = $2;
#    $duration = $DURATION + ($line[$RACES] - 1); # Allow extra hour for each additional race
  } 
  elsif ($line[$START] =~ /TBA|TBC|-/i || $line[$START] eq '') {
    $hour = $TBAhour;
    $min = $TBAmin;
    $event = $event . $TBAstring;
 #   $duration = $TBAduration;
  }
  elsif ($line[$START] =~ /N\/?A/i) {
    $hour = $NAhour;
    $min = $NAmin;
 #   $duration = $TBAduration;
  }
  else { warn "BAD RECORD \"$_\" at line $.\nCannot parse time.\n"; next; }

  # allow extra hour if more than one race
  my $duration = ($event =~ /\d\/\d/) ? $DURATION + 1 : $DURATION;

  # get highwater time
  my $highwater = $line[$HW];
  
  #print the record
#  my $highwater = sprintf("%04d", $line[$HW]);

  if ($opt_d) {
    printDBA($num, $month, $hour, $min, $duration, $event, $highwater);
  } elsif ($opt_v) {
    printVCAL($num, $month, $hour, $min, $duration, $event, $highwater);
  } else {
    printICAL($num, $month, $hour, $min, $duration, $event, $highwater);
  }
}

#print trailer
printEOCAL() unless $opt_d;


# SUBROUTINES -----------------------------------------

# write header
# format is:
# day/month/year	hour:minute	duration	details (HW=highwater)
sub printDBAhdr () {
  print "#WYC racing schedule $YEAR\n";
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
PRODID:Richard Jones wyc.pl generated
X-WR-TIMEZONE:Europe/London
CALSCALE:GREGORIAN
EOH
}

# Print .dba entry
sub printDBA ($$$$$$$){
  my($num, $month, $hour, $min, $duration, $event, $highwater) = @_;
  my $hour = $1;
  my $min = $2;
  print "$num/$month/$YEAR\t";
  print "$hour:$min\t$duration\t";
  print  "WYC $event";
  printf ", HW=%s", $highwater unless $highwater eq '';
  print "\n";
  print "$note\n.\n" if defined $opt_n;
}  

# Print vCalendar entry
sub printVCAL ($$$$$$$){
  my($num, $month, $hour, $min, $duration, $event, $highwater) = @_;
  my $hw = $highwater eq '' ? '' : ", HW=$highwater";
  # specify times as local if we span DST
  my $start = $hour.$min.'00';
  my $day = sprintf "%4d%02d%02d", $YEAR, $month, $num;
  my $end_hour = $hour + $duration;
  my $end;
  if ($end_hour >= 24) {        # assume $hour+$DURATION < 2400
    warn "Event spans midnight! $event, $hour, $duration\n";
    $end = "235900";
  } else {
    $end = sprintf "%02d%s00", $end_hour, $min;
  }
  my $alarm_hour = $hour - $ADVANCE;
  my $alarm;
  if ($alarm_hour < 0) {        # assume nothing starts before ADVANCE:00!
    warn "Alarm set for previous day: $event\n";
    $alarm = "000000";
  } else { 
    $alarm = sprintf "%02d%s00", $alarm_hour, $min;
  }
  # sanity check
  sanity($day, $start, $end, $alarm);
  print <<"EOV"
BEGIN:VEVENT
SUMMARY:WYC $event$hw
DTSTAMP;$TZ:$DTSTAMP
DTSTART;$TZ:$day$T$start
DTEND;$TZ:$day$T$end
END:VEVENT
EOV
#DESCRIPTION;QUOTED-PRINTABLE:$note
#DALARM:$day$T$alarm$Z
}


# Print iCalendar entry
sub printICAL ($$$$$$$){
  my($num, $month, $hour, $min, $duration, $event, $highwater) = @_;
  my $hw = $highwater eq '' ? '' : ", HW=$highwater";
  # specify times as local if we span DST
  my $start = $hour.$min.'00';
  my $day = sprintf "%4d%02d%02d", $YEAR, $month, $num;
  my $end_hour = $hour + $duration;
  my $end;
  if ($end_hour >= 24) {        # assume $hour+$DURATION < 2400
    warn "Event spans midnight! $event, $hour, $duration\n";
    $end = "235900";
  } else {
    $end = sprintf "%02d%s00", $end_hour, $min;
  }
 my $alarm_hour = $hour - $ADVANCE;
 my $alarm;
 if ($alarm_hour < 0) {         # assume nothing starts before ADVANCE:00!
   warn "Alarm set for previous day: $event\n";
   $alarm = "000000";
 } else { 
   $alarm = sprintf "%02d%s00", $alarm_hour, $min;
 }
  # sanity check
  sanity($day, $start, $end, $alarm);
  my $uid = uuid();
  print <<"EOI"
BEGIN:VEVENT
CREATED;$TZ:$DTSTAMP
UID:$uid
DTSTAMP;$TZ:$DTSTAMP
DTSTART;$TZ:$day$T$start
DTEND;$TZ:$day$T$end
SUMMARY:WYC $event$hw
END:VEVENT
EOI
#  print "DESCRIPTION;QUOTED-PRINTABLE:$note\n" unless $note eq '';
#DALARM:$day$T$alarm$Z
}


# Calendar and vCalendar trailer
sub printEOCAL() {
  print "END:VCALENDAR\n";
}

# Sanity check on lengths
sub sanity($$$$) {
  my ($day, $start, $end, $alarm) = @_;
  die "Bad length for day $day\n"     unless length($day) == 8;   # yyyymmdd 
  die "Bad length for start $start\n" unless length($start) == 6; # hhmm00
  die "Bad length for end $end\n"     unless length($end) == 6;   # hhmm00
  die "Bad length for alarm $alarm\n" unless length($alarm) == 6; # hhmm00
}

=begin comment
Improvements
0. Add to the spreadsheet a CALENDAR column with values e.g. Main, Cadets, 
   Both. Use this either to filter in Excel or add it to the user interface
   and filter in the tool.
1. Ensure that times in an format (hhmm, hh:mm, any more?) can be parsed.
2. Choose date from either a pair of DAY and MONTH columns or a single DATE
   column (e.g. Mon 12 Aug) which would need to be parsed and allow variations.
3. Rather than hard-wired column numbers, get these from the user (but beware
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

6. Operate on the .xlsx directly rather than on a .csv.
=end comment
