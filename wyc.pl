#!/usr/bin/perl

# Convert WYC racing schedule into format that can be loaded into a calendar.
# Two formats are provided, one suitable for converting to .dba with convdb on Windows,
# and a vCalendar format for importing directly into calendars that support this
# format.

use strict;
use Getopt::Long;
use Data::GUID;
use DateTime;
use vars qw($opt_h $opt_n $opt_d $opt_z);

my $usage = "Usage: wyc.pl [-h] [-v] [-z] [-n [note]] [year] < in.csv > out.vcs";

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

# 2021 format
my $DAY = 7; # format e.g Sat 06 Mar
my $HW = 9;
my $START = 10;
my $EVENT = 11; 
my $RACES = 12; # number of races
my $RACE_NO = 13; # e.g. 1 or 5&6
my $CB = 14; # (CB) or blank

my $VERSION = '2.0';
my $TZ = 'TZID=Europe/London';
my $DURATION = 2; # hours for a single race
my $ADVANCE = 2; # 2 hours warning

my $TBA = 'TBC'; # not announced yet
my $TBAhour = '10'; # so set the period to be 10.00 - 16.00
my $TBAmin = '00'; 
my $TBAduration = 6;
my $TBAstring = ' (times TBC)'; # string to add to TBA events

# This note provide a colour and icon for REJ's DateBk5 calendar
#my $NOTE = '##@@PC@@@A@@@@@@@@p=0D=0A';
my $note = '';

sub help($);
sub printDBAhdr();
sub printVCALhdr();
sub printDBA($$$$$$$);
sub printVCAL($$$$$$$);
sub printEOVCAL();
sub sanity($$$$);



GetOptions("h"   => \$opt_h,
           "n:s" => \$opt_n,
           "d"   => \$opt_d,
           "z"   => \$opt_z
	   );
if ($opt_h) {
  help($usage);
  exit 0;
}

$note = $opt_n if defined $opt_n;

my $YEAR = (@ARGV == 1) ?
             shift @ARGV :
             1900 + (localtime)[5];
die "Bad year: \"$YEAR\"" if ($YEAR < 2006 || $YEAR > 9999);

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
  printVCALhdr();
}

# Process the source file
while (<>) {
  next if $. < 2; # discard header
  chomp;
  $_ =~ s/"//g;         # excel sometimes puts ".." around entries
  my @line = split /,/;
  my $date = $line[$DAY];

#  # update the month
#  my $dateUC = uc $date;
#  if (exists $months{$dateUC}) {
#    $month = $months{$dateUC};
#    next;
#  }

  # convert the day
  # unfortunately the WYC data is unreliably formatted here
  my $day;
  my $num;
  my $month;
  if ($date =~ /^([A-Za-z]+)\s*([0-9]+)\s*([A-Za-z]+)/) {
  	$day = $1;
  	$num = $2;
        $month = $months{substr(uc $3,0, $MONTHS_PREFIX_LEN)};
  }
  else {warn "header or BAD RECORD \"$_\" at line $.\ndate=\"$date\"\nCannot parse date.\n"; next; }
  
  #convert the time
  # there is no WYC consistency here either :-(
  my $hour;
  my $min;
  my $event = $line[$EVENT];
  $event = "$event $line[$RACE_NO]" unless $line[$RACE_NO] eq '';
  $event = "$event $line[$CB]" unless $line[$CB] eq '';
  $event =~ s/\s+,/,/g;
  my $duration;
  #TODO time before 0900 are sometimes recorded with only 3 digits
  if ($line[$START] =~ /^(\d\d):?(\d\d)/) {
    $hour = $1;
    $min = $2;
    $duration = $DURATION + ($line[$RACES] - 1); # Allow extra hour for each additional race
  } 
  elsif ($line[$START] eq $TBA || $line[$START] eq '') {
    $hour = $TBAhour;
    $min = $TBAmin;
    $event = $event . $TBAstring;
    $duration = $TBAduration;
  }
  else { warn "BAD RECORD \"$_\" at line $.\nCannot parse time.\n"; next; }
  
  #print the record
  my $highwater = sprintf("%04d", $line[$HW]);

  if ($opt_d) {
    printDBA($num, $month, $hour, $min, $duration, $event, $highwater);
  } else {
    printVCAL($num, $month, $hour, $min, $duration, $event, $highwater);
  }
}

#print trailer
printEOVCAL() unless $opt_d;


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

# vCalendar header
sub printVCALhdr() {
  print <<"EOVH"
BEGIN:VCALENDAR
VERSION:$VERSION
PRODID:Richard Jones wyc.pl generated
EOVH
}

# Print .dba entry
sub printDBA ($$$$$$$){
  my($num, $month, $hour, $min, $duration, $event, $highwater) = @_;
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
  if ($alarm_hour < 0) {       # assume nothing starts before ADVANCE:00!
    warn "Alarm set for previous day: $event\n";
    $alarm = "000000";
  } else { 
    $alarm = sprintf "%02d%s00", $alarm_hour, $min;
  }
  # sanity check
  sanity($day, $start, $end, $alarm);
  my $guid = Data::GUID->new; #todo
  print <<"EOV"
BEGIN:VEVENT
SUMMARY:WYC $event$hw
DESCRIPTION;QUOTED-PRINTABLE:$note
UID:$guid
DTSTAMP;$TZ:$DTSTAMP
DTSTART;$TZ:$day$T$start
DTEND;$TZ:$day$T$end
END:VEVENT
EOV
#DALARM:$day$T$alarm$Z
}

# vCalendar trailer
sub printEOVCAL() {
  print "END:VCALENDAR\n";
}

# Print help message
sub help($) {
  my $usage = shift;
  print <<"EOH"
$usage
Convert CSV output grabbed by Excel from WYC racing schedule
to a file suitable for input to vCalendar or convdb.
Options:
  -h 		Print this help.
  -n [note]	Add note to each entry in the calendar. 
  -d		Use .dba format rather than vCalendar.
  -z		Use UTC rather than local time.
EOH
} 
  
# Sanity check on lengths
sub sanity($$$$) {
  my ($day, $start, $end, $alarm) = @_;
  die "Bad length for day $day\n"     unless length($day) == 8;   # yyyymmdd 
  die "Bad length for start $start\n" unless length($start) == 6; # hhmm00
  die "Bad length for end $end\n"     unless length($end) == 6;   # hhmm00
  die "Bad length for alarm $alarm\n" unless length($alarm) == 6; # hhmm00
}
