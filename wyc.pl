#!/usr/bin/perl

# Convert WYC racing schedule into format that can be loaded into Palm DateBook
# Two formats are provided, one suitable for converting to .dba with convdb on Windows,
# and one vCal format for importing directly into the MacOS X Palm Desktop.

use strict;
use Getopt::Long;
use vars qw($opt_h $opt_n $opt_v $opt_z);

my $usage = "Usage: wyc.pl [-h] [-v] [-z] [-n [note]] [year] < in.csv > out.txt";

my %months = (
  "JANUARY" => 1,
  "FEBRUARY" => 2,
  "MARCH" => 3,
  "APRIL" => 4,
  "MAY" => 5,
  "JUNE" => 6,
  "JULY" => 7,
  "AUGUST" => 8,
  "SEPTEMBER" => 9,
  "OCTOBER" => 10,
  "NOVEMBER" => 11,
  "DECEMBER" => 12 );

my $DAY = 0;
my $HW = 1;
my $START = 2;
my $EVENT = 3; 
my $DURATION = 3;
my $ADVANCE = 2; # 2 hours warning

# This note provide a colour and icon for REJ's DateBk5 calendar
my $NOTE = '##@@PC@@@A@@@@@@@@p=0D=0A';

sub help($);
sub printDBAhdr();
sub printVCALhdr();
sub printDBA($$$$$$);
sub printVCAL($$$$$$);
sub printEOVCAL();



GetOptions("h"   => \$opt_h,
           "n:s" => \$opt_n,
           "v"   => \$opt_v,
           "z"   => \$opt_z
	   );
if ($opt_h) {
  help($usage);
  exit 0;
}

if ($opt_n) {
  $NOTE = $opt_n unless $opt_n eq '';
}

my $YEAR = (@ARGV == 1) ?
             shift @ARGV :
             1900 + (localtime)[5];
die "Bad year: $YEAR" if ($YEAR < 2006 || $YEAR > 9999);
my $month = 1;

#skip source file headers
while (<>) {
  my @line = split /,/;
  my $date = $line[$DAY];
  my $dateUC = uc $date;
  if (exists $months{$dateUC}) {
    $month = $months{$dateUC};
    last;
  }
}

# Print header
if ($opt_v) {
  printVCALhdr();
} else {
  printDBAhdr();
}

# Process the source file
while (<>) {
  chomp;
  my @line = split /,/;
  my $date = $line[$DAY];

  # update the month
  my $dateUC = uc $date;
  if (exists $months{$dateUC}) {
    $month = $months{$dateUC};
    next;
  }

  # convert the day
  # unfortunately the WYC data is unreliably formatted here
  my $day;
  my $num;
  if ($date =~ /^([A-Za-z]+)\s*([0-9]+)/) {
  	$day = $1;
  	$num = $2;
  }
  else {warn 'BAD RECORD '.$_; next; }
  
  #convert the time
  # there is no WYC consistency here either :-(
  my $hour;
  my $min;
  if ($line[$START] =~ /^(\d\d):?(\d\d)/) {
    $hour = $1;
    $min = $2;
  } 
  else { warn 'BAD RECORD '.$_; next; }
  
  #print the record
  my $highwater = sprintf "%04d", $line[$HW];
  if ($opt_v) {
    printVCAL($num, $month, $hour, $min, $line[$EVENT], $highwater);
  } else {
    printDBA($num, $month, $hour, $min, $line[$EVENT], $highwater);
  }
}

#print trailer
printEOVCAL() if $opt_v;


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

# vCal header
sub printVCALhdr() {
  print <<"EOVH"
BEGIN:VCALENDAR
PRODID:Richard Jones wyc.pl generated
TZ:+00
VERSION:1.0
EOVH
}

# Print .dba entry
sub printDBA ($$$$$$){
  my($num, $month, $hour, $min, $event, $highwater) = @_;
  print "$num/$month/$YEAR\t";
  print "$hour:$min\t$DURATION\t";
  print  "WYC $event";
  printf ", HW=%s", $highwater unless $highwater eq '';
  print "\n";
  print "$NOTE\n.\n" if defined $opt_n;
}  

# Print vCal entry
sub printVCAL ($$$$$$){
  my($num, $month, $hour, $min, $event, $highwater) = @_;
  my $hw = $highwater eq '' ? '' : ", HW=$highwater";
  my $T = 'T';
  # 'Z' means UTC rather than local time
  my $OOZ = defined $opt_z ? '00Z' : '00';
  my $note = defined $opt_n ? $NOTE : ''; 
  my $start = $hour.$min;
  my $day = sprintf "%4d%02d%02d", $YEAR, $month, $num;
  my $end_hour = $hour + $DURATION;
  my $end;
  if ($end_hour >= 24) {        # assume $hour+$DURATION < 2400
    warn "Event spans midnight! $event\n";
    $end = "2359";
  } else {
    $end = sprintf "%02d%s", $end_hour, $min;
  }
  my $alarm_hour = $hour - $ADVANCE;
  my $alarm;
  if ($alarm_hour < 0) {       # assume nothing starts before ADVANCE:00!
    warn "Alarm set for previous day: $event\n";
    $alarm = "0000";
  } else { 
    $alarm = sprintf "%02d%s", $alarm_hour, $min;
  }
  print <<"EOV"
BEGIN:VEVENT
SUMMARY:WYC $event$hw
DESCRIPTION;QUOTED-PRINTABLE:$note
DTSTART:$day$T$start$OOZ
DTEND:$day$T$end$OOZ
DALARM:$day$T$alarm$OOZ
END:VEVENT
EOV
}

# vCal trailer
sub printEOVCAL() {
  print "END:VCALENDAR\n";
}

# Print help message
sub help($) {
  my $usage = shift;
  print <<"EOH"
$usage
Convert CSV output grabbed by Excel from WYC racing schedule
to file suitable for input to vCal or convdb.
Options:
  -h 		Print this help.
  -n [note]	Add note to each entry in the calendar. If no note given, use default for REJ's DateBk5.
  -v		Use vCal format rather than .dba.
  -z		Use UTC rather than local time
EOH
} 
  
