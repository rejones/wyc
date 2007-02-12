#!/usr/bin/perl

# Convert WYC racing schedule into format that can be loaded into Palm DateBook
# Two formats are provided, one suitable for converting to .dba with convdb on Windows,
# and one vCal format for importing directly into the MacOS X Palm Desktop.

use strict;
use Getopt::Long;
use vars qw($opt_h $opt_n $opt_v);

my $usage = "Usage: wyc.pl [-h] [-v] [-n [note]] [year] < in.csv > out.txt";

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

# This note provide a colour and icon for REJ's DateBk5 calendar
my $NOTE = '##@@PC@@@A@@@@@@@@p=0D=0A';

sub help($);
sub printDBAhdr();
sub printDBA($$$$$$);



GetOptions("h"   => \$opt_h,
           "n:s" => \$opt_n,
           "v"   => \$opt_v);
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
  if (exists $months{$date}) {
    $month = $months{$date};
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
  my @line = split /,/;
  my $date = $line[$DAY];

  # update the month
  if (exists $months{$date}) {
    $month = $months{$date};
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
  if ($opt_v) {
    printVCAL($num, $month, $hour, $min, $line[$EVENT], $line[$HW]);
  } else {
    printDBA($num, $month, $hour, $min, $line[$EVENT], $line[$HW]);
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
  my $OOZ = '00Z';
  my $note = defined $opt_n ? $NOTE : ''; 
  my $start = $hour.$min;
  my $day = sprintf "%4d%02d%02d", $YEAR, $month, $num;
  my $end = sprintf "%02d%s", $hour+$DURATION, $min; # assume $hour+$DURATION < 2400
  my $alarm = sprintf "%02d%s", $hour-1, $min;       # assume nothing starts before 01:00!
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
to file suitable for input to convdb.
Options:
  -h 		Print this help.
  -n [note]	Add note to each entry in the calendar. If no note given, use default for REJ's DateBk5.
  -v		Use vCal format rather than .dba.
EOH
} 
  
