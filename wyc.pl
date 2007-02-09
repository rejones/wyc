#!/usr/bin/perl

# Convert WYC racing schedule into format suitable for converting to .dba with convdb

use strict;

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

my $YEAR = 2006;
my $month = 1;


#skip headers
while (<>) {
  my @line = split /,/;
  my $date = $line[$DAY];
  if (exists $months{$date}) {
    $month = $months{$date};
    last;
  }
}
# write header
print "#WYC racing schedule $YEAR\n";
printf "%s\t%s\t%s\t%s\n", '%d/%m/%y', '%h:%i', '%t', '%v';
while (<>) {
  my @line = split /,/;
  my $date = $line[$DAY];

  # update the month
  if (exists $months{$date}) {
    $month = $months{$date};
    next;
  }

  # convert the day
  my ($day, $num, $x) = split /\s+/, $date, 3;
  print "$num/$month/$YEAR\t";
  
  #convert the time
  my ($hour, $min, $xx) = split /:/, $line[$START], 3;
  if ($min eq '') { warn 'BAD RECORD '.$_; next; }
  print "$hour:$min\t3\t";


  # the event
  print  "WYC $line[$EVENT] ";
  printf "(HW=%s)", $line[$HW] unless $line[$HW] eq '';
  print "\n";
}
  
  
  
