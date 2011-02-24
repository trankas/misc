#!/usr/bin/perl -w

use strict;
use Spreadsheet::ParseExcel;

my $excel = new Spreadsheet::ParseExcel;

die "filename to $0 required" unless @ARGV;

my $book = $excel->Parse($ARGV[0]);

my $worksheet = $book->worksheet(0);

my ($row_min, $row_max) = $worksheet->row_range();
my ($col_min, $col_max) = $worksheet->col_range();

for my $row ($row_min .. $row_max) {
	#for my $col ($col_min .. $col_max) {
		my $cell1 = $worksheet->get_cell($row, 1);
		my $cell2 = $worksheet->get_cell($row, 2);
		next unless $cell1 or $cell2;
		
		print "Row, col $row, 1\n";
		print "Value = ", $cell1->value() , "\n";
		print "Unformatted = ", $cell1->unformatted(), "\n";
		print "Row, col $row, 2\n";
		print "Value = ", $cell2->value() , "\n";
		print "Unformatted = ", $cell2->unformatted(), "\n";
		print "\n";
	#}
}
