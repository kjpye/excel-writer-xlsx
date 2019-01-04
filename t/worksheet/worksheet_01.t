###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use v6.c+;

use lib 't/lib';
use TestFunctions; # qw(_expected_to_aref _got_to_aref _is_deep_diff _new_worksheet);

use Test;

plan 1;

###############################################################################
#
# Tests setup.
#
my $got = '';
my $got-fh;
my $caption;
my $worksheet;

my $*data = Q:to/END/;
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
END

###############################################################################
#
# Test the assemble-xml-file() method.
#
$caption = " \tWorksheet: assemble-xml-file()";

$worksheet = new-worksheet($got);

$worksheet.select();
$worksheet.assemble-xml-file();

my @expected = expected-to-aref();
my @got      = |got-to-aref( $got );
dd @expected;
dd @got;

ok @got eqv @expected, $caption;
