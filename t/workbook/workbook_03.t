###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use v6.c+;

use lib 't/lib';
use TestFunctions; # qw(_expected_to_aref _got_to_aref _is_deep_diff _new_workbook);

use Test;

plan 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got = '';
my $got-fh;
my $caption;
my $workbook;


my $*data = Q:to/END/;
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
  </bookViews>
  <sheets>
    <sheet name="Non Default Name" sheetId="1" r:id="rId1"/>
    <sheet name="Another Name" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>
END

###############################################################################
#
# Test the assemble-xml-file() method.
#
$caption = " \tWorkbook: _assemble_xml_file()";

$workbook = new-workbook($*data, $got, $got-fh);
$workbook.add-worksheet('Non Default Name');
$workbook.add-worksheet('Another Name');
$workbook.assemble-xml-file();

my @expected = expected-to-aref();
my @got      = |got-to-aref( $got-fh );
ok @got eqv @expected, $caption;


