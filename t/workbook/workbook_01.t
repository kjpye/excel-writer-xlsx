###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use v6;

use lib 't/lib';
use TestFunctions;
#use TestFunctions qw(expected-to-aref got-to-aref is-deep-diff new-workbook);

#use Test::More tests => 1;
use Test;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
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
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>
END


###############################################################################
#
# Test the assemble_xml_file() method.
#
$caption = " \tWorkbook: assemble-xml-file()";

note "Creating new workbook";
$workbook = new-workbook($*data);
note "Adding worksheet";
$workbook.add-worksheet();
note "assemble-xml-file";
$workbook.assemble-xml-file();
note "created";

$expected = expected-to-aref();
dd $expected;
dd $got;
$got      = got-to-aref( $got );
dd $got; fail "got";

is-deep-diff( $got, $expected, $caption );
