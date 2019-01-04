use v6.d+;

unit package TestFunctions;

###############################################################################
#
# TestFunctions - Helper functions for Excel::Writer::XLSX test cases.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

#use Exporter;
#use Test::More;
use IO::String;
use Excel::Writer::XLSX;


#our @ISA         = qw(Exporter);
#our @EXPORT      = ();
#our %EXPORT_TAGS = ();
#our @EXPORT_OK   = qw(
#  _expected_to_aref
#  _expected_vml_to_aref
#  _got_to_aref
#  _is_deep_diff
#  _new_object
#  _new_worksheet
#  _new_workbook
#  _new_style
#  _compare_xlsx_files
#);

our $VERSION = '0.06';


###############################################################################
#
# Turn the embedded XML in the __DATA__ section of the calling test program
# into an array ref for comparison testing. Also performs some minor string
# formatting to make comparison easier with got-to-aref().
#
# The XML data in the testcases is taken from Excel 2007 files with formatting
# via "xmllint --format".
#
sub expected-to-aref is export {

    my @data;

    # Ignore warning for files that don't have a 'main::DATA'.
#    no warnings 'once';

    for $*data.lines() -> $line is copy {
        next unless $line ~~ /\S/;    # Skip blank lines.
        $line ~~ s/^\s+//;           # Remove leading whitespace from XML.
        @data.push: $line;
    }

    @data;
}


###############################################################################
#
# Turn the embedded VML in the __DATA__ section of the calling test program
# into an array ref for comparison testing.
#
sub expected-vml-to-aref is export {

    # Ignore warning for files that don't have a 'main::DATA'.
    #no warnings 'once';

    my $vml-str = main::DATA.IO.slurp;

    vml-str-to-array( $vml-str );
}


###############################################################################
#
# Convert an XML string returned by the XMLWriter subclasses into an
# array ref for comparison testing with expected-to-aref().
#
sub got-to-aref($string) is export {

    my $xml-str = ~$string;
dd $xml-str;
    # Remove the newlines after the XML declaration and any others.
    $xml-str ~~ s:g/\n//;

    # Split the XML into chunks at element boundaries.
    #$xml-str.split: /(?<=>)(?=<)/;
    $xml-str.split: /<after \>><before \<>/;
}

###############################################################################
#
# xml-str-to-array()
#
# Convert an XML string into an array for comparison testing.
#
sub xml-str-to-array ($xml-str) is export {

    got-to-aref( $xml-str );

    #s{ />$}{/>} for @xml;
}

###############################################################################
#
# vml-str-to-array()
#
# Convert an Excel generated VML string into an array for comparison testing.
#
# The VML data in the testcases is taken from Excel 2007 files. The data has
# to be massaged significantly to make it suitable for comparison.
#
# Excel::Writer::XLSX produced VML can be parsed as ordinary XML.
#
sub vml-str-to-array($vml-str) is export {

    my @vml = $vml-str.split: /[\r\n]+/; # FIX

    $vml-str = '';

    for @vml {

        .chomp;
        next unless /\S/;    # Skip blank lines.

        s/^\s+//;            # Remove leading whitespace.
        s/\s+$//;            # Remove trailing whitespace.
        s:g/\'/"/;            # Convert VMLs attribute quotes.

        $_ ~= " "  if /\"$/;  # Add space between attributes.
        $_ ~= "\n" if /\>$/;  # Add newline after element end.

        s:g/'><'/>\n</;         # Split multiple elements.

        .chomp if $_ eq "<x:Anchor>\n";    # Put all of Anchor on one line.

        $vml-str ~= $_;
    }

    $vml-str.split: "\n";
}


###############################################################################
#
# compare-xlsx-files()
#
# Compare two XLSX files by extracting the XML files from each archive and
# comparing them.
#
# This is used to compare an "expected" file produced by Excel with a "got"
# file produced by Excel::Writer::XLSX.
#
# In order to compare the XLSX files we convert the data in each XML file.
# contained in the zip archive into arrays of XML elements to make identifying
# differences easier.
#
# This function returns 3 elements suitable for is-deep-diff() comparison:
#    return ( $got-aref, $expected-aref, $caption)
#
sub compare-xlsx-files($got-filename, $exp-filename, $ignore-members, $ignore-elements) is export {

    my $got-zip         = Archive::Zip.new();
    my $exp-zip         = Archive::Zip.new();
    my @got-xml;
    my @exp-xml;

    # Suppress Archive::Zip error reporting. We will handle errors.
    Archive::Zip::setErrorHandler( sub { } );

    # Test the $got file exists.
    if $got-zip.read( $got-filename ) != 0 {
        my $error = 'Excel::Write::XML generated file not found.';
        return ( [$error], [$got-filename], " compare-xlsx-files(). Files." );
    }

    # Test the $exp file exists.
    if $exp-zip.read( $exp-filename ) != 0 {
        my $error = "Excel generated comparison file not found.";
        return ( [$error], [$exp-filename], " compare-xlsx-files(). Files." );
    }

    # The zip "members" are the files in the XLSX container.
    my @got-members = $got-zip.memberNames().sort;
    my @exp-members = $exp-zip.memberNames().sort;

    # Ignore some test specific filenames.
    if $ignore-members.defined && @($ignore-members) {
        my @ignore-members = @$ignore-members;

        @got-members = @got-members.grep({!/@ignore-members/});
        @exp-members = @exp-members.grep({!/@ignore-members/});
    }

    # Check that each XLSX container has the same file members.
    if !arrays-equal( @got-members, @exp-members ) {
        return ( @got-members, @exp-members,
            ' compare-xlsx-files(): Members.' );
    }

    # Compare each file in the XLSX containers.
    for @exp-members -> $filename {
        my $got-xml-str = $got-zip.contents( $filename );
        my $exp-xml-str = $exp-zip.contents( $filename );

        # Remove dates and user specific data from the core.xml data.
        if $filename eq 'docProps/core.xml' {
            $exp-xml-str ~~ s:g/' '? John//;
            $exp-xml-str ~~ s:g/\d\d\d\d\-\d\d\-\d\dT\d\d\:\d\d:\d\dZ//;
            $got-xml-str ~~ s:g/\d\d\d\d\-\d\d\-\d\dT\d\d\:\d\d:\d\dZ//;
        }

        # Remove workbookView dimensions which are almost always different.
        if $filename eq 'xl/workbook.xml' {
            $exp-xml-str ~~ s{\<workbookView<-[^]>*\>} = '<workbookView/>';
            $got-xml-str ~~ s{\<workbookView<-[>]>*\>} = '<workbookView/>';
        }

        # Remove the calcPr elements which may have different Excel version ids.
        if $filename eq 'xl/workbook.xml' {
            $exp-xml-str ~~ s{\<calcPr<-[>]>*\>} = '<calcPr/>';
            $got-xml-str ~~ s{\<calcPr<-[>]>*\>} = '<calcPr/>';
        }

        # Remove printer specific settings from Worksheet pageSetup elements.
        if $filename ~~ /xl\/worksheets\/sheet\d\.xml/ {
            $exp-xml-str ~~ s/'horizontalDpi="200" '//;
            $exp-xml-str ~~ s/'verticalDpi="200" '//;
            $exp-xml-str ~~ s/(\<pageSetup<-[>]>*) ' ' r\:id\=\"rId1\"/$0/;
        }

        # Remove Chart pageMargin dimensions which are almost always different.
        if $filename ~~ /xl\/charts\/chart\d\.xml/ {
# TODO
#            $exp-xml-str ~~ s{\<c\:pageMargins<-[>]>*>} = '<c:pageMargins/>';
#            $got-xml-str ~~ s{\<c\:pageMargins<-[>]>*>} = '<c:pageMargins/>';
        }

        if $filename.ends-with: '.vml' {
            @got-xml = xml-str-to-array( $got-xml-str );
            @exp-xml = vml-str-to-array( $exp-xml-str );
        }
        else {
            @got-xml = xml-str-to-array( $got-xml-str );
            @exp-xml = xml-str-to-array( $exp-xml-str );
        }

        # Ignore test specific XML elements for defined filenames.
        if $ignore-elements.defined && $ignore-elements{$filename}.exists
        {
            my @ignore-elements = @( $ignore-elements{$filename} );

            if +@ignore-elements {
                @got-xml = @got-xml.grep( { !/@ignore-elements/ } );
                @exp-xml = @exp-xml.grep( { !/@ignore-elements/ } );
            }
        }

        # Reorder the XML elements in the XLSX relationship files.
        if $filename eq '[Content_Types].xml' || $filename ~~ /.rels$/ {
            @got-xml = sort-rel-file-data( |@got-xml );
            @exp-xml = sort-rel-file-data( |@exp-xml );
        }

        # Comparison of the XML elements in each file.
        if !arrays-equal( @got-xml, @exp-xml ) {
            return @got-xml, @exp-xml, " comparexlsxfiles(): $filename";
        }
    }

    # Files were the same. Return values that will evaluate to a test pass.
    return ['ok'], ['ok'], ' compare-xlsx-files()';
}


###############################################################################
#
# arrays-equal()
#
# Compare two array refs for equality.
#
sub arrays-equal($exp, $got) is export {

    if +$exp != +$got {
        return 0;
    }

    for ^+$exp -> $i {
        if $exp[$i] ne $got[$i] {
            return 0;
        }
    }

    return 1;
}


###############################################################################
#
# sort-rel-file-data()
#
# Re-order the relationship elements in an array of XLSX XML rel (relationship)
# data. This is necessary for comparison since Excel can produce the elements
# in a semi-random order.
#
sub sort-rel-file-data($header, $tail, *@xml-elements) is export {

    # Sort the relationship elements.
    @xml-elements .= sort;

    $header, @xml-elements, $tail;
}


###############################################################################
#
# Use Test::Differences::eq-or-diff() where available or else fall back to
# using Test::More::is-deeply().
#
sub is-deep-diff($got, $expected, $caption) is export {
#
#    eval {
#        require Test::Differences;
#        Test::Differences->import();
#    };
#
#    if ( !$@ ) {
#        eq-or-diff( $got, $expected, $caption, { context => 1 } );
#    }
#    else {
       #is-deeply( $got, $expected, $caption );
#    }
#
}


###############################################################################
#
# Create a new XML writer sub-classed object based on a class name and bind
# the output to the supplied scalar ref for testing. Calls to the objects XML
# writing subs will add the output to the scalar.
#
#TODO
#sub new-object($got-ref, $class) {
#
#    open my $gotfh, '>', $got-ref or die "Failed to open filehandle: $!";
#
#    my $object = $class->new( $got-fh );
#
#    return $object;
#}


###############################################################################
#
# Create a new Worksheet object and bind the output to the supplied scalar ref.
#
#TODO
#sub new-worksheet {
#
#    my $got-ref = shift;
#
#    return new-object( $got-ref, 'Excel::Writer::XLSX::Worksheet' );
#}


###############################################################################
#
# Create a new Style object and bind the output to the supplied scalar ref.
#
#TODO
#sub new-style {
#
#    my $got-ref = shift;
#
#    return new-object( $got-ref, 'Excel::Writer::XLSX::Package::Styles' );
#}


###############################################################################
#
# Create a new Workbook object and bind the output to the supplied scalar ref.
# This is slightly different than the previous cases since the constructor
# requires a filename/filehandle.
#
#TODO
sub new-workbook($got-ref, $buffer, $fh is rw) is export {

    $fh = IO::String.new(:$buffer) or die "Failed to open filehandle: $!";
    my $tmp-fh = IO::String.new();

    my $workbook = Excel::Writer::XLSX.new(filename => $tmp-fh);
    $workbook.fh = $fh;

note "Got new workbook";
    return $workbook;
}
