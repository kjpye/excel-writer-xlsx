unit class Excel::Writer::XLSX::Package::ContentTypes;

###############################################################################
#
# Excel::Writer::XLSX::Package::ContentTypes - A class for writing the Excel
# XLS [Content_Types] file.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use v6.c;
#NYI use strict;
#NYI use warnings;
#NYI use Carp;
#use Excel::Writer::XLSX::Package::XMLwriter;

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';


###############################################################################
#
# Package data.
#
###############################################################################

my $app_package  = 'application/vnd.openxmlformats-package.';
my $app_document = 'application/vnd.openxmlformats-officedocument.';

has @!defaults = (
    [ 'rels', $app_package ~ 'relationships+xml' ],
    [ 'xml',  'application/xml' ],
);

has @!overrides = (
    [ '/docProps/app.xml',    $app_document ~ 'extended-properties+xml' ],
    [ '/docProps/core.xml',   $app_package  ~ 'core-properties+xml' ],
    [ '/xl/styles.xml',       $app_document ~ 'spreadsheetml.styles+xml' ],
    [ '/xl/theme/theme1.xml', $app_document ~ 'theme+xml' ],
    [ '/xl/workbook.xml',     $app_document ~ 'spreadsheetml.sheet.main+xml' ],
);


###############################################################################
#
# Public and private API methods.
#
###############################################################################

###############################################################################
#
# new()
#
# Constructor.
#
#NYI sub new {

#NYI     my $class = shift;
#NYI     my $fh    = shift;
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );

#NYI     $self->{_defaults}  = [@defaults];
#NYI     $self->{_overrides} = [@overrides];

#NYI     bless $self, $class;

#NYI     return $self;
#NYI }


###############################################################################
#
# _assemble_xml_file()
#
# Assemble and write the XML file.
#
method assemble_xml_file {
    self.xml_declaration;
    self.write_types();
    self.write_defaults();
    self.write_overrides();

    self.xml_end_tag( 'Types' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# _add_default()
#
# Add elements to the ContentTypes defaults.
#
method add_default($part-name, $content-type) {
    @!defaults.push: [ $part-name, $content-type ];
}


###############################################################################
#
# _add_override()
#
# Add elements to the ContentTypes overrides.
#
method add_override($part-name, $content-type) {
    @!overrides.push: [ $part-name, $content-type ];
}


###############################################################################
#
# _add_worksheet_name()
#
# Add the name of a worksheet to the ContentTypes overrides.
#
method add_worksheet_name($worksheet-name) {
    $worksheet-name = "/xl/worksheets/$worksheet-name.xml";

    self.add_override( $worksheet-name,
        $app_document ~ 'spreadsheetml.worksheet+xml' );
}


###############################################################################
#
# _add_chartsheet_name()
#
# Add the name of a chartsheet to the ContentTypes overrides.
#
method add_chartsheet_name($chartsheet-name) {
    $chartsheet-name = "/xl/chartsheets/$chartsheet-name.xml";

    self.add_override( $chartsheet-name,
        $app_document ~ 'spreadsheetml.chartsheet+xml' );
}


###############################################################################
#
# _add_chart_name()
#
# Add the name of a chart to the ContentTypes overrides.
#
method add_chart_name($chart-name) {
    $chart-name = "/xl/charts/$chart-name.xml";

    self.add_override( $chart-name, $app_document ~ 'drawingml.chart+xml' );
}


###############################################################################
#
# _add_drawing_name()
#
# Add the name of a drawing to the ContentTypes overrides.
#
method add_drawing_name($drawing-name) {
    $drawing-name = "/xl/drawings/$drawing-name.xml";

    self.add_override( $drawing-name, $app_document ~ 'drawing+xml' );
}


###############################################################################
#
# _add_vml_name()
#
# Add the name of a VML drawing to the ContentTypes defaults.
#
method add_vml_name {
    self.add_default( 'vml', $app_document ~ 'vmlDrawing' );
}


###############################################################################
#
# _add_comment_name()
#
# Add the name of a comment to the ContentTypes overrides.
#
method add_comment_name($comment-name) {
    $comment-name = "/xl/$comment-name.xml";

    self.add_override( $comment-name,
        $app_document ~ 'spreadsheetml.comments+xml' );
}

###############################################################################
#
# _Add_shared_strings()
#
# Add the sharedStrings link to the ContentTypes overrides.
#
method add_shared_strings {
    self.add_override( '/xl/sharedStrings.xml',
        $app_document ~ 'spreadsheetml.sharedStrings+xml' );
}


###############################################################################
#
# _add_calc_chain()
#
# Add the calcChain link to the ContentTypes overrides.
#
method add_calc_chain {
    self.add_override( '/xl/calcChain.xml',
        $app_document ~ 'spreadsheetml.calcChain+xml' );
}


###############################################################################
#
# _add_image_types()
#
# Add the image default types.
	# 
method add_image_types(*%types) {
    for %types ->$type {
        self.add_default( $type, 'image/' ~ $type );
    }
}


###############################################################################
#
# _add_table_name()
#
# Add the name of a table to the ContentTypes overrides.
#
method add_table_name($table-name) {
    $table-name = "/xl/tables/$table-name.xml";

    self.add_override( $table-name,
        $app_document ~ 'spreadsheetml.table+xml' );
}


###############################################################################
#
# _add_vba_project()
#
# Add a vbaProject to the ContentTypes defaults.
#
method add_vba_project {
    # Change the workbook.xml content-type from xlsx to xlsm.
    for @!overrides -> $aref {
        if $aref[0] eq '/xl/workbook.xml' {
            $aref[1] = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
        }
    }

    self.add_default( 'bin', 'application/vnd.ms-office.vbaProject' );
}


###############################################################################
#
# _add_custom_properties()
#
# Add the custom properties to the ContentTypes overrides.
#
method add_custom_properties {
    my $custom = "/docProps/custom.xml";

    self.add_override( $custom, $app_document ~ 'custom-properties+xml' );
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _write_defaults()
#
# Write out all of the <Default> types.
#
method write_defaults {
    for @!defaults -> $aref {
        #<<<
        self.xml_empty_tag(
            'Default',
            'Extension',   $aref[0],
            'ContentType', $aref[1] );
        #>>>
    }
}


###############################################################################
#
# _write_overrides()
#
# Write out all of the <Override> types.
#
method write_overrides {
    for @!overrides -> $aref {
        #<<<
        self.xml_empty_tag(
            'Override',
            'PartName',    $aref[0],
            'ContentType', $aref[1] );
        #>>>
    }
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_types()
#
# Write the <Types> element.
#
method write_types {
    my $xmlns = 'http://schemas.openxmlformats.org/package/2006/content-types';

    my @attributes = ( 'xmlns' => $xmlns, );

    self.xml_start_tag( 'Types', @attributes );
}

###############################################################################
#
# _write_default()
#
# Write the <Default> element.
#
method write_default($extension, $content-type) {
    my @attributes = (
        'Extension'   => $extension,
        'ContentType' => $content-type,
    );

    self.xml_empty_tag( 'Default', @attributes );
}


###############################################################################
#
# _write_override()
#
# Write the <Override> element.
#
method write_override($part-name, $content-type, $writer) {
    my @attributes = (
        'PartName'    => $part-name,
        'ContentType' => $content-type,
    );

    self.xml_empty_tag( 'Override', @attributes );
}

=begin pod
=pod

=head1 NAME

Excel::Writer::XLSX::Package::ContentTypes - A class for writing the Excel XLSX [Content_Types] file.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

(c) MM-MMXVII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
=end pod
