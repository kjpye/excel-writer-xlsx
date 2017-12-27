unit class Excel::Writer::XLSX::Package::Packager;

###############################################################################
#
# Packager - A class for creating the Excel XLSX package.
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
#NYI use Exporter;
#NYI use Carp;
#NYI use File::Copy;
#NYI use Excel::Writer::XLSX::Package::App;
#NYI use Excel::Writer::XLSX::Package::Comments;
#NYI use Excel::Writer::XLSX::Package::ContentTypes;
#NYI use Excel::Writer::XLSX::Package::Core;
#NYI use Excel::Writer::XLSX::Package::Custom;
#NYI use Excel::Writer::XLSX::Package::Relationships;
#NYI use Excel::Writer::XLSX::Package::SharedStrings;
#NYI use Excel::Writer::XLSX::Package::Styles;
#NYI use Excel::Writer::XLSX::Package::Table;
#NYI use Excel::Writer::XLSX::Package::Theme;
#NYI use Excel::Writer::XLSX::Package::VML;
#NYI 
#NYI our @ISA     = qw(Exporter);
#NYI our $VERSION = '0.96';

has $!package_dir      = '';
has $!workbook;
has $!worksheet_count  = 0;
has $!chartsheet_count = 0;
has $!chart_count      = 0;
has $!drawing_count    = 0;
has $!table_count      = 0;
has @!named_ranges     = [];
has $!num-vml-files;
has $!num-comment-files;

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
#NYI 
#NYI     my $class = shift;
#NYI     my $fh    = shift;
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter.new( $fh );
#NYI 
#NYI     $self->{_package_dir}      = '';
#NYI     $self->{_workbook};
#NYI     $self->{_worksheet_count}  = 0;
#NYI     $self->{_chartsheet_count} = 0;
#NYI     $self->{_chart_count}      = 0;
#NYI     $self->{_drawing_count}    = 0;
#NYI     $self->{_table_count}      = 0;
#NYI     $self->{_named_ranges}     = [];
#NYI 
#NYI 
#NYI     bless $self, $class;
#NYI 
#NYI     return $self;
#NYI }


###############################################################################
#
# _set_package_dir()
#
# Set the XLSX OPC package directory.
#
method set_package_dir($dir) {
    $!package_dir = $dir;
}


###############################################################################
#
# _add_workbook()
#
# Add the Excel::Writer::XLSX::Workbook object to the package.
#
method add_workbook($workbook) {
    $!workbook          = $workbook;
    $!chart_count       = $workbook.num-charts;
    $!drawing_count     = $workbook.num-drawings;
    $!num-vml-files     = $workbook.num-vml-files;
    $!num-comment-files = $workbook.num-comment-files;
    @!named_ranges      = $workbook.named-ranges;

    for %$workbook.worksheets -> $worksheet {
        if $worksheet.is_chartsheet {
            $!chartsheet_count++;
        }
        else {
            $!worksheet_count++;
        }
    }
}


###############################################################################
#
# _create_package()
#
# Write the xml files that make up the XLXS OPC package.
#
method create_package {
    self.write_worksheet_files();
    self.write_chartsheet_files();
    self.write_workbook_file();
    self.write_chart_files();
    self.write_drawing_files();
    self.write_vml_files();
    self.write_comment_files();
    self.write_table_files();
    self.write_shared_strings_file();
    self.write_app_file();
    self.write_core_file();
    self.write_custom_file();
    self.write_content_types_file();
    self.write_styles_file();
    self.write_theme_file();
    self.write_root_rels_file();
    self.write_workbook_rels_file();
    self.write_worksheet_rels_files();
    self.write_chartsheet_rels_files();
    self.write_drawing_rels_files();
    self.add_image_files();
    self.add_vba_project();
}


###############################################################################
#
# _write_workbook_file()
#
# Write the workbook.xml file.
#
method write_workbook_file {
    my $dir      = $!package_dir;
    my $workbook = $!workbook;

    _mkdir( $dir ~ '/xl' );

    $workbook.set_xml_writer( $dir ~ '/xl/workbook.xml' );
    $workbook.assemble_xml_file();
}


###############################################################################
#
# _write_worksheet_files()
#
# Write the worksheet files.
#
method write_worksheet_files {
    my $dir  = $!package_dir;

    _mkdir( $dir ~ '/xl' );
    _mkdir( $dir ~ '/xl/worksheets' );

    my $index = 1;
    for $!workbook.worksheets -> $worksheet {
        next if $worksheet.is_chartsheet;

        $worksheet.set_xml_writer(
            $dir ~ '/xl/worksheets/sheet' ~ $index++ ~ '.xml' );
        $worksheet.assemble_xml_file();

    }
}


###############################################################################
#
# _write_chartsheet_files()
#
# Write the chartsheet files.
#
method write_chartsheet_files {
    my $dir  = $!package_dir;

    my $index = 1;
    for $!workbook.worksheets -> $worksheet {
        next unless $worksheet.is_chartsheet;

        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/chartsheets' );

        $worksheet.set_xml_writer(
            $dir ~ '/xl/chartsheets/sheet' ~ $index++ ~ '.xml' );
        $worksheet.assemble_xml_file();

    }
}


###############################################################################
#
# _write_chart_files()
#
# Write the chart files.
#
method write_chart_files {
    my $dir  = $!package_dir;

    return unless $!workbook.charts;

    _mkdir( $dir ~ '/xl' );
    _mkdir( $dir ~ '/xl/charts' );

    my $index = 1;
    for $!workbook.charts -> $chart {
        $chart.set_xml_writer(
            $dir ~ '/xl/charts/chart' ~ $index++ ~ '.xml' );
        $chart.assemble_xml_file();

    }
}


###############################################################################
#
# _write_drawing_files()
#
# Write the drawing files.
#
method write_drawing_files {
    my $dir  = $!package_dir;

    return unless $!drawing_count;

    _mkdir( $dir ~ '/xl' );
    _mkdir( $dir ~ '/xl/drawings' );

    my $index = 1;
    for $!workbook.drawings -> $drawing {
        $drawing.set_xml_writer(
            $dir ~ '/xl/drawings/drawing' ~ $index++ ~ '.xml' );
        $drawing.assemble_xml_file();
    }
}


###############################################################################
#
# _write_vml_files()
#
# Write the comment VML files.
#
method write_vml_files {
    my $dir  = $!package_dir;

    my $index = 1;
    for $!workbook.worksheets -> $worksheet {

        next if !$worksheet.has_vml and !$worksheet.has_header_vml;

        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/drawings' );

        if $worksheet.has_vml {
            my $vml = Excel::Writer::XLSX::Package::VML.new();

            $vml.set_xml_writer(
                $dir ~ '/xl/drawings/vmlDrawing' ~ $index ~ '.vml' );
            $vml.assemble_xml_file(
                $worksheet.vml_data_id,    $worksheet.vml_shape_id,
                $worksheet.comments_array, $worksheet.buttons_array,
                Nil
            );

            $index++;
        }

        if $worksheet.has_header_vml {
            my $vml = Excel::Writer::XLSX::Package::VML.new();

            $vml.writer(
                $dir ~ '/xl/drawings/vmlDrawing' ~ $index ~ '.vml' );
            $vml.assemble_xml_file(
                $worksheet.vml_header_id,
                $worksheet.vml_header_id * 1024,
                Nil, Nil, $worksheet.header_images_array
            );

            self.write_vml_drawing_rels_file($worksheet, $index);

            $index++;
        }
    }
}


###############################################################################
#
# _write_comment_files()
#
# Write the comment files.
#
method write_comment_files {
    my $dir  = $!package_dir;

    my $index = 1;
    for $!workbook.worksheets -> $worksheet {
        next unless $worksheet.has_comments;

        my $comment = Excel::Writer::XLSX::Package::Comments.new();

        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/drawings' );

        $comment.set_xml_writer( $dir ~ '/xl/comments' ~ $index++ ~ '.xml' );
        $comment.assemble_xml_file( $worksheet.comments_array );
    }
}


###############################################################################
#
# _write_shared_strings_file()
#
# Write the sharedStrings.xml file.
#
method write_shared_strings_file {

    my $dir  = $!package_dir;
    my $sst  = Excel::Writer::XLSX::Package::SharedStrings.new();

    my $total    = $!workbook.str_total;
    my $unique   = $!workbook.str_unique;
    my $sst_data = $!workbook.str_array;

    return unless $total > 0;

    _mkdir( $dir ~ '/xl' );

    $sst.set_string_count( $total );
    $sst.set_unique_count( $unique );
    $sst.add_strings( $sst_data );

    $sst.set_xml_writer( $dir ~ '/xl/sharedStrings.xml' );
    $sst.assemble_xml_file();
}


###############################################################################
#
# _write_app_file()
#
# Write the app.xml file.
#
method write_app_file {
    my $dir        = $!package_dir;
    my $properties = $!workbook.doc_properties;
    my $app        = Excel::Writer::XLSX::Package::App.new();

    _mkdir( $dir ~ '/docProps' );

    # Add the Worksheet heading pairs.
    $app.add_heading_pair( [ 'Worksheets', $!worksheet_count ] );

    # Add the Chartsheet heading pairs.
    $app.add_heading_pair( [ 'Charts', $!chartsheet_count ] );

    # Add the Worksheet parts.
    for $!workbook.worksheets -> $worksheet {
        next if $worksheet.is_chartsheet;
        $app.add_part_name( $worksheet.get_name() );
    }

    # Add the Chartsheet parts.
    for $!workbook.worksheets -> $worksheet {
        next unless $worksheet.is_chartsheet;
        $app.add_part_name( $worksheet.get_name() );
    }

    # Add the Named Range heading pairs.
    if my $range_count = @!named_ranges.elems {
        $app.add_heading_pair( [ 'Named Ranges', $range_count ] );
    }

    # Add the Named Ranges parts.
    for @!named_ranges -> $named-range {
        $app.add_part_name( $named-range );
    }

    $app.set_properties( $properties );

    $app.set_xml_writer( $dir ~ '/docProps/app.xml' );
    $app.assemble_xml_file();
}


###############################################################################
#
# _write_core_file()
#
# Write the core.xml file.
#
method write_core_file {
    my $dir        = $!package_dir;
    my $properties = $!workbook.doc_properties;
    my $core       = Excel::Writer::XLSX::Package::Core.new();

    _mkdir( $dir ~ '/docProps' );

    $core.set_properties( $properties );
    $core.set_xml_writer( $dir ~ '/docProps/core.xml' );
    $core.assemble_xml_file();
}


###############################################################################
#
# _write_custom_file()
#
# Write the custom.xml file.
#
method write_custom_file {
    my $dir        = $!package_dir;
    my $properties = $!workbook.custom_properties;
    my $custom     = Excel::Writer::XLSX::Package::Custom.new();

    return if !$properties;

    _mkdir( $dir ~ '/docProps' );

    $custom.set_properties( $properties );
    $custom.set_xml_writer( $dir ~ '/docProps/custom.xml' );
    $custom.assemble_xml_file();
}


###############################################################################
#
# _write_content_types_file()
#
# Write the ContentTypes.xml file.
#
method write_content_types_file {
    my $dir     = $!package_dir;
    my $content = Excel::Writer::XLSX::Package::ContentTypes.new();

    $content.add_image_types( $!workbook.image_types );

    my $worksheet_index  = 1;
    my $chartsheet_index = 1;
    for $!workbook.worksheets -> $worksheet {
        if $worksheet.is_chartsheet {
            $content.add_chartsheet_name( 'sheet' ~ $chartsheet_index++ );
        }
        else {
            $content.add_worksheet_name( 'sheet' ~ $worksheet_index++ );
        }
    }

    for 1 .. $!chart_count -> $i {
        $content.add_chart_name( 'chart' ~ $i );
    }

    for 1 .. $!drawing_count -> $i {
        $content.add_drawing_name( 'drawing' ~ $i );
    }

    if $!num-vml-files {
        $content.add_vml_name();
    }

    for 1 .. $!table_count -> $i {
        $content.add_table_name( 'table' ~ $i );
    }

    for 1 .. $!num-comment-files -> $i {
        $content.add_comment_name( 'comments' ~ $i );
    }

    # Add the sharedString rel if there is string data in the workbook.
    if $!workbook.str_total {
        $content.add_shared_strings();
    }

    # Add vbaProject if present.
    if $!workbook.vba_project {
        $content.add_vba_project();
    }

    # Add the custom properties if present.
    if $!workbook.custom_properties {
        $content.add_custom_properties();
    }

    $content.set_xml_writer( $dir ~ '/[Content_Types].xml' );
    $content.assemble_xml_file();
}


###############################################################################
#
# _write_styles_file()
#
# Write the style xml file.
#
method write_styles_file {
    my $dir              = $!package_dir;
    my $xf_formats       = $!workbook.xf_formats;
    my $palette          = $!workbook.palette;
    my $font_count       = $!workbook.font_count;
    my $num_format_count = $!workbook.num_format_count;
    my $border_count     = $!workbook.border_count;
    my $fill_count       = $!workbook.fill_count;
    my $custom_colors    = $!workbook.custom_colors;
    my $dxf_formats      = $!workbook.dxf_formats;

    my $rels = Excel::Writer::XLSX::Package::Styles.new();

    _mkdir( $dir ~ '/xl' );

    $rels.set_style_properties(
        $xf_formats,
        $palette,
        $font_count,
        $num_format_count,
        $border_count,
        $fill_count,
        $custom_colors,
        $dxf_formats,

    );

    $rels.set_xml_writer( $dir ~ '/xl/styles.xml' );
    $rels.assemble_xml_file();
}


###############################################################################
#
# _write_theme_file()
#
# Write the style xml file.
#
method write_theme_file {
    my $dir  = $!package_dir;
    my $rels = Excel::Writer::XLSX::Package::Theme.new();

    _mkdir( $dir ~ '/xl' );
    _mkdir( $dir ~ '/xl/theme' );

    $rels.set_xml_writer( $dir ~ '/xl/theme/theme1.xml' );
    $rels.assemble_xml_file();
}


###############################################################################
#
# _write_table_files()
#
# Write the table files.
#
method write_table_files {
    my $dir  = $!package_dir;

    my $index = 1;
    for $!workbook.worksheets -> $worksheet {
        my @table_props = $worksheet.tables;

        next unless @table_props;

        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/tables' );

        for @table_props -> $table-props {

            my $table = Excel::Writer::XLSX::Package::Table.new();

            $table.set_xml_writer(
                $dir ~ '/xl/tables/table' ~ $index++ ~ '.xml' );

            $table.set_properties( $table-props );

            $table.assemble_xml_file();

            $!table_count++;
        }
    }
}


###############################################################################
#
# _write_root_rels_file()
#
# Write the _rels/.rels xml file.
#
method write_root_rels_file {
    my $dir  = $!package_dir;
    my $rels = Excel::Writer::XLSX::Package::Relationships.new();

    _mkdir( $dir ~ '/_rels' );

    $rels.add_document_relationship( '/officeDocument', 'xl/workbook.xml' );

    $rels.add_package_relationship( '/metadata/core-properties',
        'docProps/core.xml' );

    $rels.add_document_relationship( '/extended-properties',
        'docProps/app.xml' );

    if $!workbook.custom_properties {
        $rels.add_document_relationship( '/custom-properties',
            'docProps/custom.xml' );
    }

    $rels.set_xml_writer( $dir ~ '/_rels/.rels' );
    $rels.assemble_xml_file();
}


###############################################################################
#
# _write_workbook_rels_file()
#
# Write the _rels/.rels xml file.
#
method write_workbook_rels_file {
    my $dir  = $!package_dir;
    my $rels = Excel::Writer::XLSX::Package::Relationships.new();

    _mkdir( $dir ~ '/xl' );
    _mkdir( $dir ~ '/xl/_rels' );

    my $worksheet_index  = 1;
    my $chartsheet_index = 1;

    for $!workbook.worksheets -> $worksheet {
        if $worksheet.is_chartsheet {
            $rels.add_document_relationship( '/chartsheet',
                'chartsheets/sheet' ~ $chartsheet_index++ ~ '.xml' );
        }
        else {
            $rels.add_document_relationship( '/worksheet',
                'worksheets/sheet' ~ $worksheet_index++ ~ '.xml' );
        }
    }

    $rels.add_document_relationship( '/theme',  'theme/theme1.xml' );
    $rels.add_document_relationship( '/styles', 'styles.xml' );

    # Add the sharedString rel if there is string data in the workbook.
    if $!workbook.str_total {
        $rels.add_document_relationship( '/sharedStrings',
            'sharedStrings.xml' );
    }

    # Add vbaProject if present.
    if $!workbook.vba_project {
        $rels.add_ms_package_relationship( '/vbaProject', 'vbaProject.bin' );
    }

    $rels.set_xml_writer( $dir ~ '/xl/_rels/workbook.xml.rels' );
    $rels.assemble_xml_file();
}


###############################################################################
#
# _write_worksheet_rels_files()
#
# Write the worksheet .rels files for worksheets that contain links to external
# data such as hyperlinks or drawings.
#
method write_worksheet_rels_files {
    my $dir  = $!package_dir;

    my $index = 0;
    for $!workbook.worksheets -> $worksheet {

        next if $worksheet.is_chartsheet;

        $index++;

        my @external_links = (
            $worksheet.external_hyper_links,
            $worksheet.external_drawing_links,
            $worksheet.external_vml_links,
            $worksheet.external_table_links,
            $worksheet.external_comment_links,
        );

        next unless @external_links;

        # Create the worksheet .rels dirs.
        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/worksheets' );
        _mkdir( $dir ~ '/xl/worksheets/_rels' );

        my $rels = Excel::Writer::XLSX::Package::Relationships.new();

        for @external_links -> $link-data {
            $rels.add_worksheet_relationship( $link-data );
        }

        # Create the .rels file such as /xl/worksheets/_rels/sheet1.xml.rels.
        $rels.set_xml_writer(
            $dir ~ '/xl/worksheets/_rels/sheet' ~ $index ~ '.xml.rels' );
        $rels.assemble_xml_file();
    }
}


###############################################################################
#
# _write_chartsheet_rels_files()
#
# Write the chartsheet .rels files for links to drawing files.
#
method write_chartsheet_rels_files {
    my $dir  = $!package_dir;


    my $index = 0;
    for $!workbook.worksheets -> $worksheet {

        next unless $worksheet.is_chartsheet;

        $index++;

        my @external_links = $worksheet.external_drawing_links;

        next unless @external_links;

        # Create the chartsheet .rels dir.
        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/chartsheets' );
        _mkdir( $dir ~ '/xl/chartsheets/_rels' );

        my $rels = Excel::Writer::XLSX::Package::Relationships.new();

        for @external_links -> $link-data {
            $rels.add_worksheet_relationship( $link-data );
        }

        # Create the .rels file such as /xl/chartsheets/_rels/sheet1.xml.rels.
        $rels.set_xml_writer(
            $dir ~ '/xl/chartsheets/_rels/sheet' ~ $index ~ '.xml.rels' );
        $rels.assemble_xml_file();
    }
}


###############################################################################
#
# _write_drawing_rels_files()
#
# Write the drawing .rels files for worksheets that contain charts or drawings.
#
method write_drawing_rels_files {
    my $dir  = $!package_dir;


    my $index = 0;
    for $!workbook.worksheets -> $worksheet {

        if $worksheet.drawing_links || $worksheet.has_shapes {
            $index++;
        }

        next unless $worksheet.drawing_links;

        # Create the drawing .rels dir.
        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/drawings' );
        _mkdir( $dir ~ '/xl/drawings/_rels' );

        my $rels = Excel::Writer::XLSX::Package::Relationships.new();

        for $worksheet.drawing_links -> $drawing-data {
            $rels.add_document_relationship( $drawing-data );
        }

        # Create the .rels file such as /xl/drawings/_rels/sheet1.xml.rels.
        $rels.set_xml_writer(
            $dir ~ '/xl/drawings/_rels/drawing' ~ $index ~ '.xml.rels' );
        $rels.assemble_xml_file();
    }
}


###############################################################################
#
# _write_vml_drawing_rels_files()
#
# Write the vmlDdrawing .rels files for worksheets with images in header or
# footers.
#
method write_vml_drawing_rels_file($worksheet, $index) {
    my $dir       = $!package_dir;


    # Create the drawing .rels dir.
    _mkdir( $dir ~ '/xl' );
    _mkdir( $dir ~ '/xl/drawings' );
    _mkdir( $dir ~ '/xl/drawings/_rels' );

    my $rels = Excel::Writer::XLSX::Package::Relationships.new();

    for $worksheet.vml_drawing_links -> $drawing-data {
        $rels.add_document_relationship( $drawing-data );
    }

    # Create the .rels file such as /xl/drawings/_rels/vmlDrawing1.vml.rels.
    $rels.set_xml_writer(
        $dir ~ '/xl/drawings/_rels/vmlDrawing' ~ $index ~ '.vml.rels' );
    $rels.assemble_xml_file();

}


###############################################################################
#
# _add_image_files()
#
# Write the /xl/media/image?.xml files.
#
method add_image_files {
    my $dir      = $!package_dir;
    my $workbook = $!workbook;
    my $index    = 1;

    for $workbook.images -> $image {
        my $filename  = $image[0];
        my $extension = '.' ~ $image[1];

        _mkdir( $dir ~ '/xl' );
        _mkdir( $dir ~ '/xl/media' );

        copy( $filename, $dir ~ '/xl/media/image' ~ $index++ ~ $extension );
    }
}


###############################################################################
#
# _add_vba_project()
#
# Write the vbaProject.bin file.
#
method add_vba_project {
    my $dir         = $!package_dir;
    my $vba_project = $!workbook.vba_project;

    return unless $vba_project;

    _mkdir( $dir ~ '/xl' );

    copy( $vba_project, $dir ~ '/xl/vbaProject.bin' );
}


###############################################################################
#
# _mkdir()
#
# Wrapper function for Perl's mkdir to allow error trapping.
#
sub _mkdir($dir) {

    return if $dir.IO ~~ :e;

    my $ret = mkdir( $dir );

    if !$ret {
        fail "Couldn't create sub directory $dir: $!";
    }
}

=begin pod
=pod

=head1 NAME

Packager - A class for creating the Excel XLSX package.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX> to create an Excel XLSX container file.

From Wikipedia: I<The Open Packaging Conventions (OPC) is a container-file technology initially created by Microsoft to store a combination of XML and non-XML files that together form a single entity such as an Open XML Paper Specification (OpenXPS) document>. L<http://en.wikipedia.org/wiki/Open_Packaging_Conventions>.

At its simplest an Excel XLSX file contains the following elements:

     ____ [Content_Types].xml
    |
    |____ docProps
    | |____ app.xml
    | |____ core.xml
    |
    |____ xl
    | |____ workbook.xml
    | |____ worksheets
    | | |____ sheet1.xml
    | |
    | |____ styles.xml
    | |
    | |____ theme
    | | |____ theme1.xml
    | |
    | |_____rels
    |   |____ workbook.xml.rels
    |
    |_____rels
      |____ .rels


The C<Excel::Writer::XLSX::Package::Packager> class co-ordinates the classes that represent the elements of the package and writes them into the XLSX file.

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
