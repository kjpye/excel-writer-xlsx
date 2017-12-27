unit class Excel::Writer::XLSX::Package::App;

###############################################################################
#
# App - A class for writing the Excel XLSX app.xml file.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use v6.c;
#use Excel::Writer::XLSX::Package::XMLwriter;

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';


###############################################################################
#
# Public and private API methods.
#
###############################################################################

has @!heading-pairs;
has %!properties;
has @!part-names;

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

#NYI     $self->{_part-names}    = [];
#NYI     $self->{_heading-pairs} = [];
#NYI     $self->{_properties}    = {};

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
    self.write_properties();
    self.write_application();
    self.write_doc_security();
    self.write_scale_crop();
    self.write_heading_pairs();
    self.write_titles_of_parts();
    self.write_manager();
    self.write_company();
    self.write_links_up_to_date();
    self.write_shared_doc();
    self.write_hyperlink_base();
    self.write_hyperlinks_changed();
    self.write_app_version();

    self.xml_end_tag( 'Properties' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# _add_part_name()
#
# Add the name of a workbook Part such as 'Sheet1' or 'Print_Titles'.
#
method add_part_name($part-name) {
    @!part-names.push: $part-name;
}


###############################################################################
#
# _add_heading_pair()
#
# Add the name of a workbook Heading Pair such as 'Worksheets', 'Charts' or
# 'Named Ranges'.
#
method add_heading_pair($heading-pair) {
    return unless $heading-pair[1];  # Ignore empty pairs such as chartsheets.

    my @vector = (
        [ 'lpstr', $heading-pair[0] ],    # Data name
        [ 'i4',    $heading-pair[1] ],    # Data size
    );

    @!heading-pairs.push: @vector;
}


###############################################################################
#
# _set_properties()
#
# Set the document properties.
#
method set_properties($properties) {
    %!properties = $properties;
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_properties()
#
# Write the <Properties> element.
#
method write_properties {
    my $schema   = 'http://schemas.openxmlformats.org/officeDocument/2006/';
    my $xmlns    = $schema ~ 'extended-properties';
    my $xmlns_vt = $schema ~ 'docPropsVTypes';

    my @attributes = (
        'xmlns'    => $xmlns,
        'xmlns:vt' => $xmlns_vt,
    );

    self.xml_start_tag( 'Properties', @attributes );
}

###############################################################################
#
# _write_application()
#
# Write the <Application> element.
#
method write_application {
    my $data = 'Microsoft Excel';

    self.xml_data_element( 'Application', $data );
}


###############################################################################
#
# _write_doc_security()
#
# Write the <DocSecurity> element.
#
method write_doc_security {
    my $data = 0;

    self.xml_data_element( 'DocSecurity', $data );
}


###############################################################################
#
# _write_scale_crop()
#
# Write the <ScaleCrop> element.
#
method write_scale_crop {
    my $data = 'false';

    self.xml_data_element( 'ScaleCrop', $data );
}


###############################################################################
#
# _write_heading_pairs()
#
# Write the <HeadingPairs> element.
#
method write_heading_pairs {
    self.xml_start_tag( 'HeadingPairs' );

    self.write_vt_vector( 'variant', @!heading-pairs );

    self.xml_end_tag( 'HeadingPairs' );
}


###############################################################################
#
# _write_titles_of_parts()
#
# Write the <TitlesOfParts> element.
#
method write_titles_of_parts {
    self.xml_start_tag( 'TitlesOfParts' );

    my @parts_data;

    for @!part-names -> $part-name {
        @parts_data.push: [ 'lpstr', $part-name ];
    }

    self.write_vt_vector( 'lpstr', @parts_data );

    self.xml_end_tag( 'TitlesOfParts' );
}


###############################################################################
#
# _write_vt_vector()
#
# Write the <vt:vector> element.
#
method write_vt_vector($base-type, $data) {
    my $size      = $data.elems;

    my @attributes = (
        'size'     => $size,
        'baseType' => $base-type,
    );

    self.xml_start_tag( 'vt:vector', @attributes );

    for $data -> $aref {
        self.xml_start_tag( 'vt:variant' ) if $base-type eq 'variant';
        self.write_vt_data( $aref );
        self.xml_end_tag( 'vt:variant' ) if $base-type eq 'variant';
    }

    self.xml_end_tag( 'vt:vector' );
}


##############################################################################
#
# _write_vt_data()
#
# Write the <vt:*> elements such as <vt:lpstr> and <vt:if>.
#
method write_vt_data($type, $data) {
    self.xml_data_element( "vt:$type", $data );
}


###############################################################################
#
# _write_company()
#
# Write the <Company> element.
#
method write_company {
    my $data = %!properties<company> || '';

    self.xml_data_element( 'Company', $data );
}


###############################################################################
#
# _write_manager()
#
# Write the <Manager> element.
#
method write_manager {
    my $data = %!properties<manager>;

    return unless $data;

    self.xml_data_element( 'Manager', $data );
}


###############################################################################
#
# _write_links_up_to_date()
#
# Write the <LinksUpToDate> element.
#
method write_links_up_to_date {
    my $data = 'false';

    self.xml_data_element( 'LinksUpToDate', $data );
}


###############################################################################
#
# _write_shared_doc()
#
# Write the <SharedDoc> element.
#
method write_shared_doc {
    my $data = 'false';

    self.xml_data_element( 'SharedDoc', $data );
}


###############################################################################
#
# _write_hyperlink_base()
#
# Write the <HyperlinkBase> element.
#
method write_hyperlink_base {
    my $data = %!properties<hyperlink_base>;

    return unless $data;

    self.xml_data_element( 'HyperlinkBase', $data );
}


###############################################################################
#
# _write_hyperlinks_changed()
#
# Write the <HyperlinksChanged> element.
#
method write_hyperlinks_changed {
    my $data = 'false';

    self.xml_data_element( 'HyperlinksChanged', $data );
}


###############################################################################
#
# _write_app_version()
#
# Write the <AppVersion> element.
#
method write_app_version {
    my $data = '12.0000';

    self.xml_data_element( 'AppVersion', $data );
}


=begin pod


__END__

=pod

=head1 NAME

App - A class for writing the Excel XLSX app.xml file.

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
