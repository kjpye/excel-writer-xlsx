unit class Excel::Writer::XLSX::Package::Core;

###############################################################################
#
# Core - A class for writing the Excel XLSX core.xml file.
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


has %!properties;

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
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );

#NYI     $self->{_properties} = {};
#NYI     $self->{_createtime}  = [ gmtime() ];

#NYI     bless $self, $class;

#NYI     return $self;
#NYI }


###############################################################################
#
# assemble-xml-file()
#
# Assemble and write the XML file.
#
method assemble-xml-file {
    self.xml_declaration;
    self.write_cp_core_properties();
    self.write_dc_title();
    self.write_dc_subject();
    self.write_dc_creator();
    self.write_cp_keywords();
    self.write_dc_description();
    self.write_cp_last_modified_by();
    self.write_dcterms_created();
    self.write_dcterms_modified();
    self.write_cp_category();
    self.write_cp_content_status();

    self.xml_end_tag( 'cp:coreProperties' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# _set_properties()
#
# Set the document properties.
#
method set_properties($properties) {
    self.properties = $properties;
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _datetime_to_iso8601_date()
#
# Convert a gmtime/localtime() date to a ISO 8601 style "2010-01-01T00:00:00Z"
# date. Excel always treats this as a utc date/time.
#
method datetime_to_iso8601_date($gmtime) {
    $gmtime ||= self.createtime;

    my ( $seconds, $minutes, $hours, $day, $month, $year ) = $gmtime;

    $month++;
    $year += 1900;

    my $date = sprintf "%4d-%02d-%02dT%02d:%02d:%02dZ", $year, $month, $day,
      $hours, $minutes, $seconds;
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_cp_core_properties()
#
# Write the <cp:coreProperties> element.
#
method write_cp_core_properties {
    my $xmlns_cp =
      'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
    my $xmlns_dc       = 'http://purl.org/dc/elements/1.1/';
    my $xmlns_dcterms  = 'http://purl.org/dc/terms/';
    my $xmlns_dcmitype = 'http://purl.org/dc/dcmitype/';
    my $xmlns_xsi      = 'http://www.w3.org/2001/XMLSchema-instance';

    my @attributes = (
        'xmlns:cp'       => $xmlns_cp,
        'xmlns:dc'       => $xmlns_dc,
        'xmlns:dcterms'  => $xmlns_dcterms,
        'xmlns:dcmitype' => $xmlns_dcmitype,
        'xmlns:xsi'      => $xmlns_xsi,
    );

    self.xml_start_tag( 'cp:coreProperties', @attributes );
}


###############################################################################
#
# _write_dc_creator()
#
# Write the <dc:creator> element.
#
method write_dc_creator {
    my $data = %!properties<author> || '';

    self.xml_data_element( 'dc:creator', $data );
}


###############################################################################
#
# _write_cp_last_modified_by()
#
# Write the <cp:lastModifiedBy> element.
#
method write_cp_last_modified_by {
    my $data = %!properties<author> || '';

    self.xml_data_element( 'cp:lastModifiedBy', $data );
}


###############################################################################
#
# _write_dcterms_created()
#
# Write the <dcterms:created> element.
#
method write_dcterms_created {
    my $date     = %!properties<created>;
    my $xsi_type = 'dcterms:W3CDTF';

    $date = self.datetime_to_iso8601_date( $date );

    my @attributes = ( 'xsi:type' => $xsi_type, );

    self.xml_data_element( 'dcterms:created', $date, @attributes );
}


###############################################################################
#
# _write_dcterms_modified()
#
# Write the <dcterms:modified> element.
#
method write_dcterms_modified {
    my $date     = %!properties<created>;
    my $xsi_type = 'dcterms:W3CDTF';

    $date = self.datetime_to_iso8601_date( $date );

    my @attributes = ( 'xsi:type' => $xsi_type, );

    self.xml_data_element( 'dcterms:modified', $date, @attributes );
}


##############################################################################
#
# _write_dc_title()
#
# Write the <dc:title> element.
#
method write_dc_title {
    my $data = %!properties<title>;

    return unless $data;

    self.xml_data_element( 'dc:title', $data );
}


##############################################################################
#
# _write_dc_subject()
#
# Write the <dc:subject> element.
#
method write_dc_subject {
    my $data = %!properties<subject>;

    return unless $data;

    self.xml_data_element( 'dc:subject', $data );
}


##############################################################################
#
# _write_cp_keywords()
#
# Write the <cp:keywords> element.
#
method write_cp_keywords {
    my $data = %!properties<keywords>;

    return unless $data;

    self.xml_data_element( 'cp:keywords', $data );
}


##############################################################################
#
# _write_dc_description()
#
# Write the <dc:description> element.
#
method write_dc_description {
    my $data = %!properties<comments>;

    return unless $data;

    self.xml_data_element( 'dc:description', $data );
}


##############################################################################
#
# _write_cp_category()
#
# Write the <cp:category> element.
#
method write_cp_category {
    my $data = %!properties<category>;

    return unless $data;

    self.xml_data_element( 'cp:category', $data );
}


##############################################################################
#
# _write_cp_content_status()
#
# Write the <cp:contentStatus> element.
#
method write_cp_content_status {
    my $data = %!properties<status>;

    return unless $data;

    self.xml_data_element( 'cp:contentStatus', $data );
}

=begin pod

=head1 NAME

Core - A class for writing the Excel XLSX core.xml file.

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
