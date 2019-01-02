use v6.c+;

unit class Excel::Writer::XLSX::Package::Custom;

###############################################################################
#
# Custom - A class for writing the Excel XLSX custom.xml file for custom
# workbook properties.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

#use Excel::Writer::XLSX::Package::XMLwriter;

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';


###############################################################################
#
# Public and private API methods.
#
###############################################################################

has %!properties;
has $!pid = 0;

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

#NYI     $self->{_properties} = [];
#NYI     $self->{_pid}        = 1;

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

    self.xml_end_tag( 'Properties' );

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
    my $xmlns    = $schema ~ 'custom-properties';
    my $xmlns_vt = $schema ~ 'docPropsVTypes';

    my @attributes = (
        'xmlns'    => $xmlns,
        'xmlns:vt' => $xmlns_vt,
    );

    self.xml_start_tag( 'Properties', @attributes );

    for %!properties.keys -> $property {

        # Write the property element.
        self.write_property( $property );
    }
}

##############################################################################
#
# _write_property()
#
# Write the <property> element.
#
method write_property($property) {
    my $fmtid    = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

    $!pid++;

    my ( $name, $value, $type ) = $property;


    my @attributes = (
        'fmtid' => $fmtid,
        'pid'   => $!pid,
        'name'  => $name,
    );

    self.xml_start_tag( 'property', @attributes );

    if $type eq 'date' {

        # Write the vt:filetime element.
        self.write_vt_filetime( $value );
    }
    elsif $type eq 'number' {

        # Write the vt:r8 element.
        self.write_vt_r8( $value );
    }
    elsif $type eq 'number_int' {

        # Write the vt:i4 element.
        self.write_vt_i4( $value );
    }
    elsif $type eq 'bool' {

        # Write the vt:bool element.
        self.write_vt_bool( $value );
    }
    else {

        # Write the vt:lpwstr element.
        self.write_vt_lpwstr( $value );
    }


    self.xml_end_tag( 'property' );
}


##############################################################################
#
# _write_vt_lpwstr()
#
# Write the <vt:lpwstr> element.
#
method write_vt_lpwstr($data) {
    self.xml_data_element( 'vt:lpwstr', $data );
}


##############################################################################
#
# _write_vt_i4()
#
# Write the <vt:i4> element.
#
method write_vt_i4($data) {
    self.xml_data_element( 'vt:i4', $data );
}


##############################################################################
#
# _write_vt_r8()
#
# Write the <vt:r8> element.
#
method write_vt_r8($data) {
    self.xml_data_element( 'vt:r8', $data );
}


##############################################################################
#
# _write_vt_bool()
#
# Write the <vt:bool> element.
#
method write_vt_bool($data) {
    if $data {
        $data = 'true';
    }
    else {
        $data = 'false';
    }

    self.xml_data_element( 'vt:bool', $data );
}

##############################################################################
#
# _write_vt_filetime()
#
# Write the <vt:filetime> element.
#
method write_vt_filetime($data) {
    self.xml_data_element( 'vt:filetime', $data );
}


=begin pod


__END__

=pod

=head1 NAME

Custom - A class for writing the Excel XLSX custom.xml file.

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
