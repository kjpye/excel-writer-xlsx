unit class Excel::Writer::XLSX::Package::SharedStrings;

###############################################################################
#
# SharedStrings - A class for writing the Excel XLSX sharedStrings file.
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
#NYI use Encode;
#use Excel::Writer::XLSX::Package::XMLwriter;

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';


###############################################################################
#
# Public and private API methods.
#
###############################################################################

has @!strings;
has $!string_count;
has $!unique_count;

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

#NYI     $self->{_strings}      = [];
#NYI     $self->{_string_count} = 0;
#NYI     $self->{_unique_count} = 0;

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

    # Write the sst table.
    self.write_sst( $!string_count, $!unique_count );

    # Write the sst strings.
    self.write_sst_strings();

    # Close the sst tag.
    self.xml_end_tag( 'sst' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# _set_string_count()
#
# Set the total sst string count.
#
method set_string_count($count) {
    $!string_count = $count;
}


###############################################################################
#
# _set_unique_count()
#
# Set the total of unique sst strings.
#
method set_unique_count($count) {
    $!unique_count = $count;
}


###############################################################################
#
# _add_strings()
#
# Add the array ref of strings to be written.
#
method add_strings($strings) {
    @!strings= $strings;
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


##############################################################################
#
# _write_sst()
#
# Write the <sst> element.
#
method write_sst($count, $unique-count) {
    my $schema       = 'http://schemas.openxmlformats.org';
    my $xmlns        = $schema ~ '/spreadsheetml/2006/main';

    my @attributes = (
        'xmlns'       => $xmlns,
        'count'       => $count,
        'uniqueCount' => $unique-count,
    );

    self.xml_start_tag( 'sst', @attributes );
}


###############################################################################
#
# _write_sst_strings()
#
# Write the sst string elements.
#
method write_sst_strings {
    for @!strings -> $string {
        self.write_si( $string );
    }
}


##############################################################################
#
# _write_si()
#
# Write the <si> element.
#
method write_si($string) {
    my @attributes = ();

    # Excel escapes control characters with _xHHHH_ and also escapes any
    # literal strings of that type by encoding the leading underscore. So
    # "\0" -> _x0000_ and "_x0000_" -> _x005F_x0000_.
    # The following substitutions deal with those cases.

    # Escape the escape.
    $string ~~ s:g/('_x' <[0..9 a..f A..F]> ** 4 '_')/_x005F$0/;

    # Convert control character to the _xHHHH_ escape.
    $string ~~ s:g/(<[\x00..\x08 \x0B..\x1F]>)/{sprintf "_x%04X_", ord($0)}/;


    # Add attribute to preserve leading or trailing whitespace.
    if $string ~~ /^\s/ || $string ~~ /\s$/ {
        push @attributes, ( 'xml:space' => 'preserve' );
    }


    # Write any rich strings without further tags.
    if $string ~~ /^'<r>'/ && $string ~~ /'</r>'$/ {

        # Prevent utf8 strings from getting double encoded.
        #TODO $string = decode_utf8( $string );

        self.xml_rich_si_element( $string );
    }
    else {
        self.xml_si_element( $string, @attributes );
    }

}


=begin pod


__END__

=pod

=head1 NAME

SharedStrings - A class for writing the Excel XLSX sharedStrings.xml file.

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
