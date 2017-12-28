unit class Excel::Writer::XLSX::Package::Relationships;

###############################################################################
#
# Relationships - A class for writing the Excel XLSX Rels file.
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

our $schema_root     = 'http://schemas.openxmlformats.org';
our $package_schema  = $schema_root ~ '/package/2006/relationships';
our $document_schema = $schema_root ~ '/officeDocument/2006/relationships';

###############################################################################
#
# Public and private API methods.
#
###############################################################################

has $!id = 1;
has @!rels;

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

#NYI     $self->{_rels} = [];
#NYI     $self->{_id}   = 1;

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
    self.write_relationships();
}


###############################################################################
#
# _add_document_relationship()
#
# Add container relationship to XLSX .rels xml files.
#
method add_document_relationship($type, $target) {
    $type   = $document_schema ~ $type;

    @!rels.push: [ $type, $target ];
}


###############################################################################
#
# _add_package_relationship()
#
# Add container relationship to XLSX .rels xml files.
#
method add_package_relationship($type, $target) {
    $type   = $package_schema ~ $type;

    @!rels.push: [ $type, $target ];
}


###############################################################################
#
# _add_ms_package_relationship()
#
# Add container relationship to XLSX .rels xml files. Uses MS schema.
#
method add_ms_package_relationship($type, $target) {
    my $schema = 'http://schemas.microsoft.com/office/2006/relationships';

    $type   = $schema ~ $type;

    @!rels.push: [ $type, $target ];
}


###############################################################################
#
# _add_worksheet_relationship()
#
# Add worksheet relationship to sheet.rels xml files.
#
method add_worksheet_relationship($type, $target, $target-mode) {
    $type   = $document_schema ~ $type;

    @!rels.push: [ $type, $target, $target-mode ];
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
# _write_relationships()
#
# Write the <Relationships> element.
#
method write_relationships {
    my @attributes = ( 'xmlns' => $package_schema, );

    self.xml_start_tag( 'Relationships', @attributes );

    for @!rels -> $rel {
        self.write_relationship( $rel );
    }

    self.xml_end_tag( 'Relationships' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


##############################################################################
#
# _write_relationship()
#
# Write the <Relationship> element.
#
method write_relationship($type, $target, $target-mode) {
    my @attributes = (
        'Id'     => 'rId' ~ $!id++,
        'Type'   => $type,
        'Target' => $target,
    );

    @attributes.push: ( 'TargetMode' => $target-mode ) if $target-mode;

    self.xml_empty_tag( 'Relationship', @attributes );
}


=begin pod


__END__

=pod

=head1 NAME

Relationships - A class for writing the Excel XLSX Rels file.

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
