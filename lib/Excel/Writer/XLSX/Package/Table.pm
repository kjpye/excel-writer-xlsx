use v6.c+;

unit class Excel::Writer::XLSX::Package::Table;

###############################################################################
#
# Table - A class for writing the Excel XLSX Table file.
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

#NYI     $self->{_properties} = {};

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

    # Write the table element.
    self.write_table();

    # Write the autoFilter element.
    self.write_auto_filter();

    # Write the tableColumns element.
    self.write_table_columns();

    # Write the tableStyleInfo element.
    self.write_table_style_info();


    # Close the table tag.
    self.xml_end_tag( 'table' );

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


##############################################################################
#
# _write_table()
#
# Write the <table> element.
#
method write_table {
    my $schema           = 'http://schemas.openxmlformats.org/';
    my $xmlns            = $schema ~ 'spreadsheetml/2006/main';
    my $id               = %!properties<id>;
    my $name             = %!properties<name>;
    my $display-name     = %!properties<name>;
    my $ref              = %!properties<range>;
    my $totals-row-shown = %!properties<totals_row_shown>;
    my $header-row-count = %!properties<header_row_count>;

    my @attributes = (
        'xmlns'       => $xmlns,
        'id'          => $id,
        'name'        => $name,
        'displayName' => $display-name,
        'ref'         => $ref,
    );

    @attributes.push: ( 'headerRowCount' => 0 ) if !$header-row-count;

    if $totals-row-shown {
        @attributes.push: ( 'totalsRowCount' => 1 );
    }
    else {
        @attributes.push: ( 'totalsRowShown' => 0 );
    }


    self.xml_start_tag( 'table', @attributes );
}


##############################################################################
#
# _write_auto_filter()
#
# Write the <autoFilter> element.
#
method write_auto_filter {
    my $autofilter = %!properties<autofilter>;

    return unless $autofilter;

    my @attributes = ( 'ref' => $autofilter, );

    self.xml_empty_tag( 'autoFilter', @attributes );
}


##############################################################################
#
# _write_table_columns()
#
# Write the <tableColumns> element.
#
method write_table_columns {
    my @columns = %!properties<columns>;

    my $count = +@columns;

    my @attributes = ( 'count' => $count, );

    self.xml_start_tag( 'tableColumns', @attributes );

    for @columns -> $col-data {

        # Write the tableColumn element.
        self.write_table_column( $col-data );
    }

    self.xml_end_tag( 'tableColumns' );
}


##############################################################################
#
# _write_table_column()
#
# Write the <tableColumn> element.
#
method write_table_column($col-data) {
    my @attributes = (
        'id'   => $col-data<id>,
        'name' => $col-data<name>,
    );


    if $col-data<total_string> {
        @attributes.push: ( totalsRowLabel => $col-data<total_string> );
    }
    elsif $col-data<total_function> {
        @attributes.push: ( totalsRowFunction => $col-data<total_function> );
    }


    if $col-data<format>.defined {
        @attributes.push: ( dataDxfId => $col-data<format> );
    }

    if $col-data<formula> {
        self.xml_start_tag( 'tableColumn', @attributes );

        # Write the calculatedColumnFormula element.
        self.write_calculated_column_formula( $col-data<formula> );

        self.xml_end_tag( 'tableColumn' );
    }
    else {
        self.xml_empty_tag( 'tableColumn', @attributes );
    }

}


##############################################################################
#
# _write_table_style_info()
#
# Write the <tableStyleInfo> element.
#
method write_table_style_info {
    my $props = %!properties;

    my $name                = $props<style>;
    my $show-first-column   = $props<show_first_col>;
    my $show-last-column    = $props<show_last_col>;
    my $show-row-stripes    = $props<show_row_stripes>;
    my $show-column-stripes = $props<show_col_stripes>;

    my @attributes = (
        'name'              => $name,
        'showFirstColumn'   => $show-first-column,
        'showLastColumn'    => $show-last-column,
        'showRowStripes'    => $show-row-stripes,
        'showColumnStripes' => $show-column-stripes,
    );

    self.xml_empty_tag( 'tableStyleInfo', @attributes );
}


##############################################################################
#
# _write_calculated_column_formula()
#
# Write the <calculatedColumnFormula> element.
#
method write_calculated_column_formula($formula) {
    self.xml_data_element( 'calculatedColumnFormula', $formula );
}


=begin pod


__END__

=pod

=head1 NAME

Table - A class for writing the Excel XLSX Table file.

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
