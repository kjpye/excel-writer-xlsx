#NYI package Excel::Writer::XLSX::Chartsheet;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Chartsheet - A class for writing the Excel XLSX Chartsheet files.
#NYI #
#NYI # Used in conjunction with Excel::Writer::XLSX
#NYI #
#NYI # Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#NYI #
#NYI # Documentation after __END__
#NYI #
#NYI 
#NYI # perltidy with the following options: -mbl=2 -pt=0 -nola
#NYI 
#NYI use 5.008002;
#NYI use strict;
#NYI use warnings;
#NYI use Exporter;
#NYI use Excel::Writer::XLSX::Worksheet;
#NYI 
#NYI our @ISA     = qw(Excel::Writer::XLSX::Worksheet);
#NYI our $VERSION = '0.96';
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Public and private API methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # new()
#NYI #
#NYI # Constructor.
#NYI #
#NYI sub new {
#NYI 
#NYI     my $class = shift;
#NYI     my $self  = Excel::Writer::XLSX::Worksheet->new( @_ );
#NYI 
#NYI     $self->{_drawing}           = 1;
#NYI     $self->{_is_chartsheet}     = 1;
#NYI     $self->{_chart}             = undef;
#NYI     $self->{_charts}            = [1];
#NYI     $self->{_zoom_scale_normal} = 0;
#NYI     $self->{_orientation}       = 0;
#NYI 
#NYI     bless $self, $class;
#NYI 
#NYI     return $self;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _assemble_xml_file()
#NYI #
#NYI # Assemble and write the XML file.
#NYI #
#NYI sub _assemble_xml_file {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_declaration;
#NYI 
#NYI     # Write the root chartsheet element.
#NYI     $self->_write_chartsheet();
#NYI 
#NYI     # Write the worksheet properties.
#NYI     $self->_write_sheet_pr();
#NYI 
#NYI     # Write the sheet view properties.
#NYI     $self->_write_sheet_views();
#NYI 
#NYI     # Write the sheetProtection element.
#NYI     $self->_write_sheet_protection();
#NYI 
#NYI     # Write the printOptions element.
#NYI     $self->_write_print_options();
#NYI 
#NYI     # Write the worksheet page_margins.
#NYI     $self->_write_page_margins();
#NYI 
#NYI     # Write the worksheet page setup.
#NYI     $self->_write_page_setup();
#NYI 
#NYI     # Write the headerFooter element.
#NYI     $self->_write_header_footer();
#NYI 
#NYI     # Write the drawing element.
#NYI     $self->_write_drawings();
#NYI 
#NYI     # Close the worksheet tag.
#NYI     $self->xml_end_tag( 'chartsheet' );
#NYI 
#NYI     # Close the XML writer filehandle.
#NYI     $self->xml_get_fh()->close();
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Public methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI # Over-ride parent protect() method to protect both worksheet and chart.
#NYI sub protect {
#NYI 
#NYI     my $self     = shift;
#NYI     my $password = shift || '';
#NYI     my $options  = shift || {};
#NYI 
#NYI     $self->{_chart}->{_protection} = 1;
#NYI 
#NYI     $options->{sheet}     = 0;
#NYI     $options->{content}   = 1;
#NYI     $options->{scenarios} = 1;
#NYI 
#NYI     $self->SUPER::protect( $password, $options );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Encapsulated Chart methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI sub add_series         { return shift->{_chart}->add_series( @_ ) }
#NYI sub combine            { return shift->{_chart}->combine( @_ ) }
#NYI sub set_x_axis         { return shift->{_chart}->set_x_axis( @_ ) }
#NYI sub set_y_axis         { return shift->{_chart}->set_y_axis( @_ ) }
#NYI sub set_x2_axis        { return shift->{_chart}->set_x2_axis( @_ ) }
#NYI sub set_y2_axis        { return shift->{_chart}->set_y2_axis( @_ ) }
#NYI sub set_title          { return shift->{_chart}->set_title( @_ ) }
#NYI sub set_legend         { return shift->{_chart}->set_legend( @_ ) }
#NYI sub set_plotarea       { return shift->{_chart}->set_plotarea( @_ ) }
#NYI sub set_chartarea      { return shift->{_chart}->set_chartarea( @_ ) }
#NYI sub set_style          { return shift->{_chart}->set_style( @_ ) }
#NYI sub show_blanks_as     { return shift->{_chart}->show_blanks_as( @_ ) }
#NYI sub show_hidden_data   { return shift->{_chart}->show_hidden_data( @_ ) }
#NYI sub set_size           { return shift->{_chart}->set_size( @_ ) }
#NYI sub set_table          { return shift->{_chart}->set_table( @_ ) }
#NYI sub set_up_down_bars   { return shift->{_chart}->set_up_down_bars( @_ ) }
#NYI sub set_drop_lines     { return shift->{_chart}->set_drop_lines( @_ ) }
#NYI sub set_high_low_lines { return shift->{_chart}->high_low_lines( @_ ) }
#NYI 
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Internal methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_chart()
#NYI #
#NYI # Set up chart/drawings.
#NYI #
#NYI sub _prepare_chart {
#NYI 
#NYI     my $self       = shift;
#NYI     my $index      = shift;
#NYI     my $chart_id   = shift;
#NYI     my $drawing_id = shift;
#NYI 
#NYI     $self->{_chart}->{_id} = $chart_id -1;
#NYI 
#NYI     my $drawing = Excel::Writer::XLSX::Drawing->new();
#NYI     $self->{_drawing} = $drawing;
#NYI     $self->{_drawing}->{_orientation} = $self->{_orientation};
#NYI 
#NYI     push @{ $self->{_external_drawing_links} },
#NYI       [ '/drawing', '../drawings/drawing' . $drawing_id . '.xml' ];
#NYI 
#NYI     push @{ $self->{_drawing_links} },
#NYI       [ '/chart', '../charts/chart' . $chart_id . '.xml' ];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # XML writing methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_chartsheet()
#NYI #
#NYI # Write the <chartsheet> element. This is the root element of Chartsheet.
#NYI #
#NYI sub _write_chartsheet {
#NYI 
#NYI     my $self                   = shift;
#NYI     my $schema                 = 'http://schemas.openxmlformats.org/';
#NYI     my $xmlns                  = $schema . 'spreadsheetml/2006/main';
#NYI     my $xmlns_r                = $schema . 'officeDocument/2006/relationships';
#NYI     my $xmlns_mc               = $schema . 'markup-compatibility/2006';
#NYI     my $xmlns_mv               = 'urn:schemas-microsoft-com:mac:vml';
#NYI     my $mc_ignorable           = 'mv';
#NYI     my $mc_preserve_attributes = 'mv:*';
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns'   => $xmlns,
#NYI         'xmlns:r' => $xmlns_r,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'chartsheet', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_sheet_pr()
#NYI #
#NYI # Write the <sheetPr> element for Sheet level properties.
#NYI #
#NYI sub _write_sheet_pr {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI 
#NYI     push @attributes, ( 'filterMode' => 1 ) if $self->{_filter_on};
#NYI 
#NYI     if ( $self->{_fit_page} || $self->{_tab_color} ) {
#NYI         $self->xml_start_tag( 'sheetPr', @attributes );
#NYI         $self->_write_tab_color();
#NYI         $self->_write_page_set_up_pr();
#NYI         $self->xml_end_tag( 'sheetPr' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'sheetPr', @attributes );
#NYI     }
#NYI }
#NYI 
#NYI 1;
#NYI 
#NYI 
#NYI __END__
#NYI 
#NYI =pod
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Chartsheet - A class for writing the Excel XLSX Chartsheet files.
#NYI 
#NYI =head1 SYNOPSIS
#NYI 
#NYI See the documentation for L<Excel::Writer::XLSX>.
#NYI 
#NYI =head1 DESCRIPTION
#NYI 
#NYI This module is used in conjunction with L<Excel::Writer::XLSX>.
#NYI 
#NYI =head1 AUTHOR
#NYI 
#NYI John McNamara jmcnamara@cpan.org
#NYI 
#NYI =head1 COPYRIGHT
#NYI 
#NYI (c) MM-MMXVII, John McNamara.
#NYI 
#NYI All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
#NYI 
#NYI =head1 LICENSE
#NYI 
#NYI Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.
#NYI 
#NYI =head1 DISCLAIMER OF WARRANTY
#NYI 
#NYI See the documentation for L<Excel::Writer::XLSX>.
#NYI 
#NYI =cut
