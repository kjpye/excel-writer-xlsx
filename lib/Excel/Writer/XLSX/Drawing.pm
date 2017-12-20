#NYI package Excel::Writer::XLSX::Drawing;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Drawing - A class for writing the Excel XLSX drawing.xml file.
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
#NYI use Carp;
#NYI use Excel::Writer::XLSX::Package::XMLwriter;
#NYI use Excel::Writer::XLSX::Worksheet;
#NYI 
#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
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
#NYI     my $fh    = shift;
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );
#NYI 
#NYI     $self->{_drawings}    = [];
#NYI     $self->{_embedded}    = 0;
#NYI     $self->{_orientation} = 0;
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
#NYI     # Write the xdr:wsDr element.
#NYI     $self->_write_drawing_workspace();
#NYI 
#NYI     if ( $self->{_embedded} ) {
#NYI 
#NYI         my $index = 0;
#NYI         for my $dimensions ( @{ $self->{_drawings} } ) {
#NYI 
#NYI             # Write the xdr:twoCellAnchor element.
#NYI             $self->_write_two_cell_anchor( ++$index, @$dimensions );
#NYI         }
#NYI 
#NYI     }
#NYI     else {
#NYI         my $index = 0;
#NYI 
#NYI         # Write the xdr:absoluteAnchor element.
#NYI         $self->_write_absolute_anchor( ++$index );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'xdr:wsDr' );
#NYI 
#NYI     # Close the XML writer filehandle.
#NYI     $self->xml_get_fh()->close();
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _add_drawing_object()
#NYI #
#NYI # Add a chart, image or shape sub object to the drawing.
#NYI #
#NYI sub _add_drawing_object {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     push @{ $self->{_drawings} }, [@_];
#NYI }
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
#NYI # XML writing methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_drawing_workspace()
#NYI #
#NYI # Write the <xdr:wsDr> element.
#NYI #
#NYI sub _write_drawing_workspace {
#NYI 
#NYI     my $self      = shift;
#NYI     my $schema    = 'http://schemas.openxmlformats.org/drawingml/';
#NYI     my $xmlns_xdr = $schema . '2006/spreadsheetDrawing';
#NYI     my $xmlns_a   = $schema . '2006/main';
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns:xdr' => $xmlns_xdr,
#NYI         'xmlns:a'   => $xmlns_a,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'xdr:wsDr', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_two_cell_anchor()
#NYI #
#NYI # Write the <xdr:twoCellAnchor> element.
#NYI #
#NYI sub _write_two_cell_anchor {
#NYI 
#NYI     my $self            = shift;
#NYI     my $index           = shift;
#NYI     my $type            = shift;
#NYI     my $col_from        = shift;
#NYI     my $row_from        = shift;
#NYI     my $col_from_offset = shift;
#NYI     my $row_from_offset = shift;
#NYI     my $col_to          = shift;
#NYI     my $row_to          = shift;
#NYI     my $col_to_offset   = shift;
#NYI     my $row_to_offset   = shift;
#NYI     my $col_absolute    = shift;
#NYI     my $row_absolute    = shift;
#NYI     my $width           = shift;
#NYI     my $height          = shift;
#NYI     my $description     = shift;
#NYI     my $shape           = shift;
#NYI 
#NYI     my @attributes = ();
#NYI 
#NYI 
#NYI     # Add attribute for images.
#NYI     if ( $type == 2 ) {
#NYI         push @attributes, ( editAs => 'oneCell' );
#NYI     }
#NYI 
#NYI     # Add editAs attribute for shapes.
#NYI     push @attributes, ( editAs => $shape->{_editAs} ) if $shape->{_editAs};
#NYI 
#NYI     $self->xml_start_tag( 'xdr:twoCellAnchor', @attributes );
#NYI 
#NYI     # Write the xdr:from element.
#NYI     $self->_write_from(
#NYI         $col_from,
#NYI         $row_from,
#NYI         $col_from_offset,
#NYI         $row_from_offset,
#NYI 
#NYI     );
#NYI 
#NYI     # Write the xdr:from element.
#NYI     $self->_write_to(
#NYI         $col_to,
#NYI         $row_to,
#NYI         $col_to_offset,
#NYI         $row_to_offset,
#NYI 
#NYI     );
#NYI 
#NYI     if ( $type == 1 ) {
#NYI 
#NYI         # Graphic frame.
#NYI 
#NYI         # Write the xdr:graphicFrame element for charts.
#NYI         $self->_write_graphic_frame( $index, $description );
#NYI     }
#NYI     elsif ( $type == 2 ) {
#NYI 
#NYI         # Write the xdr:pic element.
#NYI         $self->_write_pic( $index, $col_absolute, $row_absolute, $width,
#NYI             $height, $description );
#NYI     }
#NYI     else {
#NYI 
#NYI         # Write the xdr:sp element for shapes.
#NYI         $self->_write_sp( $index, $col_absolute, $row_absolute, $width, $height,
#NYI             $shape );
#NYI     }
#NYI 
#NYI     # Write the xdr:clientData element.
#NYI     $self->_write_client_data();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:twoCellAnchor' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_absolute_anchor()
#NYI #
#NYI # Write the <xdr:absoluteAnchor> element.
#NYI #
#NYI sub _write_absolute_anchor {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:absoluteAnchor' );
#NYI 
#NYI     # Different co-ordinates for horizonatal (= 0) and vertical (= 1).
#NYI     if ( $self->{_orientation} == 0 ) {
#NYI 
#NYI         # Write the xdr:pos element.
#NYI         $self->_write_pos( 0, 0 );
#NYI 
#NYI         # Write the xdr:ext element.
#NYI         $self->_write_ext( 9308969, 6078325 );
#NYI 
#NYI     }
#NYI     else {
#NYI 
#NYI         # Write the xdr:pos element.
#NYI         $self->_write_pos( 0, -47625 );
#NYI 
#NYI         # Write the xdr:ext element.
#NYI         $self->_write_ext( 6162675, 6124575 );
#NYI 
#NYI     }
#NYI 
#NYI 
#NYI     # Write the xdr:graphicFrame element.
#NYI     $self->_write_graphic_frame( $index );
#NYI 
#NYI     # Write the xdr:clientData element.
#NYI     $self->_write_client_data();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:absoluteAnchor' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_from()
#NYI #
#NYI # Write the <xdr:from> element.
#NYI #
#NYI sub _write_from {
#NYI 
#NYI     my $self       = shift;
#NYI     my $col        = shift;
#NYI     my $row        = shift;
#NYI     my $col_offset = shift;
#NYI     my $row_offset = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:from' );
#NYI 
#NYI     # Write the xdr:col element.
#NYI     $self->_write_col( $col );
#NYI 
#NYI     # Write the xdr:colOff element.
#NYI     $self->_write_col_off( $col_offset );
#NYI 
#NYI     # Write the xdr:row element.
#NYI     $self->_write_row( $row );
#NYI 
#NYI     # Write the xdr:rowOff element.
#NYI     $self->_write_row_off( $row_offset );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:from' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_to()
#NYI #
#NYI # Write the <xdr:to> element.
#NYI #
#NYI sub _write_to {
#NYI 
#NYI     my $self       = shift;
#NYI     my $col        = shift;
#NYI     my $row        = shift;
#NYI     my $col_offset = shift;
#NYI     my $row_offset = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:to' );
#NYI 
#NYI     # Write the xdr:col element.
#NYI     $self->_write_col( $col );
#NYI 
#NYI     # Write the xdr:colOff element.
#NYI     $self->_write_col_off( $col_offset );
#NYI 
#NYI     # Write the xdr:row element.
#NYI     $self->_write_row( $row );
#NYI 
#NYI     # Write the xdr:rowOff element.
#NYI     $self->_write_row_off( $row_offset );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:to' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_col()
#NYI #
#NYI # Write the <xdr:col> element.
#NYI #
#NYI sub _write_col {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'xdr:col', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_col_off()
#NYI #
#NYI # Write the <xdr:colOff> element.
#NYI #
#NYI sub _write_col_off {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'xdr:colOff', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_row()
#NYI #
#NYI # Write the <xdr:row> element.
#NYI #
#NYI sub _write_row {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'xdr:row', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_row_off()
#NYI #
#NYI # Write the <xdr:rowOff> element.
#NYI #
#NYI sub _write_row_off {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'xdr:rowOff', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_pos()
#NYI #
#NYI # Write the <xdr:pos> element.
#NYI #
#NYI sub _write_pos {
#NYI 
#NYI     my $self = shift;
#NYI     my $x    = shift;
#NYI     my $y    = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'x' => $x,
#NYI         'y' => $y,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'xdr:pos', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_ext()
#NYI #
#NYI # Write the <xdr:ext> element.
#NYI #
#NYI sub _write_ext {
#NYI 
#NYI     my $self = shift;
#NYI     my $cx   = shift;
#NYI     my $cy   = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'cx' => $cx,
#NYI         'cy' => $cy,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'xdr:ext', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_graphic_frame()
#NYI #
#NYI # Write the <xdr:graphicFrame> element.
#NYI #
#NYI sub _write_graphic_frame {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI     my $name  = shift;
#NYI     my $macro = '';
#NYI 
#NYI     my @attributes = ( 'macro' => $macro );
#NYI 
#NYI     $self->xml_start_tag( 'xdr:graphicFrame', @attributes );
#NYI 
#NYI     # Write the xdr:nvGraphicFramePr element.
#NYI     $self->_write_nv_graphic_frame_pr( $index, $name );
#NYI 
#NYI     # Write the xdr:xfrm element.
#NYI     $self->_write_xfrm();
#NYI 
#NYI     # Write the a:graphic element.
#NYI     $self->_write_atag_graphic( $index );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:graphicFrame' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_nv_graphic_frame_pr()
#NYI #
#NYI # Write the <xdr:nvGraphicFramePr> element.
#NYI #
#NYI sub _write_nv_graphic_frame_pr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI     my $name  = shift;
#NYI 
#NYI     if ( !$name ) {
#NYI         $name = 'Chart ' . $index;
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'xdr:nvGraphicFramePr' );
#NYI 
#NYI     # Write the xdr:cNvPr element.
#NYI     $self->_write_c_nv_pr( $index + 1, $name );
#NYI 
#NYI     # Write the xdr:cNvGraphicFramePr element.
#NYI     $self->_write_c_nv_graphic_frame_pr();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:nvGraphicFramePr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_nv_pr()
#NYI #
#NYI # Write the <xdr:cNvPr> element.
#NYI #
#NYI sub _write_c_nv_pr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $id    = shift;
#NYI     my $name  = shift;
#NYI     my $descr = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'id'   => $id,
#NYI         'name' => $name,
#NYI     );
#NYI 
#NYI     # Add description attribute for images.
#NYI     if ( defined $descr ) {
#NYI         push @attributes, ( descr => $descr );
#NYI     }
#NYI 
#NYI     $self->xml_empty_tag( 'xdr:cNvPr', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_nv_graphic_frame_pr()
#NYI #
#NYI # Write the <xdr:cNvGraphicFramePr> element.
#NYI #
#NYI sub _write_c_nv_graphic_frame_pr {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     if ( $self->{_embedded} ) {
#NYI         $self->xml_empty_tag( 'xdr:cNvGraphicFramePr' );
#NYI     }
#NYI     else {
#NYI         $self->xml_start_tag( 'xdr:cNvGraphicFramePr' );
#NYI 
#NYI         # Write the a:graphicFrameLocks element.
#NYI         $self->_write_a_graphic_frame_locks();
#NYI 
#NYI         $self->xml_end_tag( 'xdr:cNvGraphicFramePr' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_graphic_frame_locks()
#NYI #
#NYI # Write the <a:graphicFrameLocks> element.
#NYI #
#NYI sub _write_a_graphic_frame_locks {
#NYI 
#NYI     my $self   = shift;
#NYI     my $no_grp = 1;
#NYI 
#NYI     my @attributes = ( 'noGrp' => $no_grp );
#NYI 
#NYI     $self->xml_empty_tag( 'a:graphicFrameLocks', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_xfrm()
#NYI #
#NYI # Write the <xdr:xfrm> element.
#NYI #
#NYI sub _write_xfrm {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:xfrm' );
#NYI 
#NYI     # Write the xfrmOffset element.
#NYI     $self->_write_xfrm_offset();
#NYI 
#NYI     # Write the xfrmOffset element.
#NYI     $self->_write_xfrm_extension();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:xfrm' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_xfrm_offset()
#NYI #
#NYI # Write the <a:off> xfrm sub-element.
#NYI #
#NYI sub _write_xfrm_offset {
#NYI 
#NYI     my $self = shift;
#NYI     my $x    = 0;
#NYI     my $y    = 0;
#NYI 
#NYI     my @attributes = (
#NYI         'x' => $x,
#NYI         'y' => $y,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'a:off', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_xfrm_extension()
#NYI #
#NYI # Write the <a:ext> xfrm sub-element.
#NYI #
#NYI sub _write_xfrm_extension {
#NYI 
#NYI     my $self = shift;
#NYI     my $x    = 0;
#NYI     my $y    = 0;
#NYI 
#NYI     my @attributes = (
#NYI         'cx' => $x,
#NYI         'cy' => $y,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'a:ext', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_atag_graphic()
#NYI #
#NYI # Write the <a:graphic> element.
#NYI #
#NYI sub _write_atag_graphic {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:graphic' );
#NYI 
#NYI     # Write the a:graphicData element.
#NYI     $self->_write_atag_graphic_data( $index );
#NYI 
#NYI     $self->xml_end_tag( 'a:graphic' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_atag_graphic_data()
#NYI #
#NYI # Write the <a:graphicData> element.
#NYI #
#NYI sub _write_atag_graphic_data {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI     my $uri   = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
#NYI 
#NYI     my @attributes = ( 'uri' => $uri, );
#NYI 
#NYI     $self->xml_start_tag( 'a:graphicData', @attributes );
#NYI 
#NYI     # Write the c:chart element.
#NYI     $self->_write_c_chart( 'rId' . $index );
#NYI 
#NYI     $self->xml_end_tag( 'a:graphicData' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_chart()
#NYI #
#NYI # Write the <c:chart> element.
#NYI #
#NYI sub _write_c_chart {
#NYI 
#NYI     my $self    = shift;
#NYI     my $r_id    = shift;
#NYI     my $schema  = 'http://schemas.openxmlformats.org/';
#NYI     my $xmlns_c = $schema . 'drawingml/2006/chart';
#NYI     my $xmlns_r = $schema . 'officeDocument/2006/relationships';
#NYI 
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns:c' => $xmlns_c,
#NYI         'xmlns:r' => $xmlns_r,
#NYI         'r:id'    => $r_id,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'c:chart', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_client_data()
#NYI #
#NYI # Write the <xdr:clientData> element.
#NYI #
#NYI sub _write_client_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_empty_tag( 'xdr:clientData' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_sp()
#NYI #
#NYI # Write the <xdr:sp> element.
#NYI #
#NYI sub _write_sp {
#NYI 
#NYI     my $self         = shift;
#NYI     my $index        = shift;
#NYI     my $col_absolute = shift;
#NYI     my $row_absolute = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $shape        = shift;
#NYI 
#NYI     if ( $shape->{_connect} ) {
#NYI         my @attributes = ( macro => '' );
#NYI         $self->xml_start_tag( 'xdr:cxnSp', @attributes );
#NYI 
#NYI         # Write the xdr:nvCxnSpPr element.
#NYI         $self->_write_nv_cxn_sp_pr( $index, $shape );
#NYI 
#NYI         # Write the xdr:spPr element.
#NYI         $self->_write_xdr_sp_pr( $index, $col_absolute, $row_absolute, $width,
#NYI             $height, $shape );
#NYI 
#NYI         $self->xml_end_tag( 'xdr:cxnSp' );
#NYI     }
#NYI     else {
#NYI 
#NYI         # Add attribute for shapes.
#NYI         my @attributes = ( macro => '', textlink => '' );
#NYI         $self->xml_start_tag( 'xdr:sp', @attributes );
#NYI 
#NYI         # Write the xdr:nvSpPr element.
#NYI         $self->_write_nv_sp_pr( $index, $shape );
#NYI 
#NYI         # Write the xdr:spPr element.
#NYI         $self->_write_xdr_sp_pr( $index, $col_absolute, $row_absolute, $width,
#NYI             $height, $shape );
#NYI 
#NYI         # Write the xdr:txBody element.
#NYI         if ( $shape->{_text} ) {
#NYI             $self->_write_txBody( $col_absolute, $row_absolute, $width, $height,
#NYI                 $shape );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'xdr:sp' );
#NYI     }
#NYI }
#NYI ##############################################################################
#NYI #
#NYI # _write_nv_cxn_sp_pr()
#NYI #
#NYI # Write the <xdr:nvCxnSpPr> element.
#NYI #
#NYI sub _write_nv_cxn_sp_pr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI     my $shape = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:nvCxnSpPr' );
#NYI 
#NYI     $shape->{_name} = join( ' ', $shape->{_type}, $index )
#NYI       unless defined $shape->{_name};
#NYI     $self->_write_c_nv_pr( $shape->{_id}, $shape->{_name} );
#NYI 
#NYI     $self->xml_start_tag( 'xdr:cNvCxnSpPr' );
#NYI 
#NYI     my @attributes = ( noChangeShapeType => '1' );
#NYI     $self->xml_empty_tag( 'a:cxnSpLocks', @attributes );
#NYI 
#NYI     if ( $shape->{_start} ) {
#NYI         @attributes =
#NYI           ( 'id' => $shape->{_start}, 'idx' => $shape->{_start_index} );
#NYI         $self->xml_empty_tag( 'a:stCxn', @attributes );
#NYI     }
#NYI 
#NYI     if ( $shape->{_end} ) {
#NYI         @attributes = ( 'id' => $shape->{_end}, 'idx' => $shape->{_end_index} );
#NYI         $self->xml_empty_tag( 'a:endCxn', @attributes );
#NYI     }
#NYI     $self->xml_end_tag( 'xdr:cNvCxnSpPr' );
#NYI     $self->xml_end_tag( 'xdr:nvCxnSpPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_nv_sp_pr()
#NYI #
#NYI # Write the <xdr:NvSpPr> element.
#NYI #
#NYI sub _write_nv_sp_pr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI     my $shape = shift;
#NYI 
#NYI     my @attributes = ();
#NYI 
#NYI     $self->xml_start_tag( 'xdr:nvSpPr' );
#NYI 
#NYI     my $shape_name = $shape->{_type} . ' ' . $index;
#NYI 
#NYI     $self->_write_c_nv_pr( $shape->{_id}, $shape_name );
#NYI 
#NYI     @attributes = ( 'txBox' => 1 ) if $shape->{_txBox};
#NYI 
#NYI     $self->xml_start_tag( 'xdr:cNvSpPr', @attributes );
#NYI 
#NYI     @attributes = ( noChangeArrowheads => '1' );
#NYI 
#NYI     $self->xml_empty_tag( 'a:spLocks', @attributes );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:cNvSpPr' );
#NYI     $self->xml_end_tag( 'xdr:nvSpPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_pic()
#NYI #
#NYI # Write the <xdr:pic> element.
#NYI #
#NYI sub _write_pic {
#NYI 
#NYI     my $self         = shift;
#NYI     my $index        = shift;
#NYI     my $col_absolute = shift;
#NYI     my $row_absolute = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $description  = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:pic' );
#NYI 
#NYI     # Write the xdr:nvPicPr element.
#NYI     $self->_write_nv_pic_pr( $index, $description );
#NYI 
#NYI     # Write the xdr:blipFill element.
#NYI     $self->_write_blip_fill( $index );
#NYI 
#NYI     # Pictures are rectangle shapes by default.
#NYI     my $shape = { _type => 'rect' };
#NYI 
#NYI     # Write the xdr:spPr element.
#NYI     $self->_write_sp_pr( $col_absolute, $row_absolute, $width, $height,
#NYI         $shape );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:pic' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_nv_pic_pr()
#NYI #
#NYI # Write the <xdr:nvPicPr> element.
#NYI #
#NYI sub _write_nv_pic_pr {
#NYI 
#NYI     my $self        = shift;
#NYI     my $index       = shift;
#NYI     my $description = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:nvPicPr' );
#NYI 
#NYI     # Write the xdr:cNvPr element.
#NYI     $self->_write_c_nv_pr( $index + 1, 'Picture ' . $index, $description );
#NYI 
#NYI     # Write the xdr:cNvPicPr element.
#NYI     $self->_write_c_nv_pic_pr();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:nvPicPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_nv_pic_pr()
#NYI #
#NYI # Write the <xdr:cNvPicPr> element.
#NYI #
#NYI sub _write_c_nv_pic_pr {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:cNvPicPr' );
#NYI 
#NYI     # Write the a:picLocks element.
#NYI     $self->_write_a_pic_locks();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:cNvPicPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_pic_locks()
#NYI #
#NYI # Write the <a:picLocks> element.
#NYI #
#NYI sub _write_a_pic_locks {
#NYI 
#NYI     my $self             = shift;
#NYI     my $no_change_aspect = 1;
#NYI 
#NYI     my @attributes = ( 'noChangeAspect' => $no_change_aspect );
#NYI 
#NYI     $self->xml_empty_tag( 'a:picLocks', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_blip_fill()
#NYI #
#NYI # Write the <xdr:blipFill> element.
#NYI #
#NYI sub _write_blip_fill {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI 
#NYI     $self->xml_start_tag( 'xdr:blipFill' );
#NYI 
#NYI     # Write the a:blip element.
#NYI     $self->_write_a_blip( $index );
#NYI 
#NYI     # Write the a:stretch element.
#NYI     $self->_write_a_stretch();
#NYI 
#NYI     $self->xml_end_tag( 'xdr:blipFill' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_blip()
#NYI #
#NYI # Write the <a:blip> element.
#NYI #
#NYI sub _write_a_blip {
#NYI 
#NYI     my $self    = shift;
#NYI     my $index   = shift;
#NYI     my $schema  = 'http://schemas.openxmlformats.org/officeDocument/';
#NYI     my $xmlns_r = $schema . '2006/relationships';
#NYI     my $r_embed = 'rId' . $index;
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns:r' => $xmlns_r,
#NYI         'r:embed' => $r_embed,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'a:blip', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_stretch()
#NYI #
#NYI # Write the <a:stretch> element.
#NYI #
#NYI sub _write_a_stretch {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:stretch' );
#NYI 
#NYI     # Write the a:fillRect element.
#NYI     $self->_write_a_fill_rect();
#NYI 
#NYI     $self->xml_end_tag( 'a:stretch' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_fill_rect()
#NYI #
#NYI # Write the <a:fillRect> element.
#NYI #
#NYI sub _write_a_fill_rect {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_empty_tag( 'a:fillRect' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_sp_pr()
#NYI #
#NYI # Write the <xdr:spPr> element, for charts.
#NYI #
#NYI sub _write_sp_pr {
#NYI 
#NYI     my $self         = shift;
#NYI     my $col_absolute = shift;
#NYI     my $row_absolute = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $shape        = shift || {};
#NYI 
#NYI     $self->xml_start_tag( 'xdr:spPr' );
#NYI 
#NYI     # Write the a:xfrm element.
#NYI     $self->_write_a_xfrm( $col_absolute, $row_absolute, $width, $height );
#NYI 
#NYI     # Write the a:prstGeom element.
#NYI     $self->_write_a_prst_geom( $shape );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:spPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_xdr_sp_pr()
#NYI #
#NYI # Write the <xdr:spPr> element for shapes.
#NYI #
#NYI sub _write_xdr_sp_pr {
#NYI 
#NYI     my $self         = shift;
#NYI     my $index        = shift;
#NYI     my $col_absolute = shift;
#NYI     my $row_absolute = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $shape        = shift;
#NYI 
#NYI     my @attributes = ( 'bwMode' => 'auto' );
#NYI 
#NYI     $self->xml_start_tag( 'xdr:spPr', @attributes );
#NYI 
#NYI     # Write the a:xfrm element.
#NYI     $self->_write_a_xfrm( $col_absolute, $row_absolute, $width, $height,
#NYI         $shape );
#NYI 
#NYI     # Write the a:prstGeom element.
#NYI     $self->_write_a_prst_geom( $shape );
#NYI 
#NYI     my $fill = $shape->{_fill};
#NYI 
#NYI     if ( length $fill > 1 ) {
#NYI 
#NYI         # Write the a:solidFill element.
#NYI         $self->_write_a_solid_fill( $fill );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'a:noFill' );
#NYI     }
#NYI 
#NYI     # Write the a:ln element.
#NYI     $self->_write_a_ln( $shape );
#NYI 
#NYI     $self->xml_end_tag( 'xdr:spPr' );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_xfrm()
#NYI #
#NYI # Write the <a:xfrm> element.
#NYI #
#NYI sub _write_a_xfrm {
#NYI 
#NYI     my $self         = shift;
#NYI     my $col_absolute = shift;
#NYI     my $row_absolute = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $shape        = shift || {};
#NYI     my @attributes   = ();
#NYI 
#NYI     my $rotation = $shape->{_rotation} || 0;
#NYI     $rotation *= 60000;
#NYI 
#NYI     push( @attributes, ( 'rot'   => $rotation ) ) if $rotation;
#NYI     push( @attributes, ( 'flipH' => 1 ) )         if $shape->{_flip_h};
#NYI     push( @attributes, ( 'flipV' => 1 ) )         if $shape->{_flip_v};
#NYI 
#NYI     $self->xml_start_tag( 'a:xfrm', @attributes );
#NYI 
#NYI     # Write the a:off element.
#NYI     $self->_write_a_off( $col_absolute, $row_absolute );
#NYI 
#NYI     # Write the a:ext element.
#NYI     $self->_write_a_ext( $width, $height );
#NYI 
#NYI     $self->xml_end_tag( 'a:xfrm' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_off()
#NYI #
#NYI # Write the <a:off> element.
#NYI #
#NYI sub _write_a_off {
#NYI 
#NYI     my $self = shift;
#NYI     my $x    = shift;
#NYI     my $y    = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'x' => $x,
#NYI         'y' => $y,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'a:off', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_ext()
#NYI #
#NYI # Write the <a:ext> element.
#NYI #
#NYI sub _write_a_ext {
#NYI 
#NYI     my $self = shift;
#NYI     my $cx   = shift;
#NYI     my $cy   = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'cx' => $cx,
#NYI         'cy' => $cy,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'a:ext', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_prst_geom()
#NYI #
#NYI # Write the <a:prstGeom> element.
#NYI #
#NYI sub _write_a_prst_geom {
#NYI 
#NYI     my $self = shift;
#NYI     my $shape = shift || {};
#NYI 
#NYI     my @attributes = ();
#NYI 
#NYI     @attributes = ( 'prst' => $shape->{_type} ) if $shape->{_type};
#NYI 
#NYI     $self->xml_start_tag( 'a:prstGeom', @attributes );
#NYI 
#NYI     # Write the a:avLst element.
#NYI     $self->_write_a_av_lst( $shape );
#NYI 
#NYI     $self->xml_end_tag( 'a:prstGeom' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_av_lst()
#NYI #
#NYI # Write the <a:avLst> element.
#NYI #
#NYI sub _write_a_av_lst {
#NYI 
#NYI     my $self        = shift;
#NYI     my $shape       = shift || {};
#NYI     my $adjustments = [];
#NYI 
#NYI     if ( defined $shape->{_adjustments} ) {
#NYI         $adjustments = $shape->{_adjustments};
#NYI     }
#NYI 
#NYI     if ( @$adjustments ) {
#NYI         $self->xml_start_tag( 'a:avLst' );
#NYI 
#NYI         my $i = 0;
#NYI         foreach my $adj ( @{$adjustments} ) {
#NYI             $i++;
#NYI 
#NYI             # Only connectors have multiple adjustments.
#NYI             my $suffix = $shape->{_connect} ? $i : '';
#NYI 
#NYI             # Scale Adjustments: 100,000 = 100%.
#NYI             my $adj_int = int( $adj * 1000 );
#NYI 
#NYI             my @attributes =
#NYI               ( name => 'adj' . $suffix, fmla => "val $adj_int" );
#NYI 
#NYI             $self->xml_empty_tag( 'a:gd', @attributes );
#NYI         }
#NYI         $self->xml_end_tag( 'a:avLst' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'a:avLst' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_solid_fill()
#NYI #
#NYI # Write the <a:solidFill> element.
#NYI #
#NYI sub _write_a_solid_fill {
#NYI 
#NYI     my $self = shift;
#NYI     my $rgb  = shift;
#NYI 
#NYI     $rgb = '000000' unless defined $rgb;
#NYI 
#NYI     my @attributes = ( 'val' => $rgb );
#NYI 
#NYI     $self->xml_start_tag( 'a:solidFill' );
#NYI 
#NYI     $self->xml_empty_tag( 'a:srgbClr', @attributes );
#NYI 
#NYI     $self->xml_end_tag( 'a:solidFill' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_ln()
#NYI #
#NYI # Write the <a:ln> element.
#NYI #
#NYI sub _write_a_ln {
#NYI 
#NYI     my $self = shift;
#NYI     my $shape = shift || {};
#NYI 
#NYI     my $weight = $shape->{_line_weight};
#NYI 
#NYI     my @attributes = ( 'w' => $weight * 9525 );
#NYI 
#NYI     $self->xml_start_tag( 'a:ln', @attributes );
#NYI 
#NYI     my $line = $shape->{_line};
#NYI 
#NYI     if ( length $line > 1 ) {
#NYI 
#NYI         # Write the a:solidFill element.
#NYI         $self->_write_a_solid_fill( $line );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'a:noFill' );
#NYI     }
#NYI 
#NYI     if ( $shape->{_line_type} ) {
#NYI 
#NYI         @attributes = ( 'val' => $shape->{_line_type} );
#NYI         $self->xml_empty_tag( 'a:prstDash', @attributes );
#NYI     }
#NYI 
#NYI     if ( $shape->{_connect} ) {
#NYI         $self->xml_empty_tag( 'a:round' );
#NYI     }
#NYI     else {
#NYI         @attributes = ( 'lim' => 800000 );
#NYI         $self->xml_empty_tag( 'a:miter', @attributes );
#NYI     }
#NYI 
#NYI     $self->xml_empty_tag( 'a:headEnd' );
#NYI     $self->xml_empty_tag( 'a:tailEnd' );
#NYI 
#NYI     $self->xml_end_tag( 'a:ln' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_txBody
#NYI #
#NYI # Write the <xdr:txBody> element.
#NYI #
#NYI sub _write_txBody {
#NYI 
#NYI     my $self         = shift;
#NYI     my $col_absolute = shift;
#NYI     my $row_absolute = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $shape        = shift;
#NYI 
#NYI     my @attributes = (
#NYI         vertOverflow => "clip",
#NYI         wrap         => "square",
#NYI         lIns         => "27432",
#NYI         tIns         => "22860",
#NYI         rIns         => "27432",
#NYI         bIns         => "22860",
#NYI         anchor       => $shape->{_valign},
#NYI         upright      => "1",
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'xdr:txBody' );
#NYI     $self->xml_empty_tag( 'a:bodyPr', @attributes );
#NYI     $self->xml_empty_tag( 'a:lstStyle' );
#NYI 
#NYI     $self->xml_start_tag( 'a:p' );
#NYI 
#NYI     my $rotation = $shape->{_format}->{_rotation};
#NYI     $rotation = 0 unless defined $rotation;
#NYI     $rotation *= 60000;
#NYI 
#NYI     @attributes = ( algn => $shape->{_align}, rtl => $rotation );
#NYI     $self->xml_start_tag( 'a:pPr', @attributes );
#NYI 
#NYI     @attributes = ( sz => "1000" );
#NYI     $self->xml_empty_tag( 'a:defRPr', @attributes );
#NYI 
#NYI     $self->xml_end_tag( 'a:pPr' );
#NYI     $self->xml_start_tag( 'a:r' );
#NYI 
#NYI     my $size = $shape->{_format}->{_size};
#NYI     $size = 8 unless defined $size;
#NYI     $size *= 100;
#NYI 
#NYI     my $bold = $shape->{_format}->{_bold};
#NYI     $bold = 0 unless defined $bold;
#NYI 
#NYI     my $italic = $shape->{_format}->{_italic};
#NYI     $italic = 0 unless defined $italic;
#NYI 
#NYI     my $underline = $shape->{_format}->{_underline};
#NYI     $underline = $underline ? 'sng' : 'none';
#NYI 
#NYI     my $strike = $shape->{_format}->{_font_strikeout};
#NYI     $strike = $strike ? 'Strike' : 'noStrike';
#NYI 
#NYI     @attributes = (
#NYI         lang     => "en-US",
#NYI         sz       => $size,
#NYI         b        => $bold,
#NYI         i        => $italic,
#NYI         u        => $underline,
#NYI         strike   => $strike,
#NYI         baseline => 0,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'a:rPr', @attributes );
#NYI 
#NYI     my $color = $shape->{_format}->{_color};
#NYI     if ( defined $color ) {
#NYI         $color = $shape->_get_palette_color( $color );
#NYI         $color =~ s/^FF//;    # Remove leading FF from rgb for shape color.
#NYI     }
#NYI     else {
#NYI         $color = '000000';
#NYI     }
#NYI 
#NYI     $self->_write_a_solid_fill( $color );
#NYI 
#NYI     my $font = $shape->{_format}->{_font};
#NYI     $font = 'Calibri' unless defined $font;
#NYI     @attributes = ( typeface => $font );
#NYI     $self->xml_empty_tag( 'a:latin', @attributes );
#NYI 
#NYI     $self->xml_empty_tag( 'a:cs', @attributes );
#NYI 
#NYI     $self->xml_end_tag( 'a:rPr' );
#NYI 
#NYI     $self->xml_data_element( 'a:t', $shape->{_text} );
#NYI 
#NYI     $self->xml_end_tag( 'a:r' );
#NYI     $self->xml_end_tag( 'a:p' );
#NYI     $self->xml_end_tag( 'xdr:txBody' );
#NYI 
#NYI }
#NYI 
#NYI 
#NYI 1;
#NYI __END__
#NYI 
#NYI =pod
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Drawing - A class for writing the Excel XLSX drawing.xml file.
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
