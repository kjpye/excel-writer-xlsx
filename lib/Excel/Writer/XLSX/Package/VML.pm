unit class Excel::Writer::XLSX::Package::VML;

###############################################################################
#
# VML - A class for writing the Excel XLSX VML files.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use v6.c;
#use Excel::Writer::XLSX::Package::XMLwriter;

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';


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

#NYI     bless $self, $class;

#NYI     return $self;
#NYI }


###############################################################################
#
# _assemble_xml_file()
#
# Assemble and write the XML file.
#
method assemble_xml_file($data-id, $vml-shape-id, $comments-data, $buttons-data, $header-images-data, $z-index = 1) {
    self.write_xml_namespace;

    # Write the o:shapelayout element.
    self.write_shapelayout( $data-id );

    if $buttons-data.defined && $buttons-data {

        # Write the v:shapetype element.
        self.write_button_shapetype();

        for $buttons-data -> $button {

            # Write the v:shape element.
            self.write_button_shape( ++$vml-shape-id, $z-index++, $button );
        }
    }

    if $comments-data.defined && $comments-data {

        # Write the v:shapetype element.
        self.write_comment_shapetype();

        for $comments-data -> $comment {

            # Write the v:shape element.
            self.write_comment_shape( ++$vml-shape-id, $z-index++,
                $comment );
        }
    }

    if $header-images-data.defined && $header-images-data {

        # Write the v:shapetype element.
        self.write_image_shapetype();

        my $index = 1;
        for $header-images-data -> $image {

            # Write the v:shape element.
            self.write_image_shape( ++$vml-shape-id, $index++, $image );
        }
    }


    self.xml_end_tag( 'xml' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _pixels_to_points()
#
# Convert comment vertices from pixels to points.
#
method pixels_to_points($vertices) {
    my (
        $col-start, $row-start, $x1,    $y1,
        $col-end,   $row-end,   $x2,    $y2,
        $left,      $top,       $width, $height
    ) = $vertices;

#TODO
#    for my $pixels ( $left, $top, $width, $height ) {
#        $pixels *= 0.75;
#    }

    return ( $left, $top, $width, $height );
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_xml_namespace()
#
# Write the <xml> element. This is the root element of VML.
#
method write_xml_namespace {
    my $schema  = 'urn:schemas-microsoft-com:';
    my $xmlns   = $schema ~ 'vml';
    my $xmlns_o = $schema ~ 'office:office';
    my $xmlns_x = $schema ~ 'office:excel';

    my @attributes = (
        'xmlns:v' => $xmlns,
        'xmlns:o' => $xmlns_o,
        'xmlns:x' => $xmlns_x,
    );

    self.xml_start_tag( 'xml', @attributes );
}


##############################################################################
#
# _write_shapelayout()
#
# Write the <o:shapelayout> element.
#
method write_shapelayout($data-id) {
    my $ext     = 'edit';

    my @attributes = ( 'v:ext' => $ext );

    self.xml_start_tag( 'o:shapelayout', @attributes );

    # Write the o:idmap element.
    self.write_idmap( $data-id );

    self.xml_end_tag( 'o:shapelayout' );
}


##############################################################################
#
# _write_idmap()
#
# Write the <o:idmap> element.
#
method write_idmap($data-id) {
    my $ext     = 'edit';

    my @attributes = (
        'v:ext' => $ext,
        'data'  => $data-id,
    );

    self.xml_empty_tag( 'o:idmap', @attributes );
}


##############################################################################
#
# _write_comment_shapetype()
#
# Write the <v:shapetype> element.
#
method write_comment_shapetype {
    my $id        = '_x0000_t202';
    my $coordsize = '21600,21600';
    my $spt       = 202;
    my $path      = 'm,l,21600r21600,l21600,xe';

    my @attributes = (
        'id'        => $id,
        'coordsize' => $coordsize,
        'o:spt'     => $spt,
        'path'      => $path,
    );

    self.xml_start_tag( 'v:shapetype', @attributes );

    # Write the v:stroke element.
    self.write_stroke();

    # Write the v:path element.
    self.write_comment_path( 't', 'rect' );

    self.xml_end_tag( 'v:shapetype' );
}


##############################################################################
#
# _write_button_shapetype()
#
# Write the <v:shapetype> element.
#
method write_button_shapetype {
    my $id        = '_x0000_t201';
    my $coordsize = '21600,21600';
    my $spt       = 201;
    my $path      = 'm,l,21600r21600,l21600,xe';

    my @attributes = (
        'id'        => $id,
        'coordsize' => $coordsize,
        'o:spt'     => $spt,
        'path'      => $path,
    );

    self.xml_start_tag( 'v:shapetype', @attributes );

    # Write the v:stroke element.
    self.write_stroke();

    # Write the v:path element.
    self.write_button_path( 't', 'rect' );

    # Write the o:lock element.
    self.write_shapetype_lock();

    self.xml_end_tag( 'v:shapetype' );
}


##############################################################################
#
# _write_image_shapetype()
#
# Write the <v:shapetype> element.
#
method write_image_shapetype {
    my $id               = '_x0000_t75';
    my $coordsize        = '21600,21600';
    my $spt              = 75;
    my $o_preferrelative = 't';
    my $path             = 'm@4@5l@4@11@9@11@9@5xe';
    my $filled           = 'f';
    my $stroked          = 'f';

    my @attributes = (
        'id'               => $id,
        'coordsize'        => $coordsize,
        'o:spt'            => $spt,
        'o:preferrelative' => $o_preferrelative,
        'path'             => $path,
        'filled'           => $filled,
        'stroked'          => $stroked,
    );

    self.xml_start_tag( 'v:shapetype', @attributes );

    # Write the v:stroke element.
    self.write_stroke();

    # Write the v:formulas element.
    self.write_formulas();

    # Write the v:path element.
    self.write_image_path();

    # Write the o:lock element.
    self.write_aspect_ratio_lock();

    self.xml_end_tag( 'v:shapetype' );
}


##############################################################################
#
# _write_stroke()
#
# Write the <v:stroke> element.
#
method write_stroke {
    my $joinstyle = 'miter';

    my @attributes = ( 'joinstyle' => $joinstyle );

    self.xml_empty_tag( 'v:stroke', @attributes );
}


##############################################################################
#
# _write_comment_path()
#
# Write the <v:path> element.
#
method write_comment_path($gradientshapeok, $connecttype) {
    my @attributes      = ();

    @attributes.push: ( 'gradientshapeok' => 't' ) if $gradientshapeok;
    @attributes.push: ( 'o:connecttype' => $connecttype );

    self.xml_empty_tag( 'v:path', @attributes );
}


##############################################################################
#
# _write_button_path()
#
# Write the <v:path> element.
#
method write_button_path {
    my $shadowok    = 'f';
    my $extrusionok = 'f';
    my $strokeok    = 'f';
    my $fillok      = 'f';
    my $connecttype = 'rect';

    my @attributes = (
        'shadowok'      => $shadowok,
        'o:extrusionok' => $extrusionok,
        'strokeok'      => $strokeok,
        'fillok'        => $fillok,
        'o:connecttype' => $connecttype,
    );

    self.xml_empty_tag( 'v:path', @attributes );
}


##############################################################################
#
# _write_image_path()
#
# Write the <v:path> element.
#
method write_image_path {
    my $extrusionok     = 'f';
    my $gradientshapeok = 't';
    my $connecttype     = 'rect';

    my @attributes = (
        'o:extrusionok'   => $extrusionok,
        'gradientshapeok' => $gradientshapeok,
        'o:connecttype'   => $connecttype,
    );

    self.xml_empty_tag( 'v:path', @attributes );
}


##############################################################################
#
# _write_shapetype_lock()
#
# Write the <o:lock> element.
#
method write_shapetype_lock {
    my $ext       = 'edit';
    my $shapetype = 't';

    my @attributes = (
        'v:ext'     => $ext,
        'shapetype' => $shapetype,
    );

    self.xml_empty_tag( 'o:lock', @attributes );
}


##############################################################################
#
# _write_rotation_lock()
#
# Write the <o:lock> element.
#
method write_rotation_lock {
    my $ext      = 'edit';
    my $rotation = 't';

    my @attributes = (
        'v:ext'    => $ext,
        'rotation' => $rotation,
    );

    self.xml_empty_tag( 'o:lock', @attributes );
}


##############################################################################
#
# _write_aspect_ratio_lock()
#
# Write the <o:lock> element.
#
method write_aspect_ratio_lock {
    my $ext         = 'edit';
    my $aspectratio = 't';

    my @attributes = (
        'v:ext'       => $ext,
        'aspectratio' => $aspectratio,
    );

    self.xml_empty_tag( 'o:lock', @attributes );
}

##############################################################################
#
# _write_comment_shape()
#
# Write the <v:shape> element.
#
method write_comment_shape($id, $z-index, $comment) {
    my $type       = '#_x0000_t202';
    my $insetmode  = 'auto';
    my $visibility = 'hidden';

    # Set the shape index.
    $id = '_x0000_s' ~ $id;

    # Get the comment parameters
    my $row       = $comment[0];
    my $col       = $comment[1];
    my $string    = $comment[2];
    my $author    = $comment[3];
    my $visible   = $comment[4];
    my $fillcolor = $comment[5];
    my $vertices  = $comment[6];

    my ( $left, $top, $width, $height ) = self.pixels_to_points( $vertices );

    # Set the visibility.
    $visibility = 'visible' if $visible;

    my $style =
        'position:absolute;'
      ~ 'margin-left:'
      ~ $left ~ 'pt;'
      ~ 'margin-top:'
      ~ $top ~ 'pt;'
      ~ 'width:'
      ~ $width ~ 'pt;'
      ~ 'height:'
      ~ $height ~ 'pt;'
      ~ 'z-index:'
      ~ $z-index ~ ';'
      ~ 'visibility:'
      ~ $visibility;


    my @attributes = (
        'id'          => $id,
        'type'        => $type,
        'style'       => $style,
        'fillcolor'   => $fillcolor,
        'o:insetmode' => $insetmode,
    );

    self.xml_start_tag( 'v:shape', @attributes );

    # Write the v:fill element.
    self.write_comment_fill();

    # Write the v:shadow element.
    self.write_shadow();

    # Write the v:path element.
    self.write_comment_path( Nil, 'none' );

    # Write the v:textbox element.
    self.write_comment_textbox();

    # Write the x:ClientData element.
    self.write_comment_client_data( $row, $col, $visible, $vertices );

    self.xml_end_tag( 'v:shape' );
}


##############################################################################
#
# _write_button_shape()
#
# Write the <v:shape> element.
#
method write_button_shape($id, $z-index, $button) {
    my $type       = '#_x0000_t201';

    # Set the shape index.
    $id = '_x0000_s' ~ $id;

    # Get the button parameters
    my $row       = $button<row>;
    my $col       = $button<col>;
    my $vertices  = $button<vertices>;

    my ( $left, $top, $width, $height ) = self.pixels_to_points( $vertices );

    my $style =
        'position:absolute;'
      ~ 'margin-left:'
      ~ $left ~ 'pt;'
      ~ 'margin-top:'
      ~ $top ~ 'pt;'
      ~ 'width:'
      ~ $width ~ 'pt;'
      ~ 'height:'
      ~ $height ~ 'pt;'
      ~ 'z-index:'
      ~ $z-index ~ ';'
      ~ 'mso-wrap-style:tight';


    my @attributes = (
        'id'          => $id,
        'type'        => $type,
        'style'       => $style,
        'o:button'    => 't',
        'fillcolor'   => 'buttonFace [67]',
        'strokecolor' => 'windowText [64]',
        'o:insetmode' => 'auto',
    );

    self.xml_start_tag( 'v:shape', @attributes );

    # Write the v:fill element.
    self.write_button_fill();

    # Write the o:lock element.
    self.write_rotation_lock();

    # Write the v:textbox element.
    self.write_button_textbox( $button<font> );

    # Write the x:ClientData element.
    self.write_button_client_data( $button );

    self.xml_end_tag( 'v:shape' );
}


##############################################################################
#
# _write_image_shape()
#
# Write the <v:shape> element.
#
method write_image_shape($id, $index, $image-data) {
    my $type       = '#_x0000_t75';

    # Set the shape index.
    $id = '_x0000_s' ~ $id;

    # Get the image parameters
    my $width    = $image-data[0];
    my $height   = $image-data[1];
    my $name     = $image-data[2];
    my $position = $image-data[3];
    my $x_dpi    = $image-data[4];
    my $y_dpi    = $image-data[5];

    # Scale the height/width by the resolution, relative to 72dpi.
    $width  = $width  * 72 / $x_dpi;
    $height = $height * 72 / $y_dpi;

    # Excel uses a rounding based around 72 and 96 dpi.
    $width  = 72/96 * ($width  * 96/72 + 0.25).int;
    $height = 72/96 * ($height * 96/72 + 0.25).int;

    my $style =
        'position:absolute;'
      ~ 'margin-left:0;'
      ~ 'margin-top:0;'
      ~ 'width:'
      ~ $width ~ 'pt;'
      ~ 'height:'
      ~ $height ~ 'pt;'
      ~ 'z-index:'
      ~ $index;

    my @attributes = (
        'id'     => $position,
        'o:spid' => $id,
        'type'   => $type,
        'style'  => $style,
    );

    self.xml_start_tag( 'v:shape', @attributes );

    # Write the v:imagedata element.
    self.write_imagedata( $index, $name );

    # Write the o:lock element.
    self.write_rotation_lock();

    self.xml_end_tag( 'v:shape' );
}

##############################################################################
#
# _write_comment_fill()
#
# Write the <v:fill> element.
#
method write_comment_fill {
    my $color_2 = '#ffffe1';

    my @attributes = ( 'color2' => $color_2 );

    self.xml_empty_tag( 'v:fill', @attributes );
}


##############################################################################
#
# _write_button_fill()
#
# Write the <v:fill> element.
#
method write_button_fill {
    my $color_2          = 'buttonFace [67]';
    my $detectmouseclick = 't';

    my @attributes = (
        'color2'             => $color_2,
        'o:detectmouseclick' => $detectmouseclick,
    );

    self.xml_empty_tag( 'v:fill', @attributes );
}


##############################################################################
#
# _write_shadow()
#
# Write the <v:shadow> element.
#
method write_shadow {
    my $on       = 't';
    my $color    = 'black';
    my $obscured = 't';

    my @attributes = (
        'on'       => $on,
        'color'    => $color,
        'obscured' => $obscured,
    );

    self.xml_empty_tag( 'v:shadow', @attributes );
}


##############################################################################
#
# _write_comment_textbox()
#
# Write the <v:textbox> element.
#
method write_comment_textbox {
    my $style = 'mso-direction-alt:auto';

    my @attributes = ( 'style' => $style );

    self.xml_start_tag( 'v:textbox', @attributes );

    # Write the div element.
    self.write_div( 'left' );

    self.xml_end_tag( 'v:textbox' );
}


##############################################################################
#
# _write_button_textbox()
#
# Write the <v:textbox> element.
#
method write_button_textbox($font) {
    my $style = 'mso-direction-alt:auto';

    my @attributes = ( 'style' => $style, 'o:singleclick' => 'f' );

    self.xml_start_tag( 'v:textbox', @attributes );

    # Write the div element.
    self.write_div( 'center', $font );

    self.xml_end_tag( 'v:textbox' );
}


##############################################################################
#
# _write_div()
#
# Write the <div> element.
#
method write_div($align, $font) {
    my $style = 'text-align:' ~ $align;

    my @attributes = ( 'style' => $style );

    self.xml_start_tag( 'div', @attributes );

    if $font {

        # Write the font element.
        self.write_font( $font );
    }

    self.xml_end_tag( 'div' );
}

##############################################################################
#
# _write_font()
#
# Write the <font> element.
#
method write_font($font) {
    my $caption = $font<caption>;
    my $face    = 'Calibri';
    my $size    = 220;
    my $color   = '#000000';

    my @attributes = (
        'face'  => $face,
        'size'  => $size,
        'color' => $color,
    );

    self.xml_data_element( 'font', $caption, @attributes );
}


##############################################################################
#
# _write_comment_client_data()
#
# Write the <x:ClientData> element.
#
method write_comment_client_data($row, $col, $visible, $vertices) {
    my $object_type = 'Note';

    my @attributes = ( 'ObjectType' => $object_type );

    self.xml_start_tag( 'x:ClientData', @attributes );

    # Write the x:MoveWithCells element.
    self.write_move_with_cells();

    # Write the x:SizeWithCells element.
    self.write_size_with_cells();

    # Write the x:Anchor element.
    self.write_anchor( $vertices );

    # Write the x:AutoFill element.
    self.write_auto_fill();

    # Write the x:Row element.
    self.write_row( $row );

    # Write the x:Column element.
    self.write_column( $col );

    # Write the x:Visible element.
    self.write_visible() if $visible;

    self.xml_end_tag( 'x:ClientData' );
}


##############################################################################
#
# _write_button_client_data()
#
# Write the <x:ClientData> element.
#
method write_button_client_data($button) {
    my $row       = $button<row>;
    my $col       = $button<col>;
    my $macro     = $button<macro>;
    my $vertices  = $button<vertices>;


    my $object_type = 'Button';

    my @attributes = ( 'ObjectType' => $object_type );

    self.xml_start_tag( 'x:ClientData', @attributes );

    # Write the x:Anchor element.
    self.write_anchor( $vertices );

    # Write the x:PrintObject element.
    self.write_print_object();

    # Write the x:AutoFill element.
    self.write_auto_fill();

    # Write the x:FmlaMacro element.
    self.write_fmla_macro( $macro );

    # Write the x:TextHAlign element.
    self.write_text_halign();

    # Write the x:TextVAlign element.
    self.write_text_valign();

    self.xml_end_tag( 'x:ClientData' );
}


##############################################################################
#
# _write_move_with_cells()
#
# Write the <x:MoveWithCells> element.
#
method write_move_with_cells {
    self.xml_empty_tag( 'x:MoveWithCells' );
}


##############################################################################
#
# _write_size_with_cells()
#
# Write the <x:SizeWithCells> element.
#
method write_size_with_cells {
    self.xml_empty_tag( 'x:SizeWithCells' );
}


##############################################################################
#
# _write_visible()
#
# Write the <x:Visible> element.
#
method write_visible {
    self.xml_empty_tag( 'x:Visible' );
}


##############################################################################
#
# _write_anchor()
#
# Write the <x:Anchor> element.
#
method write_anchor($vertices) {
    my ( $col_start, $row_start, $x1, $y1, $col_end, $row_end, $x2, $y2 ) =
      $vertices;

    my $data = join ", ",
      ( $col_start, $x1, $row_start, $y1, $col_end, $x2, $row_end, $y2 );

    self.xml_data_element( 'x:Anchor', $data );
}


##############################################################################
#
# _write_auto_fill()
#
# Write the <x:AutoFill> element.
#
method write_auto_fill {
    my $data = 'False';

    self.xml_data_element( 'x:AutoFill', $data );
}


##############################################################################
#
# _write_row()
#
# Write the <x:Row> element.
#
method write_row($data) {
    self.xml_data_element( 'x:Row', $data );
}


##############################################################################
#
# _write_column()
#
# Write the <x:Column> element.
#
method write_column($data) {
    self.xml_data_element( 'x:Column', $data );
}


##############################################################################
#
# _write_print_object()
#
# Write the <x:PrintObject> element.
#
method write_print_object {
    my $data = 'False';

    self.xml_data_element( 'x:PrintObject', $data );
}


##############################################################################
#
# _write_text_halign()
#
# Write the <x:TextHAlign> element.
#
method write_text_halign {
    my $data = 'Center';

    self.xml_data_element( 'x:TextHAlign', $data );
}


##############################################################################
#
# _write_text_valign()
#
# Write the <x:TextVAlign> element.
#
method write_text_valign {
    my $data = 'Center';

    self.xml_data_element( 'x:TextVAlign', $data );
}


##############################################################################
#
# _write_fmla_macro()
#
# Write the <x:FmlaMacro> element.
#
method write_fmla_macro($data) {
    self.xml_data_element( 'x:FmlaMacro', $data );
}

##############################################################################
#
# _write_imagedata()
#
# Write the <v:imagedata> element.
#
method write_imagedata($index, $o-title) {
    my @attributes = (
        'o:relid' => 'rId' ~ $index,
        'o:title' => $o-title,
    );

    self.xml_empty_tag( 'v:imagedata', @attributes );
}



##############################################################################
#
# _write_formulas()
#
# Write the <v:formulas> element.
#
method write_formulas {
    self.xml_start_tag( 'v:formulas' );

    # Write the v:f elements.
    self.write_f('if lineDrawn pixelLineWidth 0');
    self.write_f('sum @0 1 0');
    self.write_f('sum 0 0 @1');
    self.write_f('prod @2 1 2');
    self.write_f('prod @3 21600 pixelWidth');
    self.write_f('prod @3 21600 pixelHeight');
    self.write_f('sum @0 0 1');
    self.write_f('prod @6 1 2');
    self.write_f('prod @7 21600 pixelWidth');
    self.write_f('sum @8 21600 0');
    self.write_f('prod @7 21600 pixelHeight');
    self.write_f('sum @10 21600 0');

    self.xml_end_tag( 'v:formulas' );
}


##############################################################################
#
# _write_f()
#
# Write the <v:f> element.
#
method write_f($eqn) {
    my @attributes = ( 'eqn' => $eqn );

    self.xml_empty_tag( 'v:f', @attributes );
}

=begin pod

__END__

=pod

=head1 NAME

VML - A class for writing the Excel XLSX VML files.

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
