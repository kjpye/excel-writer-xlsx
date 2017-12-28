unit class Excel::Writer::XLSX::Package::Styles;

###############################################################################
#
# Styles - A class for writing the Excel XLSX styles file.
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

has @!xf_formats;
has @!palette          = [];
has $!font_count       = 0;
has $!num_format_count = 0;
has $!border_count     = 0;
has $!fill_count       = 0;
has @!custom_colors    = [];
has @!dxf_formats      = [];
has $!has_hyperlink    = 0;
has @!xfformats;

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

#NYI     $self->{_xf_formats}       = undef;
#NYI     $self->{_palette}          = [];
#NYI     $self->{_font_count}       = 0;
#NYI     $self->{_num_format_count} = 0;
#NYI     $self->{_border_count}     = 0;
#NYI     $self->{_fill_count}       = 0;
#NYI     $self->{_custom_colors}    = [];
#NYI     $self->{_dxf_formats}      = [];
#NYI     $self->{_has_hyperlink}    = 0;

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

    # Add the style sheet.
    self.write_style_sheet();

    # Write the number formats.
    self.write_num_fmts();

    # Write the fonts.
    self.write_fonts();

    # Write the fills.
    self.write_fills();

    # Write the borders element.
    self.write_borders();

    # Write the cellStyleXfs element.
    self.write_cell_style_xfs();

    # Write the cellXfs element.
    self.write_cell_xfs();

    # Write the cellStyles element.
    self.write_cell_styles();

    # Write the dxfs element.
    self.write_dxfs();

    # Write the tableStyles element.
    self.write_table_styles();

    # Write the colors element.
    self.write_colors();

    # Close the style sheet tag.
    self.xml_end_tag( 'styleSheet' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# _set_style_properties()
#
# Pass in the Format objects and other properties used to set the styles.
#
method set_style_properties(*@args) {
    @!xf_formats       = @args.shift;
    @!palette          = @args.shift;
    $!font_count       = @args.shift;
    $!num_format_count = @args.shift;
    $!border_count     = @args.shift;
    $!fill_count       = @args.shift;
    @!custom_colors    = @args.shift;
    @!dxf_formats      = @args.shift;
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _get_palette_color()
#
# Convert from an Excel internal colour index to a XML style #RRGGBB index
# based on the default or user defined values in the Workbook palette.
#
method get_palette_color($index) {
    my $palette = @!palette;

    # Handle colours in #XXXXXX RGB format.
    if $index ~~ m:i/^\#(<[0..9 A..F]> ** 6)$/ {
        return "FF" . uc( $0 );
    }

    # Adjust the colour index.
    $index -= 8;

    # Palette is passed in from the Workbook class.
    my @rgb = @!palette[$index];

    return sprintf "FF%02X%02X%02X", @rgb[0, 1, 2];
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


##############################################################################
#
# _write_style_sheet()
#
# Write the <styleSheet> element.
#
method write_style_sheet {
    my $xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    my @attributes = ( 'xmlns' => $xmlns );

    self.xml_start_tag( 'styleSheet', @attributes );
}


##############################################################################
#
# _write_num_fmts()
#
# Write the <numFmts> element.
#
method write_num_fmts {
    my $count = $!num_format_count;

    return unless $count;

    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'numFmts', @attributes );

    # Write the numFmts elements.
    for @!xfformats -> $format {

        # Ignore built-in number formats, i.e., < 164.
        next unless $format<num_format_index> >= 164;
        self.write_num_fmt( $format<num_format_index>,
            $format<num_format> );
    }

    self.xml_end_tag( 'numFmts' );
}


##############################################################################
#
# _write_num_fmt()
#
# Write the <numFmt> element.
#
method write_num_fmt($num-fmt-id, $format-code) {
    my %format-codes = (
        0  => 'General',
        1  => '0',
        2  => '0.00',
        3  => '#,##0',
        4  => '#,##0.00',
        5  => '($#,##0_);($#,##0)',
        6  => '($#,##0_);[Red]($#,##0)',
        7  => '($#,##0.00_);($#,##0.00)',
        8  => '($#,##0.00_);[Red]($#,##0.00)',
        9  => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'm/d/yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',
        37 => '(#,##0_);(#,##0)',
        38 => '(#,##0_);[Red](#,##0)',
        39 => '(#,##0.00_);(#,##0.00)',
        40 => '(#,##0.00_);[Red](#,##0.00)',
        41 => '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
        42 => '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
        43 => '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
        44 => '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mm:ss.0',
        48 => '##0.0E+0',
        49 => '@',
    );

    # Set the format code for built-in number formats.
    if ( $num-fmt-id < 164 ) {
        if %format-codes{$num-fmt-id}.exists {
            $format-code = %format-codes{$num-fmt-id};
        }
        else {
            $format-code = 'General';
        }
    }

    my @attributes = (
        'numFmtId'   => $num-fmt-id,
        'formatCode' => $format-code,
    );

    self.xml_empty_tag( 'numFmt', @attributes );
}


##############################################################################
#
# _write_fonts()
#
# Write the <fonts> element.
#
method write_fonts {
    my $count = $!font_count;

    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'fonts', @attributes );

    # Write the font elements for format objects that have them.
    for @!xf_formats -> $format {
        self.write_font( $format ) if $format<has_font>;
    }

    self.xml_end_tag( 'fonts' );
}


##############################################################################
#
# _write_font()
#
# Write the <font> element.
#
method write_font($format, $dxf-format) {
    self.xml_start_tag( 'font' );

    # The condense and extend elements are mainly used in dxf formats.
    self.write_condense() if $format<font_condense>;
    self.write_extend()   if $format<font_extend>;

    self.xml_empty_tag( 'b' )       if $format<bold>;
    self.xml_empty_tag( 'i' )       if $format<italic>;
    self.xml_empty_tag( 'strike' )  if $format<font_strikeout>;
    self.xml_empty_tag( 'outline' ) if $format<font_outline>;
    self.xml_empty_tag( 'shadow' )  if $format<font_shadow>;

    # Handle the underline variants.
    self.write_underline( $format<underline> ) if $format<underline>;

    self.write_vert_align( 'superscript' ) if $format<font_script> == 1;
    self.write_vert_align( 'subscript' )   if $format<font_script> == 2;

    if !$dxf-format {
        self.xml_empty_tag( 'sz', 'val', $format<size> );
    }

    my $theme = $format<theme>;


    if $theme == -1 {
        # Ignore for excel2003_style.
    }
    elsif $theme {
        self.write_color( 'theme' => $theme );
    }
    elsif my $index = $format<color_indexed> {
        self.write_color( 'indexed' => $index );
    }
    elsif my $color = $format<color> {
        $color = self.get_palette_color( $color );

        self.write_color( 'rgb' => $color );
    }
    elsif !$dxf-format {
        self.write_color( 'theme' => 1 );
    }

    if !$dxf-format {
        self.xml_empty_tag( 'name',   'val', $format<font> );

        if $format<font_family> {
            self.xml_empty_tag( 'family', 'val', $format<font_family> );
        }

        if $format<font_charset> {
            self.xml_empty_tag( 'charset', 'val', $format<font_charset> );
        }

        if $format<font> eq 'Calibri' && !$format<hyperlink> {
            self.xml_empty_tag(
                'scheme',
                'val' => $format<font_scheme>
            );
        }

        if $format<hyperlink> {
            $!has_hyperlink = 1;
        }
    }

    self.xml_end_tag( 'font' );
}


###############################################################################
#
# _write_underline()
#
# Write the underline font element.
#
method write_underline($underline) {
    my @attributes;

    # Handle the underline variants.
    if $underline == 2 {
        @attributes = ( val => 'double' );
    }
    elsif $underline == 33 {
        @attributes = ( val => 'singleAccounting' );
    }
    elsif $underline == 34 {
        @attributes = ( val => 'doubleAccounting' );
    }
    else {
        @attributes = ();    # Default to single underline.
    }

    self.xml_empty_tag( 'u', @attributes );

}


##############################################################################
#
# _write_vert_align()
#
# Write the <vertAlign> font sub-element.
#
method write_vert_align($val) {
    my @attributes = ( 'val' => $val );

    self.xml_empty_tag( 'vertAlign', @attributes );
}


##############################################################################
#
# _write_color()
#
# Write the <color> element.
#
method write_color($name, $value) {
    my @attributes = ( $name => $value );

    self.xml_empty_tag( 'color', @attributes );
}


##############################################################################
#
# _write_fills()
#
# Write the <fills> element.
#
method write_fills {
    my $count = $!fill_count;

    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'fills', @attributes );

    # Write the default fill element.
    self.write_default_fill( 'none' );
    self.write_default_fill( 'gray125' );

    # Write the fill elements for format objects that have them.
    for @!xf_formats -> $format {
        self.write_fill( $format ) if $format<has_fill>;
    }

    self.xml_end_tag( 'fills' );
}


##############################################################################
#
# _write_default_fill()
#
# Write the <fill> element for the default fills.
#
method write_default_fill($pattern-type) {
    self.xml_start_tag( 'fill' );

    self.xml_empty_tag( 'patternFill', 'patternType', $pattern-type );

    self.xml_end_tag( 'fill' );
}


##############################################################################
#
# _write_fill()
#
# Write the <fill> element.
#
method write_fill($format, $dxf-format) {
    my $pattern    = $format<pattern>;
    my $bg-color   = $format<bg_color>;
    my $fg-color   = $format<fg_color>;

    # Colors for dxf formats are handled differently from normal formats since
    # the normal format reverses the meaning of BG and FG for solid fills.
    if $dxf-format {
        $bg-color = $format<dxf_bg_color>;
        $fg-color = $format<dxf_fg_color>;
    }


    my @patterns = <
      none
      solid
      mediumGray
      darkGray
      lightGray
      darkHorizontal
      darkVertical
      darkDown
      darkUp
      darkGrid
      darkTrellis
      lightHorizontal
      lightVertical
      lightDown
      lightUp
      lightGrid
      lightTrellis
      gray125
      gray0625

    >;


    self.xml_start_tag( 'fill' );

    # The "none" pattern is handled differently for dxf formats.
    if $dxf-format && $format<pattern> <= 1 {
        self.xml_start_tag( 'patternFill' );
    }
    else {
        self.xml_start_tag(
            'patternFill',
            'patternType',
            @patterns[ $format<pattern> ]

        );
    }

    if $fg-color {
        $fg-color = self.get_palette_color( $fg-color );
        self.xml_empty_tag( 'fgColor', 'rgb' => $fg-color );
    }

    if $bg-color {
        $bg-color = self.get_palette_color( $bg-color );
        self.xml_empty_tag( 'bgColor', 'rgb' => $bg-color );
    }
    else {
        if !$dxf-format {
            self.xml_empty_tag( 'bgColor', 'indexed' => 64 );
        }
    }

    self.xml_end_tag( 'patternFill' );
    self.xml_end_tag( 'fill' );
}


##############################################################################
#
# _write_borders()
#
# Write the <borders> element.
#
method write_borders {
    my $count = $!border_count;

    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'borders', @attributes );

    # Write the border elements for format objects that have them.
    for @!xf_formats -> $format {
        self.write_border( $format ) if $format<has_border>;
    }

    self.xml_end_tag( 'borders' );
}


##############################################################################
#
# _write_border()
#
# Write the <border> element.
#
method write_border($format, $dxf-format) {
    my @attributes = ();


    # Diagonal borders add attributes to the <border> element.
    if $format<diag_type> == 1 {
        push @attributes, ( diagonalUp => 1 );
    }
    elsif $format<diag_type> == 2 {
        push @attributes, ( diagonalDown => 1 );
    }
    elsif $format<diag_type> == 3 {
        push @attributes, ( diagonalUp   => 1 );
        push @attributes, ( diagonalDown => 1 );
    }

    # Ensure that a default diag border is set if the diag type is set.
    if $format<diag_type> && !$format<diag_border> {
        $format<diag_border> = 1;
    }

    # Write the start border tag.
    self.xml_start_tag( 'border', @attributes );

    # Write the <border> sub elements.
    self.write_sub_border(
        'left',
        $format<left>,
        $format<left_color>

    );

    self.write_sub_border(
        'right',
        $format<right>,
        $format<right_color>

    );

    self.write_sub_border(
        'top',
        $format<top>,
        $format<top_color>

    );

    self.write_sub_border(
        'bottom',
        $format<bottom>,
        $format<bottom_color>

    );

    # Condition DXF formats don't allow diagonal borders
    if !$dxf-format {
        self.write_sub_border(
            'diagonal',
            $format<diag_border>,
            $format<diag_color>

        );
    }

    if $dxf-format {
        self.write_sub_border( 'vertical' );
        self.write_sub_border( 'horizontal' );
    }

    self.xml_end_tag( 'border' );
}


##############################################################################
#
# _write_sub_border()
#
# Write the <border> sub elements such as <right>, <top>, etc.
#
method write_sub_border($type, $style, $color) {
    my @attributes;

    if !$style {
        self.xml_empty_tag( $type );
        return;
    }

    my @border-styles = <
      none
      thin
      medium
      dashed
      dotted
      thick
      double
      hair
      mediumDashed
      dashDot
      mediumDashDot
      dashDotDot
      mediumDashDotDot
      slantDashDot

    >;


    @attributes.push: ( style => @border-styles[$style] );

    self.xml_start_tag( $type, @attributes );

    if $color {
        $color = self.get_palette_color( $color );
        self.xml_empty_tag( 'color', 'rgb' => $color );
    }
    else {
        self.xml_empty_tag( 'color', 'auto' => 1 );
    }

    self.xml_end_tag( $type );
}


##############################################################################
#
# _write_cell_style_xfs()
#
# Write the <cellStyleXfs> element.
#
method write_cell_style_xfs {
    my $count = 1;

    if $!has_hyperlink {
        $count = 2;
    }

    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'cellStyleXfs', @attributes );

    # Write the style_xf element.
    self.write_style_xf(0, 0);

    if $!has_hyperlink {
        self.write_style_xf(1, 1);
    }

    self.xml_end_tag( 'cellStyleXfs' );
}


##############################################################################
#
# _write_cell_xfs()
#
# Write the <cellXfs> element.
#
method write_cell_xfs {
    my @formats = @!xf_formats;

    # Workaround for when the last format is used for the comment font
    # and shouldn't be used for cellXfs.
    my $last_format = @formats[*-1];

    if $last_format<font_only> {
        @formats.pop;
    }

    my $count = +@formats;
    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'cellXfs', @attributes );

    # Write the xf elements.
    for @formats -> $format {
        self.write_xf( $format );
    }

    self.xml_end_tag( 'cellXfs' );
}


##############################################################################
#
# _write_style_xf()
#
# Write the style <xf> element.
#
method write_style_xf($is-hyperlink, $font-id) {
    my $num-fmt-id   = 0;
    my $fill-id      = 0;
    my $border-id    = 0;

    my @attributes = (
        'numFmtId' => $num-fmt-id,
        'fontId'   => $font-id,
        'fillId'   => $fill-id,
        'borderId' => $border-id,
    );

    if $is-hyperlink {
        @attributes.push: ( 'applyNumberFormat' => 0 );
        @attributes.push: ( 'applyFill'         => 0 );
        @attributes.push: ( 'applyBorder'       => 0 );
        @attributes.push: ( 'applyAlignment'    => 0 );
        @attributes.push: ( 'applyProtection'   => 0 );

        self.xml_start_tag( 'xf', @attributes );
        self.xml_empty_tag( 'alignment',  ( 'vertical', 'top' ) );
        self.xml_empty_tag( 'protection', ( 'locked',   0 ) );
        self.xml_end_tag( 'xf' );
    }
    else {
        self.xml_empty_tag( 'xf', @attributes );
    }
}


##############################################################################
#
# _write_xf()
#
# Write the <xf> element.
#
method write_xf($format) {
    my $num_fmt_id  = $format<num_format_index>;
    my $font_id     = $format<font_index>;
    my $fill_id     = $format<fill_index>;
    my $border_id   = $format<border_index>;
    my $xf_id       = $format<xf_id>;
    my $has_align   = 0;
    my $has_protect = 0;

    my @attributes = (
        'numFmtId' => $num_fmt_id,
        'fontId'   => $font_id,
        'fillId'   => $fill_id,
        'borderId' => $border_id,
        'xfId'     => $xf_id,
    );


    if $format<num_format_index> > 0 {
        @attributes.push: ( 'applyNumberFormat' => 1 );
    }

    # Add applyFont attribute if XF format uses a font element.
    if $format<font_index> > 0 && !$format<hyperlink> {
        @attributes.push: ( 'applyFont' => 1 );
    }

    # Add applyFill attribute if XF format uses a fill element.
    if $format<fill_index> > 0 {
        @attributes.push: ( 'applyFill' => 1 );
    }

    # Add applyBorder attribute if XF format uses a border element.
    if $format<border_index> > 0 {
        @attributes.push: ( 'applyBorder' => 1 );
    }

    # Check if XF format has alignment properties set.
    my ( $apply_align, @align ) = $format.get_align_properties();

    # Check if an alignment sub-element should be written.
    $has_align = 1 if $apply_align && @align;

    # We can also have applyAlignment without a sub-element.
    if $apply_align || $format.hyperlink {
        @attributes.push: ( 'applyAlignment' => 1 );
    }

    # Check for cell protection properties.
    my @protection = $format.get_protection_properties();

    if @protection || $format.hyperlink {
        @attributes.push: ( 'applyProtection' => 1 );

        if !$format.hyperlink {
            $has_protect = 1;
        }
    }

    # Write XF with sub-elements if required.
    if $has_align || $has_protect {
        self.xml_start_tag( 'xf', @attributes );
        self.xml_empty_tag( 'alignment',  @align )      if $has_align;
        self.xml_empty_tag( 'protection', @protection ) if $has_protect;
        self.xml_end_tag( 'xf' );
    }
    else {
        self.xml_empty_tag( 'xf', @attributes );
    }
}


##############################################################################
#
# _write_cell_styles()
#
# Write the <cellStyles> element.
#
method write_cell_styles {
    my $count = 1;

    if $!has_hyperlink {
        $count = 2;
    }

    my @attributes = ( 'count' => $count );

    self.xml_start_tag( 'cellStyles', @attributes );

    # Write the cellStyle element.
    if $!has_hyperlink {
        self.write_cell_style('Hyperlink', 1, 8);
    }

    self.write_cell_style('Normal', 0, 0);

    self.xml_end_tag( 'cellStyles' );
}


##############################################################################
#
# _write_cell_style()
#
# Write the <cellStyle> element.
#
method write_cell_style($name, $xf-id, $builtin-id) {
    my @attributes = (
        'name'      => $name,
        'xfId'      => $xf-id,
        'builtinId' => $builtin-id,
    );

    self.xml_empty_tag( 'cellStyle', @attributes );
}


##############################################################################
#
# _write_dxfs()
#
# Write the <dxfs> element.
#
method write_dxfs {
    my $formats = @!dxf_formats;

    my $count = +$formats;

    my @attributes = ( 'count' => $count );

    if $count {
        self.xml_start_tag( 'dxfs', @attributes );

        # Write the font elements for format objects that have them.
        for @!dxf_formats -> $format {
            self.xml_start_tag( 'dxf' );
            self.write_font( $format, 1 ) if $format.has_dxf_font;

            if $format.num_format_index {
                self.write_num_fmt( $format.num_format_index,
                    $format.num_format );
            }

            self.write_fill( $format, 1 ) if $format.has_dxf_fill;
            self.write_border( $format, 1 ) if $format.has_dxf_border;
            self.xml_end_tag( 'dxf' );
        }

        self.xml_end_tag( 'dxfs' );
    }
    else {
        self.xml_empty_tag( 'dxfs', @attributes );
    }

}


##############################################################################
#
# _write_table_styles()
#
# Write the <tableStyles> element.
#
method write_table_styles {
    my $count               = 0;
    my $default_table_style = 'TableStyleMedium9';
    my $default_pivot_style = 'PivotStyleLight16';

    my @attributes = (
        'count'             => $count,
        'defaultTableStyle' => $default_table_style,
        'defaultPivotStyle' => $default_pivot_style,
    );

    self.xml_empty_tag( 'tableStyles', @attributes );
}


##############################################################################
#
# _write_colors()
#
# Write the <colors> element.
#
method write_colors {
    my @custom_colors = @!custom_colors;

    return unless @custom_colors;

    self.xml_start_tag( 'colors' );
    self.write_mru_colors( @custom_colors );
    self.xml_end_tag( 'colors' );
}


##############################################################################
#
# _write_mru_colors()
#
# Write the <mruColors> element for the most recently used colours.
#
method write_mru_colors(@custom-colors) {
    # Limit the mruColors to the last 10.
    my $count = +@custom-colors;
    if $count > 10 {
        @custom-colors.splice: 0, ( $count - 10 );
    }

    self.xml_start_tag( 'mruColors' );

    # Write the custom colors in reverse order.
    for @custom-colors.reverse -> $color {
        self.write_color( 'rgb' => $color );
    }

    self.xml_end_tag( 'mruColors' );
}


##############################################################################
#
# _write_condense()
#
# Write the <condense> element.
#
method write_condense {
    my $val  = 0;

    my @attributes = ( 'val' => $val );

    self.xml_empty_tag( 'condense', @attributes );
}


##############################################################################
#
# _write_extend()
#
# Write the <extend> element.
#
method write_extend {
    my $val  = 0;

    my @attributes = ( 'val' => $val );

    self.xml_empty_tag( 'extend', @attributes );
}


=begin pod


__END__

=pod

=head1 NAME

Styles - A class for writing the Excel XLSX styles file.

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
