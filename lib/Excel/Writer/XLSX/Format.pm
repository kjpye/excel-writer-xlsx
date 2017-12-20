#NYI package Excel::Writer::XLSX::Format;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Format - A class for defining Excel formatting.
#NYI #
#NYI #
#NYI # Used in conjunction with Excel::Writer::XLSX
#NYI #
#NYI # Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#NYI #
#NYI # Documentation after __END__
#NYI #
#NYI 
#NYI use 5.008002;
#NYI use Exporter;
#NYI use strict;
#NYI use warnings;
#NYI use Carp;
#NYI 
#NYI 
#NYI our @ISA     = qw(Exporter);
#NYI our $VERSION = '0.96';
#NYI our $AUTOLOAD;
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # new()
#NYI #
#NYI # Constructor
#NYI #
#NYI sub new {
#NYI 
#NYI     my $class = shift;
#NYI 
#NYI     my $self = {
#NYI         _xf_format_indices  => shift,
#NYI         _dxf_format_indices => shift,
#NYI         _xf_index           => undef,
#NYI         _dxf_index          => undef,
#NYI 
#NYI         _num_format       => 0,
#NYI         _num_format_index => 0,
#NYI         _font_index       => 0,
#NYI         _has_font         => 0,
#NYI         _has_dxf_font     => 0,
#NYI         _font             => 'Calibri',
#NYI         _size             => 11,
#NYI         _bold             => 0,
#NYI         _italic           => 0,
#NYI         _color            => 0x0,
#NYI         _underline        => 0,
#NYI         _font_strikeout   => 0,
#NYI         _font_outline     => 0,
#NYI         _font_shadow      => 0,
#NYI         _font_script      => 0,
#NYI         _font_family      => 2,
#NYI         _font_charset     => 0,
#NYI         _font_scheme      => 'minor',
#NYI         _font_condense    => 0,
#NYI         _font_extend      => 0,
#NYI         _theme            => 0,
#NYI         _hyperlink        => 0,
#NYI         _xf_id            => 0,
#NYI 
#NYI         _hidden => 0,
#NYI         _locked => 1,
#NYI 
#NYI         _text_h_align  => 0,
#NYI         _text_wrap     => 0,
#NYI         _text_v_align  => 0,
#NYI         _text_justlast => 0,
#NYI         _rotation      => 0,
#NYI 
#NYI         _fg_color     => 0x00,
#NYI         _bg_color     => 0x00,
#NYI         _pattern      => 0,
#NYI         _has_fill     => 0,
#NYI         _has_dxf_fill => 0,
#NYI         _fill_index   => 0,
#NYI         _fill_count   => 0,
#NYI 
#NYI         _border_index   => 0,
#NYI         _has_border     => 0,
#NYI         _has_dxf_border => 0,
#NYI         _border_count   => 0,
#NYI 
#NYI         _bottom       => 0,
#NYI         _bottom_color => 0x0,
#NYI         _diag_border  => 0,
#NYI         _diag_color   => 0x0,
#NYI         _diag_type    => 0,
#NYI         _left         => 0,
#NYI         _left_color   => 0x0,
#NYI         _right        => 0,
#NYI         _right_color  => 0x0,
#NYI         _top          => 0,
#NYI         _top_color    => 0x0,
#NYI 
#NYI         _indent        => 0,
#NYI         _shrink        => 0,
#NYI         _merge_range   => 0,
#NYI         _reading_order => 0,
#NYI         _just_distrib  => 0,
#NYI         _color_indexed => 0,
#NYI         _font_only     => 0,
#NYI 
#NYI     };
#NYI 
#NYI     bless $self, $class;
#NYI 
#NYI     # Set properties passed to Workbook::add_format()
#NYI     $self->set_format_properties(@_) if @_;
#NYI 
#NYI     return $self;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # copy($format)
#NYI #
#NYI # Copy the attributes of another Excel::Writer::XLSX::Format object.
#NYI #
#NYI sub copy {
#NYI     my $self  = shift;
#NYI     my $other = $_[0];
#NYI 
#NYI 
#NYI     return unless defined $other;
#NYI     return unless ( ref( $self ) eq ref( $other ) );
#NYI 
#NYI     # Store properties that we don't want over-ridden.
#NYI     my $xf_index           = $self->{_xf_index};
#NYI     my $dxf_index          = $self->{_dxf_index};
#NYI     my $xf_format_indices  = $self->{_xf_format_indices};
#NYI     my $dxf_format_indices = $self->{_dxf_format_indices};
#NYI     my $palette            = $self->{_palette};
#NYI 
#NYI     # Copy properties.
#NYI     %$self             = %$other;
#NYI 
#NYI     # Restore original properties.
#NYI     $self->{_xf_index}           = $xf_index;
#NYI     $self->{_dxf_index}          = $dxf_index;
#NYI     $self->{_xf_format_indices}  = $xf_format_indices;
#NYI     $self->{_dxf_format_indices} = $dxf_format_indices;
#NYI     $self->{_palette}            = $palette;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_align_properties()
#NYI #
#NYI # Return properties for an Style xf <alignment> sub-element.
#NYI #
#NYI sub get_align_properties {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @align;    # Attributes to return
#NYI 
#NYI     # Check if any alignment options in the format have been changed.
#NYI     my $changed =
#NYI       (      $self->{_text_h_align} != 0
#NYI           || $self->{_text_v_align} != 0
#NYI           || $self->{_indent} != 0
#NYI           || $self->{_rotation} != 0
#NYI           || $self->{_text_wrap} != 0
#NYI           || $self->{_shrink} != 0
#NYI           || $self->{_reading_order} != 0 ) ? 1 : 0;
#NYI 
#NYI     return unless $changed;
#NYI 
#NYI 
#NYI 
#NYI     # Indent is only allowed for horizontal left, right and distributed. If it
#NYI     # is defined for any other alignment or no alignment has been set then
#NYI     # default to left alignment.
#NYI     if (   $self->{_indent}
#NYI         && $self->{_text_h_align} != 1
#NYI         && $self->{_text_h_align} != 3
#NYI         && $self->{_text_h_align} != 7 )
#NYI     {
#NYI         $self->{_text_h_align} = 1;
#NYI     }
#NYI 
#NYI     # Check for properties that are mutually exclusive.
#NYI     $self->{_shrink}       = 0 if $self->{_text_wrap};
#NYI     $self->{_shrink}       = 0 if $self->{_text_h_align} == 4;    # Fill
#NYI     $self->{_shrink}       = 0 if $self->{_text_h_align} == 5;    # Justify
#NYI     $self->{_shrink}       = 0 if $self->{_text_h_align} == 7;    # Distributed
#NYI     $self->{_just_distrib} = 0 if $self->{_text_h_align} != 7;    # Distributed
#NYI     $self->{_just_distrib} = 0 if $self->{_indent};
#NYI 
#NYI     my $continuous = 'centerContinuous';
#NYI 
#NYI     push @align, 'horizontal', 'left'        if $self->{_text_h_align} == 1;
#NYI     push @align, 'horizontal', 'center'      if $self->{_text_h_align} == 2;
#NYI     push @align, 'horizontal', 'right'       if $self->{_text_h_align} == 3;
#NYI     push @align, 'horizontal', 'fill'        if $self->{_text_h_align} == 4;
#NYI     push @align, 'horizontal', 'justify'     if $self->{_text_h_align} == 5;
#NYI     push @align, 'horizontal', $continuous   if $self->{_text_h_align} == 6;
#NYI     push @align, 'horizontal', 'distributed' if $self->{_text_h_align} == 7;
#NYI 
#NYI     push @align, 'justifyLastLine', 1 if $self->{_just_distrib};
#NYI 
#NYI     # Property 'vertical' => 'bottom' is a default. It sets applyAlignment
#NYI     # without an alignment sub-element.
#NYI     push @align, 'vertical', 'top'         if $self->{_text_v_align} == 1;
#NYI     push @align, 'vertical', 'center'      if $self->{_text_v_align} == 2;
#NYI     push @align, 'vertical', 'justify'     if $self->{_text_v_align} == 4;
#NYI     push @align, 'vertical', 'distributed' if $self->{_text_v_align} == 5;
#NYI 
#NYI     push @align, 'indent',       $self->{_indent}   if $self->{_indent};
#NYI     push @align, 'textRotation', $self->{_rotation} if $self->{_rotation};
#NYI 
#NYI     push @align, 'wrapText',     1 if $self->{_text_wrap};
#NYI     push @align, 'shrinkToFit',  1 if $self->{_shrink};
#NYI 
#NYI     push @align, 'readingOrder', 1 if $self->{_reading_order} == 1;
#NYI     push @align, 'readingOrder', 2 if $self->{_reading_order} == 2;
#NYI 
#NYI     return $changed, @align;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_protection_properties()
#NYI #
#NYI # Return properties for an Excel XML <Protection> element.
#NYI #
#NYI sub get_protection_properties {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attribs;
#NYI 
#NYI     push @attribs, 'locked', 0 if !$self->{_locked};
#NYI     push @attribs, 'hidden', 1 if $self->{_hidden};
#NYI 
#NYI     return @attribs;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_format_key()
#NYI #
#NYI # Returns a unique hash key for the Format object.
#NYI #
#NYI sub get_format_key {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $key = join ':',
#NYI       (
#NYI         $self->get_font_key(), $self->get_border_key,
#NYI         $self->get_fill_key(), $self->get_alignment_key(),
#NYI         $self->{_num_format},  $self->{_locked},
#NYI         $self->{_hidden}
#NYI       );
#NYI 
#NYI     return $key;
#NYI }
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_font_key()
#NYI #
#NYI # Returns a unique hash key for a font. Used by Workbook.
#NYI #
#NYI sub get_font_key {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $key = join ':', (
#NYI         $self->{_bold},
#NYI         $self->{_color},
#NYI         $self->{_font_charset},
#NYI         $self->{_font_family},
#NYI         $self->{_font_outline},
#NYI         $self->{_font_script},
#NYI         $self->{_font_shadow},
#NYI         $self->{_font_strikeout},
#NYI         $self->{_font},
#NYI         $self->{_italic},
#NYI         $self->{_size},
#NYI         $self->{_underline},
#NYI         $self->{_theme},
#NYI 
#NYI     );
#NYI 
#NYI     return $key;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_border_key()
#NYI #
#NYI # Returns a unique hash key for a border style. Used by Workbook.
#NYI #
#NYI sub get_border_key {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $key = join ':', (
#NYI         $self->{_bottom},
#NYI         $self->{_bottom_color},
#NYI         $self->{_diag_border},
#NYI         $self->{_diag_color},
#NYI         $self->{_diag_type},
#NYI         $self->{_left},
#NYI         $self->{_left_color},
#NYI         $self->{_right},
#NYI         $self->{_right_color},
#NYI         $self->{_top},
#NYI         $self->{_top_color},
#NYI 
#NYI     );
#NYI 
#NYI     return $key;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_fill_key()
#NYI #
#NYI # Returns a unique hash key for a fill style. Used by Workbook.
#NYI #
#NYI sub get_fill_key {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $key = join ':', (
#NYI         $self->{_pattern},
#NYI         $self->{_bg_color},
#NYI         $self->{_fg_color},
#NYI 
#NYI     );
#NYI 
#NYI     return $key;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_alignment_key()
#NYI #
#NYI # Returns a unique hash key for alignment formats.
#NYI #
#NYI sub get_alignment_key {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $key = join ':', (
#NYI         $self->{_text_h_align},
#NYI         $self->{_text_v_align},
#NYI         $self->{_indent},
#NYI         $self->{_rotation},
#NYI         $self->{_text_wrap},
#NYI         $self->{_shrink},
#NYI         $self->{_reading_order},
#NYI 
#NYI     );
#NYI 
#NYI     return $key;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_xf_index()
#NYI #
#NYI # Returns the index used by Worksheet->_XF()
#NYI #
#NYI sub get_xf_index {
#NYI     my $self = shift;
#NYI 
#NYI     if ( defined $self->{_xf_index} ) {
#NYI         return $self->{_xf_index};
#NYI     }
#NYI     else {
#NYI         my $key  = $self->get_format_key();
#NYI         my $indices_href = ${ $self->{_xf_format_indices} };
#NYI 
#NYI         if ( exists $indices_href->{$key} ) {
#NYI             return $indices_href->{$key};
#NYI         }
#NYI         else {
#NYI             my $index = 1 + scalar keys %$indices_href;
#NYI             $indices_href->{$key} = $index;
#NYI             $self->{_xf_index} = $index;
#NYI             return $index;
#NYI         }
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_dxf_index()
#NYI #
#NYI # Returns the index used by Worksheet->_XF()
#NYI #
#NYI sub get_dxf_index {
#NYI     my $self = shift;
#NYI 
#NYI     if ( defined $self->{_dxf_index} ) {
#NYI         return $self->{_dxf_index};
#NYI     }
#NYI     else {
#NYI         my $key  = $self->get_format_key();
#NYI         my $indices_href = ${ $self->{_dxf_format_indices} };
#NYI 
#NYI         if ( exists $indices_href->{$key} ) {
#NYI             return $indices_href->{$key};
#NYI         }
#NYI         else {
#NYI             my $index = scalar keys %$indices_href;
#NYI             $indices_href->{$key} = $index;
#NYI             $self->{_dxf_index} = $index;
#NYI             return $index;
#NYI         }
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_color()
#NYI #
#NYI # Used in conjunction with the set_xxx_color methods to convert a color
#NYI # string into a number. Color range is 0..63 but we will restrict it
#NYI # to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
#NYI #
#NYI sub _get_color {
#NYI 
#NYI     my %colors = (
#NYI         aqua    => 0x0F,
#NYI         cyan    => 0x0F,
#NYI         black   => 0x08,
#NYI         blue    => 0x0C,
#NYI         brown   => 0x10,
#NYI         magenta => 0x0E,
#NYI         fuchsia => 0x0E,
#NYI         gray    => 0x17,
#NYI         grey    => 0x17,
#NYI         green   => 0x11,
#NYI         lime    => 0x0B,
#NYI         navy    => 0x12,
#NYI         orange  => 0x35,
#NYI         pink    => 0x21,
#NYI         purple  => 0x14,
#NYI         red     => 0x0A,
#NYI         silver  => 0x16,
#NYI         white   => 0x09,
#NYI         yellow  => 0x0D,
#NYI     );
#NYI 
#NYI     # Return RGB style colors for processing later.
#NYI     if ( $_[0] =~ m/^#[0-9A-F]{6}$/i ) {
#NYI         return $_[0];
#NYI     }
#NYI 
#NYI     # Return the default color if undef,
#NYI     return 0x00 unless defined $_[0];
#NYI 
#NYI     # or the color string converted to an integer,
#NYI     return $colors{ lc( $_[0] ) } if exists $colors{ lc( $_[0] ) };
#NYI 
#NYI     # or the default color if string is unrecognised,
#NYI     return 0x00 if ( $_[0] =~ m/\D/ );
#NYI 
#NYI     # or an index < 8 mapped into the correct range,
#NYI     return $_[0] + 8 if $_[0] < 8;
#NYI 
#NYI     # or the default color if arg is outside range,
#NYI     return 0x00 if $_[0] > 63;
#NYI 
#NYI     # or an integer in the valid range
#NYI     return $_[0];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_type()
#NYI #
#NYI # Set the XF object type as 0 = cell XF or 0xFFF5 = style XF.
#NYI #
#NYI sub set_type {
#NYI 
#NYI     my $self = shift;
#NYI     my $type = $_[0];
#NYI 
#NYI     if (defined $_[0] and $_[0] eq 0) {
#NYI         $self->{_type} = 0x0000;
#NYI     }
#NYI     else {
#NYI         $self->{_type} = 0xFFF5;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_align()
#NYI #
#NYI # Set cell alignment.
#NYI #
#NYI sub set_align {
#NYI 
#NYI     my $self     = shift;
#NYI     my $location = $_[0];
#NYI 
#NYI     return if not defined $location;    # No default
#NYI     return if $location =~ m/\d/;       # Ignore numbers
#NYI 
#NYI     $location = lc( $location );
#NYI 
#NYI     $self->set_text_h_align( 1 ) if $location eq 'left';
#NYI     $self->set_text_h_align( 2 ) if $location eq 'centre';
#NYI     $self->set_text_h_align( 2 ) if $location eq 'center';
#NYI     $self->set_text_h_align( 3 ) if $location eq 'right';
#NYI     $self->set_text_h_align( 4 ) if $location eq 'fill';
#NYI     $self->set_text_h_align( 5 ) if $location eq 'justify';
#NYI     $self->set_text_h_align( 6 ) if $location eq 'center_across';
#NYI     $self->set_text_h_align( 6 ) if $location eq 'centre_across';
#NYI     $self->set_text_h_align( 6 ) if $location eq 'merge';              # Legacy.
#NYI     $self->set_text_h_align( 7 ) if $location eq 'distributed';
#NYI     $self->set_text_h_align( 7 ) if $location eq 'equal_space';        # S::PE.
#NYI     $self->set_text_h_align( 7 ) if $location eq 'justify_distributed';
#NYI 
#NYI     $self->{_just_distrib} = 1 if $location eq 'justify_distributed';
#NYI 
#NYI     $self->set_text_v_align( 1 ) if $location eq 'top';
#NYI     $self->set_text_v_align( 2 ) if $location eq 'vcentre';
#NYI     $self->set_text_v_align( 2 ) if $location eq 'vcenter';
#NYI     $self->set_text_v_align( 3 ) if $location eq 'bottom';
#NYI     $self->set_text_v_align( 4 ) if $location eq 'vjustify';
#NYI     $self->set_text_v_align( 5 ) if $location eq 'vdistributed';
#NYI     $self->set_text_v_align( 5 ) if $location eq 'vequal_space';    # S::PE.
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_valign()
#NYI #
#NYI # Set vertical cell alignment. This is required by the set_properties() method
#NYI # to differentiate between the vertical and horizontal properties.
#NYI #
#NYI sub set_valign {
#NYI 
#NYI     my $self = shift;
#NYI     $self->set_align( @_ );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_center_across()
#NYI #
#NYI # Implements the Excel5 style "merge".
#NYI #
#NYI sub set_center_across {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->set_text_h_align( 6 );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_merge()
#NYI #
#NYI # This was the way to implement a merge in Excel5. However it should have been
#NYI # called "center_across" and not "merge".
#NYI # This is now deprecated. Use set_center_across() or better merge_range().
#NYI #
#NYI #
#NYI sub set_merge {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->set_text_h_align( 6 );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_bold()
#NYI #
#NYI #
#NYI sub set_bold {
#NYI 
#NYI     my $self = shift;
#NYI     my $bold = defined $_[0] ? $_[0] : 1;
#NYI 
#NYI     $self->{_bold} = $bold ? 1 : 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_border($style)
#NYI #
#NYI # Set cells borders to the same style
#NYI #
#NYI sub set_border {
#NYI 
#NYI     my $self  = shift;
#NYI     my $style = $_[0];
#NYI 
#NYI     $self->set_bottom( $style );
#NYI     $self->set_top( $style );
#NYI     $self->set_left( $style );
#NYI     $self->set_right( $style );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_border_color($color)
#NYI #
#NYI # Set cells border to the same color
#NYI #
#NYI sub set_border_color {
#NYI 
#NYI     my $self  = shift;
#NYI     my $color = $_[0];
#NYI 
#NYI     $self->set_bottom_color( $color );
#NYI     $self->set_top_color( $color );
#NYI     $self->set_left_color( $color );
#NYI     $self->set_right_color( $color );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_rotation($angle)
#NYI #
#NYI # Set the rotation angle of the text. An alignment property.
#NYI #
#NYI sub set_rotation {
#NYI 
#NYI     my $self     = shift;
#NYI     my $rotation = $_[0];
#NYI 
#NYI     # Argument should be a number
#NYI     return if $rotation !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/;
#NYI 
#NYI     # The arg type can be a double but the Excel dialog only allows integers.
#NYI     $rotation = int $rotation;
#NYI 
#NYI     if ( $rotation == 270 ) {
#NYI         $rotation = 255;
#NYI     }
#NYI     elsif ( $rotation >= -90 and $rotation <= 90 ) {
#NYI         $rotation = -$rotation + 90 if $rotation < 0;
#NYI     }
#NYI     else {
#NYI         carp "Rotation $rotation outside range: -90 <= angle <= 90";
#NYI         $rotation = 0;
#NYI     }
#NYI 
#NYI     $self->{_rotation} = $rotation;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_hyperlink()
#NYI #
#NYI # Set the properties for the hyperlink style.
#NYI #
#NYI sub set_hyperlink {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->{_hyperlink} = 1;
#NYI     $self->{_xf_id}     = 1;
#NYI 
#NYI     $self->set_underline( 1 );
#NYI     $self->set_theme( 10 );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_format_properties()
#NYI #
#NYI # Convert hashes of properties to method calls.
#NYI #
#NYI sub set_format_properties {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my %properties = @_;    # Merge multiple hashes into one
#NYI 
#NYI     while ( my ( $key, $value ) = each( %properties ) ) {
#NYI 
#NYI         # Strip leading "-" from Tk style properties e.g. -color => 'red'.
#NYI         $key =~ s/^-//;
#NYI 
#NYI         # Create a sub to set the property.
#NYI         my $sub = \&{"set_$key"};
#NYI         $sub->( $self, $value );
#NYI     }
#NYI }
#NYI 
#NYI # Renamed rarely used set_properties() to set_format_properties() to avoid
#NYI # confusion with Workbook method of the same name. The following acts as an
#NYI # alias for any code that uses the old name.
#NYI *set_properties = *set_format_properties;
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # AUTOLOAD. Deus ex machina.
#NYI #
#NYI # Dynamically create set methods that aren't already defined.
#NYI #
#NYI sub AUTOLOAD {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Ignore calls to DESTROY
#NYI     return if $AUTOLOAD =~ /::DESTROY$/;
#NYI 
#NYI     # Check for a valid method names, i.e. "set_xxx_yyy".
#NYI     $AUTOLOAD =~ /.*::set(\w+)/ or die "Unknown method: $AUTOLOAD\n";
#NYI 
#NYI     # Match the attribute, i.e. "_xxx_yyy".
#NYI     my $attribute = $1;
#NYI 
#NYI     # Check that the attribute exists
#NYI     exists $self->{$attribute} or die "Unknown method: $AUTOLOAD\n";
#NYI 
#NYI     # The attribute value
#NYI     my $value;
#NYI 
#NYI 
#NYI     # There are two types of set methods: set_property() and
#NYI     # set_property_color(). When a method is AUTOLOADED we store a new anonymous
#NYI     # sub in the appropriate slot in the symbol table. The speeds up subsequent
#NYI     # calls to the same method.
#NYI     #
#NYI     no strict 'refs';    # To allow symbol table hackery
#NYI 
#NYI     if ( $AUTOLOAD =~ /.*::set\w+color$/ ) {
#NYI 
#NYI         # For "set_property_color" methods
#NYI         $value = _get_color( $_[0] );
#NYI 
#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self = shift;
#NYI 
#NYI             $self->{$attribute} = _get_color( $_[0] );
#NYI         };
#NYI     }
#NYI     else {
#NYI 
#NYI         $value = $_[0];
#NYI         $value = 1 if not defined $value;    # The default value is always 1
#NYI 
#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self  = shift;
#NYI             my $value = shift;
#NYI 
#NYI             $value = 1 if not defined $value;
#NYI             $self->{$attribute} = $value;
#NYI         };
#NYI     }
#NYI 
#NYI 
#NYI     $self->{$attribute} = $value;
#NYI }
#NYI 
#NYI 
#NYI 1;
#NYI 
#NYI 
#NYI __END__
#NYI 
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Format - A class for defining Excel formatting.
#NYI 
#NYI =head1 SYNOPSIS
#NYI 
#NYI See the documentation for L<Excel::Writer::XLSX>
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
