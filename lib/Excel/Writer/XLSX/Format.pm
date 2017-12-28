unit class Excel::Writer::XLSX::Format;

###############################################################################
#
# Format - A class for defining Excel formatting.
#
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use v6.c;
#NYI our @ISA     = qw(Exporter);
#NYI our $VERSION = '0.96';
#NYI our $AUTOLOAD;

has @!xf_format_indices;
has @!dxf_format_indices;
has @!xf_index;
has @!dxf_index;

has $!num_format       = 0;
has $!num_format_index = 0;
has $!font_index       = 0;
has $!has_font         = 0;
has $!has_dxf_font     = 0;
has $!font             = 'Calibri';
has $!size             = 11;
has $!bold             = 0;
has $!italic           = 0;
has $!color            = 0x0;
has $!underline        = 0;
has $!font_strikeout   = 0;
has $!font_outline     = 0;
has $!font_shadow      = 0;
has $!font_script      = 0;
has $!font_family      = 2;
has $!font_charset     = 0;
has $!font_scheme      = 'minor';
has $!font_condense    = 0;
has $!font_extend      = 0;
has $!theme            = 0;
has $!hyperlink        = 0;
has $!xf_id            = 0;

has $!hidden = 0;
has $!locked = 1;

has $!text_h_align  = 0;
has $!text_wrap     = 0;
has $!text_v_align  = 0;
has $!text_justlast = 0;
has $!rotation      = 0;

has $!fg_color     = 0x00;
has $!bg_color     = 0x00;
has $!pattern      = 0;
has $!has_fill     = 0;
has $!has_dxf_fill = 0;
has $!fill_index   = 0;
has $!fill_count   = 0;

has $!border_index   = 0;
has $!has_border     = 0;
has $!has_dxf_border = 0;
has $!border_count   = 0;

has $!bottom       = 0;
has $!bottom_color = 0x0;
has $!diag_border  = 0;
has $!diag_color   = 0x0;
has $!diag_type    = 0;
has $!left         = 0;
has $!left_color   = 0x0;
has $!right        = 0;
has $!right_color  = 0x0;
has $!top          = 0;
has $!top_color    = 0x0;

has $!indent        = 0;
has $!shrink        = 0;
has $!merge_range   = 0;
has $!reading_order = 0;
has $!just_distrib  = 0;
has $!color_indexed = 0;
has $!font_only     = 0;

# Added because they're used, but not previously defined
has $!palette;
has $!type;

###############################################################################
#
# new()
#
# Constructor
#
#NYI sub new {

#NYI     my $class = shift;

#NYI     my $self = {

#NYI     };

#NYI     bless $self, $class;

#NYI     # Set properties passed to Workbook::add_format()
#NYI     $self->set_format_properties(@_) if @_;

#NYI     return $self;
#NYI }


###############################################################################
#
# copy($format)
#
# Copy the attributes of another Excel::Writer::XLSX::Format object.
#
method copy($other) {
    return unless $other.defined;
    #TODO return unless $other.^WHAT eq (Excel::Writer::XLSX::Format);

    # Store properties that we don't want over-ridden.
    my @xf_index           = @!xf_index;
    my @dxf_index          = @!dxf_index;
    my @xf_format_indices  = @!xf_format_indices;
    my @dxf_format_indices = @!dxf_format_indices;
    my $palette            = $!palette;

    # Copy properties.
    #TODO %$self             = %$other;

    # Restore original properties.
    @!xf_index           = @xf_index;
    @!dxf_index          = @dxf_index;
    @!xf_format_indices  = @xf_format_indices;
    @!dxf_format_indices = @dxf_format_indices;
    $!palette            = $palette;
}


###############################################################################
#
# get_align_properties()
#
# Return properties for an Style xf <alignment> sub-element.
#
method get_align_properties {
    my @align;    # Attributes to return

    # Check if any alignment options in the format have been changed.
    my $changed =
      (      $!text_h_align != 0
          || $!text_v_align != 0
          || $!indent != 0
          || $!rotation != 0
          || $!text_wrap != 0
          || $!shrink != 0
          || $!reading_order != 0 ) ?? 1 !! 0;

    return unless $changed;

    # Indent is only allowed for horizontal left, right and distributed. If it
    # is defined for any other alignment or no alignment has been set then
    # default to left alignment.
    if   $!indent
      && $!text_h_align != 1
      && $!text_h_align != 3
      && $!text_h_align != 7
    {
        $!text_h_align = 1;
    }

    # Check for properties that are mutually exclusive.
    $!shrink       = 0 if $!text_wrap;
    $!shrink       = 0 if $!text_h_align == 4;    # Fill
    $!shrink       = 0 if $!text_h_align == 5;    # Justify
    $!shrink       = 0 if $!text_h_align == 7;    # Distributed
    $!just_distrib = 0 if $!text_h_align != 7;    # Distributed
    $!just_distrib = 0 if $!indent;

    my $continuous = 'centerContinuous';

    @align.push: 'horizontal', 'left'        if $!text_h_align == 1;
    @align.push: 'horizontal', 'center'      if $!text_h_align == 2;
    @align.push: 'horizontal', 'right'       if $!text_h_align == 3;
    @align.push: 'horizontal', 'fill'        if $!text_h_align == 4;
    @align.push: 'horizontal', 'justify'     if $!text_h_align == 5;
    @align.push: 'horizontal', $continuous   if $!text_h_align == 6;
    @align.push: 'horizontal', 'distributed' if $!text_h_align == 7;

    @align.push: 'justifyLastLine', 1 if $!just_distrib;

    # Property 'vertical' => 'bottom' is a default. It sets applyAlignment
    # without an alignment sub-element.
    @align.push: 'vertical', 'top'         if $!text_v_align == 1;
    @align.push: 'vertical', 'center'      if $!text_v_align == 2;
    @align.push: 'vertical', 'justify'     if $!text_v_align == 4;
    @align.push: 'vertical', 'distributed' if $!text_v_align == 5;

    @align.push: 'indent',       $!indent   if $!indent;
    @align.push: 'textRotation', $!rotation if $!rotation;

    @align.push: 'wrapText',     1 if $!text_wrap;
    @align.push: 'shrinkToFit',  1 if $!shrink;

    @align.push: 'readingOrder', 1 if $!reading_order == 1;
    @align.push: 'readingOrder', 2 if $!reading_order == 2;

    return $changed, @align;
}


###############################################################################
#
# get_protection_properties()
#
# Return properties for an Excel XML <Protection> element.
#
method get_protection_properties {
    my @attribs;

    push @attribs, 'locked', 0 if ! $!locked;
    push @attribs, 'hidden', 1 if   $!hidden;

    return @attribs;
}


###############################################################################
#
# get_format_key()
#
# Returns a unique hash key for the Format object.
#
method get_format_key {
    my $key = join ':',
      (
        self.get_font_key(), self.get_border_key,
        self.get_fill_key(), self.get_alignment_key(),
        $!num_format,  $!locked,
        $!hidden
      );

    return $key;
}

###############################################################################
#
# get_font_key()
#
# Returns a unique hash key for a font. Used by Workbook.
#
method get_font_key {
    my $key = join ':', (
        $!bold,
        $!color,
        $!font_charset,
        $!font_family,
        $!font_outline,
        $!font_script,
        $!font_shadow,
        $!font_strikeout,
        $!font,
        $!italic,
        $!size,
        $!underline,
        $!theme,

    );

    return $key;
}


###############################################################################
#
# get_border_key()
#
# Returns a unique hash key for a border style. Used by Workbook.
#
method get_border_key {
    my $key = join ':', (
        $!bottom,
        $!bottom_color,
        $!diag_border,
        $!diag_color,
        $!diag_type,
        $!left,
        $!left_color,
        $!right,
        $!right_color,
        $!top,
        $!top_color,

    );

    return $key;
}


###############################################################################
#
# get_fill_key()
#
# Returns a unique hash key for a fill style. Used by Workbook.
#
method get_fill_key {
    my $key = join ':', (
        $!pattern,
        $!bg_color,
        $!fg_color,

    );

    return $key;
}


###############################################################################
#
# get_alignment_key()
#
# Returns a unique hash key for alignment formats.
#
method get_alignment_key {
    my $key = join ':', (
        $!text_h_align,
        $!text_v_align,
        $!indent,
        $!rotation,
        $!text_wrap,
        $!shrink,
        $!reading_order,

    );

    return $key;
}


###############################################################################
#
# get_xf_index()
#
# Returns the index used by Worksheet->_XF()
#
method get_xf_index {
    if @!xf_index.defined {
        return @!xf_index;
    }
    else {
        my $key  = self.get_format_key();
        my %indices_href = @!xf_format_indices;

        if %indices_href{$key}.exists {
            return %indices_href{$key};
        }
        else {
            my $index = 1 + %indices_href.keys.elems;
            %indices_href{$key} = $index;
            @!xf_index = $index;
            return $index;
        }
    }
}


###############################################################################
#
# get_dxf_index()
#
# Returns the index used by Worksheet->_XF()
#
method get_dxf_index {
    if @!dxf_index.defined {
        return @!dxf_index;
    }
    else {
        my $key  = self.get_format_key();
        my %indices_href = @!dxf_format_indices;

        if %indices_href{$key}.exists {
            return %indices_href{$key};
        }
        else {
            my $index = %indices_href.keys.elems;
            %indices_href{$key} = $index;
            @!dxf_index = $index;
            return $index;
        }
    }
}


###############################################################################
#
# _get_color()
#
# Used in conjunction with the set_xxx_color methods to convert a color
# string into a number. Color range is 0..63 but we will restrict it
# to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
#
method get_color($color?) {

    my %colors = (
        aqua    => 0x0F,
        cyan    => 0x0F,
        black   => 0x08,
        blue    => 0x0C,
        brown   => 0x10,
        magenta => 0x0E,
        fuchsia => 0x0E,
        gray    => 0x17,
        grey    => 0x17,
        green   => 0x11,
        lime    => 0x0B,
        navy    => 0x12,
        orange  => 0x35,
        pink    => 0x21,
        purple  => 0x14,
        red     => 0x0A,
        silver  => 0x16,
        white   => 0x09,
        yellow  => 0x0D,
    );

    # Return RGB style colors for processing later.
    if $color ~~ m:i/^\#<[0..9 A..F]> ** 6 $/ {
        return $color;
    }

    # Return the default color if undef,
    return 0x00 unless $color.defined;

    # or the color string converted to an integer,
    return %colors{ lc( $color ) } if %colors{ lc( $color ) }.exists;

    # or the default color if string is unrecognised,
    return 0x00 if ( $color ~~ /\D/ );

    # or an index < 8 mapped into the correct range,
    return $color + 8 if $color < 8;

    # or the default color if arg is outside range,
    return 0x00 if $color > 63;

    # or an integer in the valid range
    return $color;
}


###############################################################################
#
# set_type()
#
# Set the XF object type as 0 = cell XF or 0xFFF5 = style XF.
#
method set_type($type) {
    if ($type.defined and $type eq 0) {
        $!type = 0x0000;
    }
    else {
        $!type = 0xFFF5;
    }
}


###############################################################################
#
# set_align()
#
# Set cell alignment.
#
method set_align($location) {
    return if not $location.defined;    # No default
    return if $location ~~ /\d/;       # Ignore numbers

    $location .= lc;

    self.set_text_h_align( 1 ) if $location eq 'left';
    self.set_text_h_align( 2 ) if $location eq 'centre';
    self.set_text_h_align( 2 ) if $location eq 'center';
    self.set_text_h_align( 3 ) if $location eq 'right';
    self.set_text_h_align( 4 ) if $location eq 'fill';
    self.set_text_h_align( 5 ) if $location eq 'justify';
    self.set_text_h_align( 6 ) if $location eq 'center_across';
    self.set_text_h_align( 6 ) if $location eq 'centre_across';
    self.set_text_h_align( 6 ) if $location eq 'merge';              # Legacy.
    self.set_text_h_align( 7 ) if $location eq 'distributed';
    self.set_text_h_align( 7 ) if $location eq 'equal_space';        # S::PE.
    self.set_text_h_align( 7 ) if $location eq 'justify_distributed';

    $!just_distrib = 1 if $location eq 'justify_distributed';

    self.set_text_v_align( 1 ) if $location eq 'top';
    self.set_text_v_align( 2 ) if $location eq 'vcentre';
    self.set_text_v_align( 2 ) if $location eq 'vcenter';
    self.set_text_v_align( 3 ) if $location eq 'bottom';
    self.set_text_v_align( 4 ) if $location eq 'vjustify';
    self.set_text_v_align( 5 ) if $location eq 'vdistributed';
    self.set_text_v_align( 5 ) if $location eq 'vequal_space';    # S::PE.
}


###############################################################################
#
# set_valign()
#
# Set vertical cell alignment. This is required by the set_properties() method
# to differentiate between the vertical and horizontal properties.
#
method set_valign(*@args) {
    self.set_align( @args );
}


###############################################################################
#
# set_center_across()
#
# Implements the Excel5 style "merge".
#
method set_center_across {
    self.set_text_h_align( 6 );
}


###############################################################################
#
# set_merge()
#
# This was the way to implement a merge in Excel5. However it should have been
# called "center_across" and not "merge".
# This is now deprecated. Use set_center_across() or better merge_range().
#
#
method set_merge {
    self.set_text_h_align( 6 );
}


###############################################################################
#
# set_bold()
#
#
method set_bold($bold = 1) {
    $!bold = $bold;
}


###############################################################################
#
# set_border($style)
#
# Set cells borders to the same style
#
method set_border($style) {
    self.set_bottom( $style );
    self.set_top(    $style );
    self.set_left(   $style );
    self.set_right(  $style );
}


###############################################################################
#
# set_border_color($color)
#
# Set cells border to the same color
#
method set_border_color($color) {
    self.set_bottom_color( $color );
    self.set_top_color( $color );
    self.set_left_color( $color );
    self.set_right_color( $color );
}


###############################################################################
#
# set_rotation($angle)
#
# Set the rotation angle of the text. An alignment property.
#
method set_rotation($rotation) {
    # Argument should be a number
    return if $rotation !~~ /^
			        (<[+-]>?) <before \d|\.\d>
                                \d*
                                (\.\d*)?
                                (<[Ee]>(<[+-]>?\d+))?
                            $/;

    # The arg type can be a double but the Excel dialog only allows integers.
    $rotation .= int;

    if $rotation == 270 {
        $rotation = 255;
    }
    elsif -90 <= $rotation <= 90 {
        $rotation = -$rotation + 90 if $rotation < 0;
    }
    else {
        note "Rotation $rotation outside range: -90 <= angle <= 90";
        $rotation = 0;
    }

    $!rotation = $rotation;
}


###############################################################################
#
# set_hyperlink()
#
# Set the properties for the hyperlink style.
#
method set_hyperlink {
    $!hyperlink = 1;
    $!xf_id     = 1;

    self.set_underline( 1 );
    self.set_theme( 10 );
}


###############################################################################
#
# set_format_properties()
#
# Convert hashes of properties to method calls.
#
method set_format_properties(*%properties) { # Merge multiple hashes into one

    for %properties.kv -> $key, $value {

        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        $key ~~ s/^\-//;
# TODO
#        # Create a sub to set the property.
#        my $sub = \&{"set_$key"};
#        $sub->( $self, $value );
    }
}

# Renamed rarely used set_properties() to set_format_properties() to avoid
# confusion with Workbook method of the same name. The following acts as an
# alias for any code that uses the old name.
#TODO *set_properties = *set_format_properties;


###############################################################################
#
# AUTOLOAD. Deus ex machina.
#
# Dynamically create set methods that aren't already defined.
#
#NYI sub AUTOLOAD {

#NYI     my $self = shift;

#NYI     # Ignore calls to DESTROY
#NYI     return if $AUTOLOAD =~ /::DESTROY$/;

#NYI     # Check for a valid method names, i.e. "set_xxx_yyy".
#NYI     $AUTOLOAD =~ /.*::set(\w+)/ or die "Unknown method: $AUTOLOAD\n";

#NYI     # Match the attribute, i.e. "_xxx_yyy".
#NYI     my $attribute = $1;

#NYI     # Check that the attribute exists
#NYI     exists $self->{$attribute} or die "Unknown method: $AUTOLOAD\n";

#NYI     # The attribute value
#NYI     my $value;


#NYI     # There are two types of set methods: set_property() and
#NYI     # set_property_color(). When a method is AUTOLOADED we store a new anonymous
#NYI     # sub in the appropriate slot in the symbol table. The speeds up subsequent
#NYI     # calls to the same method.
#NYI     #
#NYI     no strict 'refs';    # To allow symbol table hackery

#NYI     if ( $AUTOLOAD =~ /.*::set\w+color$/ ) {

#NYI         # For "set_property_color" methods
#NYI         $value = _get_color( $_[0] );

#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self = shift;

#NYI             $self->{$attribute} = _get_color( $_[0] );
#NYI         };
#NYI     }
#NYI     else {

#NYI         $value = $_[0];
#NYI         $value = 1 if not defined $value;    # The default value is always 1

#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self  = shift;
#NYI             my $value = shift;

#NYI             $value = 1 if not defined $value;
#NYI             $self->{$attribute} = $value;
#NYI         };
#NYI     }


#NYI     $self->{$attribute} = $value;
#NYI }


=begin pod


__END__


=head1 NAME

Format - A class for defining Excel formatting.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

(c) MM-MMXVII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
=end pod
