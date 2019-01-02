use v6.c+;

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

#NYI our @ISA     = qw(Exporter);
#NYI our $VERSION = '0.96';
#NYI our $AUTOLOAD;

has @!xf-format-indices;
has @!dxf-format-indices;
has @!xf-index;
has @!dxf-index;

has $.num-format       = 0;
has $.num-format-index = 0;
has $.font-index       = 0;
has $.has-font         = 0;
has $.has-dxf-font     = 0;
has $.font             = 'Calibri';
has $.size             = 11;
has $.bold             = 0;
has $.italic           = 0;
has $.color            = 0x0;
has $.underline        = 0;
has $.font-strikeout   = 0;
has $.font-outline     = 0;
has $.font-shadow      = 0;
has $.font-script      = 0;
has $.font-family      = 2;
has $.font-charset     = 0;
has $.font-scheme      = 'minor';
has $.font-condense    = 0;
has $.font-extend      = 0;
has $.theme            = 0;
has $.hyperlink        = 0;
has $.xf-id            = 0;

has $!hidden = 0;
has $!locked = 1;

has $!text-h-align  = 0;
has $!text-wrap     = 0;
has $!text-v-align  = 0;
has $!text-justlast = 0;
has $!rotation      = 0;

has $!fg-color     = 0x00;
has $!bg-color     = 0x00;
has $!pattern      = 0;
has $!has-fill     = 0;
has $!has-dxf-fill = 0;
has $!fill-index   = 0;
has $!fill-count   = 0;

has $!border-index   = 0;
has $!has-border     = 0;
has $!has-dxf-border = 0;
has $!border-count   = 0;

has $!bottom       = 0;
has $!bottom-color = 0x0;
has $!diag-border  = 0;
has $!diag-color   = 0x0;
has $!diag-type    = 0;
has $!left         = 0;
has $!left-color   = 0x0;
has $!right        = 0;
has $!right-color  = 0x0;
has $!top          = 0;
has $!top-color    = 0x0;

has $!indent        = 0;
has $!shrink        = 0;
has $!merge-range   = 0;
has $!reading-order = 0;
has $!just-distrib  = 0;
has $!color-indexed = 0;
has $!font-only     = 0;

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

#NYI     # Set properties passed to Workbook::add-format()
#NYI     $self->set-format-properties(@_) if @_;

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
    my @xf-index           = @!xf-index;
    my @dxf-index          = @!dxf-index;
    my @xf-format-indices  = @!xf-format-indices;
    my @dxf-format-indices = @!dxf-format-indices;
    my $palette            = $!palette;

    # Copy properties.
    #TODO %$self             = %$other;

    # Restore original properties.
    @!xf-index           = @xf-index;
    @!dxf-index          = @dxf-index;
    @!xf-format-indices  = @xf-format-indices;
    @!dxf-format-indices = @dxf-format-indices;
    $!palette            = $palette;
}


###############################################################################
#
# get-align-properties()
#
# Return properties for an Style xf <alignment> sub-element.
#
method get-align-properties {
    my @align;    # Attributes to return

    # Check if any alignment options in the format have been changed.
    my $changed =
      (      $!text-h-align != 0
          || $!text-v-align != 0
          || $!indent != 0
          || $!rotation != 0
          || $!text-wrap != 0
          || $!shrink != 0
          || $!reading-order != 0 ) ?? 1 !! 0;

    return unless $changed;

    # Indent is only allowed for horizontal left, right and distributed. If it
    # is defined for any other alignment or no alignment has been set then
    # default to left alignment.
    if   $!indent
      && $!text-h-align != 1
      && $!text-h-align != 3
      && $!text-h-align != 7
    {
        $!text-h-align = 1;
    }

    # Check for properties that are mutually exclusive.
    $!shrink       = 0 if $!text-wrap;
    $!shrink       = 0 if $!text-h-align == 4;    # Fill
    $!shrink       = 0 if $!text-h-align == 5;    # Justify
    $!shrink       = 0 if $!text-h-align == 7;    # Distributed
    $!just-distrib = 0 if $!text-h-align != 7;    # Distributed
    $!just-distrib = 0 if $!indent;

    my $continuous = 'centerContinuous';

    @align.push: 'horizontal', 'left'        if $!text-h-align == 1;
    @align.push: 'horizontal', 'center'      if $!text-h-align == 2;
    @align.push: 'horizontal', 'right'       if $!text-h-align == 3;
    @align.push: 'horizontal', 'fill'        if $!text-h-align == 4;
    @align.push: 'horizontal', 'justify'     if $!text-h-align == 5;
    @align.push: 'horizontal', $continuous   if $!text-h-align == 6;
    @align.push: 'horizontal', 'distributed' if $!text-h-align == 7;

    @align.push: 'justifyLastLine', 1 if $!just-distrib;

    # Property 'vertical' => 'bottom' is a default. It sets applyAlignment
    # without an alignment sub-element.
    @align.push: 'vertical', 'top'         if $!text-v-align == 1;
    @align.push: 'vertical', 'center'      if $!text-v-align == 2;
    @align.push: 'vertical', 'justify'     if $!text-v-align == 4;
    @align.push: 'vertical', 'distributed' if $!text-v-align == 5;

    @align.push: 'indent',       $!indent   if $!indent;
    @align.push: 'textRotation', $!rotation if $!rotation;

    @align.push: 'wrapText',     1 if $!text-wrap;
    @align.push: 'shrinkToFit',  1 if $!shrink;

    @align.push: 'readingOrder', 1 if $!reading-order == 1;
    @align.push: 'readingOrder', 2 if $!reading-order == 2;

    return $changed, @align;
}


###############################################################################
#
# get-protection-properties()
#
# Return properties for an Excel XML <Protection> element.
#
method get-protection-properties {
    my @attribs;

    push @attribs, 'locked', 0 if ! $!locked;
    push @attribs, 'hidden', 1 if   $!hidden;

    return @attribs;
}


###############################################################################
#
# get-format-key()
#
# Returns a unique hash key for the Format object.
#
method get-format-key {
    my $key = join ':',
      (
        self.get-font-key(), self.get-border-key,
        self.get-fill-key(), self.get-alignment-key(),
        $!num-format,  $!locked,
        $!hidden
      );

    return $key;
}

###############################################################################
#
# get-font-key()
#
# Returns a unique hash key for a font. Used by Workbook.
#
method get-font-key {
    my $key = join ':', (
        $!bold,
        $!color,
        $!font-charset,
        $!font-family,
        $!font-outline,
        $!font-script,
        $!font-shadow,
        $!font-strikeout,
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
# get-border-key()
#
# Returns a unique hash key for a border style. Used by Workbook.
#
method get-border-key {
    my $key = join ':', (
        $!bottom,
        $!bottom-color,
        $!diag-border,
        $!diag-color,
        $!diag-type,
        $!left,
        $!left-color,
        $!right,
        $!right-color,
        $!top,
        $!top-color,

    );

    return $key;
}


###############################################################################
#
# get-fill-key()
#
# Returns a unique hash key for a fill style. Used by Workbook.
#
method get-fill-key {
    my $key = join ':', (
        $!pattern,
        $!bg-color,
        $!fg-color,

    );

    return $key;
}


###############################################################################
#
# get-alignment-key()
#
# Returns a unique hash key for alignment formats.
#
method get-alignment-key {
    my $key = join ':', (
        $!text-h-align,
        $!text-v-align,
        $!indent,
        $!rotation,
        $!text-wrap,
        $!shrink,
        $!reading-order,

    );

    return $key;
}


###############################################################################
#
# get-xf-index()
#
# Returns the index used by Worksheet->_XF()
#
method get-xf-index {
    if @!xf-index.defined {
        return @!xf-index;
    }
    else {
        my $key  = self.get-format-key();
        my %indices-href = @!xf-format-indices;

        if %indices-href{$key}.exists {
            return %indices-href{$key};
        }
        else {
            my $index = 1 + %indices-href.keys.elems;
            %indices-href{$key} = $index;
            @!xf-index = $index;
            return $index;
        }
    }
}


###############################################################################
#
# get-dxf-index()
#
# Returns the index used by Worksheet->_XF()
#
method get-dxf-index {
    if @!dxf-index.defined {
        return @!dxf-index;
    }
    else {
        my $key  = self.get-format-key();
        my %indices-href = @!dxf-format-indices;

        if %indices-href{$key}.exists {
            return %indices-href{$key};
        }
        else {
            my $index = %indices-href.keys.elems;
            %indices-href{$key} = $index;
            @!dxf-index = $index;
            return $index;
        }
    }
}


###############################################################################
#
# get-color()
#
# Used in conjunction with the set-xxx-color methods to convert a color
# string into a number. Color range is 0..63 but we will restrict it
# to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
#
method get-color($color?) {

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
# set-type()
#
# Set the XF object type as 0 = cell XF or 0xFFF5 = style XF.
#
method set-type($type) {
    if ($type.defined and $type eq 0) {
        $!type = 0x0000;
    }
    else {
        $!type = 0xFFF5;
    }
}


###############################################################################
#
# set-align()
#
# Set cell alignment.
#
method set-align($location) {
    return if not $location.defined;    # No default
    return if $location ~~ /\d/;       # Ignore numbers

    $location .= lc;

    self.set-text-h-align( 1 ) if $location eq 'left';
    self.set-text-h-align( 2 ) if $location eq 'centre';
    self.set-text-h-align( 2 ) if $location eq 'center';
    self.set-text-h-align( 3 ) if $location eq 'right';
    self.set-text-h-align( 4 ) if $location eq 'fill';
    self.set-text-h-align( 5 ) if $location eq 'justify';
    self.set-text-h-align( 6 ) if $location eq 'center-across';
    self.set-text-h-align( 6 ) if $location eq 'centre-across';
    self.set-text-h-align( 6 ) if $location eq 'merge';              # Legacy.
    self.set-text-h-align( 7 ) if $location eq 'distributed';
    self.set-text-h-align( 7 ) if $location eq 'equal-space';        # S::PE.
    self.set-text-h-align( 7 ) if $location eq 'justify-distributed';

    $!just-distrib = 1 if $location eq 'justify-distributed';

    self.set-text-v-align( 1 ) if $location eq 'top';
    self.set-text-v-align( 2 ) if $location eq 'vcentre';
    self.set-text-v-align( 2 ) if $location eq 'vcenter';
    self.set-text-v-align( 3 ) if $location eq 'bottom';
    self.set-text-v-align( 4 ) if $location eq 'vjustify';
    self.set-text-v-align( 5 ) if $location eq 'vdistributed';
    self.set-text-v-align( 5 ) if $location eq 'vequal-space';    # S::PE.
}


###############################################################################
#
# set-valign()
#
# Set vertical cell alignment. This is required by the set-properties() method
# to differentiate between the vertical and horizontal properties.
#
method set-valign(*@args) {
    self.set-align( @args );
}


###############################################################################
#
# set-center-across()
#
# Implements the Excel5 style "merge".
#
method set-center-across {
    self.set-text-h-align( 6 );
}


###############################################################################
#
# set-merge()
#
# This was the way to implement a merge in Excel5. However it should have been
# called "center-across" and not "merge".
# This is now deprecated. Use set-center-across() or better merge-range().
#
#
method set-merge {
    self.set-text-h-align( 6 );
}


###############################################################################
#
# set-bold()
#
#
method set-bold($bold = 1) {
    $!bold = $bold;
}


###############################################################################
#
# set-border($style)
#
# Set cells borders to the same style
#
method set-border($style) {
    self.set-bottom( $style );
    self.set-top(    $style );
    self.set-left(   $style );
    self.set-right(  $style );
}


###############################################################################
#
# set-border-color($color)
#
# Set cells border to the same color
#
method set-border-color($color) {
    self.set-bottom-color( $color );
    self.set-top-color( $color );
    self.set-left-color( $color );
    self.set-right-color( $color );
}


###############################################################################
#
# set-rotation($angle)
#
# Set the rotation angle of the text. An alignment property.
#
method set-rotation($rotation) {
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
# set-hyperlink()
#
# Set the properties for the hyperlink style.
#
method set-hyperlink {
    $!hyperlink = 1;
    $!xf-id     = 1;

    self.set-underline( 1 );
    self.set-theme( 10 );
}


###############################################################################
#
# set-format-properties()
#
# Convert hashes of properties to method calls.
#
method set-format-properties(*%properties) { # Merge multiple hashes into one

    for %properties.kv -> $key, $value {

        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        $key ~~ s/^\-//;
# TODO
#        # Create a sub to set the property.
#        my $sub = \&{"set-$key"};
#        $sub->( $self, $value );
    }
}

# Renamed rarely used set-properties() to set-format-properties() to avoid
# confusion with Workbook method of the same name. The following acts as an
# alias for any code that uses the old name.
#TODO *set-properties = *set-format-properties;


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

#NYI     # Check for a valid method names, i.e. "set-xxx-yyy".
#NYI     $AUTOLOAD =~ /.*::set(\w+)/ or die "Unknown method: $AUTOLOAD\n";

#NYI     # Match the attribute, i.e. "-xxx-yyy".
#NYI     my $attribute = $1;

#NYI     # Check that the attribute exists
#NYI     exists $self->{$attribute} or die "Unknown method: $AUTOLOAD\n";

#NYI     # The attribute value
#NYI     my $value;


#NYI     # There are two types of set methods: set-property() and
#NYI     # set-property-color(). When a method is AUTOLOADED we store a new anonymous
#NYI     # sub in the appropriate slot in the symbol table. The speeds up subsequent
#NYI     # calls to the same method.
#NYI     #
#NYI     no strict 'refs';    # To allow symbol table hackery

#NYI     if ( $AUTOLOAD =~ /.*::set\w+color$/ ) {

#NYI         # For "set-property-color" methods
#NYI         $value = get-color( $_[0] );

#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self = shift;

#NYI             $self->{$attribute} = get=color( $_[0] );
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
