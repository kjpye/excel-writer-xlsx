use v6.c+;

use File::Temp; # <tempfile>;
use Archive::SimpleZip;
use Excel::Writer::XLSX::Worksheet;
# use Excel::Writer::XLSX::Chartsheet;
# use Excel::Writer::XLSX::Format;
# use Excel::Writer::XLSX::Shape;
# use Excel::Writer::XLSX::Chart;
use Excel::Writer::XLSX::Package::Packager;
use Excel::Writer::XLSX::Package::XMLwriter;
use Excel::Writer::XLSX::Utility;

unit class Excel::Writer::XLSX::Workbook is Excel::Writer::XLSX::Package::XMLwriter;

#`[

 Workbook - A class for writing Excel Workbooks.


 Used in conjunction with Excel::Writer::XLSX

 Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
 Copyright 2017-2018, Kevin.Pye,     kjpye@cpan.org

 Documentation after __END__
]

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';

###############################################################################
#
# Public and private API methods.
#
###############################################################################

has $!filename;
has $.tempdir                is rw;
has $!date1904           = 0;
has $!activesheet        = 0;
has $!firstsheet         = 0;
has $!selected           = 0;
has $!fileclosed         = 0;
has $!filehandle;
has $!internal-fh        = 0;
has $!sheet-name         = 'Sheet';
has $!chart-name         = 'Chart';
has $!sheetname-count    = 0;
has $!chartname-count    = 0;
has @.worksheets         = [];
has @!charts             = [];
has @!drawings           = [];
has %.sheetnames         = {};
has @!formats            = [];
has @!xf-formats         = [];
has %!xf-format-indices  = {};
has @!dxf-formats        = [];
has %!dxf-format-indices = {};
has @!palette            = [];
has $!font-count         = 0;
has $!num-format-count   = 0;
has @!defined-names      = [];
has @!named-ranges       = [];
has @!custom-colors      = [];
has %!doc-properties     = {};
has @!custom-properties  = [];
has @!createtime         = [ now ];
has $!num-vml-files      = 0;
has $!num-comment-files  = 0;
has $!optimization       = 0;
has $!x-window           = 240;
has $!y-window           = 15;
has $!window-width       = 16095;
has $!window-height      = 9660;
has $!tab-ratio          = 500;
has $!excel2003-style    = 0;
has $!vba-codename;

has %!default-format-properties = {};

# Structures for the shared strings data.
has $!str-total  = 0;
has $!str-unique = 0;
has %!str-table  = {};
has $!str-array  = [];

# Formula calculation default settings.
has $!calc-id      = 124519;
has $!calc-mode    = 'auto';
has $!calc-on-load = 1;

has $!vba-project;
has @!shapes;

###############################################################################
#
# new()
#
# Constructor.
#
method TWEAK (*%args) {
note "in TWEAK";
#NYI     $self.filename = $_[0] || '';
#NYI     my $options = $_[1] || {};

#NYI     if ( exists $options->{tempdir} ) {
#NYI         $self->{tempdir} = $options->{tempdir};
#NYI     }

#NYI     if ( exists $options->{date1904} ) {
#NYI         $self->{date1904} = $options->{date1904};
#NYI     }

#NYI     if ( exists $options->{optimization} ) {
#NYI         $self->{optimization} = $options->{optimization};
#NYI     }

#NYI     if ( exists $options->{default-format-properties} ) {
#NYI         $self->{default-format-properties} =
#NYI           $options->{default-format-properties};
#NYI     }

#NYI     if ( exists $options->{excel2003-style} ) {
#NYI         $self->{excel2003-style} = 1;
#NYI     }


#NYI     bless $self, $class;

# Add the default cell format.
  if $!excel2003-style {
    self.add-format( xf-index => 0, font-family => 0 );
  }
  else {
    self.add-format( xf-index => 0 );
  }

# Check for a filename unless it is an existing filehandle
  $!filename = %args<filename>;
  if ! $!filename.defined {
    fail 'Filename required by Excel::Writer::XLSX.new';
    return Nil;
  }


# If filename is a reference we assume that it is a valid filehandle.
  if $!filename ~~ (IO::Handle) {
    $!filehandle  = $!filename;
    $!internal-fh = 0;
  } elsif $!filename eq '-' {
# Support special filename/filehandle '-' for backward compatibility.
#    binmode $*IN;
    $!filehandle = $*IN;
    $!internal-fh = 0;
  } else {
      if ! $!filename {
	  fail 'Filename required by Excel::Writer::XLSX.new';
      }
    my $fh = open $!filename, :w, :bin;
    return Nil unless $fh.defined;

    $!filehandle  = $fh;
    $!internal-fh = 1;
  }


# Set colour palette.
  self.set-color-palette();

}


###############################################################################
#
# assemble-xml-file()
#
# Assemble and write the XML file.
#
method assemble-xml-file {

    # Prepare format object for passing to Style.pm.
    self.prepare-format-properties();

    self.xml-declaration();

    # Write the root workbook element.
    self.write-workbook();

    # Write the XLSX file version.
    self.write-file-version();

    # Write the workbook properties.
    self.write-workbook-pr();

    # Write the workbook view properties.
    self.write-book-views();

    # Write the worksheet names and ids.
    self.write-sheets();

    # Write the workbook defined names.
    self.write-defined-names();

    # Write the workbook calculation properties.
    self.write-calc-pr();

    # Write the workbook extension storage.
    # self.write-ext-lst();

    # Close the workbook tag.
    self.xml-end-tag( 'workbook' );

    # Close the XML writer filehandle.
    $!filehandle.close();
}


###############################################################################
#
# close()
#
# Calls finalization methods.
#
method close {

    # In case close() is called twice, by user and by DESTROY.
    return if $!fileclosed;

    # Test filehandle in case new() failed and the user didn't check.
    return Nil unless $!filehandle;

    $!fileclosed = 1;
    self.store-workbook;

    # Return the file close value.
    if $!internal-fh {
        return $!filehandle.close();
    }
    else {
        # Return true and let users deal with their own filehandles.
        return 1;
    }
}


###############################################################################
#
# DESTROY()
#
# Close the workbook if it hasn't already been explicitly closed.
#
method DESTROY {

#NYI     local ( $@, $!, $^E, $? );

    self.close() unless $!fileclosed;
}


#NYI ###############################################################################
#NYI #
#NYI # sheets(slice,...)
#NYI #
#NYI # An accessor for the worksheets[] array
#NYI #
#NYI # Returns: an optionally sliced list of the worksheet objects in a workbook.
#NYI #
#NYI sub sheets {

#NYI     my $self = shift;

#NYI     if ( @_ ) {

#NYI         # Return a slice of the array
#NYI         return @{ $self->{worksheets} }[@_];
#NYI     }
#NYI     else {

#NYI         # Return the entire list
#NYI         return @{ $self->{worksheets} };
#NYI     }
#NYI }


###############################################################################
#
# get-worksheet-by-name(name)
#
# Return a worksheet object in the workbook using the sheetname.
#
method get-worksheet-by-name($sheetname) {

    return Nil unless $sheetname.defined;

    return %!sheetnames{$sheetname};
}


#NYI ###############################################################################
#NYI #
#NYI # worksheets()
#NYI #
#NYI # An accessor for the worksheets[] array.
#NYI # This method is now deprecated. Use the sheets() method instead.
#NYI #
#NYI # Returns: an array reference
#NYI #
#NYI sub worksheets {

#NYI     my $self = shift;

#NYI     return $self->{worksheets};
#NYI }


###############################################################################
#
# add-worksheet($name)
#
# Add a new worksheet to the Excel workbook.
#
# Returns: reference to a worksheet object
#
method add-worksheet($name? is copy) {

    my $index = @!worksheets.elems;
    $name  = self.check-sheetname( $name );
    my $fh;

    # Porters take note, the following scheme of passing references to Workbook
    # data (in the \$self->{foo} cases) instead of a reference to the Workbook
    # itself is a workaround to avoid circular references between Workbook and
    # Worksheet objects. Feel free to implement this in any way the suits your
    # language.
    #
    my %init-data = (
        fh => $fh,
        name => $name,
        index =>$index,

        activesheet => $!activesheet,
        firstsheet => $!firstsheet,

        str-total => $!str-total,
        str-unique => $!str-unique,
        str-table => %!str-table,

        date1904 => $!date1904,
        palette => @!palette,
        optimization => $!optimization,
        tempdir => $!tempdir,
        excel2003-style => $!excel2003-style,

    );

    my $worksheet = Excel::Writer::XLSX::Worksheet.new( |%init-data );
dd $worksheet;
    @!worksheets[$index] = $worksheet;
    %!sheetnames{$name}  = $worksheet;
dd @!worksheets, %!sheetnames;

    return $worksheet;
}


#NYI ###############################################################################
#NYI #
#NYI # add-chart( %args )
#NYI #
#NYI # Create a chart for embedding or as a new sheet.
#NYI #
#NYI sub add-chart {

#NYI     my $self  = shift;
#NYI     my %arg   = @_;
#NYI     my $name  = '';
#NYI     my $index = @{ $self->{worksheets} };
#NYI     my $fh    = undef;

#NYI     # Type must be specified so we can create the required chart instance.
#NYI     my $type = $arg{type};
#NYI     if ( !defined $type ) {
#NYI         fail "Must define chart type in add-chart()";
#NYI     }

#NYI     # Ensure that the chart defaults to non embedded.
#NYI     my $embedded = $arg{embedded} || 0;

#NYI     # Check the worksheet name for non-embedded charts.
#NYI     if ( !$embedded ) {
#NYI         $name = $self->check-sheetname( $arg{name}, 1 );
#NYI     }


#NYI     my @init-data = (

#NYI         $fh,
#NYI         $name,
#NYI         $index,

#NYI         \$self->{activesheet},
#NYI         \$self->{firstsheet},

#NYI         \$self->{str-total},
#NYI         \$self->{str-unique},
#NYI         \$self->{str-table},

#NYI         $self->{date1904},
#NYI         $self->{palette},
#NYI         $self->{optimization},
#NYI     );


#NYI     my $chart = Excel::Writer::XLSX::Chart->factory( $type, $arg{subtype} );

#NYI     # If the chart isn't embedded let the workbook control it.
#NYI     if ( !$embedded ) {

#NYI         my $drawing    = Excel::Writer::XLSX::Drawing->new();
#NYI         my $chartsheet = Excel::Writer::XLSX::Chartsheet->new( @init-data );

#NYI         $chart->{palette} = $self->{palette};

#NYI         $chartsheet->{chart}   = $chart;
#NYI         $chartsheet->{drawing} = $drawing;

#NYI         $self->{worksheets}->[$index] = $chartsheet;
#NYI         $self->{sheetnames}->{$name} = $chartsheet;

#NYI         push @{ $self->{charts} }, $chart;

#NYI         return $chartsheet;
#NYI     }
#NYI     else {

#NYI         # Set the embedded chart name if present.
#NYI         $chart->{chart-name} = $arg{name} if $arg{name};

#NYI         # Set index to 0 so that the activate() and set-first-sheet() methods
#NYI         # point back to the first worksheet if used for embedded charts.
#NYI         $chart->{index}   = 0;
#NYI         $chart->{palette} = $self->{palette};
#NYI         $chart->set-embedded-config-data();
#NYI         push @{ $self->{charts} }, $chart;

#NYI         return $chart;
#NYI     }

#NYI }


###############################################################################
#
# check-sheetname( $name )
#
# Check for valid worksheet names. We check the length, if it contains any
# invalid characters and if the name is unique in the workbook.
#
method check-sheetname($name is copy = '', $chart = 0) {

    my $invalid-char = token { <[\[\]:*?/\\]> };

    # Increment the Sheet/Chart number used for default sheet names below.
    if $chart {
        $!chartname-count++;
    }
    else {
        $!sheetname-count++;
    }

    # Supply default Sheet/Chart name if none has been defined.
    if ! $name.defined or $name eq '' {

        if $chart {
            $name = $!chart-name ~ $!chartname-count;
        }
        else {
            $name = $!sheet-name ~ $!sheetname-count;
        }
    }

    # Check that sheet name is <= 31. Excel limit.
    fail "Sheetname $name must be <= 31 chars" if $name.chars > 31;

    # Check that sheetname doesn't contain any invalid characters
    if $name ~~ $invalid-char {
        fail 'Invalid character(' ~ $0 ~ ') in worksheet name: "' ~ $name ~ '" ([]:*?/\\ cannot be used)';
    }

    # Check that the worksheet name doesn't already exist since this is a fatal
    # error in Excel 97. The check must also exclude case insensitive matches.
    for @!worksheets -> $worksheet {
        my $name-a = $name;
        my $name-b = $worksheet.name;

        if $name-a.fc eq $name-b.fc {
            fail "Worksheet name '$name', with case ignored, is already used.";
        }
    }

    $name;
}


###############################################################################
#
# add-format(%properties)
#
# Add a new format to the Excel workbook.
#
method add-format(*%options) {

    my %init-data =
      ( |%!xf-format-indices, |%!dxf-format-indices );

    # Change default format style for Excel2003/XLS format.
    if $!excel2003-style {
        %init-data.append: ( font => 'Arial', size => 10, theme => -1 );
    }

    # Add the default format properties.
    %init-data.push: { %!default-format-properties};

    # Add the user defined properties.
    %init-data.push: %options;

    my $format = Excel::Writer::XLSX::Format.new( |%init-data );

    @!formats.push: $format;    # Store format reference

    return $format;
}


###############################################################################
#
# add-shape(%properties)
#
# Add a new shape to the Excel workbook.
#
method add-shape(*@options) {

    my $fh;
    my $shape = Excel::Writer::XLSX::Shape.new( $fh, @options );

    $shape.palette = self.palette;


    @!shapes.push: $shape;    # Store shape reference.

    $shape;
}

###############################################################################
#
# set1904()
#
# Set the date system: 0 = 1900 (the default), 1 = 1904
#
method set1904($value = 1) {

    $!date1904 = $value;
}


###############################################################################
#
# get1904()
#
# Return the date system: 0 = 1900, 1 = 1904
#
method get1904 {
  $!date1904;
}


###############################################################################
#
# set-custom-color()
#
# Change the RGB components of the elements in the colour palette.
#
multi method set-custom-color($index, $color) {
    if $color.defined and $color ~~ /^ '#' (\w\w) (\w\w) (\w\w) $ / {
        self.set-custom-color($index, $0, $1, $2);
    } else {
        fail 'illegal usage of set-custom-color';
    }
}

multi method set-custom-color($index, $red, $green?, $blue?) {

    my $aref = @!palette;

    # Check that the colour index is the right range
    if ( $index < 8 or $index > 64 ) {
        warn "Color index $index outside range: 8 <= index <= 64";
        return 0;
    }

    # Check that the colour components are in the right range
    unless 0 <= $red   <= 255
       and 0 <= $green <= 255
       and 0 <= $blue  <= 255
    {
        warn "Color component outside range: 0 <= color <= 255";
        return 0;
    }

    $index -= 8;    # Adjust colour index (wingless dragonfly)

    # Set the RGB value.
    my @rgb = ( $red, $green, $blue );
    $aref[$index] = @rgb;

    # Store the custom colors for the style.xml file.
    push @!custom-colors, sprintf "FF%02X%02X%02X", @rgb;

    $index + 8;
}


###############################################################################
#
# set-color-palette()
#
# Sets the colour palette to the Excel defaults.
#
method set-color-palette {

    @!palette = [
        [ 0x00, 0x00, 0x00, 0x00 ],    # 8
        [ 0xff, 0xff, 0xff, 0x00 ],    # 9
        [ 0xff, 0x00, 0x00, 0x00 ],    # 10
        [ 0x00, 0xff, 0x00, 0x00 ],    # 11
        [ 0x00, 0x00, 0xff, 0x00 ],    # 12
        [ 0xff, 0xff, 0x00, 0x00 ],    # 13
        [ 0xff, 0x00, 0xff, 0x00 ],    # 14
        [ 0x00, 0xff, 0xff, 0x00 ],    # 15
        [ 0x80, 0x00, 0x00, 0x00 ],    # 16
        [ 0x00, 0x80, 0x00, 0x00 ],    # 17
        [ 0x00, 0x00, 0x80, 0x00 ],    # 18
        [ 0x80, 0x80, 0x00, 0x00 ],    # 19
        [ 0x80, 0x00, 0x80, 0x00 ],    # 20
        [ 0x00, 0x80, 0x80, 0x00 ],    # 21
        [ 0xc0, 0xc0, 0xc0, 0x00 ],    # 22
        [ 0x80, 0x80, 0x80, 0x00 ],    # 23
        [ 0x99, 0x99, 0xff, 0x00 ],    # 24
        [ 0x99, 0x33, 0x66, 0x00 ],    # 25
        [ 0xff, 0xff, 0xcc, 0x00 ],    # 26
        [ 0xcc, 0xff, 0xff, 0x00 ],    # 27
        [ 0x66, 0x00, 0x66, 0x00 ],    # 28
        [ 0xff, 0x80, 0x80, 0x00 ],    # 29
        [ 0x00, 0x66, 0xcc, 0x00 ],    # 30
        [ 0xcc, 0xcc, 0xff, 0x00 ],    # 31
        [ 0x00, 0x00, 0x80, 0x00 ],    # 32
        [ 0xff, 0x00, 0xff, 0x00 ],    # 33
        [ 0xff, 0xff, 0x00, 0x00 ],    # 34
        [ 0x00, 0xff, 0xff, 0x00 ],    # 35
        [ 0x80, 0x00, 0x80, 0x00 ],    # 36
        [ 0x80, 0x00, 0x00, 0x00 ],    # 37
        [ 0x00, 0x80, 0x80, 0x00 ],    # 38
        [ 0x00, 0x00, 0xff, 0x00 ],    # 39
        [ 0x00, 0xcc, 0xff, 0x00 ],    # 40
        [ 0xcc, 0xff, 0xff, 0x00 ],    # 41
        [ 0xcc, 0xff, 0xcc, 0x00 ],    # 42
        [ 0xff, 0xff, 0x99, 0x00 ],    # 43
        [ 0x99, 0xcc, 0xff, 0x00 ],    # 44
        [ 0xff, 0x99, 0xcc, 0x00 ],    # 45
        [ 0xcc, 0x99, 0xff, 0x00 ],    # 46
        [ 0xff, 0xcc, 0x99, 0x00 ],    # 47
        [ 0x33, 0x66, 0xff, 0x00 ],    # 48
        [ 0x33, 0xcc, 0xcc, 0x00 ],    # 49
        [ 0x99, 0xcc, 0x00, 0x00 ],    # 50
        [ 0xff, 0xcc, 0x00, 0x00 ],    # 51
        [ 0xff, 0x99, 0x00, 0x00 ],    # 52
        [ 0xff, 0x66, 0x00, 0x00 ],    # 53
        [ 0x66, 0x66, 0x99, 0x00 ],    # 54
        [ 0x96, 0x96, 0x96, 0x00 ],    # 55
        [ 0x00, 0x33, 0x66, 0x00 ],    # 56
        [ 0x33, 0x99, 0x66, 0x00 ],    # 57
        [ 0x00, 0x33, 0x00, 0x00 ],    # 58
        [ 0x33, 0x33, 0x00, 0x00 ],    # 59
        [ 0x99, 0x33, 0x00, 0x00 ],    # 60
        [ 0x99, 0x33, 0x66, 0x00 ],    # 61
        [ 0x33, 0x33, 0x99, 0x00 ],    # 62
        [ 0x33, 0x33, 0x33, 0x00 ],    # 63
    ];

    0;
}

#NYI ###############################################################################
#NYI #
#NYI # define-name()
#NYI #
#NYI # Create a defined name in Excel. We handle global/workbook level names and
#NYI # local/worksheet names.
#NYI #
#NYI sub define-name {

#NYI     my $self        = shift;
#NYI     my $name        = shift;
#NYI     my $formula     = shift;
#NYI     my $sheet-index = undef;
#NYI     my $sheetname   = '';
#NYI     my $full-name   = $name;

#NYI     # Remove the = sign from the formula if it exists.
#NYI     $formula =~ s/^=//;

#NYI     # Local defined names are formatted like "Sheet1!name".
#NYI     if ( $name =~ /^(.*)!(.*)$/ ) {
#NYI         $sheetname   = $1;
#NYI         $name        = $2;
#NYI         $sheet-index = $self->get-sheet-index( $sheetname );
#NYI     }
#NYI     else {
#NYI         $sheet-index = -1;    # Use -1 to indicate global names.
#NYI     }

#NYI     # Warn if the sheet index wasn't found.
#NYI     if ( !defined $sheet-index ) {
#NYI         warn "Unknown sheet name $sheetname in defined-name()";
#NYI         return -1;
#NYI     }

#NYI     # Warn if the name contains invalid chars as defined by Excel help.
#NYI     if ( $name !~ m/^[\w\\][\w\\.]*$/ || $name =~ m/^\d/ ) {
#NYI         warn "Invalid character in name '$name' used in defined-name()";
#NYI         return -1;
#NYI     }

#NYI     # Warn if the name looks like a cell name.
#NYI     if ( $name =~ m/^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$/ ) {
#NYI         warn "Invalid name '$name' looks like a cell name in defined-name()";
#NYI         return -1;
#NYI     }

#NYI     # Warn if the name looks like a R1C1.
#NYI     if ( $name =~ m/^[rcRC]$/ || $name =~ m/^[rcRC]\d+[rcRC]\d+$/ ) {
#NYI         warn "Invalid name '$name' like a RC cell ref in defined-name()";
#NYI         return -1;
#NYI     }

#NYI     push @{ $self->{defined-names} }, [ $name, $sheet-index, $formula ];
#NYI }


###############################################################################
#
# set-size()
#
# Set the workbook size.
#
method set-size($width = 1073, $height = 644) {

    $width  ||= 1073;
    $height ||=  644;
    # Convert to twips at 96 dpi.
    $!window-width  = int( $width  * 1440 / 96 );
    $!window-height = int( $height * 1440 / 96 );
}


###############################################################################
#
# set-properties()
#
# Set the document properties such as Title, Author etc. These are written to
# property sets in the OLE container.
#
method set-properties(*%param) {

    # Ignore if no args were passed.
    return -1 unless %param;

    # List of valid input parameters.
    my %valid = (
        title          => 1,
        subject        => 1,
        author         => 1,
        keywords       => 1,
        comments       => 1,
        last-author    => 1,
        created        => 1,
        category       => 1,
        manager        => 1,
        company        => 1,
        status         => 1,
        hyperlink-base => 1,
    );

    # Check for valid input parameters.
    for %param.keys -> $parameter {
        if ( not %valid{$parameter}.defined ) {
            warn "Unknown parameter '$parameter' in set-properties()";
            return -1;
        }
    }

    # Set the creation time unless specified by the user.
    %param<created> //= @!createtime;

    %!doc-properties = %param;
}


###############################################################################
#
# set-custom-property()
#
# Set a user defined custom document property.
#
method set-custom-property($name, $value, $type?) {

    # Valid types.
    my %valid-type = (
        'text'       => 1,
        'date'       => 1,
        'number'     => 1,
        'number-int' => 1,
        'bool'       => 1,
    );

    if ! $name.defined || ! $value.defined {
        warn "The name and value parameters must be defined "
          ~ "in set-custom-property()";

        return -1;
    }

    # Determine the type for strings and numbers if it hasn't been specified.
    if !$type {
        if $value ~~ /^\d+$/ {
            $type = 'number-int';
        }
        elsif $value ~~
            /^
             <[+-]>? <before \d|\.\d>
             \d*
             [\.\d*]?
             [
               [<[Ee]>
                 [
                   <[+-]>?
                   \d+
                 ]?
               ]
             ]?
            $/
        {
            $type = 'number';
        }
        else {
            $type = 'text';
        }
    }

    # Check for valid validation types.
    if ! %valid-type{$type}.exists {
        warn "Unknown custom type '$type' in set-custom-property()";
        return -1;
    }

    #  Check for strings longer than Excel's limit of 255 chars.
    if $type eq 'text' and $value.chars > 255 {
        warn "Length of text custom value '$value' exceeds "
          ~ "Excel's limit of 255 in set-custom-property()";
        return -1;
    }
    if $name.chars > 255 {
        warn "Length of custom name '$name' exceeds "
          ~ "Excel's limit of 255 in set-custom-property()";
        return -1;
    }

    push @!custom-properties, [ $name, $value, $type ];
}



#NYI ###############################################################################
#NYI #
#NYI # add-vba-project()
#NYI #
#NYI # Add a vbaProject binary to the XLSX file.
#NYI #
#NYI sub add-vba-project {

#NYI     my $self        = shift;
#NYI     my $vba-project = shift;

#NYI     fail "No vbaProject.bin specified in add-vba-project()"
#NYI       if not $vba-project;

#NYI     fail "Couldn't locate $vba-project in add-vba-project(): $!"
#NYI       unless -e $vba-project;

#NYI     $self->{vba-project} = $vba-project;
#NYI }


#NYI ###############################################################################
#NYI #
#NYI # set-vba-name()
#NYI #
#NYI # Set the VBA name for the workbook.
#NYI #
#NYI sub set-vba-name {

#NYI     my $self         = shift;
#NYI     my $vba-codemame = shift;

#NYI     if ( $vba-codemame ) {
#NYI         $self->{vba-codename} = $vba-codemame;
#NYI     }
#NYI     else {
#NYI         $self->{vba-codename} = 'ThisWorkbook';
#NYI     }
#NYI }


###############################################################################
#
# set-calc-mode()
#
# Set the Excel calculation mode for the workbook.
#
method set-calc-mode($mode = 'auto', $calc-id?) {

    $!calc-mode = $mode;

    if $mode eq 'manual' {
        $!calc-mode    = 'manual';
        $!calc-on-load = 0;
    }
    elsif $mode eq 'auto-except-tables' {
        $!calc-mode = 'autoNoTable';
    }

    $!calc-id = $calc-id if $calc-id.defined;
}


###############################################################################
#
# store-workbook()
#
# Assemble worksheets into a workbook.
#
method store-workbook {

    my $tempdir  = File::Temp.newdir( tempdir => $!tempdir );
    my $packager = Excel::Writer::XLSX::Package::Packager.new();
    my $zip      = Archive::SimpleZip.new();


    # Add a default worksheet if none have been added.
    self.add-worksheet() if not @!worksheets;

    # Ensure that at least one worksheet has been selected.
    if ( $!activesheet == 0 ) {
        @!worksheets[0]<selected> = 1;
        @!worksheets[0]<hidden>   = 0;
    }

    # Set the active sheet.
    for @!worksheets -> $sheet {
        $sheet<active> = 1 if $sheet<index> == $!activesheet;
    }

    # Convert the SST strings data structure.
    self.prepare-sst-string-data();

    # Prepare the worksheet VML elements such as comments and buttons.
    self.prepare-vml-objects();

    # Set the defined names for the worksheets such as Print Titles.
    self.prepare-defined-names();

    # Prepare the drawings, charts and images.
    self.prepare-drawings();

    # Add cached data to charts.
    self.add-chart-data();

    # Prepare the worksheet tables.
    self.prepare-tables();

    # Package the workbook.
    $packager.add-workbook();
    $packager.set-package-dir( $tempdir );
    $packager.create-package();

    # Free up the Packager object.
    $packager = Nil;

    # Add the files to the zip archive. Due to issues with Archive::Zip in
    # taint mode we can't use addTree() so we have to build the file list
    # with File::Find and pass each one to addFile().
    my $xlsx-files = File::Find::find( dir => $tempdir, type => 'file' );

    # Store the xlsx component files with the temp dir name removed.
    for $xlsx-files -> $filename {
        my $short-name = $filename;
        $short-name ~~ s/^$tempdir '/'?//;
        $zip.addFile( $filename, $short-name );
    }


    if $!internal-fh {

        if $zip.writeToFileHandle( $!filehandle ) != 0 {
            warn 'Error writing zip container for xlsx file.';
        }
    }
    else {

        # Archive::Zip needs to rewind a filehandle to write the zip headers.
        # This won't work for arbitrary user defined filehandles so we use
        # a temp file based filehandle to create the zip archive and then
        # stream that to the filehandle.
        my $tmp-fh = tempfile( tempdir => $!tempdir );
        my $is-seekable = 1;

        if $zip.writeToFileHandle( $tmp-fh, $is-seekable ) != 0 {
            warn 'Error writing zip container for xlsx file.';
        }

        my $buffer;
        $tmp-fh.seek: 0, 0;

        while $tmp-fh.read: $buffer, 4_096 {
            # local $\ = undef;    # Protect print from -l on commandline.
            $!filehandle.print: $buffer;
        }
    }
}


###############################################################################
#
# prepare-sst-string-data()
#
# Convert the SST string data from a hash to an array.
#
#NYI sub prepare-sst-string-data {

#NYI     my $self = shift;

#NYI     my @strings;
#NYI     $#strings = $self->{str-unique} - 1;    # Pre-extend array

#NYI     while ( my $key = each %{ $self->{str-table} } ) {
#NYI         $strings[ $self->{str-table}->{$key} ] = $key;
#NYI     }

#NYI     # The SST data could be very large, free some memory (maybe).
#NYI     $self->{str-table} = undef;
#NYI     $self->{str-array} = \@strings;

#NYI }


###############################################################################
#
# prepare-format-properties()
#
# Prepare all of the format properties prior to passing them to Styles.pm.
#
method prepare-format-properties {

    # Separate format objects into XF and DXF formats.
    self.prepare-formats();

    # Set the font index for the format objects.
    self.prepare-fonts();

    # Set the number format index for the format objects.
    self.prepare-num-formats();

    # Set the border index for the format objects.
    self.prepare-borders();

    # Set the fill index for the format objects.
    self.prepare-fills();


}


###############################################################################
#
# prepare-formats()
#
# Iterate through the XF Format objects and separate them into XF and DXF
# formats.
#
method prepare-formats {

    for @!formats -> $format {
        my $xf-index  = $format<xf-index>;
        my $dxf-index = $format<dxf-index>;

        if $xf-index.defined {
            @!xf-formats[$xf-index] = $format;
        }

        if $dxf-index.defined {
            @!dxf-formats[$dxf-index] = $format;
        }
    }
}


###############################################################################
#
# set-default-xf-indices()
#
# Set the default index for each format. This is mainly used for testing.
#
#NYI sub set-default-xf-indices {

#NYI     my $self = shift;

#NYI     for my $format ( @{ $self->{formats} } ) {
#NYI         $format->get-xf-index();
#NYI     }
#NYI }


###############################################################################
#
# prepare-fonts()
#
# Iterate through the XF Format objects and give them an index to non-default
# font elements.
#
method prepare-fonts {

    my %fonts;
    my $index = 0;

    for @!xf-formats -> $format {
        my $key = $format.get-font-key();

        if %fonts{$key}.exists {

            # Font has already been used.
            $format.font-index: %fonts{$key};
            $format.has-font:   0;
        }
        else {

            # This is a new font.
            %fonts{$key}        = $index;
            $format.font-index:   $index;
            $format.has-font:     1;
            $index++;
        }
    }

    $!font-count = $index;

    # For the DXF formats we only need to check if the properties have changed.
    for @!dxf-formats -> $format {

        # The only font properties that can change for a DXF format are: color,
        # bold, italic, underline and strikethrough.
        if    $format.color
           || $format.bold
           || $format.italic
           || $format.underline
           || $format.font-strikeout
        {
            $format.has-dxf-font: 1;
        }
    }
}


###############################################################################
#
# prepare-num-formats()
#
# Iterate through the XF Format objects and give them an index to non-default
# number format elements.
#
# User defined records start from index 0xA4.
#
method prepare-num-formats {

    my %num-formats;
    my $index            = 164;
    my $num-format-count = 0;

    for |@!xf-formats, |@!dxf-formats -> $format {
        my $num-format = $format.num-format;

        # Check if $num-format is an index to a built-in number format.
        # Also check for a string of zeros, which is a valid number format
        # string but would evaluate to zero.
        #
        if $num-format ~~ m/^\d+$/ && $num-format !~~ m/^0+\d/ {

            # Index to a built-in number format.
            $format.num-format-index: $num-format;
            next;
        }


        if %num-formats{$num-format}.exists {

            # Number format has already been used.
            $format.num-format-index: %num-formats{$num-format};
        }
        else {

            # Add a new number format.
            %num-formats{$num-format} = $index;
            $format.num-format-index:   $index;
            $index++;

            # Only increase font count for XF formats (not for DXF formats).
            if $format.xf-index {
                $num-format-count++;
            }
        }
    }

    $!num-format-count = $num-format-count;
}


###############################################################################
#
# prepare-borders()
#
# Iterate through the XF Format objects and give them an index to non-default
# border elements.
#
method prepare-borders {

    my %borders;
    my $index = 0;

    for @!xf-formats -> $format {
        my $key = $format.get-border-key();

        if %borders{$key}.exists {

            # Border has already been used.
            $format.border-index: %borders{$key};
            $format.has-border:   0;
        }
        else {

            # This is a new border.
            %borders{$key}       = $index;
            $format.border-index:  $index;
            $format.has-border:    1;
            $index++;
        }
    }

 #TODO   $!border-count = $index;

    # For the DXF formats we only need to check if the properties have changed.
    for @!dxf-formats -> $format {
        my $key = $format.get-border-key();

        if $key ~~ m/<-[0:]>/ {
            $format.has-dxf-border: 1;
        }
    }

}


###############################################################################
#
# prepare-fills()
#
# Iterate through the XF Format objects and give them an index to non-default
# fill elements.
#
# The user defined fill properties start from 2 since there are 2 default
# fills: patternType="none" and patternType="gray125".
#
method prepare-fills {

    my %fills;
    my $index = 2;    # Start from 2. See above.

    # Add the default fills.
    %fills{'0:0:0'}  = 0;
    %fills{'17:0:0'} = 1;


    # Store the DXF colours separately since them may be reversed below.
    for @!dxf-formats -> $format {
        if    $format.pattern
           || $format.bg-color
           || $format.fg-color
        {
            $format.has-dxf-fill: 1;
            $format.dxf-bg-color: $format.bg-color;
            $format.dxf-fg-color: $format.fg-color;
        }
    }


    for @!xf-formats -> $format {

        # The following logical statements jointly take care of special cases
        # in relation to cell colours and patterns:
        # 1. For a solid fill (pattern == 1) Excel reverses the role of
        #    foreground and background colours, and
        # 2. If the user specifies a foreground or background colour without
        #    a pattern they probably wanted a solid fill, so we fill in the
        #    defaults.
        #
        if    $format.pattern  == 1
           && $format.bg-color ne '0'
           && $format.fg-color ne '0'
        {
            my $tmp = $format.fg-color;
            $format.fg-color: $format.bg-color;
            $format.bg-color: $tmp;
        }

        if    $format.pattern  <= 1
           && $format.bg-color ne '0'
           && $format.fg-color eq '0'
        {
            $format.fg-color: $format.bg-color;
            $format.bg-color: 0;
            $format.pattern:  1;
        }

        if    $format.pattern  <= 1
           && $format.bg-color eq '0'
           && $format.fg-color ne '0'
        {
            $format.bg-color: 0;
            $format.pattern:  1;
        }


        my $key = $format.get-fill-key();

        if %fills{$key}.exists {

            # Fill has already been used.
            $format.fill-index: %fills{$key};
            $format.has-fill:   0;
        }
        else {

            # This is a new fill.
            %fills{$key}        = $index;
            $format.fill-index:   $index;
            $format.has-fill:     1;
            $index++;
        }
    }

#TODO    $!fill-count = $index;
}


###############################################################################
#
# prepare-defined-names()
#
# Iterate through the worksheets and store any defined names in addition to
# any user defined names. Stores the defined names for the Workbook.xml and
# the named ranges for App.xml.
#
#NYI sub prepare-defined-names {

#NYI     my $self = shift;

#NYI     my @defined-names = @{ $self->{defined-names} };

#NYI     for my $sheet ( @{ $self->{worksheets} } ) {

#NYI         # Check for Print Area settings.
#NYI         if ( $sheet->{autofilter} ) {

#NYI             my $range  = $sheet->{autofilter};
#NYI             my $hidden = 1;

#NYI             # Store the defined names.
#NYI             push @defined-names,
#NYI               [ '_xlnm._FilterDatabase', $sheet->{index}, $range, $hidden ];

#NYI         }

#NYI         # Check for Print Area settings.
#NYI         if ( $sheet->{print-area} ) {

#NYI             my $range = $sheet->{print-area};

#NYI             # Store the defined names.
#NYI             push @defined-names,
#NYI               [ '_xlnm.Print_Area', $sheet->{index}, $range ];
#NYI         }

#NYI         # Check for repeat rows/cols. aka, Print Titles.
#NYI         if ( $sheet->{repeat-cols} || $sheet->{repeat-rows} ) {
#NYI             my $range = '';

#NYI             if ( $sheet->{repeat-cols} && $sheet->{repeat-rows} ) {
#NYI                 $range = $sheet->{repeat-cols} . ',' . $sheet->{repeat-rows};
#NYI             }
#NYI             else {
#NYI                 $range = $sheet->{repeat-cols} . $sheet->{repeat-rows};
#NYI             }

#NYI             # Store the defined names.
#NYI             push @defined-names,
#NYI               [ '_xlnm.Print_Titles', $sheet->{index}, $range ];
#NYI         }

#NYI     }

#NYI     @defined-names          = sort-defined-names( @defined-names );
#NYI     $self->{defined-names} = \@defined-names;
#NYI     $self->{named-ranges}  = extract-named-ranges( @defined-names );
#NYI }


###############################################################################
#
# sort-defined-names()
#
# Sort internal and user defined names in the same order as used by Excel.
# This may not be strictly necessary but unsorted elements caused a lot of
# issues in the Spreadsheet::WriteExcel binary version. Also makes
# comparison testing easier.
#
#NYI sub sort-defined-names {

#NYI     my @names = @_;

#NYI     #<<< Perltidy ignore this.

#NYI     @names = sort {
#NYI         # Primary sort based on the defined name.
#NYI         -normalise-defined-name( $a->[0] )
#NYI         cmp
#NYI         -normalise-defined-name( $b->[0] )

#NYI         ||
#NYI         # Secondary sort based on the sheet name.
#NYI         -normalise-sheet-name( $a->[2] )
#NYI         cmp
#NYI         normalise-sheet-name( $b->[2] )

#NYI     } @names;
#NYI     #>>>

#NYI     return @names;
#NYI }

# Used in the above sort routine to normalise the defined names. Removes any
# leading 'xmln.' from internal names and lowercases the strings.
#NYI sub normalise-defined-name {
#NYI     my $name = shift;

#NYI     $name =~ s/^_xlnm.//;
#NYI     $name = lc $name;

#NYI     return $name;
#NYI }

# Used in the above sort routine to normalise the worksheet names for the
# secondary sort. Removes leading quote and lowercases the strings.
#NYI sub normalise-sheet-name {
#NYI     my $name = shift;

#NYI     $name =~ s/^'//;
#NYI     $name = lc $name;

#NYI     return $name;
#NYI }


###############################################################################
#
# extract-named-ranges()
#
# Extract the named ranges from the sorted list of defined names. These are
# used in the App.xml file.
#
#NYI sub extract-named-ranges {

#NYI     my @defined-names = @_;
#NYI     my @named-ranges;

#NYI     NAME:
#NYI     for my $defined-name ( @defined-names ) {

#NYI         my $name  = $defined-name->[0];
#NYI         my $index = $defined-name->[1];
#NYI         my $range = $defined-name->[2];

#NYI         # Skip autoFilter ranges.
#NYI         next NAME if $name eq '_xlnm._FilterDatabase';

#NYI         # We are only interested in defined names with ranges.
#NYI         if ( $range =~ /^([^!]+)!/ ) {
#NYI             my $sheet-name = $1;

#NYI             # Match Print_Area and Print_Titles xlnm types.
#NYI             if ( $name =~ /^_xlnm\.(.*)$/ ) {
#NYI                 my $xlnm-type = $1;
#NYI                 $name = $sheet-name . '!' . $xlnm-type;
#NYI             }
#NYI             elsif ( $index != -1 ) {
#NYI                 $name = $sheet-name . '!' . $name;
#NYI             }

#NYI             push @named-ranges, $name;
#NYI         }
#NYI     }

#NYI     return \@named-ranges;
#NYI }


###############################################################################
#
# prepare-drawings()
#
# Iterate through the worksheets and set up any chart or image drawings.
#
#NYI sub prepare-drawings {

#NYI     my $self         = shift;
#NYI     my $chart-ref-id = 0;
#NYI     my $image-ref-id = 0;
#NYI     my $drawing-id   = 0;

#NYI     for my $sheet ( @{ $self->{worksheets} } ) {

#NYI         my $chart-count = scalar @{ $sheet->{charts} };
#NYI         my $image-count = scalar @{ $sheet->{images} };
#NYI         my $shape-count = scalar @{ $sheet->{shapes} };

#NYI         my $header-image-count = scalar @{ $sheet->{header-images} };
#NYI         my $footer-image-count = scalar @{ $sheet->{footer-images} };
#NYI         my $has-drawing        = 0;


#NYI         # Check that some image or drawing needs to be processed.
#NYI         if (   !$chart-count
#NYI             && !$image-count
#NYI             && !$shape-count
#NYI             && !$header-image-count
#NYI             && !$footer-image-count )
#NYI         {
#NYI             next;
#NYI         }

#NYI         # Don't increase the drawing-id header/footer images.
#NYI         if ( $chart-count || $image-count || $shape-count ) {
#NYI             $drawing-id++;
#NYI             $has-drawing = 1;
#NYI         }

#NYI         # Prepare the worksheet charts.
#NYI         for my $index ( 0 .. $chart-count - 1 ) {
#NYI             $chart-ref-id++;
#NYI             $sheet->prepare-chart( $index, $chart-ref-id, $drawing-id );
#NYI         }

#NYI         # Prepare the worksheet images.
#NYI         for my $index ( 0 .. $image-count - 1 ) {

#NYI             my $filename = $sheet->{images}->[$index]->[2];

#NYI             my ( $type, $width, $height, $name, $x-dpi, $y-dpi ) =
#NYI               $self->get-image-properties( $filename );

#NYI             $image-ref-id++;

#NYI             $sheet->prepare-image(
#NYI                 $index, $image-ref-id, $drawing-id,
#NYI                 $width, $height,       $name,
#NYI                 $type,  $x-dpi,        $y-dpi
#NYI             );
#NYI         }

#NYI         # Prepare the worksheet shapes.
#NYI         for ^$shape-count -> $index {
#NYI             $sheet->prepare-shape( $index, $drawing-id );
#NYI         }

#NYI         # Prepare the header images.
#NYI         for my $index ( 0 .. $header-image-count - 1 ) {

#NYI             my $filename = $sheet->{header-images}->[$index]->[0];
#NYI             my $position = $sheet->{header-images}->[$index]->[1];

#NYI             my ( $type, $width, $height, $name, $x-dpi, $y-dpi ) =
#NYI               $self->get-image-properties( $filename );

#NYI             $image-ref-id++;

#NYI             $sheet->prepare-header-image( $image-ref-id, $width, $height,
#NYI                 $name, $type, $position, $x-dpi, $y-dpi );
#NYI         }

#NYI         # Prepare the footer images.
#NYI         for my $index ( 0 .. $footer-image-count - 1 ) {

#NYI             my $filename = $sheet->{footer-images}->[$index]->[0];
#NYI             my $position = $sheet->{footer-images}->[$index]->[1];

#NYI             my ( $type, $width, $height, $name, $x-dpi, $y-dpi ) =
#NYI               $self->get-image-properties( $filename );

#NYI             $image-ref-id++;

#NYI             $sheet->prepare-header-image( $image-ref-id, $width, $height,
#NYI                 $name, $type, $position, $x-dpi, $y-dpi );
#NYI         }


#NYI         if ( $has-drawing ) {
#NYI             my $drawing = $sheet->{drawing};
#NYI             push @{ $self->{drawings} }, $drawing;
#NYI         }
#NYI     }


#NYI     # Remove charts that were created but not inserted into worksheets.
#NYI     my @chart-data;

#NYI     for my $chart ( @{ $self->{charts} } ) {
#NYI         if ( $chart->{id} != -1 ) {
#NYI             push @chart-data, $chart;
#NYI         }
#NYI     }

#NYI     # Sort the workbook charts references into the order that the were
#NYI     # written from the worksheets above.
#NYI     @chart-data = sort { $a->{id} <=> $b->{id} } @chart-data;

#NYI     $self->{charts} = \@chart-data;
#NYI     $self->{drawing-count} = $drawing-id;
#NYI }


###############################################################################
#
# prepare-vml-objects()
#
# Iterate through the worksheets and set up the VML objects.
#
#NYI sub prepare-vml-objects {

#NYI     my $self           = shift;
#NYI     my $comment-id     = 0;
#NYI     my $vml-drawing-id = 0;
#NYI     my $vml-data-id    = 1;
#NYI     my $vml-header-id  = 0;
#NYI     my $vml-shape-id   = 1024;
#NYI     my $vml-files      = 0;
#NYI     my $comment-files  = 0;
#NYI     my $has-button     = 0;

#NYI     for my $sheet ( @{ $self->{worksheets} } ) {

#NYI         next if !$sheet->{has-vml} and !$sheet->{has-header-vml};
#NYI         $vml-files = 1;


#NYI         if ( $sheet->{has-vml} ) {

#NYI             $comment-files++ if $sheet->{has-comments};
#NYI             $comment-id++    if $sheet->{has-comments};
#NYI             $vml-drawing-id++;

#NYI             my $count =
#NYI               $sheet->prepare-vml-objects( $vml-data-id, $vml-shape-id,
#NYI                 $vml-drawing-id, $comment-id );

#NYI             # Each VML file should start with a shape id incremented by 1024.
#NYI             $vml-data-id  += 1 * int(    ( 1024 + $count ) / 1024 );
#NYI             $vml-shape-id += 1024 * int( ( 1024 + $count ) / 1024 );

#NYI         }

#NYI         if ( $sheet->{has-header-vml} ) {
#NYI             $vml-header-id++;
#NYI             $vml-drawing-id++;
#NYI             $sheet->prepare-header-vml-objects( $vml-header-id,
#NYI                 $vml-drawing-id );
#NYI         }

#NYI         # Set the sheet vba-codename if it has a button and the workbook
#NYI         # has a vbaProject binary.
#NYI         if ( $sheet->{buttons-array} ) {
#NYI             $has-button = 1;

#NYI             if ( $self->{vba-project} && !$sheet->{vba-codename} ) {
#NYI                 $sheet->set-vba-name();
#NYI             }
#NYI         }

#NYI     }

#NYI     $self->{num-vml-files}     = $vml-files;
#NYI     $self->{num-comment-files} = $comment-files;

#NYI     # Add a font format for cell comments.
#NYI     if ( $comment-files > 0 ) {
#NYI         my $format = Excel::Writer::XLSX::Format->new(
#NYI             \$self->{xf-format-indices},
#NYI             \$self->{dxf-format-indices},
#NYI             font          => 'Tahoma',
#NYI             size          => 8,
#NYI             color-indexed => 81,
#NYI             font-only     => 1,
#NYI         );

#NYI         $format->get-xf-index();

#NYI         push @{ $self->{formats} }, $format;
#NYI     }

#NYI     # Set the workbook vba-codename if one of the sheets has a button and
#NYI     # the workbook has a vbaProject binary.
#NYI     if ( $has-button && $self->{vba-project} && !$self->{vba-codename} ) {
#NYI         $self->set-vba-name();
#NYI     }
#NYI }


###############################################################################
#
# prepare-tables()
#
# Set the table ids for the worksheet tables.
#
#NYI sub prepare-tables {

#NYI     my $self     = shift;
#NYI     my $table-id = 0;
#NYI     my $seen     = {};

#NYI     for my $sheet ( @{ $self->{worksheets} } ) {

#NYI         my $table-count = scalar @{ $sheet->{tables} };

#NYI         next unless $table-count;

#NYI         $sheet->prepare-tables( $table-id + 1, $seen );

#NYI         $table-id += $table-count;
#NYI     }
#NYI }


###############################################################################
#
# add-chart-data()
#
# Add "cached" data to charts to provide the numCache and strCache data for
# series and title/axis ranges.
#
#NYI sub add-chart-data {

#NYI     my $self = shift;
#NYI     my %worksheets;
#NYI     my %seen-ranges;
#NYI     my @charts;

#NYI     # Map worksheet names to worksheet objects.
#NYI     for my $worksheet ( @{ $self->{worksheets} } ) {
#NYI         $worksheets{ $worksheet->{name} } = $worksheet;
#NYI     }

#NYI     # Build an array of the worksheet charts including any combined charts.
#NYI     for my $chart ( @{ $self->{charts} } ) {
#NYI         push @charts, $chart;

#NYI         if ($chart->{combined}) {
#NYI             push @charts, $chart->{combined};
#NYI         }
#NYI     }


#NYI     CHART:
#NYI     for my $chart ( @charts ) {

#NYI         RANGE:
#NYI         while ( my ( $range, $id ) = each %{ $chart->{formula-ids} } ) {

#NYI             # Skip if the series has user defined data.
#NYI             if ( defined $chart->{formula-data}->[$id] ) {
#NYI                 if (   !exists $seen-ranges{$range}
#NYI                     || !defined $seen-ranges{$range} )
#NYI                 {
#NYI                     my $data = $chart->{formula-data}->[$id];
#NYI                     $seen-ranges{$range} = $data;
#NYI                 }
#NYI                 next RANGE;
#NYI             }

#NYI             # Check to see if the data is already cached locally.
#NYI             if ( exists $seen-ranges{$range} ) {
#NYI                 $chart->{formula-data}->[$id] = $seen-ranges{$range};
#NYI                 next RANGE;
#NYI             }

#NYI             # Convert the range formula to a sheet name and cell range.
#NYI             my ( $sheetname, @cells ) = $self->get-chart-range( $range );

#NYI             # Skip if we couldn't parse the formula.
#NYI             next RANGE if !defined $sheetname;

#NYI             # Handle non-contiguous ranges: (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5).
#NYI             # We don't try to parse the ranges. We just return an empty list.
#NYI             if ( $sheetname =~ m/^\([^,]+,/ ) {
#NYI                 $chart->{formula-data}->[$id] = [];
#NYI                 $seen-ranges{$range} = [];
#NYI                 next RANGE;
#NYI             }

#NYI             # Die if the name is unknown since it indicates a user error in
#NYI             # a chart series formula.
#NYI             if ( !exists $worksheets{$sheetname} ) {
#NYI                 die "Unknown worksheet reference '$sheetname' in range "
#NYI                   . "'$range' passed to add-series().\n";
#NYI             }

#NYI             # Find the worksheet object based on the sheet name.
#NYI             my $worksheet = $worksheets{$sheetname};

#NYI             # Get the data from the worksheet table.
#NYI             my @data = $worksheet->get-range-data( @cells );

#NYI             # Convert shared string indexes to strings.
#NYI             for my $token ( @data ) {
#NYI                 if ( ref $token ) {
#NYI                     $token = $self->{str-array}->[ $token->{sst-id} ];

#NYI                     # Ignore rich strings for now. Deparse later if necessary.
#NYI                     if ( $token =~ m{^<r>} && $token =~ m{</r>$} ) {
#NYI                         $token = '';
#NYI                     }
#NYI                 }
#NYI             }

#NYI             # Add the data to the chart.
#NYI             $chart->{formula-data}->[$id] = \@data;

#NYI             # Store range data locally to avoid lookup if seen again.
#NYI             $seen-ranges{$range} = \@data;
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# get-chart-range()
#
# Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name and cell
# range such as ( 'Sheet1', 0, 1, 4, 1 ).
#
#NYI sub get-chart-range {

#NYI     my $self  = shift;
#NYI     my $range = shift;
#NYI     my $cell_1;
#NYI     my $cell_2;
#NYI     my $sheetname;
#NYI     my $cells;

#NYI     # Split the range formula into sheetname and cells at the last '!'.
#NYI     my $pos = rindex $range, '!';
#NYI     if ( $pos > 0 ) {
#NYI         $sheetname = substr $range, 0, $pos;
#NYI         $cells = substr $range, $pos + 1;
#NYI     }
#NYI     else {
#NYI         return undef;
#NYI     }

#NYI     # Split the cell range into 2 cells or else use single cell for both.
#NYI     if ( $cells =~ ':' ) {
#NYI         ( $cell_1, $cell_2 ) = split /:/, $cells;
#NYI     }
#NYI     else {
#NYI         ( $cell_1, $cell_2 ) = ( $cells, $cells );
#NYI     }

#NYI     # Remove leading/trailing apostrophes and convert escaped quotes to single.
#NYI     $sheetname =~ s/^'//g;
#NYI     $sheetname =~ s/'$//g;
#NYI     $sheetname =~ s/''/'/g;

#NYI     my ( $row-start, $col-start ) = xl-cell-to-rowcol( $cell_1 );
#NYI     my ( $row-end,   $col-end )   = xl-cell-to-rowcol( $cell_2 );

#NYI     # Check that we have a 1D range only.
#NYI     if ( $row-start != $row-end && $col-start != $col-end ) {
#NYI         return undef;
#NYI     }

#NYI     return ( $sheetname, $row-start, $col-start, $row-end, $col-end );
#NYI }


###############################################################################
#
# store-externs()
#
# Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
# the NAME records.
#
#NYI sub store-externs {

#NYI     my $self = shift;

#NYI }


###############################################################################
#
# store-names()
#
# Write the NAME record to define the print area and the repeat rows and cols.
#
#NYI sub store-names {

#NYI     my $self = shift;

#NYI }


###############################################################################
#
# quote-sheetname()
#
# Sheetnames used in references should be quoted if they contain any spaces,
# special characters or if the look like something that isn't a sheet name.
# TODO. We need to handle more special cases.
#
#NYI sub quote-sheetname {

#NYI     my $self      = shift;
#NYI     my $sheetname = $_[0];

#NYI     if ( $sheetname =~ /^Sheet\d+$/ ) {
#NYI         return $sheetname;
#NYI     }
#NYI     else {
#NYI         return qq('$sheetname');
#NYI     }
#NYI }


###############################################################################
#
# get-image-properties()
#
# Extract information from the image file such as dimension, type, filename,
# and extension. Also keep track of previously seen images to optimise out
# any duplicates.
#
#NYI sub get-image-properties {

#NYI     my $self     = shift;
#NYI     my $filename = shift;

#NYI     my $type;
#NYI     my $width;
#NYI     my $height;
#NYI     my $x-dpi = 96;
#NYI     my $y-dpi = 96;
#NYI     my $image-name;


#NYI     ( $image-name ) = fileparse( $filename );

#NYI     # Open the image file and import the data.
#NYI     my $fh = FileHandle->new( $filename );
#NYI     fail "Couldn't import $filename: $!" unless defined $fh;
#NYI     binmode $fh;

#NYI     # Slurp the file into a string and do some size calcs.
#NYI     my $data = do { local $/; <$fh> };
#NYI     my $size = length $data;


#NYI     if ( unpack( 'x A3', $data ) eq 'PNG' ) {

#NYI         # Test for PNGs.
#NYI         ( $type, $width, $height, $x-dpi, $y-dpi ) =
#NYI           $self->process-png( $data, $filename );

#NYI         $self->{image-types}->{png} = 1;
#NYI     }
#NYI     elsif ( unpack( 'n', $data ) == 0xFFD8 ) {

#NYI         # Test for JPEG files.
#NYI         ( $type, $width, $height, $x-dpi, $y-dpi ) =
#NYI           $self->process-jpg( $data, $filename );

#NYI         $self->{image-types}->{jpeg} = 1;
#NYI     }
#NYI     elsif ( unpack( 'A2', $data ) eq 'BM' ) {

#NYI         # Test for BMPs.
#NYI         ( $type, $width, $height ) = $self->process-bmp( $data, $filename );

#NYI         $self->{image-types}->{bmp} = 1;
#NYI     }
#NYI     else {
#NYI         fail "Unsupported image format for file: $filename\n";
#NYI     }

#NYI     push @{ $self->{images} }, [ $filename, $type ];

#NYI     # Set a default dpi for images with 0 dpi.
#NYI     $x-dpi = 96 if $x-dpi == 0;
#NYI     $y-dpi = 96 if $y-dpi == 0;

#NYI     $fh->close;

#NYI     return ( $type, $width, $height, $image-name, $x-dpi, $y-dpi );
#NYI }


###############################################################################
#
# process-png()
#
# Extract width and height information from a PNG file.
#
#NYI sub process-png {

#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];

#NYI     my $type   = 'png';
#NYI     my $width  = 0;
#NYI     my $height = 0;
#NYI     my $x-dpi  = 96;
#NYI     my $y-dpi  = 96;

#NYI     my $offset      = 8;
#NYI     my $data-length = length $data;

#NYI     # Search through the image data to read the height and width in the
#NYI     # IHDR element. Also read the DPI in the pHYs element.
#NYI     while ( $offset < $data-length ) {

#NYI         my $length = unpack "N",  substr $data, $offset + 0, 4;
#NYI         my $type   = unpack "A4", substr $data, $offset + 4, 4;

#NYI         if ( $type eq "IHDR" ) {
#NYI             $width  = unpack "N", substr $data, $offset + 8,  4;
#NYI             $height = unpack "N", substr $data, $offset + 12, 4;
#NYI         }

#NYI         if ( $type eq "pHYs" ) {
#NYI             my $x-ppu = unpack "N", substr $data, $offset + 8,  4;
#NYI             my $y-ppu = unpack "N", substr $data, $offset + 12, 4;
#NYI             my $units = unpack "C", substr $data, $offset + 16, 1;

#NYI             if ( $units == 1 ) {
#NYI                 $x-dpi = $x-ppu * 0.0254;
#NYI                 $y-dpi = $y-ppu * 0.0254;
#NYI             }
#NYI         }

#NYI         $offset = $offset + $length + 12;

#NYI         last if $type eq "IEND";
#NYI     }

#NYI     if ( not defined $height ) {
#NYI         fail "$filename: no size data found in png image.\n";
#NYI     }

#NYI     return ( $type, $width, $height, $x-dpi, $y-dpi );
#NYI }


###############################################################################
#
# process-bmp()
#
# Extract width and height information from a BMP file.
#
# Most of the checks came from old Spredsheet::WriteExcel code.
#
#NYI sub process-bmp {

#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];
#NYI     my $type     = 'bmp';


#NYI     # Check that the file is big enough to be a bitmap.
#NYI     if ( length $data <= 0x36 ) {
#NYI         fail "$filename doesn't contain enough data.";
#NYI     }


#NYI     # Read the bitmap width and height. Verify the sizes.
#NYI     my ( $width, $height ) = unpack "x18 V2", $data;

#NYI     if ( $width > 0xFFFF ) {
#NYI         fail "$filename: largest image width $width supported is 65k.";
#NYI     }

#NYI     if ( $height > 0xFFFF ) {
#NYI         fail "$filename: largest image height supported is 65k.";
#NYI     }

#NYI     # Read the bitmap planes and bpp data. Verify them.
#NYI     my ( $planes, $bitcount ) = unpack "x26 v2", $data;

#NYI     if ( $bitcount != 24 ) {
#NYI         fail "$filename isn't a 24bit true color bitmap.";
#NYI     }

#NYI     if ( $planes != 1 ) {
#NYI         fail "$filename: only 1 plane supported in bitmap image.";
#NYI     }


#NYI     # Read the bitmap compression. Verify compression.
#NYI     my $compression = unpack "x30 V", $data;

#NYI     if ( $compression != 0 ) {
#NYI         fail "$filename: compression not supported in bitmap image.";
#NYI     }

#NYI     return ( $type, $width, $height );
#NYI }


###############################################################################
#
# process-jpg()
#
# Extract width and height information from a JPEG file.
#
#NYI sub process-jpg {

#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];
#NYI     my $type     = 'jpeg';
#NYI     my $x-dpi    = 96;
#NYI     my $y-dpi    = 96;
#NYI     my $width;
#NYI     my $height;

#NYI     my $offset      = 2;
#NYI     my $data-length = length $data;

#NYI     # Search through the image data to read the height and width in the
#NYI     # 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element.
#NYI     while ( $offset < $data-length ) {

#NYI         my $marker = unpack "n", substr $data, $offset + 0, 2;
#NYI         my $length = unpack "n", substr $data, $offset + 2, 2;

#NYI         if ( $marker == 0xFFC0 || $marker == 0xFFC2 ) {
#NYI             $height = unpack "n", substr $data, $offset + 5, 2;
#NYI             $width  = unpack "n", substr $data, $offset + 7, 2;
#NYI         }

#NYI         if ( $marker == 0xFFE0 ) {
#NYI             my $units     = unpack "C", substr $data, $offset + 11, 1;
#NYI             my $x-density = unpack "n", substr $data, $offset + 12, 2;
#NYI             my $y-density = unpack "n", substr $data, $offset + 14, 2;

#NYI             if ( $units == 1 ) {
#NYI                 $x-dpi = $x-density;
#NYI                 $y-dpi = $y-density;
#NYI             }

#NYI             if ( $units == 2 ) {
#NYI                 $x-dpi = $x-density * 2.54;
#NYI                 $y-dpi = $y-density * 2.54;
#NYI             }
#NYI         }

#NYI         $offset = $offset + $length + 2;
#NYI         last if $marker == 0xFFDA;
#NYI     }

#NYI     if ( not defined $height ) {
#NYI         fail "$filename: no size data found in jpeg image.\n";
#NYI     }

#NYI     return ( $type, $width, $height, $x-dpi, $y-dpi );
#NYI }


#NYI ###############################################################################
#NYI #
#NYI # get-sheet-index()
#NYI #
#NYI # Convert a sheet name to its index. Return undef otherwise.
#NYI #
#NYI sub get-sheet-index {

#NYI     my $self        = shift;
#NYI     my $sheetname   = shift;
#NYI     my $sheet-index = undef;

#NYI     $sheetname =~ s/^'//;
#NYI     $sheetname =~ s/'$//;

#NYI     if ( exists $self->{sheetnames}->{$sheetname} ) {
#NYI         return $self->{sheetnames}->{$sheetname}->{index};
#NYI     }
#NYI     else {
#NYI         return undef;
#NYI     }
#NYI }


###############################################################################
#
# set-optimization()
#
# Set the speed/memory optimisation level.
#
method set-optimization($level = 1) {

    fail "set-optimization() must be called before add-worksheet()"
      if @!worksheets.elems == 0;

    $!optimization = $level;
}


#NYI ###############################################################################
#NYI #
#NYI # Deprecated methods for backwards compatibility.
#NYI #
#NYI ###############################################################################

#NYI # No longer required by Excel::Writer::XLSX.
#NYI sub compatibility-mode { }
#NYI sub set-codepage       { }


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# write-workbook()
#
# Write <workbook> element.
#
method write-workbook {

    my $schema  = 'http://schemas.openxmlformats.org';
    my $xmlns   = $schema ~ '/spreadsheetml/2006/main';
    my $xmlns-r = $schema ~ '/officeDocument/2006/relationships';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns-r,
    );

    self.xml-start-tag( 'workbook', @attributes );
}


###############################################################################
#
# write-file-version()
#
# Write the <fileVersion> element.
#
method write-file-version {

    my $app-name      = 'xl';
    my $last-edited   = 4;
    my $lowest-edited = 4;
    my $rup-build     = 4505;

    my @attributes = (
        'appName'      => $app-name,
        'lastEdited'   => $last-edited,
        'lowestEdited' => $lowest-edited,
        'rupBuild'     => $rup-build,
    );

    if $!vba-project {
        push @attributes, codeName => '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}';
    }

    self.xml-empty-tag( 'fileVersion', @attributes );
}


###############################################################################
#
# write-workbook-pr()
#
# Write <workbookPr> element.
#
method write-workbook-pr {

    my $date1904              = $!date1904;
    my $show-ink-annotation    = 0;
    my $auto-compress-pictures = 0;
    my $default-theme-version  = 124226;
    my $codename               = $!vba-codename;
    my @attributes;

    @attributes.push: 'codeName' => $codename if $codename;
    @attributes.push: 'date1904' => 1         if $date1904;
    @attributes.push: 'defaultThemeVersion' => $default-theme-version;

    self.xml-empty-tag( 'workbookPr', |@attributes );
}


###############################################################################
#
# write-book-views()
#
# Write <bookViews> element.
#
method write-book-views {

    self.xml-start-tag( 'bookViews' );
    self.write-workbook-view();
    self.xml-end-tag( 'bookViews' );
}

###############################################################################
#
# write-workbook-view()
#
# Write <workbookView> element.
#
method write-workbook-view {
    my $x-window      = $!x-window;
    my $y-window      = $!y-window;
    my $window-width  = $!window-width;
    my $window-height = $!window-height;
    my $tab-ratio     = $!tab-ratio;
    my $active-tab    = $!activesheet;
    my $first-sheet   = $!firstsheet;

    my @attributes = (
        'xWindow'      => $x-window,
        'yWindow'      => $y-window,
        'windowWidth'  => $window-width,
        'windowHeight' => $window-height,
    );

    # Store the tabRatio attribute when it isn't the default.
    @attributes.push: tabRatio => $tab-ratio if $tab-ratio != 500;

    # Store the firstSheet attribute when it isn't the default.
    @attributes.push: firstSheet => $first-sheet + 1 if $first-sheet > 0;

    # Store the activeTab attribute when it isn't the first sheet.
    @attributes.push: activeTab => $active-tab if $active-tab > 0;

    self.xml-empty-tag( 'workbookView', |@attributes );
}

###############################################################################
#
# write-sheets()
#
# Write <sheets> element.
#
method write-sheets {

    my $id-num = 1;

    self.xml-start-tag( 'sheets' );

    for @!worksheets -> $worksheet {
        self.write-sheet( $worksheet.name, $id-num++,
            $worksheet.hidden );
    }

    self.xml-end-tag( 'sheets' );
}


###############################################################################
#
# write-sheet()
#
# Write <sheet> element.
#
method write-sheet($name, $sheet-id, $hidden) {

    my $r-id     = 'rId' ~ $sheet-id;

    my @attributes = (
        'name'    => $name,
        'sheetId' => $sheet-id,
    );

    @attributes.push: 'state' => 'hidden' if $hidden;
    @attributes.push: 'r:id' => $r-id;

    self.xml-empty-tag( 'sheet', |@attributes );
}


###############################################################################
#
# write-calc-pr()
#
# Write <calcPr> element.
#
method write-calc-pr {

    my $calc-id         = $!calc-id;
    my $concurrent-calc = 0;

    my @attributes = calcId => $calc-id;

    if $!calc-mode eq 'manual' {
        @attributes.push: 'calcMode'   => 'manual';
        @attributes.push: 'calcOnSave' => 0;
    }
    elsif $!calc-mode eq 'autoNoTable' {
        @attributes.push: calcMode => 'autoNoTable';
    }

    if $!calc-on-load {
        @attributes.push: 'fullCalcOnLoad' => 1;
    }


    self.xml-empty-tag( 'calcPr', |@attributes );
}


###############################################################################
#
# write-ext-lst()
#
# Write <extLst> element.
#
#NYI sub write-ext-lst {

#NYI     my $self = shift;

#NYI     $self->xml-start-tag( 'extLst' );
#NYI     $self->write-ext();
#NYI     $self->xml-end-tag( 'extLst' );
#NYI }


###############################################################################
#
# write-ext()
#
# Write <ext> element.
#
#NYI sub write-ext {

#NYI     my $self     = shift;
#NYI     my $xmlns-mx = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
#NYI     my $uri      = 'http://schemas.microsoft.com/office/mac/excel/2008/main';

#NYI     my @attributes = (
#NYI         'xmlns:mx' => $xmlns-mx,
#NYI         'uri'      => $uri,
#NYI     );

#NYI     $self->xml-start-tag( 'ext', @attributes );
#NYI     $self->write-mx-arch-id();
#NYI     $self->xml-end-tag( 'ext' );
#NYI }

###############################################################################
#
# write-mx-arch-id()
#
# Write <mx:ArchID> element.
#
#NYI sub write-mx-arch-id {

#NYI     my $self  = shift;
#NYI     my $Flags = 2;

#NYI     my @attributes = ( 'Flags' => $Flags, );

#NYI     $self->xml-empty-tag( 'mx:ArchID', @attributes );
#NYI }


##############################################################################
#
# write-defined-names()
#
# Write the <definedNames> element.
#
method write-defined-names {

    return unless @!defined-names;

    self.xml-start-tag( 'definedNames' );

    for @!defined-names -> $aref {
        self.write-defined-name( $aref );
    }

    self.xml-end-tag( 'definedNames' );
}


##############################################################################
#
# write-defined-name()
#
# Write the <definedName> element.
#
method write-defined-name(@data) {

    my $name   = @data[0];
    my $id     = @data[1];
    my $range  = @data[2];
    my $hidden = @data[3];

    my @attributes = 'name' => $name;

    @attributes.push: 'localSheetId' => $id if $id != -1;
    @attributes.push: 'hidden'       => 1   if $hidden;

    self.xml-data-element( 'definedName', $range, |@attributes );
}

=begin pod

=head1 NAME

Workbook - A class for writing Excel Workbooks.

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
