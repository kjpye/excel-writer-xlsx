use v6.c;
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

###############################################################################
#
# Workbook - A class for writing Excel Workbooks.
#
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
# Copyright 2017,      Kevin.Pye
#
# Documentation after __END__
#

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';

###############################################################################
#
# Public and private API methods.
#
###############################################################################

has $!filename;
has $!tempdir;
has $!date_1904          = 0;
has $!activesheet        = 0;
has $!firstsheet         = 0;
has $!selected           = 0;
has $!fileclosed         = 0;
has $!filehandle;
has $!internal_fh        = 0;
has $!sheet_name         = 'Sheet';
has $!chart_name         = 'Chart';
has $!sheetname_count    = 0;
has $!chartname_count    = 0;
has @!worksheets         = [];
has @!charts             = [];
has @!drawings           = [];
has %!sheetnames         = {};
has @!formats            = [];
has @!xf_formats         = [];
has %!xf_format_indices  = {};
has @!dxf_formats        = [];
has %!dxf_format_indices = {};
has @!palette            = [];
has $!font_count         = 0;
has $!num_format_count   = 0;
has @!defined_names      = [];
has @!named_ranges       = [];
has @!custom_colors      = [];
has %!doc_properties     = {};
has @!custom_properties  = [];
has @!createtime         = [ now ];
has $!num_vml_files      = 0;
has $!num_comment_files  = 0;
has $!optimization       = 0;
has $!x_window           = 240;
has $!y_window           = 15;
has $!window_width       = 16095;
has $!window_height      = 9660;
has $!tab_ratio          = 500;
has $!excel2003_style    = 0;

has %!default_format_properties = {};

# Structures for the shared strings data.
has $!str_total  = 0;
has $!str_unique = 0;
has %!str_table  = {};
has $!str_array  = [];

# Formula calculation default settings.
has $!calc_id      = 124519;
has $!calc_mode    = 'auto';
has $!calc_on_load = 1;

has $!vba_project;
has @!shapes;

###############################################################################
#
# new()
#
# Constructor.
#
method TWEAK (*@args) {

#NYI     $self.filename = $_[0] || '';
#NYI     my $options = $_[1] || {};

#NYI     if ( exists $options->{tempdir} ) {
#NYI         $self->{_tempdir} = $options->{tempdir};
#NYI     }

#NYI     if ( exists $options->{date_1904} ) {
#NYI         $self->{_date_1904} = $options->{date_1904};
#NYI     }

#NYI     if ( exists $options->{optimization} ) {
#NYI         $self->{_optimization} = $options->{optimization};
#NYI     }

#NYI     if ( exists $options->{default_format_properties} ) {
#NYI         $self->{_default_format_properties} =
#NYI           $options->{default_format_properties};
#NYI     }

#NYI     if ( exists $options->{excel2003_style} ) {
#NYI         $self->{_excel2003_style} = 1;
#NYI     }


#NYI     bless $self, $class;

#NYI     # Add the default cell format.
#NYI     if ( $self->{_excel2003_style} ) {
#NYI         $self->add_format( xf_index => 0, font_family => 0 );
#NYI     }
#NYI     else {
#NYI         $self->add_format( xf_index => 0 );
#NYI     }

#NYI     # Check for a filename unless it is an existing filehandle
#NYI     if ( not ref $self->{_filename} and $self->{_filename} eq '' ) {
#NYI         warn 'Filename required by Excel::Writer::XLSX->new()';
#NYI         return Nil;
#NYI     }


#NYI     # If filename is a reference we assume that it is a valid filehandle.
#NYI     if ( ref $self->{_filename} ) {

#NYI         $self->{_filehandle}  = $self->{_filename};
#NYI         $self->{_internal_fh} = 0;
#NYI     }
#NYI     elsif ( $self->{_filename} eq '-' ) {

#NYI         # Support special filename/filehandle '-' for backward compatibility.
#NYI         binmode STDOUT;
#NYI         $self->{_filehandle}  = \*STDOUT;
#NYI         $self->{_internal_fh} = 0;
#NYI     }
#NYI     else {
#NYI         my $fh = IO::File->new( $self->{_filename}, 'w' );

#NYI         return undef unless defined $fh;

#NYI         $self->{_filehandle}  = $fh;
#NYI         $self->{_internal_fh} = 1;
#NYI     }


#NYI     # Set colour palette.
#NYI     $self->set_color_palette();

#NYI     return $self;
}


###############################################################################
#
# assemble_xml_file()
#
# Assemble and write the XML file.
#
method assemble_xml_file {

    # Prepare format object for passing to Style.pm.
    self.prepare_format_properties();

    self.xml_declaration;

    # Write the root workbook element.
    self.write_workbook();

    # Write the XLSX file version.
    self.write_file_version();

    # Write the workbook properties.
    self.write_workbook_pr();

    # Write the workbook view properties.
    self.write_book_views();

    # Write the worksheet names and ids.
    self.write_sheets();

    # Write the workbook defined names.
    self.write_defined_names();

    # Write the workbook calculation properties.
    self.write_calc_pr();

    # Write the workbook extension storage.
    # self.write_ext_lst();

    # Close the workbook tag.
    self.xml_end_tag( 'workbook' );

    # Close the XML writer filehandle.
    self.xml_get_fh().close();
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
    self.store_workbook;

    # Return the file close value.
    if $!internal_fh {
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
#NYI # An accessor for the _worksheets[] array
#NYI #
#NYI # Returns: an optionally sliced list of the worksheet objects in a workbook.
#NYI #
#NYI sub sheets {

#NYI     my $self = shift;

#NYI     if ( @_ ) {

#NYI         # Return a slice of the array
#NYI         return @{ $self->{_worksheets} }[@_];
#NYI     }
#NYI     else {

#NYI         # Return the entire list
#NYI         return @{ $self->{_worksheets} };
#NYI     }
#NYI }


###############################################################################
#
# get_worksheet_by_name(name)
#
# Return a worksheet object in the workbook using the sheetname.
#
method get_worksheet_by_name($sheetname) {

    return Nil unless $sheetname.defined;

    return %!sheetnames{$sheetname};
}


#NYI ###############################################################################
#NYI #
#NYI # worksheets()
#NYI #
#NYI # An accessor for the _worksheets[] array.
#NYI # This method is now deprecated. Use the sheets() method instead.
#NYI #
#NYI # Returns: an array reference
#NYI #
#NYI sub worksheets {

#NYI     my $self = shift;

#NYI     return $self->{_worksheets};
#NYI }


###############################################################################
#
# add_worksheet($name)
#
# Add a new worksheet to the Excel workbook.
#
# Returns: reference to a worksheet object
#
method add_worksheet($name? is copy) {

    my $index = @!worksheets.elems;
    $name  = self.check_sheetname( $name );
    my $fh;

    # Porters take note, the following scheme of passing references to Workbook
    # data (in the \$self->{_foo} cases) instead of a reference to the Workbook
    # itself is a workaround to avoid circular references between Workbook and
    # Worksheet objects. Feel free to implement this in any way the suits your
    # language.
    #
    my %init_data = (
        fh => $fh,
        name => $name,
        index =>$index,

        activesheet => $!activesheet,
        firstsheet => $!firstsheet,

        str_total => $!str_total,
        str_unique => $!str_unique,
        str_table => %!str_table,

        date_1904 => $!date_1904,
        palette => @!palette,
        optimization => $!optimization,
        tempdir => $!tempdir,
        excel2003_style => $!excel2003_style,

    );

    my $worksheet = Excel::Writer::XLSX::Worksheet.new( |%init_data.Map );
    @!worksheets[$index] = $worksheet;
    %!sheetnames{$name}  = $worksheet;

    return $worksheet;
}


#NYI ###############################################################################
#NYI #
#NYI # add_chart( %args )
#NYI #
#NYI # Create a chart for embedding or as a new sheet.
#NYI #
#NYI sub add_chart {

#NYI     my $self  = shift;
#NYI     my %arg   = @_;
#NYI     my $name  = '';
#NYI     my $index = @{ $self->{_worksheets} };
#NYI     my $fh    = undef;

#NYI     # Type must be specified so we can create the required chart instance.
#NYI     my $type = $arg{type};
#NYI     if ( !defined $type ) {
#NYI         fail "Must define chart type in add_chart()";
#NYI     }

#NYI     # Ensure that the chart defaults to non embedded.
#NYI     my $embedded = $arg{embedded} || 0;

#NYI     # Check the worksheet name for non-embedded charts.
#NYI     if ( !$embedded ) {
#NYI         $name = $self->_check_sheetname( $arg{name}, 1 );
#NYI     }


#NYI     my @init_data = (

#NYI         $fh,
#NYI         $name,
#NYI         $index,

#NYI         \$self->{_activesheet},
#NYI         \$self->{_firstsheet},

#NYI         \$self->{_str_total},
#NYI         \$self->{_str_unique},
#NYI         \$self->{_str_table},

#NYI         $self->{_date_1904},
#NYI         $self->{_palette},
#NYI         $self->{_optimization},
#NYI     );


#NYI     my $chart = Excel::Writer::XLSX::Chart->factory( $type, $arg{subtype} );

#NYI     # If the chart isn't embedded let the workbook control it.
#NYI     if ( !$embedded ) {

#NYI         my $drawing    = Excel::Writer::XLSX::Drawing->new();
#NYI         my $chartsheet = Excel::Writer::XLSX::Chartsheet->new( @init_data );

#NYI         $chart->{_palette} = $self->{_palette};

#NYI         $chartsheet->{_chart}   = $chart;
#NYI         $chartsheet->{_drawing} = $drawing;

#NYI         $self->{_worksheets}->[$index] = $chartsheet;
#NYI         $self->{_sheetnames}->{$name} = $chartsheet;

#NYI         push @{ $self->{_charts} }, $chart;

#NYI         return $chartsheet;
#NYI     }
#NYI     else {

#NYI         # Set the embedded chart name if present.
#NYI         $chart->{_chart_name} = $arg{name} if $arg{name};

#NYI         # Set index to 0 so that the activate() and set_first_sheet() methods
#NYI         # point back to the first worksheet if used for embedded charts.
#NYI         $chart->{_index}   = 0;
#NYI         $chart->{_palette} = $self->{_palette};
#NYI         $chart->_set_embedded_config_data();
#NYI         push @{ $self->{_charts} }, $chart;

#NYI         return $chart;
#NYI     }

#NYI }


###############################################################################
#
# _check_sheetname( $name )
#
# Check for valid worksheet names. We check the length, if it contains any
# invalid characters and if the name is unique in the workbook.
#
method check_sheetname($name is copy = '', $chart = 0) {

    $name //= '';
    my $invalid_char = token { <[\[\]:*?/\\]> };

    # Increment the Sheet/Chart number used for default sheet names below.
    if $chart {
        $!chartname_count++;
    }
    else {
        $!sheetname_count++;
    }

    # Supply default Sheet/Chart name if none has been defined.
    if $name eq '' {

        if $chart {
            $name = $!chart_name ~ $!chartname_count;
        }
        else {
            $name = $!sheet_name ~ $!sheetname_count;
        }
    }

    # Check that sheet name is <= 31. Excel limit.
    fail "Sheetname $name must be <= 31 chars" if $name.chars > 31;

    # Check that sheetname doesn't contain any invalid characters
    if $name ~~ $invalid_char {
        fail 'Invalid character []:*?/\\ in worksheet name: ' ~ $name;
    }

    # Check that the worksheet name doesn't already exist since this is a fatal
    # error in Excel 97. The check must also exclude case insensitive matches.
    for @!worksheets -> $worksheet {
        my $name_a = $name;
        my $name_b = $worksheet.name;

        if ( fc( $name_a ) eq fc( $name_b ) ) {
            fail "Worksheet name '$name', with case ignored, is already used.";
        }
    }

    return $name;
}


###############################################################################
#
# add_format(%properties)
#
# Add a new format to the Excel workbook.
#
method add_format(*%options) {

    my %init_data =
      ( |%!xf_format_indices, |%!dxf_format_indices );

    # Change default format style for Excel2003/XLS format.
    if $!excel2003_style {
        %init_data.append: ( font => 'Arial', size => 10, theme => -1 );
    }

    # Add the default format properties.
    %init_data.push: { %!default_format_properties};

    # Add the user defined properties.
    %init_data.push: %options;

    my $format = Excel::Writer::XLSX::Format.new( |%init_data );

    @!formats.push: $format;    # Store format reference

    return $format;
}


###############################################################################
#
# add_shape(%properties)
#
# Add a new shape to the Excel workbook.
#
method add_shape(*@options) {

    my $fh;
    my $shape = Excel::Writer::XLSX::Shape.new( $fh, @options );

    $shape.palette = self.palette;


    @!shapes.push: $shape;    # Store shape reference.

    return $shape;
}

###############################################################################
#
# set_1904()
#
# Set the date system: 0 = 1900 (the default), 1 = 1904
#
method set_1904($value = 1) {

    $!date_1904 = $value;
}


###############################################################################
#
# get_1904()
#
# Return the date system: 0 = 1900, 1 = 1904
#
method get_1904 {
  $!date_1904;
}


###############################################################################
#
# set_custom_color()
#
# Change the RGB components of the elements in the colour palette.
#
method set_custom_color($index, $red, $green?, $blue?) {

    # Match a HTML #xxyyzz style parameter
    if $red.defined and $red ~~ /^ '#' (\w\w) (\w\w) (\w\w)/ {
      ($red, $green, $blue) = ($0, $1, $2);
    }

    my $aref = @!palette;

    # Check that the colour index is the right range
    if ( $index < 8 or $index > 64 ) {
        warn "Color index $index outside range: 8 <= index <= 64";
        return 0;
    }

    # Check that the colour components are in the right range
    unless 0 <= $red   < 0 <= 255
       and 0 <= $green < 0 <= 255
       and 0 <= $blue  < 0 <= 25 
    {
        warn "Color component outside range: 0 <= color <= 255";
        return 0;
    }

    $index -= 8;    # Adjust colour index (wingless dragonfly)

    # Set the RGB value.
    my @rgb = ( $red, $green, $blue );
    $aref[$index] = @rgb;

    # Store the custom colors for the style.xml file.
    push @!custom_colors, sprintf "FF%02X%02X%02X", @rgb;

    return $index + 8;
}


###############################################################################
#
# set_color_palette()
#
# Sets the colour palette to the Excel defaults.
#
method set_color_palette {

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

    return 0;
}


#NYI ###############################################################################
#NYI #
#NYI # set_tempdir()
#NYI #
#NYI # Change the default temp directory.
#NYI #
#NYI sub set_tempdir {

#NYI     my $self = shift;
#NYI     my $dir  = shift;

#NYI     fail "$dir is not a valid directory" if defined $dir and not -d $dir;

#NYI     $self->{_tempdir} = $dir;

#NYI }


#NYI ###############################################################################
#NYI #
#NYI # define_name()
#NYI #
#NYI # Create a defined name in Excel. We handle global/workbook level names and
#NYI # local/worksheet names.
#NYI #
#NYI sub define_name {

#NYI     my $self        = shift;
#NYI     my $name        = shift;
#NYI     my $formula     = shift;
#NYI     my $sheet_index = undef;
#NYI     my $sheetname   = '';
#NYI     my $full_name   = $name;

#NYI     # Remove the = sign from the formula if it exists.
#NYI     $formula =~ s/^=//;

#NYI     # Local defined names are formatted like "Sheet1!name".
#NYI     if ( $name =~ /^(.*)!(.*)$/ ) {
#NYI         $sheetname   = $1;
#NYI         $name        = $2;
#NYI         $sheet_index = $self->_get_sheet_index( $sheetname );
#NYI     }
#NYI     else {
#NYI         $sheet_index = -1;    # Use -1 to indicate global names.
#NYI     }

#NYI     # Warn if the sheet index wasn't found.
#NYI     if ( !defined $sheet_index ) {
#NYI         warn "Unknown sheet name $sheetname in defined_name()";
#NYI         return -1;
#NYI     }

#NYI     # Warn if the name contains invalid chars as defined by Excel help.
#NYI     if ( $name !~ m/^[\w\\][\w\\.]*$/ || $name =~ m/^\d/ ) {
#NYI         warn "Invalid character in name '$name' used in defined_name()";
#NYI         return -1;
#NYI     }

#NYI     # Warn if the name looks like a cell name.
#NYI     if ( $name =~ m/^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$/ ) {
#NYI         warn "Invalid name '$name' looks like a cell name in defined_name()";
#NYI         return -1;
#NYI     }

#NYI     # Warn if the name looks like a R1C1.
#NYI     if ( $name =~ m/^[rcRC]$/ || $name =~ m/^[rcRC]\d+[rcRC]\d+$/ ) {
#NYI         warn "Invalid name '$name' like a RC cell ref in defined_name()";
#NYI         return -1;
#NYI     }

#NYI     push @{ $self->{_defined_names} }, [ $name, $sheet_index, $formula ];
#NYI }


###############################################################################
#
# set_size()
#
# Set the workbook size.
#
method set_size($width, $height) {

    if !$width {
        $!window_width = 16095;
    }
    else {
        # Convert to twips at 96 dpi.
        $!window_width = int( $width * 1440 / 96 );
    }

    if !$height {
        $!window_height = 9660;
    }
    else {
        # Convert to twips at 96 dpi.
        $!window_height = int( $height * 1440 / 96 );
    }
}


###############################################################################
#
# set_properties()
#
# Set the document properties such as Title, Author etc. These are written to
# property sets in the OLE container.
#
method set_properties(*%param) {

    # Ignore if no args were passed.
    return -1 unless %param;

    # List of valid input parameters.
    my %valid = (
        title          => 1,
        subject        => 1,
        author         => 1,
        keywords       => 1,
        comments       => 1,
        last_author    => 1,
        created        => 1,
        category       => 1,
        manager        => 1,
        company        => 1,
        status         => 1,
        hyperlink_base => 1,
    );

    # Check for valid input parameters.
    for %param.keys -> $parameter {
        if ( not %valid{$parameter}.defined ) {
            warn "Unknown parameter '$parameter' in set_properties()";
            return -1;
        }
    }

    # Set the creation time unless specified by the user.
    if ( ! %param<created>.exists ) {
        %param<created> = @!createtime;
    }


    %!doc_properties = %param;
}


###############################################################################
#
# set_custom_property()
#
# Set a user defined custom document property.
#
method set_custom_property($name, $value, $type?) {

    # Valid types.
    my %valid_type = (
        'text'       => 1,
        'date'       => 1,
        'number'     => 1,
        'number_int' => 1,
        'bool'       => 1,
    );

    if ! $name.defined || ! $value.defined {
        warn "The name and value parameters must be defined "
          ~ "in set_custom_property()";

        return -1;
    }

    # Determine the type for strings and numbers if it hasn't been specified.
    if !$type {
        if $value ~~ /^\d+$/ {
            $type = 'number_int';
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
    if ! %valid_type{$type}.exists {
        warn "Unknown custom type '$type' in set_custom_property()";
        return -1;
    }

    #  Check for strings longer than Excel's limit of 255 chars.
    if $type eq 'text' and $value.chars > 255 {
        warn "Length of text custom value '$value' exceeds "
          ~ "Excel's limit of 255 in set_custom_property()";
        return -1;
    }
    if $name.chars > 255 {
        warn "Length of custom name '$name' exceeds "
          ~ "Excel's limit of 255 in set_custom_property()";
        return -1;
    }

    push @!custom_properties, [ $name, $value, $type ];
}



#NYI ###############################################################################
#NYI #
#NYI # add_vba_project()
#NYI #
#NYI # Add a vbaProject binary to the XLSX file.
#NYI #
#NYI sub add_vba_project {

#NYI     my $self        = shift;
#NYI     my $vba_project = shift;

#NYI     fail "No vbaProject.bin specified in add_vba_project()"
#NYI       if not $vba_project;

#NYI     fail "Couldn't locate $vba_project in add_vba_project(): $!"
#NYI       unless -e $vba_project;

#NYI     $self->{_vba_project} = $vba_project;
#NYI }


#NYI ###############################################################################
#NYI #
#NYI # set_vba_name()
#NYI #
#NYI # Set the VBA name for the workbook.
#NYI #
#NYI sub set_vba_name {

#NYI     my $self         = shift;
#NYI     my $vba_codemame = shift;

#NYI     if ( $vba_codemame ) {
#NYI         $self->{_vba_codename} = $vba_codemame;
#NYI     }
#NYI     else {
#NYI         $self->{_vba_codename} = 'ThisWorkbook';
#NYI     }
#NYI }


###############################################################################
#
# set_calc_mode()
#
# Set the Excel calcuation mode for the workbook.
#
method set_calc_mode($mode = 'auto', $calc-id?) {

    $!calc_mode = $mode;

    if $mode eq 'manual' {
        $!calc_mode    = 'manual';
        $!calc_on_load = 0;
    }
    elsif $mode eq 'auto_except_tables' {
        $!calc_mode = 'autoNoTable';
    }

    $!calc_id = $calc-id if $calc-id.defined;
}


###############################################################################
#
# _store_workbook()
#
# Assemble worksheets into a workbook.
#
method store_workbook {

    my $tempdir  = File::Temp.newdir( tempdir => $!tempdir );
    my $packager = Excel::Writer::XLSX::Package::Packager.new();
    my $zip      = Archive::SimpleZip.new();


    # Add a default worksheet if none have been added.
    self.add_worksheet() if not @!worksheets;

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
    self.prepare_sst_string_data();

    # Prepare the worksheet VML elements such as comments and buttons.
    self.prepare_vml_objects();

    # Set the defined names for the worksheets such as Print Titles.
    self.prepare_defined_names();

    # Prepare the drawings, charts and images.
    self.prepare_drawings();

    # Add cached data to charts.
    self.add_chart_data();

    # Prepare the worksheet tables.
    self.prepare_tables();

    # Package the workbook.
    $packager.add_workbook();
    $packager.set_package_dir( $tempdir );
    $packager.create_package();

    # Free up the Packager object.
    $packager = Nil;

    # Add the files to the zip archive. Due to issues with Archive::Zip in
    # taint mode we can't use addTree() so we have to build the file list
    # with File::Find and pass each one to addFile().
    my $xlsx_files = File::Find::find( dir => $tempdir, type => 'file' );

    # Store the xlsx component files with the temp dir name removed.
    for $xlsx_files -> $filename {
        my $short_name = $filename;
        $short_name ~~ s/^$tempdir '/'?//;
        $zip.addFile( $filename, $short_name );
    }


    if $!internal_fh {

        if $zip.writeToFileHandle( $!filehandle ) != 0 {
            warn 'Error writing zip container for xlsx file.';
        }
    }
    else {

        # Archive::Zip needs to rewind a filehandle to write the zip headers.
        # This won't work for arbitrary user defined filehandles so we use
        # a temp file based filehandle to create the zip archive and then
        # stream that to the filehandle.
        my $tmp_fh = tempfile( tempdir => $!tempdir );
        my $is_seekable = 1;

        if $zip.writeToFileHandle( $tmp_fh, $is_seekable ) != 0 {
            warn 'Error writing zip container for xlsx file.';
        }

        my $buffer;
        $tmp_fh.seek: 0, 0;

        while $tmp_fh.read: $buffer, 4_096 {
            # local $\ = undef;    # Protect print from -l on commandline.
            $!filehandle.print: $buffer;
        }
    }
}


###############################################################################
#
# _prepare_sst_string_data()
#
# Convert the SST string data from a hash to an array.
#
#NYI sub _prepare_sst_string_data {

#NYI     my $self = shift;

#NYI     my @strings;
#NYI     $#strings = $self->{_str_unique} - 1;    # Pre-extend array

#NYI     while ( my $key = each %{ $self->{_str_table} } ) {
#NYI         $strings[ $self->{_str_table}->{$key} ] = $key;
#NYI     }

#NYI     # The SST data could be very large, free some memory (maybe).
#NYI     $self->{_str_table} = undef;
#NYI     $self->{_str_array} = \@strings;

#NYI }


###############################################################################
#
# _prepare_format_properties()
#
# Prepare all of the format properties prior to passing them to Styles.pm.
#
#NYI sub _prepare_format_properties {

#NYI     my $self = shift;

#NYI     # Separate format objects into XF and DXF formats.
#NYI     $self->_prepare_formats();

#NYI     # Set the font index for the format objects.
#NYI     $self->_prepare_fonts();

#NYI     # Set the number format index for the format objects.
#NYI     $self->_prepare_num_formats();

#NYI     # Set the border index for the format objects.
#NYI     $self->_prepare_borders();

#NYI     # Set the fill index for the format objects.
#NYI     $self->_prepare_fills();


#NYI }


###############################################################################
#
# _prepare_formats()
#
# Iterate through the XF Format objects and separate them into XF and DXF
# formats.
#
#NYI sub _prepare_formats {

#NYI     my $self = shift;

#NYI     for my $format ( @{ $self->{_formats} } ) {
#NYI         my $xf_index  = $format->{_xf_index};
#NYI         my $dxf_index = $format->{_dxf_index};

#NYI         if ( defined $xf_index ) {
#NYI             $self->{_xf_formats}->[$xf_index] = $format;
#NYI         }

#NYI         if ( defined $dxf_index ) {
#NYI             $self->{_dxf_formats}->[$dxf_index] = $format;
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# _set_default_xf_indices()
#
# Set the default index for each format. This is mainly used for testing.
#
#NYI sub _set_default_xf_indices {

#NYI     my $self = shift;

#NYI     for my $format ( @{ $self->{_formats} } ) {
#NYI         $format->get_xf_index();
#NYI     }
#NYI }


###############################################################################
#
# _prepare_fonts()
#
# Iterate through the XF Format objects and give them an index to non-default
# font elements.
#
#NYI sub _prepare_fonts {

#NYI     my $self = shift;

#NYI     my %fonts;
#NYI     my $index = 0;

#NYI     for my $format ( @{ $self->{_xf_formats} } ) {
#NYI         my $key = $format->get_font_key();

#NYI         if ( exists $fonts{$key} ) {

#NYI             # Font has already been used.
#NYI             $format->{_font_index} = $fonts{$key};
#NYI             $format->{_has_font}   = 0;
#NYI         }
#NYI         else {

#NYI             # This is a new font.
#NYI             $fonts{$key}           = $index;
#NYI             $format->{_font_index} = $index;
#NYI             $format->{_has_font}   = 1;
#NYI             $index++;
#NYI         }
#NYI     }

#NYI     $self->{_font_count} = $index;

#NYI     # For the DXF formats we only need to check if the properties have changed.
#NYI     for my $format ( @{ $self->{_dxf_formats} } ) {

#NYI         # The only font properties that can change for a DXF format are: color,
#NYI         # bold, italic, underline and strikethrough.
#NYI         if (   $format->{_color}
#NYI             || $format->{_bold}
#NYI             || $format->{_italic}
#NYI             || $format->{_underline}
#NYI             || $format->{_font_strikeout} )
#NYI         {
#NYI             $format->{_has_dxf_font} = 1;
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# _prepare_num_formats()
#
# Iterate through the XF Format objects and give them an index to non-default
# number format elements.
#
# User defined records start from index 0xA4.
#
#NYI sub _prepare_num_formats {

#NYI     my $self = shift;

#NYI     my %num_formats;
#NYI     my $index            = 164;
#NYI     my $num_format_count = 0;

#NYI     for my $format ( @{ $self->{_xf_formats} }, @{ $self->{_dxf_formats} } ) {
#NYI         my $num_format = $format->{_num_format};

#NYI         # Check if $num_format is an index to a built-in number format.
#NYI         # Also check for a string of zeros, which is a valid number format
#NYI         # string but would evaluate to zero.
#NYI         #
#NYI         if ( $num_format =~ m/^\d+$/ && $num_format !~ m/^0+\d/ ) {

#NYI             # Index to a built-in number format.
#NYI             $format->{_num_format_index} = $num_format;
#NYI             next;
#NYI         }


#NYI         if ( exists( $num_formats{$num_format} ) ) {

#NYI             # Number format has already been used.
#NYI             $format->{_num_format_index} = $num_formats{$num_format};
#NYI         }
#NYI         else {

#NYI             # Add a new number format.
#NYI             $num_formats{$num_format} = $index;
#NYI             $format->{_num_format_index} = $index;
#NYI             $index++;

#NYI             # Only increase font count for XF formats (not for DXF formats).
#NYI             if ( $format->{_xf_index} ) {
#NYI                 $num_format_count++;
#NYI             }
#NYI         }
#NYI     }

#NYI     $self->{_num_format_count} = $num_format_count;
#NYI }


###############################################################################
#
# _prepare_borders()
#
# Iterate through the XF Format objects and give them an index to non-default
# border elements.
#
#NYI sub _prepare_borders {

#NYI     my $self = shift;

#NYI     my %borders;
#NYI     my $index = 0;

#NYI     for my $format ( @{ $self->{_xf_formats} } ) {
#NYI         my $key = $format->get_border_key();

#NYI         if ( exists $borders{$key} ) {

#NYI             # Border has already been used.
#NYI             $format->{_border_index} = $borders{$key};
#NYI             $format->{_has_border}   = 0;
#NYI         }
#NYI         else {

#NYI             # This is a new border.
#NYI             $borders{$key}           = $index;
#NYI             $format->{_border_index} = $index;
#NYI             $format->{_has_border}   = 1;
#NYI             $index++;
#NYI         }
#NYI     }

#NYI     $self->{_border_count} = $index;

#NYI     # For the DXF formats we only need to check if the properties have changed.
#NYI     for my $format ( @{ $self->{_dxf_formats} } ) {
#NYI         my $key = $format->get_border_key();

#NYI         if ( $key =~ m/[^0:]/ ) {
#NYI             $format->{_has_dxf_border} = 1;
#NYI         }
#NYI     }

#NYI }


###############################################################################
#
# _prepare_fills()
#
# Iterate through the XF Format objects and give them an index to non-default
# fill elements.
#
# The user defined fill properties start from 2 since there are 2 default
# fills: patternType="none" and patternType="gray125".
#
#NYI sub _prepare_fills {

#NYI     my $self = shift;

#NYI     my %fills;
#NYI     my $index = 2;    # Start from 2. See above.

#NYI     # Add the default fills.
#NYI     $fills{'0:0:0'}  = 0;
#NYI     $fills{'17:0:0'} = 1;


#NYI     # Store the DXF colours separately since them may be reversed below.
#NYI     for my $format ( @{ $self->{_dxf_formats} } ) {
#NYI         if (   $format->{_pattern}
#NYI             || $format->{_bg_color}
#NYI             || $format->{_fg_color} )
#NYI         {
#NYI             $format->{_has_dxf_fill} = 1;
#NYI             $format->{_dxf_bg_color} = $format->{_bg_color};
#NYI             $format->{_dxf_fg_color} = $format->{_fg_color};
#NYI         }
#NYI     }


#NYI     for my $format ( @{ $self->{_xf_formats} } ) {

#NYI         # The following logical statements jointly take care of special cases
#NYI         # in relation to cell colours and patterns:
#NYI         # 1. For a solid fill (_pattern == 1) Excel reverses the role of
#NYI         #    foreground and background colours, and
#NYI         # 2. If the user specifies a foreground or background colour without
#NYI         #    a pattern they probably wanted a solid fill, so we fill in the
#NYI         #    defaults.
#NYI         #
#NYI         if (   $format->{_pattern} == 1
#NYI             && $format->{_bg_color} ne '0'
#NYI             && $format->{_fg_color} ne '0' )
#NYI         {
#NYI             my $tmp = $format->{_fg_color};
#NYI             $format->{_fg_color} = $format->{_bg_color};
#NYI             $format->{_bg_color} = $tmp;
#NYI         }

#NYI         if (   $format->{_pattern} <= 1
#NYI             && $format->{_bg_color} ne '0'
#NYI             && $format->{_fg_color} eq '0' )
#NYI         {
#NYI             $format->{_fg_color} = $format->{_bg_color};
#NYI             $format->{_bg_color} = 0;
#NYI             $format->{_pattern}  = 1;
#NYI         }

#NYI         if (   $format->{_pattern} <= 1
#NYI             && $format->{_bg_color} eq '0'
#NYI             && $format->{_fg_color} ne '0' )
#NYI         {
#NYI             $format->{_bg_color} = 0;
#NYI             $format->{_pattern}  = 1;
#NYI         }


#NYI         my $key = $format->get_fill_key();

#NYI         if ( exists $fills{$key} ) {

#NYI             # Fill has already been used.
#NYI             $format->{_fill_index} = $fills{$key};
#NYI             $format->{_has_fill}   = 0;
#NYI         }
#NYI         else {

#NYI             # This is a new fill.
#NYI             $fills{$key}           = $index;
#NYI             $format->{_fill_index} = $index;
#NYI             $format->{_has_fill}   = 1;
#NYI             $index++;
#NYI         }
#NYI     }

#NYI     $self->{_fill_count} = $index;


#NYI }


###############################################################################
#
# _prepare_defined_names()
#
# Iterate through the worksheets and store any defined names in addition to
# any user defined names. Stores the defined names for the Workbook.xml and
# the named ranges for App.xml.
#
#NYI sub _prepare_defined_names {

#NYI     my $self = shift;

#NYI     my @defined_names = @{ $self->{_defined_names} };

#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {

#NYI         # Check for Print Area settings.
#NYI         if ( $sheet->{_autofilter} ) {

#NYI             my $range  = $sheet->{_autofilter};
#NYI             my $hidden = 1;

#NYI             # Store the defined names.
#NYI             push @defined_names,
#NYI               [ '_xlnm._FilterDatabase', $sheet->{_index}, $range, $hidden ];

#NYI         }

#NYI         # Check for Print Area settings.
#NYI         if ( $sheet->{_print_area} ) {

#NYI             my $range = $sheet->{_print_area};

#NYI             # Store the defined names.
#NYI             push @defined_names,
#NYI               [ '_xlnm.Print_Area', $sheet->{_index}, $range ];
#NYI         }

#NYI         # Check for repeat rows/cols. aka, Print Titles.
#NYI         if ( $sheet->{_repeat_cols} || $sheet->{_repeat_rows} ) {
#NYI             my $range = '';

#NYI             if ( $sheet->{_repeat_cols} && $sheet->{_repeat_rows} ) {
#NYI                 $range = $sheet->{_repeat_cols} . ',' . $sheet->{_repeat_rows};
#NYI             }
#NYI             else {
#NYI                 $range = $sheet->{_repeat_cols} . $sheet->{_repeat_rows};
#NYI             }

#NYI             # Store the defined names.
#NYI             push @defined_names,
#NYI               [ '_xlnm.Print_Titles', $sheet->{_index}, $range ];
#NYI         }

#NYI     }

#NYI     @defined_names          = _sort_defined_names( @defined_names );
#NYI     $self->{_defined_names} = \@defined_names;
#NYI     $self->{_named_ranges}  = _extract_named_ranges( @defined_names );
#NYI }


###############################################################################
#
# _sort_defined_names()
#
# Sort internal and user defined names in the same order as used by Excel.
# This may not be strictly necessary but unsorted elements caused a lot of
# issues in the Spreadsheet::WriteExcel binary version. Also makes
# comparison testing easier.
#
#NYI sub _sort_defined_names {

#NYI     my @names = @_;

#NYI     #<<< Perltidy ignore this.

#NYI     @names = sort {
#NYI         # Primary sort based on the defined name.
#NYI         _normalise_defined_name( $a->[0] )
#NYI         cmp
#NYI         _normalise_defined_name( $b->[0] )

#NYI         ||
#NYI         # Secondary sort based on the sheet name.
#NYI         _normalise_sheet_name( $a->[2] )
#NYI         cmp
#NYI         _normalise_sheet_name( $b->[2] )

#NYI     } @names;
#NYI     #>>>

#NYI     return @names;
#NYI }

# Used in the above sort routine to normalise the defined names. Removes any
# leading '_xmln.' from internal names and lowercases the strings.
#NYI sub _normalise_defined_name {
#NYI     my $name = shift;

#NYI     $name =~ s/^_xlnm.//;
#NYI     $name = lc $name;

#NYI     return $name;
#NYI }

# Used in the above sort routine to normalise the worksheet names for the
# secondary sort. Removes leading quote and lowercases the strings.
#NYI sub _normalise_sheet_name {
#NYI     my $name = shift;

#NYI     $name =~ s/^'//;
#NYI     $name = lc $name;

#NYI     return $name;
#NYI }


###############################################################################
#
# _extract_named_ranges()
#
# Extract the named ranges from the sorted list of defined names. These are
# used in the App.xml file.
#
#NYI sub _extract_named_ranges {

#NYI     my @defined_names = @_;
#NYI     my @named_ranges;

#NYI     NAME:
#NYI     for my $defined_name ( @defined_names ) {

#NYI         my $name  = $defined_name->[0];
#NYI         my $index = $defined_name->[1];
#NYI         my $range = $defined_name->[2];

#NYI         # Skip autoFilter ranges.
#NYI         next NAME if $name eq '_xlnm._FilterDatabase';

#NYI         # We are only interested in defined names with ranges.
#NYI         if ( $range =~ /^([^!]+)!/ ) {
#NYI             my $sheet_name = $1;

#NYI             # Match Print_Area and Print_Titles xlnm types.
#NYI             if ( $name =~ /^_xlnm\.(.*)$/ ) {
#NYI                 my $xlnm_type = $1;
#NYI                 $name = $sheet_name . '!' . $xlnm_type;
#NYI             }
#NYI             elsif ( $index != -1 ) {
#NYI                 $name = $sheet_name . '!' . $name;
#NYI             }

#NYI             push @named_ranges, $name;
#NYI         }
#NYI     }

#NYI     return \@named_ranges;
#NYI }


###############################################################################
#
# _prepare_drawings()
#
# Iterate through the worksheets and set up any chart or image drawings.
#
#NYI sub _prepare_drawings {

#NYI     my $self         = shift;
#NYI     my $chart_ref_id = 0;
#NYI     my $image_ref_id = 0;
#NYI     my $drawing_id   = 0;

#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {

#NYI         my $chart_count = scalar @{ $sheet->{_charts} };
#NYI         my $image_count = scalar @{ $sheet->{_images} };
#NYI         my $shape_count = scalar @{ $sheet->{_shapes} };

#NYI         my $header_image_count = scalar @{ $sheet->{_header_images} };
#NYI         my $footer_image_count = scalar @{ $sheet->{_footer_images} };
#NYI         my $has_drawing        = 0;


#NYI         # Check that some image or drawing needs to be processed.
#NYI         if (   !$chart_count
#NYI             && !$image_count
#NYI             && !$shape_count
#NYI             && !$header_image_count
#NYI             && !$footer_image_count )
#NYI         {
#NYI             next;
#NYI         }

#NYI         # Don't increase the drawing_id header/footer images.
#NYI         if ( $chart_count || $image_count || $shape_count ) {
#NYI             $drawing_id++;
#NYI             $has_drawing = 1;
#NYI         }

#NYI         # Prepare the worksheet charts.
#NYI         for my $index ( 0 .. $chart_count - 1 ) {
#NYI             $chart_ref_id++;
#NYI             $sheet->_prepare_chart( $index, $chart_ref_id, $drawing_id );
#NYI         }

#NYI         # Prepare the worksheet images.
#NYI         for my $index ( 0 .. $image_count - 1 ) {

#NYI             my $filename = $sheet->{_images}->[$index]->[2];

#NYI             my ( $type, $width, $height, $name, $x_dpi, $y_dpi ) =
#NYI               $self->_get_image_properties( $filename );

#NYI             $image_ref_id++;

#NYI             $sheet->_prepare_image(
#NYI                 $index, $image_ref_id, $drawing_id,
#NYI                 $width, $height,       $name,
#NYI                 $type,  $x_dpi,        $y_dpi
#NYI             );
#NYI         }

#NYI         # Prepare the worksheet shapes.
#NYI         for my $index ( 0 .. $shape_count - 1 ) {
#NYI             $sheet->_prepare_shape( $index, $drawing_id );
#NYI         }

#NYI         # Prepare the header images.
#NYI         for my $index ( 0 .. $header_image_count - 1 ) {

#NYI             my $filename = $sheet->{_header_images}->[$index]->[0];
#NYI             my $position = $sheet->{_header_images}->[$index]->[1];

#NYI             my ( $type, $width, $height, $name, $x_dpi, $y_dpi ) =
#NYI               $self->_get_image_properties( $filename );

#NYI             $image_ref_id++;

#NYI             $sheet->_prepare_header_image( $image_ref_id, $width, $height,
#NYI                 $name, $type, $position, $x_dpi, $y_dpi );
#NYI         }

#NYI         # Prepare the footer images.
#NYI         for my $index ( 0 .. $footer_image_count - 1 ) {

#NYI             my $filename = $sheet->{_footer_images}->[$index]->[0];
#NYI             my $position = $sheet->{_footer_images}->[$index]->[1];

#NYI             my ( $type, $width, $height, $name, $x_dpi, $y_dpi ) =
#NYI               $self->_get_image_properties( $filename );

#NYI             $image_ref_id++;

#NYI             $sheet->_prepare_header_image( $image_ref_id, $width, $height,
#NYI                 $name, $type, $position, $x_dpi, $y_dpi );
#NYI         }


#NYI         if ( $has_drawing ) {
#NYI             my $drawing = $sheet->{_drawing};
#NYI             push @{ $self->{_drawings} }, $drawing;
#NYI         }
#NYI     }


#NYI     # Remove charts that were created but not inserted into worksheets.
#NYI     my @chart_data;

#NYI     for my $chart ( @{ $self->{_charts} } ) {
#NYI         if ( $chart->{_id} != -1 ) {
#NYI             push @chart_data, $chart;
#NYI         }
#NYI     }

#NYI     # Sort the workbook charts references into the order that the were
#NYI     # written from the worksheets above.
#NYI     @chart_data = sort { $a->{_id} <=> $b->{_id} } @chart_data;

#NYI     $self->{_charts} = \@chart_data;
#NYI     $self->{_drawing_count} = $drawing_id;
#NYI }


###############################################################################
#
# _prepare_vml_objects()
#
# Iterate through the worksheets and set up the VML objects.
#
#NYI sub _prepare_vml_objects {

#NYI     my $self           = shift;
#NYI     my $comment_id     = 0;
#NYI     my $vml_drawing_id = 0;
#NYI     my $vml_data_id    = 1;
#NYI     my $vml_header_id  = 0;
#NYI     my $vml_shape_id   = 1024;
#NYI     my $vml_files      = 0;
#NYI     my $comment_files  = 0;
#NYI     my $has_button     = 0;

#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {

#NYI         next if !$sheet->{_has_vml} and !$sheet->{_has_header_vml};
#NYI         $vml_files = 1;


#NYI         if ( $sheet->{_has_vml} ) {

#NYI             $comment_files++ if $sheet->{_has_comments};
#NYI             $comment_id++    if $sheet->{_has_comments};
#NYI             $vml_drawing_id++;

#NYI             my $count =
#NYI               $sheet->_prepare_vml_objects( $vml_data_id, $vml_shape_id,
#NYI                 $vml_drawing_id, $comment_id );

#NYI             # Each VML file should start with a shape id incremented by 1024.
#NYI             $vml_data_id  += 1 * int(    ( 1024 + $count ) / 1024 );
#NYI             $vml_shape_id += 1024 * int( ( 1024 + $count ) / 1024 );

#NYI         }

#NYI         if ( $sheet->{_has_header_vml} ) {
#NYI             $vml_header_id++;
#NYI             $vml_drawing_id++;
#NYI             $sheet->_prepare_header_vml_objects( $vml_header_id,
#NYI                 $vml_drawing_id );
#NYI         }

#NYI         # Set the sheet vba_codename if it has a button and the workbook
#NYI         # has a vbaProject binary.
#NYI         if ( $sheet->{_buttons_array} ) {
#NYI             $has_button = 1;

#NYI             if ( $self->{_vba_project} && !$sheet->{_vba_codename} ) {
#NYI                 $sheet->set_vba_name();
#NYI             }
#NYI         }

#NYI     }

#NYI     $self->{_num_vml_files}     = $vml_files;
#NYI     $self->{_num_comment_files} = $comment_files;

#NYI     # Add a font format for cell comments.
#NYI     if ( $comment_files > 0 ) {
#NYI         my $format = Excel::Writer::XLSX::Format->new(
#NYI             \$self->{_xf_format_indices},
#NYI             \$self->{_dxf_format_indices},
#NYI             font          => 'Tahoma',
#NYI             size          => 8,
#NYI             color_indexed => 81,
#NYI             font_only     => 1,
#NYI         );

#NYI         $format->get_xf_index();

#NYI         push @{ $self->{_formats} }, $format;
#NYI     }

#NYI     # Set the workbook vba_codename if one of the sheets has a button and
#NYI     # the workbook has a vbaProject binary.
#NYI     if ( $has_button && $self->{_vba_project} && !$self->{_vba_codename} ) {
#NYI         $self->set_vba_name();
#NYI     }
#NYI }


###############################################################################
#
# _prepare_tables()
#
# Set the table ids for the worksheet tables.
#
#NYI sub _prepare_tables {

#NYI     my $self     = shift;
#NYI     my $table_id = 0;
#NYI     my $seen     = {};

#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {

#NYI         my $table_count = scalar @{ $sheet->{_tables} };

#NYI         next unless $table_count;

#NYI         $sheet->_prepare_tables( $table_id + 1, $seen );

#NYI         $table_id += $table_count;
#NYI     }
#NYI }


###############################################################################
#
# _add_chart_data()
#
# Add "cached" data to charts to provide the numCache and strCache data for
# series and title/axis ranges.
#
#NYI sub _add_chart_data {

#NYI     my $self = shift;
#NYI     my %worksheets;
#NYI     my %seen_ranges;
#NYI     my @charts;

#NYI     # Map worksheet names to worksheet objects.
#NYI     for my $worksheet ( @{ $self->{_worksheets} } ) {
#NYI         $worksheets{ $worksheet->{_name} } = $worksheet;
#NYI     }

#NYI     # Build an array of the worksheet charts including any combined charts.
#NYI     for my $chart ( @{ $self->{_charts} } ) {
#NYI         push @charts, $chart;

#NYI         if ($chart->{_combined}) {
#NYI             push @charts, $chart->{_combined};
#NYI         }
#NYI     }


#NYI     CHART:
#NYI     for my $chart ( @charts ) {

#NYI         RANGE:
#NYI         while ( my ( $range, $id ) = each %{ $chart->{_formula_ids} } ) {

#NYI             # Skip if the series has user defined data.
#NYI             if ( defined $chart->{_formula_data}->[$id] ) {
#NYI                 if (   !exists $seen_ranges{$range}
#NYI                     || !defined $seen_ranges{$range} )
#NYI                 {
#NYI                     my $data = $chart->{_formula_data}->[$id];
#NYI                     $seen_ranges{$range} = $data;
#NYI                 }
#NYI                 next RANGE;
#NYI             }

#NYI             # Check to see if the data is already cached locally.
#NYI             if ( exists $seen_ranges{$range} ) {
#NYI                 $chart->{_formula_data}->[$id] = $seen_ranges{$range};
#NYI                 next RANGE;
#NYI             }

#NYI             # Convert the range formula to a sheet name and cell range.
#NYI             my ( $sheetname, @cells ) = $self->_get_chart_range( $range );

#NYI             # Skip if we couldn't parse the formula.
#NYI             next RANGE if !defined $sheetname;

#NYI             # Handle non-contiguous ranges: (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5).
#NYI             # We don't try to parse the ranges. We just return an empty list.
#NYI             if ( $sheetname =~ m/^\([^,]+,/ ) {
#NYI                 $chart->{_formula_data}->[$id] = [];
#NYI                 $seen_ranges{$range} = [];
#NYI                 next RANGE;
#NYI             }

#NYI             # Die if the name is unknown since it indicates a user error in
#NYI             # a chart series formula.
#NYI             if ( !exists $worksheets{$sheetname} ) {
#NYI                 die "Unknown worksheet reference '$sheetname' in range "
#NYI                   . "'$range' passed to add_series().\n";
#NYI             }

#NYI             # Find the worksheet object based on the sheet name.
#NYI             my $worksheet = $worksheets{$sheetname};

#NYI             # Get the data from the worksheet table.
#NYI             my @data = $worksheet->_get_range_data( @cells );

#NYI             # Convert shared string indexes to strings.
#NYI             for my $token ( @data ) {
#NYI                 if ( ref $token ) {
#NYI                     $token = $self->{_str_array}->[ $token->{sst_id} ];

#NYI                     # Ignore rich strings for now. Deparse later if necessary.
#NYI                     if ( $token =~ m{^<r>} && $token =~ m{</r>$} ) {
#NYI                         $token = '';
#NYI                     }
#NYI                 }
#NYI             }

#NYI             # Add the data to the chart.
#NYI             $chart->{_formula_data}->[$id] = \@data;

#NYI             # Store range data locally to avoid lookup if seen again.
#NYI             $seen_ranges{$range} = \@data;
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# _get_chart_range()
#
# Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name and cell
# range such as ( 'Sheet1', 0, 1, 4, 1 ).
#
#NYI sub _get_chart_range {

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

#NYI     my ( $row_start, $col_start ) = xl_cell_to_rowcol( $cell_1 );
#NYI     my ( $row_end,   $col_end )   = xl_cell_to_rowcol( $cell_2 );

#NYI     # Check that we have a 1D range only.
#NYI     if ( $row_start != $row_end && $col_start != $col_end ) {
#NYI         return undef;
#NYI     }

#NYI     return ( $sheetname, $row_start, $col_start, $row_end, $col_end );
#NYI }


###############################################################################
#
# _store_externs()
#
# Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
# the NAME records.
#
#NYI sub _store_externs {

#NYI     my $self = shift;

#NYI }


###############################################################################
#
# _store_names()
#
# Write the NAME record to define the print area and the repeat rows and cols.
#
#NYI sub _store_names {

#NYI     my $self = shift;

#NYI }


###############################################################################
#
# _quote_sheetname()
#
# Sheetnames used in references should be quoted if they contain any spaces,
# special characters or if the look like something that isn't a sheet name.
# TODO. We need to handle more special cases.
#
#NYI sub _quote_sheetname {

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
# _get_image_properties()
#
# Extract information from the image file such as dimension, type, filename,
# and extension. Also keep track of previously seen images to optimise out
# any duplicates.
#
#NYI sub _get_image_properties {

#NYI     my $self     = shift;
#NYI     my $filename = shift;

#NYI     my $type;
#NYI     my $width;
#NYI     my $height;
#NYI     my $x_dpi = 96;
#NYI     my $y_dpi = 96;
#NYI     my $image_name;


#NYI     ( $image_name ) = fileparse( $filename );

#NYI     # Open the image file and import the data.
#NYI     my $fh = FileHandle->new( $filename );
#NYI     fail "Couldn't import $filename: $!" unless defined $fh;
#NYI     binmode $fh;

#NYI     # Slurp the file into a string and do some size calcs.
#NYI     my $data = do { local $/; <$fh> };
#NYI     my $size = length $data;


#NYI     if ( unpack( 'x A3', $data ) eq 'PNG' ) {

#NYI         # Test for PNGs.
#NYI         ( $type, $width, $height, $x_dpi, $y_dpi ) =
#NYI           $self->_process_png( $data, $filename );

#NYI         $self->{_image_types}->{png} = 1;
#NYI     }
#NYI     elsif ( unpack( 'n', $data ) == 0xFFD8 ) {

#NYI         # Test for JPEG files.
#NYI         ( $type, $width, $height, $x_dpi, $y_dpi ) =
#NYI           $self->_process_jpg( $data, $filename );

#NYI         $self->{_image_types}->{jpeg} = 1;
#NYI     }
#NYI     elsif ( unpack( 'A2', $data ) eq 'BM' ) {

#NYI         # Test for BMPs.
#NYI         ( $type, $width, $height ) = $self->_process_bmp( $data, $filename );

#NYI         $self->{_image_types}->{bmp} = 1;
#NYI     }
#NYI     else {
#NYI         fail "Unsupported image format for file: $filename\n";
#NYI     }

#NYI     push @{ $self->{_images} }, [ $filename, $type ];

#NYI     # Set a default dpi for images with 0 dpi.
#NYI     $x_dpi = 96 if $x_dpi == 0;
#NYI     $y_dpi = 96 if $y_dpi == 0;

#NYI     $fh->close;

#NYI     return ( $type, $width, $height, $image_name, $x_dpi, $y_dpi );
#NYI }


###############################################################################
#
# _process_png()
#
# Extract width and height information from a PNG file.
#
#NYI sub _process_png {

#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];

#NYI     my $type   = 'png';
#NYI     my $width  = 0;
#NYI     my $height = 0;
#NYI     my $x_dpi  = 96;
#NYI     my $y_dpi  = 96;

#NYI     my $offset      = 8;
#NYI     my $data_length = length $data;

#NYI     # Search through the image data to read the height and width in the
#NYI     # IHDR element. Also read the DPI in the pHYs element.
#NYI     while ( $offset < $data_length ) {

#NYI         my $length = unpack "N",  substr $data, $offset + 0, 4;
#NYI         my $type   = unpack "A4", substr $data, $offset + 4, 4;

#NYI         if ( $type eq "IHDR" ) {
#NYI             $width  = unpack "N", substr $data, $offset + 8,  4;
#NYI             $height = unpack "N", substr $data, $offset + 12, 4;
#NYI         }

#NYI         if ( $type eq "pHYs" ) {
#NYI             my $x_ppu = unpack "N", substr $data, $offset + 8,  4;
#NYI             my $y_ppu = unpack "N", substr $data, $offset + 12, 4;
#NYI             my $units = unpack "C", substr $data, $offset + 16, 1;

#NYI             if ( $units == 1 ) {
#NYI                 $x_dpi = $x_ppu * 0.0254;
#NYI                 $y_dpi = $y_ppu * 0.0254;
#NYI             }
#NYI         }

#NYI         $offset = $offset + $length + 12;

#NYI         last if $type eq "IEND";
#NYI     }

#NYI     if ( not defined $height ) {
#NYI         fail "$filename: no size data found in png image.\n";
#NYI     }

#NYI     return ( $type, $width, $height, $x_dpi, $y_dpi );
#NYI }


###############################################################################
#
# _process_bmp()
#
# Extract width and height information from a BMP file.
#
# Most of the checks came from old Spredsheet::WriteExcel code.
#
#NYI sub _process_bmp {

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
# _process_jpg()
#
# Extract width and height information from a JPEG file.
#
#NYI sub _process_jpg {

#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];
#NYI     my $type     = 'jpeg';
#NYI     my $x_dpi    = 96;
#NYI     my $y_dpi    = 96;
#NYI     my $width;
#NYI     my $height;

#NYI     my $offset      = 2;
#NYI     my $data_length = length $data;

#NYI     # Search through the image data to read the height and width in the
#NYI     # 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element.
#NYI     while ( $offset < $data_length ) {

#NYI         my $marker = unpack "n", substr $data, $offset + 0, 2;
#NYI         my $length = unpack "n", substr $data, $offset + 2, 2;

#NYI         if ( $marker == 0xFFC0 || $marker == 0xFFC2 ) {
#NYI             $height = unpack "n", substr $data, $offset + 5, 2;
#NYI             $width  = unpack "n", substr $data, $offset + 7, 2;
#NYI         }

#NYI         if ( $marker == 0xFFE0 ) {
#NYI             my $units     = unpack "C", substr $data, $offset + 11, 1;
#NYI             my $x_density = unpack "n", substr $data, $offset + 12, 2;
#NYI             my $y_density = unpack "n", substr $data, $offset + 14, 2;

#NYI             if ( $units == 1 ) {
#NYI                 $x_dpi = $x_density;
#NYI                 $y_dpi = $y_density;
#NYI             }

#NYI             if ( $units == 2 ) {
#NYI                 $x_dpi = $x_density * 2.54;
#NYI                 $y_dpi = $y_density * 2.54;
#NYI             }
#NYI         }

#NYI         $offset = $offset + $length + 2;
#NYI         last if $marker == 0xFFDA;
#NYI     }

#NYI     if ( not defined $height ) {
#NYI         fail "$filename: no size data found in jpeg image.\n";
#NYI     }

#NYI     return ( $type, $width, $height, $x_dpi, $y_dpi );
#NYI }


#NYI ###############################################################################
#NYI #
#NYI # _get_sheet_index()
#NYI #
#NYI # Convert a sheet name to its index. Return undef otherwise.
#NYI #
#NYI sub _get_sheet_index {

#NYI     my $self        = shift;
#NYI     my $sheetname   = shift;
#NYI     my $sheet_index = undef;

#NYI     $sheetname =~ s/^'//;
#NYI     $sheetname =~ s/'$//;

#NYI     if ( exists $self->{_sheetnames}->{$sheetname} ) {
#NYI         return $self->{_sheetnames}->{$sheetname}->{_index};
#NYI     }
#NYI     else {
#NYI         return undef;
#NYI     }
#NYI }


###############################################################################
#
# set_optimization()
#
# Set the speed/memory optimisation level.
#
method set_optimization($level = 1) {

    fail "set_optimization() must be called before add_worksheet()"
      if @!worksheets.elems == 0;

    $!optimization = $level;
}


#NYI ###############################################################################
#NYI #
#NYI # Deprecated methods for backwards compatibility.
#NYI #
#NYI ###############################################################################

#NYI # No longer required by Excel::Writer::XLSX.
#NYI sub compatibility_mode { }
#NYI sub set_codepage       { }


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_workbook()
#
# Write <workbook> element.
#
method write_workbook {

    my $schema  = 'http://schemas.openxmlformats.org';
    my $xmlns   = $schema ~ '/spreadsheetml/2006/main';
    my $xmlns_r = $schema ~ '/officeDocument/2006/relationships';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns_r,
    );

    self.xml_start_tag( 'workbook', @attributes );
}


###############################################################################
#
# write_file_version()
#
# Write the <fileVersion> element.
#
method write_file_version {

    my $app_name      = 'xl';
    my $last_edited   = 4;
    my $lowest_edited = 4;
    my $rup_build     = 4505;

    my @attributes = (
        'appName'      => $app_name,
        'lastEdited'   => $last_edited,
        'lowestEdited' => $lowest_edited,
        'rupBuild'     => $rup_build,
    );

    if $!vba_project {
        push @attributes, codeName => '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}';
    }

    self.xml_empty_tag( 'fileVersion', @attributes );
}


###############################################################################
#
# _write_workbook_pr()
#
# Write <workbookPr> element.
#
#NYI sub _write_workbook_pr {

#NYI     my $self                   = shift;
#NYI     my $date_1904              = $self->{_date_1904};
#NYI     my $show_ink_annotation    = 0;
#NYI     my $auto_compress_pictures = 0;
#NYI     my $default_theme_version  = 124226;
#NYI     my $codename               = $self->{_vba_codename};
#NYI     my @attributes;

#NYI     push @attributes, ( 'codeName' => $codename ) if $codename;
#NYI     push @attributes, ( 'date1904' => 1 )         if $date_1904;
#NYI     push @attributes, ( 'defaultThemeVersion' => $default_theme_version );

#NYI     $self->xml_empty_tag( 'workbookPr', @attributes );
#NYI }


###############################################################################
#
# _write_book_views()
#
# Write <bookViews> element.
#
#NYI sub _write_book_views {

#NYI     my $self = shift;

#NYI     $self->xml_start_tag( 'bookViews' );
#NYI     $self->_write_workbook_view();
#NYI     $self->xml_end_tag( 'bookViews' );
#NYI }

###############################################################################
#
# _write_workbook_view()
#
# Write <workbookView> element.
#
#NYI sub _write_workbook_view {

#NYI     my $self          = shift;
#NYI     my $x_window      = $self->{_x_window};
#NYI     my $y_window      = $self->{_y_window};
#NYI     my $window_width  = $self->{_window_width};
#NYI     my $window_height = $self->{_window_height};
#NYI     my $tab_ratio     = $self->{_tab_ratio};
#NYI     my $active_tab    = $self->{_activesheet};
#NYI     my $first_sheet   = $self->{_firstsheet};

#NYI     my @attributes = (
#NYI         'xWindow'      => $x_window,
#NYI         'yWindow'      => $y_window,
#NYI         'windowWidth'  => $window_width,
#NYI         'windowHeight' => $window_height,
#NYI     );

#NYI     # Store the tabRatio attribute when it isn't the default.
#NYI     push @attributes, ( tabRatio => $tab_ratio ) if $tab_ratio != 500;

#NYI     # Store the firstSheet attribute when it isn't the default.
#NYI     push @attributes, ( firstSheet => $first_sheet + 1 ) if $first_sheet > 0;

#NYI     # Store the activeTab attribute when it isn't the first sheet.
#NYI     push @attributes, ( activeTab => $active_tab ) if $active_tab > 0;

#NYI     $self->xml_empty_tag( 'workbookView', @attributes );
#NYI }

###############################################################################
#
# _write_sheets()
#
# Write <sheets> element.
#
#NYI sub _write_sheets {

#NYI     my $self   = shift;
#NYI     my $id_num = 1;

#NYI     $self->xml_start_tag( 'sheets' );

#NYI     for my $worksheet ( @{ $self->{_worksheets} } ) {
#NYI         $self->_write_sheet( $worksheet->{_name}, $id_num++,
#NYI             $worksheet->{_hidden} );
#NYI     }

#NYI     $self->xml_end_tag( 'sheets' );
#NYI }


###############################################################################
#
# _write_sheet()
#
# Write <sheet> element.
#
#NYI sub _write_sheet {

#NYI     my $self     = shift;
#NYI     my $name     = shift;
#NYI     my $sheet_id = shift;
#NYI     my $hidden   = shift;
#NYI     my $r_id     = 'rId' . $sheet_id;

#NYI     my @attributes = (
#NYI         'name'    => $name,
#NYI         'sheetId' => $sheet_id,
#NYI     );

#NYI     push @attributes, ( 'state' => 'hidden' ) if $hidden;
#NYI     push @attributes, ( 'r:id' => $r_id );


#NYI     $self->xml_empty_tag( 'sheet', @attributes );
#NYI }


###############################################################################
#
# _write_calc_pr()
#
# Write <calcPr> element.
#
#NYI sub _write_calc_pr {

#NYI     my $self            = shift;
#NYI     my $calc_id         = $self->{_calc_id};
#NYI     my $concurrent_calc = 0;

#NYI     my @attributes = ( calcId => $calc_id );

#NYI     if ( $self->{_calc_mode} eq 'manual' ) {
#NYI         push @attributes, 'calcMode'   => 'manual';
#NYI         push @attributes, 'calcOnSave' => 0;
#NYI     }
#NYI     elsif ( $self->{_calc_mode} eq 'autoNoTable' ) {
#NYI         push @attributes, calcMode => 'autoNoTable';
#NYI     }

#NYI     if ( $self->{_calc_on_load} ) {
#NYI         push @attributes, 'fullCalcOnLoad' => 1;
#NYI     }


#NYI     $self->xml_empty_tag( 'calcPr', @attributes );
#NYI }


###############################################################################
#
# _write_ext_lst()
#
# Write <extLst> element.
#
#NYI sub _write_ext_lst {

#NYI     my $self = shift;

#NYI     $self->xml_start_tag( 'extLst' );
#NYI     $self->_write_ext();
#NYI     $self->xml_end_tag( 'extLst' );
#NYI }


###############################################################################
#
# _write_ext()
#
# Write <ext> element.
#
#NYI sub _write_ext {

#NYI     my $self     = shift;
#NYI     my $xmlns_mx = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
#NYI     my $uri      = 'http://schemas.microsoft.com/office/mac/excel/2008/main';

#NYI     my @attributes = (
#NYI         'xmlns:mx' => $xmlns_mx,
#NYI         'uri'      => $uri,
#NYI     );

#NYI     $self->xml_start_tag( 'ext', @attributes );
#NYI     $self->_write_mx_arch_id();
#NYI     $self->xml_end_tag( 'ext' );
#NYI }

###############################################################################
#
# _write_mx_arch_id()
#
# Write <mx:ArchID> element.
#
#NYI sub _write_mx_arch_id {

#NYI     my $self  = shift;
#NYI     my $Flags = 2;

#NYI     my @attributes = ( 'Flags' => $Flags, );

#NYI     $self->xml_empty_tag( 'mx:ArchID', @attributes );
#NYI }


##############################################################################
#
# _write_defined_names()
#
# Write the <definedNames> element.
#
#NYI sub _write_defined_names {

#NYI     my $self = shift;

#NYI     return unless @{ $self->{_defined_names} };

#NYI     $self->xml_start_tag( 'definedNames' );

#NYI     for my $aref ( @{ $self->{_defined_names} } ) {
#NYI         $self->_write_defined_name( $aref );
#NYI     }

#NYI     $self->xml_end_tag( 'definedNames' );
#NYI }


##############################################################################
#
# _write_defined_name()
#
# Write the <definedName> element.
#
#NYI sub _write_defined_name {

#NYI     my $self = shift;
#NYI     my $data = shift;

#NYI     my $name   = $data->[0];
#NYI     my $id     = $data->[1];
#NYI     my $range  = $data->[2];
#NYI     my $hidden = $data->[3];

#NYI     my @attributes = ( 'name' => $name );

#NYI     push @attributes, ( 'localSheetId' => $id ) if $id != -1;
#NYI     push @attributes, ( 'hidden'       => 1 )   if $hidden;

#NYI     $self->xml_data_element( 'definedName', $range, @attributes );
#NYI }

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
