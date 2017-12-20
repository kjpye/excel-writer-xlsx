#NYI package Excel::Writer::XLSX::Workbook;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Workbook - A class for writing Excel Workbooks.
#NYI #
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
#NYI use IO::File;
#NYI use File::Find;
#NYI use File::Temp qw(tempfile);
#NYI use File::Basename 'fileparse';
#NYI use Archive::Zip;
#NYI use Excel::Writer::XLSX::Worksheet;
#NYI use Excel::Writer::XLSX::Chartsheet;
#NYI use Excel::Writer::XLSX::Format;
#NYI use Excel::Writer::XLSX::Shape;
#NYI use Excel::Writer::XLSX::Chart;
#NYI use Excel::Writer::XLSX::Package::Packager;
#NYI use Excel::Writer::XLSX::Package::XMLwriter;
#NYI use Excel::Writer::XLSX::Utility qw(xl_cell_to_rowcol xl_rowcol_to_cell);
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
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter->new();
#NYI 
#NYI     $self->{_filename} = $_[0] || '';
#NYI     my $options = $_[1] || {};
#NYI 
#NYI     $self->{_tempdir}            = undef;
#NYI     $self->{_date_1904}          = 0;
#NYI     $self->{_activesheet}        = 0;
#NYI     $self->{_firstsheet}         = 0;
#NYI     $self->{_selected}           = 0;
#NYI     $self->{_fileclosed}         = 0;
#NYI     $self->{_filehandle}         = undef;
#NYI     $self->{_internal_fh}        = 0;
#NYI     $self->{_sheet_name}         = 'Sheet';
#NYI     $self->{_chart_name}         = 'Chart';
#NYI     $self->{_sheetname_count}    = 0;
#NYI     $self->{_chartname_count}    = 0;
#NYI     $self->{_worksheets}         = [];
#NYI     $self->{_charts}             = [];
#NYI     $self->{_drawings}           = [];
#NYI     $self->{_sheetnames}         = {};
#NYI     $self->{_formats}            = [];
#NYI     $self->{_xf_formats}         = [];
#NYI     $self->{_xf_format_indices}  = {};
#NYI     $self->{_dxf_formats}        = [];
#NYI     $self->{_dxf_format_indices} = {};
#NYI     $self->{_palette}            = [];
#NYI     $self->{_font_count}         = 0;
#NYI     $self->{_num_format_count}   = 0;
#NYI     $self->{_defined_names}      = [];
#NYI     $self->{_named_ranges}       = [];
#NYI     $self->{_custom_colors}      = [];
#NYI     $self->{_doc_properties}     = {};
#NYI     $self->{_custom_properties}  = [];
#NYI     $self->{_createtime}         = [ gmtime() ];
#NYI     $self->{_num_vml_files}      = 0;
#NYI     $self->{_num_comment_files}  = 0;
#NYI     $self->{_optimization}       = 0;
#NYI     $self->{_x_window}           = 240;
#NYI     $self->{_y_window}           = 15;
#NYI     $self->{_window_width}       = 16095;
#NYI     $self->{_window_height}      = 9660;
#NYI     $self->{_tab_ratio}          = 500;
#NYI     $self->{_excel2003_style}    = 0;
#NYI 
#NYI     $self->{_default_format_properties} = {};
#NYI 
#NYI     if ( exists $options->{tempdir} ) {
#NYI         $self->{_tempdir} = $options->{tempdir};
#NYI     }
#NYI 
#NYI     if ( exists $options->{date_1904} ) {
#NYI         $self->{_date_1904} = $options->{date_1904};
#NYI     }
#NYI 
#NYI     if ( exists $options->{optimization} ) {
#NYI         $self->{_optimization} = $options->{optimization};
#NYI     }
#NYI 
#NYI     if ( exists $options->{default_format_properties} ) {
#NYI         $self->{_default_format_properties} =
#NYI           $options->{default_format_properties};
#NYI     }
#NYI 
#NYI     if ( exists $options->{excel2003_style} ) {
#NYI         $self->{_excel2003_style} = 1;
#NYI     }
#NYI 
#NYI     # Structures for the shared strings data.
#NYI     $self->{_str_total}  = 0;
#NYI     $self->{_str_unique} = 0;
#NYI     $self->{_str_table}  = {};
#NYI     $self->{_str_array}  = [];
#NYI 
#NYI     # Formula calculation default settings.
#NYI     $self->{_calc_id}      = 124519;
#NYI     $self->{_calc_mode}    = 'auto';
#NYI     $self->{_calc_on_load} = 1;
#NYI 
#NYI 
#NYI     bless $self, $class;
#NYI 
#NYI     # Add the default cell format.
#NYI     if ( $self->{_excel2003_style} ) {
#NYI         $self->add_format( xf_index => 0, font_family => 0 );
#NYI     }
#NYI     else {
#NYI         $self->add_format( xf_index => 0 );
#NYI     }
#NYI 
#NYI     # Check for a filename unless it is an existing filehandle
#NYI     if ( not ref $self->{_filename} and $self->{_filename} eq '' ) {
#NYI         carp 'Filename required by Excel::Writer::XLSX->new()';
#NYI         return undef;
#NYI     }
#NYI 
#NYI 
#NYI     # If filename is a reference we assume that it is a valid filehandle.
#NYI     if ( ref $self->{_filename} ) {
#NYI 
#NYI         $self->{_filehandle}  = $self->{_filename};
#NYI         $self->{_internal_fh} = 0;
#NYI     }
#NYI     elsif ( $self->{_filename} eq '-' ) {
#NYI 
#NYI         # Support special filename/filehandle '-' for backward compatibility.
#NYI         binmode STDOUT;
#NYI         $self->{_filehandle}  = \*STDOUT;
#NYI         $self->{_internal_fh} = 0;
#NYI     }
#NYI     else {
#NYI         my $fh = IO::File->new( $self->{_filename}, 'w' );
#NYI 
#NYI         return undef unless defined $fh;
#NYI 
#NYI         $self->{_filehandle}  = $fh;
#NYI         $self->{_internal_fh} = 1;
#NYI     }
#NYI 
#NYI 
#NYI     # Set colour palette.
#NYI     $self->set_color_palette();
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
#NYI     # Prepare format object for passing to Style.pm.
#NYI     $self->_prepare_format_properties();
#NYI 
#NYI     $self->xml_declaration;
#NYI 
#NYI     # Write the root workbook element.
#NYI     $self->_write_workbook();
#NYI 
#NYI     # Write the XLSX file version.
#NYI     $self->_write_file_version();
#NYI 
#NYI     # Write the workbook properties.
#NYI     $self->_write_workbook_pr();
#NYI 
#NYI     # Write the workbook view properties.
#NYI     $self->_write_book_views();
#NYI 
#NYI     # Write the worksheet names and ids.
#NYI     $self->_write_sheets();
#NYI 
#NYI     # Write the workbook defined names.
#NYI     $self->_write_defined_names();
#NYI 
#NYI     # Write the workbook calculation properties.
#NYI     $self->_write_calc_pr();
#NYI 
#NYI     # Write the workbook extension storage.
#NYI     #$self->_write_ext_lst();
#NYI 
#NYI     # Close the workbook tag.
#NYI     $self->xml_end_tag( 'workbook' );
#NYI 
#NYI     # Close the XML writer filehandle.
#NYI     $self->xml_get_fh()->close();
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # close()
#NYI #
#NYI # Calls finalization methods.
#NYI #
#NYI sub close {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # In case close() is called twice, by user and by DESTROY.
#NYI     return if $self->{_fileclosed};
#NYI 
#NYI     # Test filehandle in case new() failed and the user didn't check.
#NYI     return undef if !defined $self->{_filehandle};
#NYI 
#NYI     $self->{_fileclosed} = 1;
#NYI     $self->_store_workbook();
#NYI 
#NYI     # Return the file close value.
#NYI     if ( $self->{_internal_fh} ) {
#NYI         return $self->{_filehandle}->close();
#NYI     }
#NYI     else {
#NYI         # Return true and let users deal with their own filehandles.
#NYI         return 1;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # DESTROY()
#NYI #
#NYI # Close the workbook if it hasn't already been explicitly closed.
#NYI #
#NYI sub DESTROY {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     local ( $@, $!, $^E, $? );
#NYI 
#NYI     $self->close() if not $self->{_fileclosed};
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # sheets(slice,...)
#NYI #
#NYI # An accessor for the _worksheets[] array
#NYI #
#NYI # Returns: an optionally sliced list of the worksheet objects in a workbook.
#NYI #
#NYI sub sheets {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     if ( @_ ) {
#NYI 
#NYI         # Return a slice of the array
#NYI         return @{ $self->{_worksheets} }[@_];
#NYI     }
#NYI     else {
#NYI 
#NYI         # Return the entire list
#NYI         return @{ $self->{_worksheets} };
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_worksheet_by_name(name)
#NYI #
#NYI # Return a worksheet object in the workbook using the sheetname.
#NYI #
#NYI sub get_worksheet_by_name {
#NYI 
#NYI     my $self      = shift;
#NYI     my $sheetname = shift;
#NYI 
#NYI     return undef if not defined $sheetname;
#NYI 
#NYI     return $self->{_sheetnames}->{$sheetname};
#NYI }
#NYI 
#NYI 
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
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return $self->{_worksheets};
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_worksheet($name)
#NYI #
#NYI # Add a new worksheet to the Excel workbook.
#NYI #
#NYI # Returns: reference to a worksheet object
#NYI #
#NYI sub add_worksheet {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = @{ $self->{_worksheets} };
#NYI     my $name  = $self->_check_sheetname( $_[0] );
#NYI     my $fh    = undef;
#NYI 
#NYI     # Porters take note, the following scheme of passing references to Workbook
#NYI     # data (in the \$self->{_foo} cases) instead of a reference to the Workbook
#NYI     # itself is a workaround to avoid circular references between Workbook and
#NYI     # Worksheet objects. Feel free to implement this in any way the suits your
#NYI     # language.
#NYI     #
#NYI     my @init_data = (
#NYI         $fh,
#NYI         $name,
#NYI         $index,
#NYI 
#NYI         \$self->{_activesheet},
#NYI         \$self->{_firstsheet},
#NYI 
#NYI         \$self->{_str_total},
#NYI         \$self->{_str_unique},
#NYI         \$self->{_str_table},
#NYI 
#NYI         $self->{_date_1904},
#NYI         $self->{_palette},
#NYI         $self->{_optimization},
#NYI         $self->{_tempdir},
#NYI         $self->{_excel2003_style},
#NYI 
#NYI     );
#NYI 
#NYI     my $worksheet = Excel::Writer::XLSX::Worksheet->new( @init_data );
#NYI     $self->{_worksheets}->[$index] = $worksheet;
#NYI     $self->{_sheetnames}->{$name} = $worksheet;
#NYI 
#NYI     return $worksheet;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_chart( %args )
#NYI #
#NYI # Create a chart for embedding or as a new sheet.
#NYI #
#NYI sub add_chart {
#NYI 
#NYI     my $self  = shift;
#NYI     my %arg   = @_;
#NYI     my $name  = '';
#NYI     my $index = @{ $self->{_worksheets} };
#NYI     my $fh    = undef;
#NYI 
#NYI     # Type must be specified so we can create the required chart instance.
#NYI     my $type = $arg{type};
#NYI     if ( !defined $type ) {
#NYI         croak "Must define chart type in add_chart()";
#NYI     }
#NYI 
#NYI     # Ensure that the chart defaults to non embedded.
#NYI     my $embedded = $arg{embedded} || 0;
#NYI 
#NYI     # Check the worksheet name for non-embedded charts.
#NYI     if ( !$embedded ) {
#NYI         $name = $self->_check_sheetname( $arg{name}, 1 );
#NYI     }
#NYI 
#NYI 
#NYI     my @init_data = (
#NYI 
#NYI         $fh,
#NYI         $name,
#NYI         $index,
#NYI 
#NYI         \$self->{_activesheet},
#NYI         \$self->{_firstsheet},
#NYI 
#NYI         \$self->{_str_total},
#NYI         \$self->{_str_unique},
#NYI         \$self->{_str_table},
#NYI 
#NYI         $self->{_date_1904},
#NYI         $self->{_palette},
#NYI         $self->{_optimization},
#NYI     );
#NYI 
#NYI 
#NYI     my $chart = Excel::Writer::XLSX::Chart->factory( $type, $arg{subtype} );
#NYI 
#NYI     # If the chart isn't embedded let the workbook control it.
#NYI     if ( !$embedded ) {
#NYI 
#NYI         my $drawing    = Excel::Writer::XLSX::Drawing->new();
#NYI         my $chartsheet = Excel::Writer::XLSX::Chartsheet->new( @init_data );
#NYI 
#NYI         $chart->{_palette} = $self->{_palette};
#NYI 
#NYI         $chartsheet->{_chart}   = $chart;
#NYI         $chartsheet->{_drawing} = $drawing;
#NYI 
#NYI         $self->{_worksheets}->[$index] = $chartsheet;
#NYI         $self->{_sheetnames}->{$name} = $chartsheet;
#NYI 
#NYI         push @{ $self->{_charts} }, $chart;
#NYI 
#NYI         return $chartsheet;
#NYI     }
#NYI     else {
#NYI 
#NYI         # Set the embedded chart name if present.
#NYI         $chart->{_chart_name} = $arg{name} if $arg{name};
#NYI 
#NYI         # Set index to 0 so that the activate() and set_first_sheet() methods
#NYI         # point back to the first worksheet if used for embedded charts.
#NYI         $chart->{_index}   = 0;
#NYI         $chart->{_palette} = $self->{_palette};
#NYI         $chart->_set_embedded_config_data();
#NYI         push @{ $self->{_charts} }, $chart;
#NYI 
#NYI         return $chart;
#NYI     }
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _check_sheetname( $name )
#NYI #
#NYI # Check for valid worksheet names. We check the length, if it contains any
#NYI # invalid characters and if the name is unique in the workbook.
#NYI #
#NYI sub _check_sheetname {
#NYI 
#NYI     my $self         = shift;
#NYI     my $name         = shift || "";
#NYI     my $chart        = shift || 0;
#NYI     my $invalid_char = qr([\[\]:*?/\\]);
#NYI 
#NYI     # Increment the Sheet/Chart number used for default sheet names below.
#NYI     if ( $chart ) {
#NYI         $self->{_chartname_count}++;
#NYI     }
#NYI     else {
#NYI         $self->{_sheetname_count}++;
#NYI     }
#NYI 
#NYI     # Supply default Sheet/Chart name if none has been defined.
#NYI     if ( $name eq "" ) {
#NYI 
#NYI         if ( $chart ) {
#NYI             $name = $self->{_chart_name} . $self->{_chartname_count};
#NYI         }
#NYI         else {
#NYI             $name = $self->{_sheet_name} . $self->{_sheetname_count};
#NYI         }
#NYI     }
#NYI 
#NYI     # Check that sheet name is <= 31. Excel limit.
#NYI     croak "Sheetname $name must be <= 31 chars" if length $name > 31;
#NYI 
#NYI     # Check that sheetname doesn't contain any invalid characters
#NYI     if ( $name =~ $invalid_char ) {
#NYI         croak 'Invalid character []:*?/\\ in worksheet name: ' . $name;
#NYI     }
#NYI 
#NYI     # Check that the worksheet name doesn't already exist since this is a fatal
#NYI     # error in Excel 97. The check must also exclude case insensitive matches.
#NYI     foreach my $worksheet ( @{ $self->{_worksheets} } ) {
#NYI         my $name_a = $name;
#NYI         my $name_b = $worksheet->{_name};
#NYI 
#NYI         if ( lc( $name_a ) eq lc( $name_b ) ) {
#NYI             croak "Worksheet name '$name', with case ignored, is already used.";
#NYI         }
#NYI     }
#NYI 
#NYI     return $name;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_format(%properties)
#NYI #
#NYI # Add a new format to the Excel workbook.
#NYI #
#NYI sub add_format {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @init_data =
#NYI       ( \$self->{_xf_format_indices}, \$self->{_dxf_format_indices} );
#NYI 
#NYI     # Change default format style for Excel2003/XLS format.
#NYI     if ( $self->{_excel2003_style} ) {
#NYI         push @init_data, ( font => 'Arial', size => 10, theme => -1 );
#NYI     }
#NYI 
#NYI     # Add the default format properties.
#NYI     push @init_data, %{ $self->{_default_format_properties} };
#NYI 
#NYI     # Add the user defined properties.
#NYI     push @init_data, @_;
#NYI 
#NYI     my $format = Excel::Writer::XLSX::Format->new( @init_data );
#NYI 
#NYI     push @{ $self->{_formats} }, $format;    # Store format reference
#NYI 
#NYI     return $format;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_shape(%properties)
#NYI #
#NYI # Add a new shape to the Excel workbook.
#NYI #
#NYI sub add_shape {
#NYI 
#NYI     my $self  = shift;
#NYI     my $fh    = undef;
#NYI     my $shape = Excel::Writer::XLSX::Shape->new( $fh, @_ );
#NYI 
#NYI     $shape->{_palette} = $self->{_palette};
#NYI 
#NYI 
#NYI     push @{ $self->{_shapes} }, $shape;    # Store shape reference.
#NYI 
#NYI     return $shape;
#NYI }
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_1904()
#NYI #
#NYI # Set the date system: 0 = 1900 (the default), 1 = 1904
#NYI #
#NYI sub set_1904 {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     if ( defined( $_[0] ) ) {
#NYI         $self->{_date_1904} = $_[0];
#NYI     }
#NYI     else {
#NYI         $self->{_date_1904} = 1;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # get_1904()
#NYI #
#NYI # Return the date system: 0 = 1900, 1 = 1904
#NYI #
#NYI sub get_1904 {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return $self->{_date_1904};
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_custom_color()
#NYI #
#NYI # Change the RGB components of the elements in the colour palette.
#NYI #
#NYI sub set_custom_color {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI 
#NYI     # Match a HTML #xxyyzz style parameter
#NYI     if ( defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
#NYI         @_ = ( $_[0], hex $1, hex $2, hex $3 );
#NYI     }
#NYI 
#NYI 
#NYI     my $index = $_[0] || 0;
#NYI     my $red   = $_[1] || 0;
#NYI     my $green = $_[2] || 0;
#NYI     my $blue  = $_[3] || 0;
#NYI 
#NYI     my $aref = $self->{_palette};
#NYI 
#NYI     # Check that the colour index is the right range
#NYI     if ( $index < 8 or $index > 64 ) {
#NYI         carp "Color index $index outside range: 8 <= index <= 64";
#NYI         return 0;
#NYI     }
#NYI 
#NYI     # Check that the colour components are in the right range
#NYI     if (   ( $red < 0 or $red > 255 )
#NYI         || ( $green < 0 or $green > 255 )
#NYI         || ( $blue < 0  or $blue > 255 ) )
#NYI     {
#NYI         carp "Color component outside range: 0 <= color <= 255";
#NYI         return 0;
#NYI     }
#NYI 
#NYI     $index -= 8;    # Adjust colour index (wingless dragonfly)
#NYI 
#NYI     # Set the RGB value.
#NYI     my @rgb = ( $red, $green, $blue );
#NYI     $aref->[$index] = [@rgb];
#NYI 
#NYI     # Store the custom colors for the style.xml file.
#NYI     push @{ $self->{_custom_colors} }, sprintf "FF%02X%02X%02X", @rgb;
#NYI 
#NYI     return $index + 8;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_color_palette()
#NYI #
#NYI # Sets the colour palette to the Excel defaults.
#NYI #
#NYI sub set_color_palette {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->{_palette} = [
#NYI         [ 0x00, 0x00, 0x00, 0x00 ],    # 8
#NYI         [ 0xff, 0xff, 0xff, 0x00 ],    # 9
#NYI         [ 0xff, 0x00, 0x00, 0x00 ],    # 10
#NYI         [ 0x00, 0xff, 0x00, 0x00 ],    # 11
#NYI         [ 0x00, 0x00, 0xff, 0x00 ],    # 12
#NYI         [ 0xff, 0xff, 0x00, 0x00 ],    # 13
#NYI         [ 0xff, 0x00, 0xff, 0x00 ],    # 14
#NYI         [ 0x00, 0xff, 0xff, 0x00 ],    # 15
#NYI         [ 0x80, 0x00, 0x00, 0x00 ],    # 16
#NYI         [ 0x00, 0x80, 0x00, 0x00 ],    # 17
#NYI         [ 0x00, 0x00, 0x80, 0x00 ],    # 18
#NYI         [ 0x80, 0x80, 0x00, 0x00 ],    # 19
#NYI         [ 0x80, 0x00, 0x80, 0x00 ],    # 20
#NYI         [ 0x00, 0x80, 0x80, 0x00 ],    # 21
#NYI         [ 0xc0, 0xc0, 0xc0, 0x00 ],    # 22
#NYI         [ 0x80, 0x80, 0x80, 0x00 ],    # 23
#NYI         [ 0x99, 0x99, 0xff, 0x00 ],    # 24
#NYI         [ 0x99, 0x33, 0x66, 0x00 ],    # 25
#NYI         [ 0xff, 0xff, 0xcc, 0x00 ],    # 26
#NYI         [ 0xcc, 0xff, 0xff, 0x00 ],    # 27
#NYI         [ 0x66, 0x00, 0x66, 0x00 ],    # 28
#NYI         [ 0xff, 0x80, 0x80, 0x00 ],    # 29
#NYI         [ 0x00, 0x66, 0xcc, 0x00 ],    # 30
#NYI         [ 0xcc, 0xcc, 0xff, 0x00 ],    # 31
#NYI         [ 0x00, 0x00, 0x80, 0x00 ],    # 32
#NYI         [ 0xff, 0x00, 0xff, 0x00 ],    # 33
#NYI         [ 0xff, 0xff, 0x00, 0x00 ],    # 34
#NYI         [ 0x00, 0xff, 0xff, 0x00 ],    # 35
#NYI         [ 0x80, 0x00, 0x80, 0x00 ],    # 36
#NYI         [ 0x80, 0x00, 0x00, 0x00 ],    # 37
#NYI         [ 0x00, 0x80, 0x80, 0x00 ],    # 38
#NYI         [ 0x00, 0x00, 0xff, 0x00 ],    # 39
#NYI         [ 0x00, 0xcc, 0xff, 0x00 ],    # 40
#NYI         [ 0xcc, 0xff, 0xff, 0x00 ],    # 41
#NYI         [ 0xcc, 0xff, 0xcc, 0x00 ],    # 42
#NYI         [ 0xff, 0xff, 0x99, 0x00 ],    # 43
#NYI         [ 0x99, 0xcc, 0xff, 0x00 ],    # 44
#NYI         [ 0xff, 0x99, 0xcc, 0x00 ],    # 45
#NYI         [ 0xcc, 0x99, 0xff, 0x00 ],    # 46
#NYI         [ 0xff, 0xcc, 0x99, 0x00 ],    # 47
#NYI         [ 0x33, 0x66, 0xff, 0x00 ],    # 48
#NYI         [ 0x33, 0xcc, 0xcc, 0x00 ],    # 49
#NYI         [ 0x99, 0xcc, 0x00, 0x00 ],    # 50
#NYI         [ 0xff, 0xcc, 0x00, 0x00 ],    # 51
#NYI         [ 0xff, 0x99, 0x00, 0x00 ],    # 52
#NYI         [ 0xff, 0x66, 0x00, 0x00 ],    # 53
#NYI         [ 0x66, 0x66, 0x99, 0x00 ],    # 54
#NYI         [ 0x96, 0x96, 0x96, 0x00 ],    # 55
#NYI         [ 0x00, 0x33, 0x66, 0x00 ],    # 56
#NYI         [ 0x33, 0x99, 0x66, 0x00 ],    # 57
#NYI         [ 0x00, 0x33, 0x00, 0x00 ],    # 58
#NYI         [ 0x33, 0x33, 0x00, 0x00 ],    # 59
#NYI         [ 0x99, 0x33, 0x00, 0x00 ],    # 60
#NYI         [ 0x99, 0x33, 0x66, 0x00 ],    # 61
#NYI         [ 0x33, 0x33, 0x99, 0x00 ],    # 62
#NYI         [ 0x33, 0x33, 0x33, 0x00 ],    # 63
#NYI     ];
#NYI 
#NYI     return 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_tempdir()
#NYI #
#NYI # Change the default temp directory.
#NYI #
#NYI sub set_tempdir {
#NYI 
#NYI     my $self = shift;
#NYI     my $dir  = shift;
#NYI 
#NYI     croak "$dir is not a valid directory" if defined $dir and not -d $dir;
#NYI 
#NYI     $self->{_tempdir} = $dir;
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # define_name()
#NYI #
#NYI # Create a defined name in Excel. We handle global/workbook level names and
#NYI # local/worksheet names.
#NYI #
#NYI sub define_name {
#NYI 
#NYI     my $self        = shift;
#NYI     my $name        = shift;
#NYI     my $formula     = shift;
#NYI     my $sheet_index = undef;
#NYI     my $sheetname   = '';
#NYI     my $full_name   = $name;
#NYI 
#NYI     # Remove the = sign from the formula if it exists.
#NYI     $formula =~ s/^=//;
#NYI 
#NYI     # Local defined names are formatted like "Sheet1!name".
#NYI     if ( $name =~ /^(.*)!(.*)$/ ) {
#NYI         $sheetname   = $1;
#NYI         $name        = $2;
#NYI         $sheet_index = $self->_get_sheet_index( $sheetname );
#NYI     }
#NYI     else {
#NYI         $sheet_index = -1;    # Use -1 to indicate global names.
#NYI     }
#NYI 
#NYI     # Warn if the sheet index wasn't found.
#NYI     if ( !defined $sheet_index ) {
#NYI         carp "Unknown sheet name $sheetname in defined_name()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     # Warn if the name contains invalid chars as defined by Excel help.
#NYI     if ( $name !~ m/^[\w\\][\w\\.]*$/ || $name =~ m/^\d/ ) {
#NYI         carp "Invalid character in name '$name' used in defined_name()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     # Warn if the name looks like a cell name.
#NYI     if ( $name =~ m/^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$/ ) {
#NYI         carp "Invalid name '$name' looks like a cell name in defined_name()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     # Warn if the name looks like a R1C1.
#NYI     if ( $name =~ m/^[rcRC]$/ || $name =~ m/^[rcRC]\d+[rcRC]\d+$/ ) {
#NYI         carp "Invalid name '$name' like a RC cell ref in defined_name()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     push @{ $self->{_defined_names} }, [ $name, $sheet_index, $formula ];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_size()
#NYI #
#NYI # Set the workbook size.
#NYI #
#NYI sub set_size {
#NYI 
#NYI     my $self   = shift;
#NYI     my $width  = shift;
#NYI     my $height = shift;
#NYI 
#NYI     if ( !$width ) {
#NYI         $self->{_window_width} = 16095;
#NYI     }
#NYI     else {
#NYI         # Convert to twips at 96 dpi.
#NYI         $self->{_window_width} = int( $width * 1440 / 96 );
#NYI     }
#NYI 
#NYI     if ( !$height ) {
#NYI         $self->{_window_height} = 9660;
#NYI     }
#NYI     else {
#NYI         # Convert to twips at 96 dpi.
#NYI         $self->{_window_height} = int( $height * 1440 / 96 );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_properties()
#NYI #
#NYI # Set the document properties such as Title, Author etc. These are written to
#NYI # property sets in the OLE container.
#NYI #
#NYI sub set_properties {
#NYI 
#NYI     my $self  = shift;
#NYI     my %param = @_;
#NYI 
#NYI     # Ignore if no args were passed.
#NYI     return -1 unless @_;
#NYI 
#NYI     # List of valid input parameters.
#NYI     my %valid = (
#NYI         title          => 1,
#NYI         subject        => 1,
#NYI         author         => 1,
#NYI         keywords       => 1,
#NYI         comments       => 1,
#NYI         last_author    => 1,
#NYI         created        => 1,
#NYI         category       => 1,
#NYI         manager        => 1,
#NYI         company        => 1,
#NYI         status         => 1,
#NYI         hyperlink_base => 1,
#NYI     );
#NYI 
#NYI     # Check for valid input parameters.
#NYI     for my $parameter ( keys %param ) {
#NYI         if ( not exists $valid{$parameter} ) {
#NYI             carp "Unknown parameter '$parameter' in set_properties()";
#NYI             return -1;
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the creation time unless specified by the user.
#NYI     if ( !exists $param{created} ) {
#NYI         $param{created} = $self->{_createtime};
#NYI     }
#NYI 
#NYI 
#NYI     $self->{_doc_properties} = \%param;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_custom_property()
#NYI #
#NYI # Set a user defined custom document property.
#NYI #
#NYI sub set_custom_property {
#NYI 
#NYI     my $self  = shift;
#NYI     my $name  = shift;
#NYI     my $value = shift;
#NYI     my $type  = shift;
#NYI 
#NYI 
#NYI     # Valid types.
#NYI     my %valid_type = (
#NYI         'text'       => 1,
#NYI         'date'       => 1,
#NYI         'number'     => 1,
#NYI         'number_int' => 1,
#NYI         'bool'       => 1,
#NYI     );
#NYI 
#NYI     if ( !defined $name || !defined $value ) {
#NYI         carp "The name and value parameters must be defined "
#NYI           . "in set_custom_property()";
#NYI 
#NYI         return -1;
#NYI     }
#NYI 
#NYI     # Determine the type for strings and numbers if it hasn't been specified.
#NYI     if ( !$type ) {
#NYI         if ( $value =~ /^\d+$/ ) {
#NYI             $type = 'number_int';
#NYI         }
#NYI         elsif ( $value =~
#NYI             /^([+-]?)(?=[0-9]|\.[0-9])[0-9]*(\.[0-9]*)?([Ee]([+-]?[0-9]+))?$/ )
#NYI         {
#NYI             $type = 'number';
#NYI         }
#NYI         else {
#NYI             $type = 'text';
#NYI         }
#NYI     }
#NYI 
#NYI     # Check for valid validation types.
#NYI     if ( !exists $valid_type{$type} ) {
#NYI         carp "Unknown custom type '$type' in set_custom_property()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     #  Check for strings longer than Excel's limit of 255 chars.
#NYI     if ( $type eq 'text' and length $value > 255 ) {
#NYI         carp "Length of text custom value '$value' exceeds "
#NYI           . "Excel's limit of 255 in set_custom_property()";
#NYI         return -1;
#NYI     }
#NYI     if ( length $value > 255 ) {
#NYI         carp "Length of custom name '$name' exceeds "
#NYI           . "Excel's limit of 255 in set_custom_property()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     push @{ $self->{_custom_properties} }, [ $name, $value, $type ];
#NYI }
#NYI 
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_vba_project()
#NYI #
#NYI # Add a vbaProject binary to the XLSX file.
#NYI #
#NYI sub add_vba_project {
#NYI 
#NYI     my $self        = shift;
#NYI     my $vba_project = shift;
#NYI 
#NYI     croak "No vbaProject.bin specified in add_vba_project()"
#NYI       if not $vba_project;
#NYI 
#NYI     croak "Couldn't locate $vba_project in add_vba_project(): $!"
#NYI       unless -e $vba_project;
#NYI 
#NYI     $self->{_vba_project} = $vba_project;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_vba_name()
#NYI #
#NYI # Set the VBA name for the workbook.
#NYI #
#NYI sub set_vba_name {
#NYI 
#NYI     my $self         = shift;
#NYI     my $vba_codemame = shift;
#NYI 
#NYI     if ( $vba_codemame ) {
#NYI         $self->{_vba_codename} = $vba_codemame;
#NYI     }
#NYI     else {
#NYI         $self->{_vba_codename} = 'ThisWorkbook';
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_calc_mode()
#NYI #
#NYI # Set the Excel caclcuation mode for the workbook.
#NYI #
#NYI sub set_calc_mode {
#NYI 
#NYI     my $self    = shift;
#NYI     my $mode    = shift || 'auto';
#NYI     my $calc_id = shift;
#NYI 
#NYI     $self->{_calc_mode} = $mode;
#NYI 
#NYI     if ( $mode eq 'manual' ) {
#NYI         $self->{_calc_mode}    = 'manual';
#NYI         $self->{_calc_on_load} = 0;
#NYI     }
#NYI     elsif ( $mode eq 'auto_except_tables' ) {
#NYI         $self->{_calc_mode} = 'autoNoTable';
#NYI     }
#NYI 
#NYI     $self->{_calc_id} = $calc_id if defined $calc_id;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _store_workbook()
#NYI #
#NYI # Assemble worksheets into a workbook.
#NYI #
#NYI sub _store_workbook {
#NYI 
#NYI     my $self     = shift;
#NYI     my $tempdir  = File::Temp->newdir( DIR => $self->{_tempdir} );
#NYI     my $packager = Excel::Writer::XLSX::Package::Packager->new();
#NYI     my $zip      = Archive::Zip->new();
#NYI 
#NYI 
#NYI     # Add a default worksheet if non have been added.
#NYI     $self->add_worksheet() if not @{ $self->{_worksheets} };
#NYI 
#NYI     # Ensure that at least one worksheet has been selected.
#NYI     if ( $self->{_activesheet} == 0 ) {
#NYI         $self->{_worksheets}->[0]->{_selected} = 1;
#NYI         $self->{_worksheets}->[0]->{_hidden}   = 0;
#NYI     }
#NYI 
#NYI     # Set the active sheet.
#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {
#NYI         $sheet->{_active} = 1 if $sheet->{_index} == $self->{_activesheet};
#NYI     }
#NYI 
#NYI     # Convert the SST strings data structure.
#NYI     $self->_prepare_sst_string_data();
#NYI 
#NYI     # Prepare the worksheet VML elements such as comments and buttons.
#NYI     $self->_prepare_vml_objects();
#NYI 
#NYI     # Set the defined names for the worksheets such as Print Titles.
#NYI     $self->_prepare_defined_names();
#NYI 
#NYI     # Prepare the drawings, charts and images.
#NYI     $self->_prepare_drawings();
#NYI 
#NYI     # Add cached data to charts.
#NYI     $self->_add_chart_data();
#NYI 
#NYI     # Prepare the worksheet tables.
#NYI     $self->_prepare_tables();
#NYI 
#NYI     # Package the workbook.
#NYI     $packager->_add_workbook( $self );
#NYI     $packager->_set_package_dir( $tempdir );
#NYI     $packager->_create_package();
#NYI 
#NYI     # Free up the Packager object.
#NYI     $packager = undef;
#NYI 
#NYI     # Add the files to the zip archive. Due to issues with Archive::Zip in
#NYI     # taint mode we can't use addTree() so we have to build the file list
#NYI     # with File::Find and pass each one to addFile().
#NYI     my @xlsx_files;
#NYI 
#NYI     my $wanted = sub { push @xlsx_files, $File::Find::name if -f };
#NYI 
#NYI     File::Find::find(
#NYI         {
#NYI             wanted          => $wanted,
#NYI             untaint         => 1,
#NYI             untaint_pattern => qr|^(.+)$|
#NYI         },
#NYI         $tempdir
#NYI     );
#NYI 
#NYI     # Store the xlsx component files with the temp dir name removed.
#NYI     for my $filename ( @xlsx_files ) {
#NYI         my $short_name = $filename;
#NYI         $short_name =~ s{^\Q$tempdir\E/?}{};
#NYI         $zip->addFile( $filename, $short_name );
#NYI     }
#NYI 
#NYI 
#NYI     if ( $self->{_internal_fh} ) {
#NYI 
#NYI         if ( $zip->writeToFileHandle( $self->{_filehandle} ) != 0 ) {
#NYI             carp 'Error writing zip container for xlsx file.';
#NYI         }
#NYI     }
#NYI     else {
#NYI 
#NYI         # Archive::Zip needs to rewind a filehandle to write the zip headers.
#NYI         # This won't work for arbitrary user defined filehandles so we use
#NYI         # a temp file based filehandle to create the zip archive and then
#NYI         # stream that to the filehandle.
#NYI         my $tmp_fh = tempfile( DIR => $self->{_tempdir} );
#NYI         my $is_seekable = 1;
#NYI 
#NYI         if ( $zip->writeToFileHandle( $tmp_fh, $is_seekable ) != 0 ) {
#NYI             carp 'Error writing zip container for xlsx file.';
#NYI         }
#NYI 
#NYI         my $buffer;
#NYI         seek $tmp_fh, 0, 0;
#NYI 
#NYI         while ( read( $tmp_fh, $buffer, 4_096 ) ) {
#NYI             local $\ = undef;    # Protect print from -l on commandline.
#NYI             print { $self->{_filehandle} } $buffer;
#NYI         }
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_sst_string_data()
#NYI #
#NYI # Convert the SST string data from a hash to an array.
#NYI #
#NYI sub _prepare_sst_string_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @strings;
#NYI     $#strings = $self->{_str_unique} - 1;    # Pre-extend array
#NYI 
#NYI     while ( my $key = each %{ $self->{_str_table} } ) {
#NYI         $strings[ $self->{_str_table}->{$key} ] = $key;
#NYI     }
#NYI 
#NYI     # The SST data could be very large, free some memory (maybe).
#NYI     $self->{_str_table} = undef;
#NYI     $self->{_str_array} = \@strings;
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_format_properties()
#NYI #
#NYI # Prepare all of the format properties prior to passing them to Styles.pm.
#NYI #
#NYI sub _prepare_format_properties {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Separate format objects into XF and DXF formats.
#NYI     $self->_prepare_formats();
#NYI 
#NYI     # Set the font index for the format objects.
#NYI     $self->_prepare_fonts();
#NYI 
#NYI     # Set the number format index for the format objects.
#NYI     $self->_prepare_num_formats();
#NYI 
#NYI     # Set the border index for the format objects.
#NYI     $self->_prepare_borders();
#NYI 
#NYI     # Set the fill index for the format objects.
#NYI     $self->_prepare_fills();
#NYI 
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_formats()
#NYI #
#NYI # Iterate through the XF Format objects and separate them into XF and DXF
#NYI # formats.
#NYI #
#NYI sub _prepare_formats {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     for my $format ( @{ $self->{_formats} } ) {
#NYI         my $xf_index  = $format->{_xf_index};
#NYI         my $dxf_index = $format->{_dxf_index};
#NYI 
#NYI         if ( defined $xf_index ) {
#NYI             $self->{_xf_formats}->[$xf_index] = $format;
#NYI         }
#NYI 
#NYI         if ( defined $dxf_index ) {
#NYI             $self->{_dxf_formats}->[$dxf_index] = $format;
#NYI         }
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _set_default_xf_indices()
#NYI #
#NYI # Set the default index for each format. This is mainly used for testing.
#NYI #
#NYI sub _set_default_xf_indices {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     for my $format ( @{ $self->{_formats} } ) {
#NYI         $format->get_xf_index();
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_fonts()
#NYI #
#NYI # Iterate through the XF Format objects and give them an index to non-default
#NYI # font elements.
#NYI #
#NYI sub _prepare_fonts {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my %fonts;
#NYI     my $index = 0;
#NYI 
#NYI     for my $format ( @{ $self->{_xf_formats} } ) {
#NYI         my $key = $format->get_font_key();
#NYI 
#NYI         if ( exists $fonts{$key} ) {
#NYI 
#NYI             # Font has already been used.
#NYI             $format->{_font_index} = $fonts{$key};
#NYI             $format->{_has_font}   = 0;
#NYI         }
#NYI         else {
#NYI 
#NYI             # This is a new font.
#NYI             $fonts{$key}           = $index;
#NYI             $format->{_font_index} = $index;
#NYI             $format->{_has_font}   = 1;
#NYI             $index++;
#NYI         }
#NYI     }
#NYI 
#NYI     $self->{_font_count} = $index;
#NYI 
#NYI     # For the DXF formats we only need to check if the properties have changed.
#NYI     for my $format ( @{ $self->{_dxf_formats} } ) {
#NYI 
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
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_num_formats()
#NYI #
#NYI # Iterate through the XF Format objects and give them an index to non-default
#NYI # number format elements.
#NYI #
#NYI # User defined records start from index 0xA4.
#NYI #
#NYI sub _prepare_num_formats {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my %num_formats;
#NYI     my $index            = 164;
#NYI     my $num_format_count = 0;
#NYI 
#NYI     for my $format ( @{ $self->{_xf_formats} }, @{ $self->{_dxf_formats} } ) {
#NYI         my $num_format = $format->{_num_format};
#NYI 
#NYI         # Check if $num_format is an index to a built-in number format.
#NYI         # Also check for a string of zeros, which is a valid number format
#NYI         # string but would evaluate to zero.
#NYI         #
#NYI         if ( $num_format =~ m/^\d+$/ && $num_format !~ m/^0+\d/ ) {
#NYI 
#NYI             # Index to a built-in number format.
#NYI             $format->{_num_format_index} = $num_format;
#NYI             next;
#NYI         }
#NYI 
#NYI 
#NYI         if ( exists( $num_formats{$num_format} ) ) {
#NYI 
#NYI             # Number format has already been used.
#NYI             $format->{_num_format_index} = $num_formats{$num_format};
#NYI         }
#NYI         else {
#NYI 
#NYI             # Add a new number format.
#NYI             $num_formats{$num_format} = $index;
#NYI             $format->{_num_format_index} = $index;
#NYI             $index++;
#NYI 
#NYI             # Only increase font count for XF formats (not for DXF formats).
#NYI             if ( $format->{_xf_index} ) {
#NYI                 $num_format_count++;
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     $self->{_num_format_count} = $num_format_count;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_borders()
#NYI #
#NYI # Iterate through the XF Format objects and give them an index to non-default
#NYI # border elements.
#NYI #
#NYI sub _prepare_borders {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my %borders;
#NYI     my $index = 0;
#NYI 
#NYI     for my $format ( @{ $self->{_xf_formats} } ) {
#NYI         my $key = $format->get_border_key();
#NYI 
#NYI         if ( exists $borders{$key} ) {
#NYI 
#NYI             # Border has already been used.
#NYI             $format->{_border_index} = $borders{$key};
#NYI             $format->{_has_border}   = 0;
#NYI         }
#NYI         else {
#NYI 
#NYI             # This is a new border.
#NYI             $borders{$key}           = $index;
#NYI             $format->{_border_index} = $index;
#NYI             $format->{_has_border}   = 1;
#NYI             $index++;
#NYI         }
#NYI     }
#NYI 
#NYI     $self->{_border_count} = $index;
#NYI 
#NYI     # For the DXF formats we only need to check if the properties have changed.
#NYI     for my $format ( @{ $self->{_dxf_formats} } ) {
#NYI         my $key = $format->get_border_key();
#NYI 
#NYI         if ( $key =~ m/[^0:]/ ) {
#NYI             $format->{_has_dxf_border} = 1;
#NYI         }
#NYI     }
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_fills()
#NYI #
#NYI # Iterate through the XF Format objects and give them an index to non-default
#NYI # fill elements.
#NYI #
#NYI # The user defined fill properties start from 2 since there are 2 default
#NYI # fills: patternType="none" and patternType="gray125".
#NYI #
#NYI sub _prepare_fills {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my %fills;
#NYI     my $index = 2;    # Start from 2. See above.
#NYI 
#NYI     # Add the default fills.
#NYI     $fills{'0:0:0'}  = 0;
#NYI     $fills{'17:0:0'} = 1;
#NYI 
#NYI 
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
#NYI 
#NYI 
#NYI     for my $format ( @{ $self->{_xf_formats} } ) {
#NYI 
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
#NYI 
#NYI         if (   $format->{_pattern} <= 1
#NYI             && $format->{_bg_color} ne '0'
#NYI             && $format->{_fg_color} eq '0' )
#NYI         {
#NYI             $format->{_fg_color} = $format->{_bg_color};
#NYI             $format->{_bg_color} = 0;
#NYI             $format->{_pattern}  = 1;
#NYI         }
#NYI 
#NYI         if (   $format->{_pattern} <= 1
#NYI             && $format->{_bg_color} eq '0'
#NYI             && $format->{_fg_color} ne '0' )
#NYI         {
#NYI             $format->{_bg_color} = 0;
#NYI             $format->{_pattern}  = 1;
#NYI         }
#NYI 
#NYI 
#NYI         my $key = $format->get_fill_key();
#NYI 
#NYI         if ( exists $fills{$key} ) {
#NYI 
#NYI             # Fill has already been used.
#NYI             $format->{_fill_index} = $fills{$key};
#NYI             $format->{_has_fill}   = 0;
#NYI         }
#NYI         else {
#NYI 
#NYI             # This is a new fill.
#NYI             $fills{$key}           = $index;
#NYI             $format->{_fill_index} = $index;
#NYI             $format->{_has_fill}   = 1;
#NYI             $index++;
#NYI         }
#NYI     }
#NYI 
#NYI     $self->{_fill_count} = $index;
#NYI 
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_defined_names()
#NYI #
#NYI # Iterate through the worksheets and store any defined names in addition to
#NYI # any user defined names. Stores the defined names for the Workbook.xml and
#NYI # the named ranges for App.xml.
#NYI #
#NYI sub _prepare_defined_names {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @defined_names = @{ $self->{_defined_names} };
#NYI 
#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {
#NYI 
#NYI         # Check for Print Area settings.
#NYI         if ( $sheet->{_autofilter} ) {
#NYI 
#NYI             my $range  = $sheet->{_autofilter};
#NYI             my $hidden = 1;
#NYI 
#NYI             # Store the defined names.
#NYI             push @defined_names,
#NYI               [ '_xlnm._FilterDatabase', $sheet->{_index}, $range, $hidden ];
#NYI 
#NYI         }
#NYI 
#NYI         # Check for Print Area settings.
#NYI         if ( $sheet->{_print_area} ) {
#NYI 
#NYI             my $range = $sheet->{_print_area};
#NYI 
#NYI             # Store the defined names.
#NYI             push @defined_names,
#NYI               [ '_xlnm.Print_Area', $sheet->{_index}, $range ];
#NYI         }
#NYI 
#NYI         # Check for repeat rows/cols. aka, Print Titles.
#NYI         if ( $sheet->{_repeat_cols} || $sheet->{_repeat_rows} ) {
#NYI             my $range = '';
#NYI 
#NYI             if ( $sheet->{_repeat_cols} && $sheet->{_repeat_rows} ) {
#NYI                 $range = $sheet->{_repeat_cols} . ',' . $sheet->{_repeat_rows};
#NYI             }
#NYI             else {
#NYI                 $range = $sheet->{_repeat_cols} . $sheet->{_repeat_rows};
#NYI             }
#NYI 
#NYI             # Store the defined names.
#NYI             push @defined_names,
#NYI               [ '_xlnm.Print_Titles', $sheet->{_index}, $range ];
#NYI         }
#NYI 
#NYI     }
#NYI 
#NYI     @defined_names          = _sort_defined_names( @defined_names );
#NYI     $self->{_defined_names} = \@defined_names;
#NYI     $self->{_named_ranges}  = _extract_named_ranges( @defined_names );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _sort_defined_names()
#NYI #
#NYI # Sort internal and user defined names in the same order as used by Excel.
#NYI # This may not be strictly necessary but unsorted elements caused a lot of
#NYI # issues in the Spreadsheet::WriteExcel binary version. Also makes
#NYI # comparison testing easier.
#NYI #
#NYI sub _sort_defined_names {
#NYI 
#NYI     my @names = @_;
#NYI 
#NYI     #<<< Perltidy ignore this.
#NYI 
#NYI     @names = sort {
#NYI         # Primary sort based on the defined name.
#NYI         _normalise_defined_name( $a->[0] )
#NYI         cmp
#NYI         _normalise_defined_name( $b->[0] )
#NYI 
#NYI         ||
#NYI         # Secondary sort based on the sheet name.
#NYI         _normalise_sheet_name( $a->[2] )
#NYI         cmp
#NYI         _normalise_sheet_name( $b->[2] )
#NYI 
#NYI     } @names;
#NYI     #>>>
#NYI 
#NYI     return @names;
#NYI }
#NYI 
#NYI # Used in the above sort routine to normalise the defined names. Removes any
#NYI # leading '_xmln.' from internal names and lowercases the strings.
#NYI sub _normalise_defined_name {
#NYI     my $name = shift;
#NYI 
#NYI     $name =~ s/^_xlnm.//;
#NYI     $name = lc $name;
#NYI 
#NYI     return $name;
#NYI }
#NYI 
#NYI # Used in the above sort routine to normalise the worksheet names for the
#NYI # secondary sort. Removes leading quote and lowercases the strings.
#NYI sub _normalise_sheet_name {
#NYI     my $name = shift;
#NYI 
#NYI     $name =~ s/^'//;
#NYI     $name = lc $name;
#NYI 
#NYI     return $name;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _extract_named_ranges()
#NYI #
#NYI # Extract the named ranges from the sorted list of defined names. These are
#NYI # used in the App.xml file.
#NYI #
#NYI sub _extract_named_ranges {
#NYI 
#NYI     my @defined_names = @_;
#NYI     my @named_ranges;
#NYI 
#NYI     NAME:
#NYI     for my $defined_name ( @defined_names ) {
#NYI 
#NYI         my $name  = $defined_name->[0];
#NYI         my $index = $defined_name->[1];
#NYI         my $range = $defined_name->[2];
#NYI 
#NYI         # Skip autoFilter ranges.
#NYI         next NAME if $name eq '_xlnm._FilterDatabase';
#NYI 
#NYI         # We are only interested in defined names with ranges.
#NYI         if ( $range =~ /^([^!]+)!/ ) {
#NYI             my $sheet_name = $1;
#NYI 
#NYI             # Match Print_Area and Print_Titles xlnm types.
#NYI             if ( $name =~ /^_xlnm\.(.*)$/ ) {
#NYI                 my $xlnm_type = $1;
#NYI                 $name = $sheet_name . '!' . $xlnm_type;
#NYI             }
#NYI             elsif ( $index != -1 ) {
#NYI                 $name = $sheet_name . '!' . $name;
#NYI             }
#NYI 
#NYI             push @named_ranges, $name;
#NYI         }
#NYI     }
#NYI 
#NYI     return \@named_ranges;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_drawings()
#NYI #
#NYI # Iterate through the worksheets and set up any chart or image drawings.
#NYI #
#NYI sub _prepare_drawings {
#NYI 
#NYI     my $self         = shift;
#NYI     my $chart_ref_id = 0;
#NYI     my $image_ref_id = 0;
#NYI     my $drawing_id   = 0;
#NYI 
#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {
#NYI 
#NYI         my $chart_count = scalar @{ $sheet->{_charts} };
#NYI         my $image_count = scalar @{ $sheet->{_images} };
#NYI         my $shape_count = scalar @{ $sheet->{_shapes} };
#NYI 
#NYI         my $header_image_count = scalar @{ $sheet->{_header_images} };
#NYI         my $footer_image_count = scalar @{ $sheet->{_footer_images} };
#NYI         my $has_drawing        = 0;
#NYI 
#NYI 
#NYI         # Check that some image or drawing needs to be processed.
#NYI         if (   !$chart_count
#NYI             && !$image_count
#NYI             && !$shape_count
#NYI             && !$header_image_count
#NYI             && !$footer_image_count )
#NYI         {
#NYI             next;
#NYI         }
#NYI 
#NYI         # Don't increase the drawing_id header/footer images.
#NYI         if ( $chart_count || $image_count || $shape_count ) {
#NYI             $drawing_id++;
#NYI             $has_drawing = 1;
#NYI         }
#NYI 
#NYI         # Prepare the worksheet charts.
#NYI         for my $index ( 0 .. $chart_count - 1 ) {
#NYI             $chart_ref_id++;
#NYI             $sheet->_prepare_chart( $index, $chart_ref_id, $drawing_id );
#NYI         }
#NYI 
#NYI         # Prepare the worksheet images.
#NYI         for my $index ( 0 .. $image_count - 1 ) {
#NYI 
#NYI             my $filename = $sheet->{_images}->[$index]->[2];
#NYI 
#NYI             my ( $type, $width, $height, $name, $x_dpi, $y_dpi ) =
#NYI               $self->_get_image_properties( $filename );
#NYI 
#NYI             $image_ref_id++;
#NYI 
#NYI             $sheet->_prepare_image(
#NYI                 $index, $image_ref_id, $drawing_id,
#NYI                 $width, $height,       $name,
#NYI                 $type,  $x_dpi,        $y_dpi
#NYI             );
#NYI         }
#NYI 
#NYI         # Prepare the worksheet shapes.
#NYI         for my $index ( 0 .. $shape_count - 1 ) {
#NYI             $sheet->_prepare_shape( $index, $drawing_id );
#NYI         }
#NYI 
#NYI         # Prepare the header images.
#NYI         for my $index ( 0 .. $header_image_count - 1 ) {
#NYI 
#NYI             my $filename = $sheet->{_header_images}->[$index]->[0];
#NYI             my $position = $sheet->{_header_images}->[$index]->[1];
#NYI 
#NYI             my ( $type, $width, $height, $name, $x_dpi, $y_dpi ) =
#NYI               $self->_get_image_properties( $filename );
#NYI 
#NYI             $image_ref_id++;
#NYI 
#NYI             $sheet->_prepare_header_image( $image_ref_id, $width, $height,
#NYI                 $name, $type, $position, $x_dpi, $y_dpi );
#NYI         }
#NYI 
#NYI         # Prepare the footer images.
#NYI         for my $index ( 0 .. $footer_image_count - 1 ) {
#NYI 
#NYI             my $filename = $sheet->{_footer_images}->[$index]->[0];
#NYI             my $position = $sheet->{_footer_images}->[$index]->[1];
#NYI 
#NYI             my ( $type, $width, $height, $name, $x_dpi, $y_dpi ) =
#NYI               $self->_get_image_properties( $filename );
#NYI 
#NYI             $image_ref_id++;
#NYI 
#NYI             $sheet->_prepare_header_image( $image_ref_id, $width, $height,
#NYI                 $name, $type, $position, $x_dpi, $y_dpi );
#NYI         }
#NYI 
#NYI 
#NYI         if ( $has_drawing ) {
#NYI             my $drawing = $sheet->{_drawing};
#NYI             push @{ $self->{_drawings} }, $drawing;
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # Remove charts that were created but not inserted into worksheets.
#NYI     my @chart_data;
#NYI 
#NYI     for my $chart ( @{ $self->{_charts} } ) {
#NYI         if ( $chart->{_id} != -1 ) {
#NYI             push @chart_data, $chart;
#NYI         }
#NYI     }
#NYI 
#NYI     # Sort the workbook charts references into the order that the were
#NYI     # written from the worksheets above.
#NYI     @chart_data = sort { $a->{_id} <=> $b->{_id} } @chart_data;
#NYI 
#NYI     $self->{_charts} = \@chart_data;
#NYI     $self->{_drawing_count} = $drawing_id;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_vml_objects()
#NYI #
#NYI # Iterate through the worksheets and set up the VML objects.
#NYI #
#NYI sub _prepare_vml_objects {
#NYI 
#NYI     my $self           = shift;
#NYI     my $comment_id     = 0;
#NYI     my $vml_drawing_id = 0;
#NYI     my $vml_data_id    = 1;
#NYI     my $vml_header_id  = 0;
#NYI     my $vml_shape_id   = 1024;
#NYI     my $vml_files      = 0;
#NYI     my $comment_files  = 0;
#NYI     my $has_button     = 0;
#NYI 
#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {
#NYI 
#NYI         next if !$sheet->{_has_vml} and !$sheet->{_has_header_vml};
#NYI         $vml_files = 1;
#NYI 
#NYI 
#NYI         if ( $sheet->{_has_vml} ) {
#NYI 
#NYI             $comment_files++ if $sheet->{_has_comments};
#NYI             $comment_id++    if $sheet->{_has_comments};
#NYI             $vml_drawing_id++;
#NYI 
#NYI             my $count =
#NYI               $sheet->_prepare_vml_objects( $vml_data_id, $vml_shape_id,
#NYI                 $vml_drawing_id, $comment_id );
#NYI 
#NYI             # Each VML file should start with a shape id incremented by 1024.
#NYI             $vml_data_id  += 1 * int(    ( 1024 + $count ) / 1024 );
#NYI             $vml_shape_id += 1024 * int( ( 1024 + $count ) / 1024 );
#NYI 
#NYI         }
#NYI 
#NYI         if ( $sheet->{_has_header_vml} ) {
#NYI             $vml_header_id++;
#NYI             $vml_drawing_id++;
#NYI             $sheet->_prepare_header_vml_objects( $vml_header_id,
#NYI                 $vml_drawing_id );
#NYI         }
#NYI 
#NYI         # Set the sheet vba_codename if it has a button and the workbook
#NYI         # has a vbaProject binary.
#NYI         if ( $sheet->{_buttons_array} ) {
#NYI             $has_button = 1;
#NYI 
#NYI             if ( $self->{_vba_project} && !$sheet->{_vba_codename} ) {
#NYI                 $sheet->set_vba_name();
#NYI             }
#NYI         }
#NYI 
#NYI     }
#NYI 
#NYI     $self->{_num_vml_files}     = $vml_files;
#NYI     $self->{_num_comment_files} = $comment_files;
#NYI 
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
#NYI 
#NYI         $format->get_xf_index();
#NYI 
#NYI         push @{ $self->{_formats} }, $format;
#NYI     }
#NYI 
#NYI     # Set the workbook vba_codename if one of the sheets has a button and
#NYI     # the workbook has a vbaProject binary.
#NYI     if ( $has_button && $self->{_vba_project} && !$self->{_vba_codename} ) {
#NYI         $self->set_vba_name();
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _prepare_tables()
#NYI #
#NYI # Set the table ids for the worksheet tables.
#NYI #
#NYI sub _prepare_tables {
#NYI 
#NYI     my $self     = shift;
#NYI     my $table_id = 0;
#NYI     my $seen     = {};
#NYI 
#NYI     for my $sheet ( @{ $self->{_worksheets} } ) {
#NYI 
#NYI         my $table_count = scalar @{ $sheet->{_tables} };
#NYI 
#NYI         next unless $table_count;
#NYI 
#NYI         $sheet->_prepare_tables( $table_id + 1, $seen );
#NYI 
#NYI         $table_id += $table_count;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _add_chart_data()
#NYI #
#NYI # Add "cached" data to charts to provide the numCache and strCache data for
#NYI # series and title/axis ranges.
#NYI #
#NYI sub _add_chart_data {
#NYI 
#NYI     my $self = shift;
#NYI     my %worksheets;
#NYI     my %seen_ranges;
#NYI     my @charts;
#NYI 
#NYI     # Map worksheet names to worksheet objects.
#NYI     for my $worksheet ( @{ $self->{_worksheets} } ) {
#NYI         $worksheets{ $worksheet->{_name} } = $worksheet;
#NYI     }
#NYI 
#NYI     # Build an array of the worksheet charts including any combined charts.
#NYI     for my $chart ( @{ $self->{_charts} } ) {
#NYI         push @charts, $chart;
#NYI 
#NYI         if ($chart->{_combined}) {
#NYI             push @charts, $chart->{_combined};
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     CHART:
#NYI     for my $chart ( @charts ) {
#NYI 
#NYI         RANGE:
#NYI         while ( my ( $range, $id ) = each %{ $chart->{_formula_ids} } ) {
#NYI 
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
#NYI 
#NYI             # Check to see if the data is already cached locally.
#NYI             if ( exists $seen_ranges{$range} ) {
#NYI                 $chart->{_formula_data}->[$id] = $seen_ranges{$range};
#NYI                 next RANGE;
#NYI             }
#NYI 
#NYI             # Convert the range formula to a sheet name and cell range.
#NYI             my ( $sheetname, @cells ) = $self->_get_chart_range( $range );
#NYI 
#NYI             # Skip if we couldn't parse the formula.
#NYI             next RANGE if !defined $sheetname;
#NYI 
#NYI             # Handle non-contiguous ranges: (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5).
#NYI             # We don't try to parse the ranges. We just return an empty list.
#NYI             if ( $sheetname =~ m/^\([^,]+,/ ) {
#NYI                 $chart->{_formula_data}->[$id] = [];
#NYI                 $seen_ranges{$range} = [];
#NYI                 next RANGE;
#NYI             }
#NYI 
#NYI             # Die if the name is unknown since it indicates a user error in
#NYI             # a chart series formula.
#NYI             if ( !exists $worksheets{$sheetname} ) {
#NYI                 die "Unknown worksheet reference '$sheetname' in range "
#NYI                   . "'$range' passed to add_series().\n";
#NYI             }
#NYI 
#NYI             # Find the worksheet object based on the sheet name.
#NYI             my $worksheet = $worksheets{$sheetname};
#NYI 
#NYI             # Get the data from the worksheet table.
#NYI             my @data = $worksheet->_get_range_data( @cells );
#NYI 
#NYI             # Convert shared string indexes to strings.
#NYI             for my $token ( @data ) {
#NYI                 if ( ref $token ) {
#NYI                     $token = $self->{_str_array}->[ $token->{sst_id} ];
#NYI 
#NYI                     # Ignore rich strings for now. Deparse later if necessary.
#NYI                     if ( $token =~ m{^<r>} && $token =~ m{</r>$} ) {
#NYI                         $token = '';
#NYI                     }
#NYI                 }
#NYI             }
#NYI 
#NYI             # Add the data to the chart.
#NYI             $chart->{_formula_data}->[$id] = \@data;
#NYI 
#NYI             # Store range data locally to avoid lookup if seen again.
#NYI             $seen_ranges{$range} = \@data;
#NYI         }
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_chart_range()
#NYI #
#NYI # Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name and cell
#NYI # range such as ( 'Sheet1', 0, 1, 4, 1 ).
#NYI #
#NYI sub _get_chart_range {
#NYI 
#NYI     my $self  = shift;
#NYI     my $range = shift;
#NYI     my $cell_1;
#NYI     my $cell_2;
#NYI     my $sheetname;
#NYI     my $cells;
#NYI 
#NYI     # Split the range formula into sheetname and cells at the last '!'.
#NYI     my $pos = rindex $range, '!';
#NYI     if ( $pos > 0 ) {
#NYI         $sheetname = substr $range, 0, $pos;
#NYI         $cells = substr $range, $pos + 1;
#NYI     }
#NYI     else {
#NYI         return undef;
#NYI     }
#NYI 
#NYI     # Split the cell range into 2 cells or else use single cell for both.
#NYI     if ( $cells =~ ':' ) {
#NYI         ( $cell_1, $cell_2 ) = split /:/, $cells;
#NYI     }
#NYI     else {
#NYI         ( $cell_1, $cell_2 ) = ( $cells, $cells );
#NYI     }
#NYI 
#NYI     # Remove leading/trailing apostrophes and convert escaped quotes to single.
#NYI     $sheetname =~ s/^'//g;
#NYI     $sheetname =~ s/'$//g;
#NYI     $sheetname =~ s/''/'/g;
#NYI 
#NYI     my ( $row_start, $col_start ) = xl_cell_to_rowcol( $cell_1 );
#NYI     my ( $row_end,   $col_end )   = xl_cell_to_rowcol( $cell_2 );
#NYI 
#NYI     # Check that we have a 1D range only.
#NYI     if ( $row_start != $row_end && $col_start != $col_end ) {
#NYI         return undef;
#NYI     }
#NYI 
#NYI     return ( $sheetname, $row_start, $col_start, $row_end, $col_end );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _store_externs()
#NYI #
#NYI # Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
#NYI # the NAME records.
#NYI #
#NYI sub _store_externs {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _store_names()
#NYI #
#NYI # Write the NAME record to define the print area and the repeat rows and cols.
#NYI #
#NYI sub _store_names {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _quote_sheetname()
#NYI #
#NYI # Sheetnames used in references should be quoted if they contain any spaces,
#NYI # special characters or if the look like something that isn't a sheet name.
#NYI # TODO. We need to handle more special cases.
#NYI #
#NYI sub _quote_sheetname {
#NYI 
#NYI     my $self      = shift;
#NYI     my $sheetname = $_[0];
#NYI 
#NYI     if ( $sheetname =~ /^Sheet\d+$/ ) {
#NYI         return $sheetname;
#NYI     }
#NYI     else {
#NYI         return qq('$sheetname');
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_image_properties()
#NYI #
#NYI # Extract information from the image file such as dimension, type, filename,
#NYI # and extension. Also keep track of previously seen images to optimise out
#NYI # any duplicates.
#NYI #
#NYI sub _get_image_properties {
#NYI 
#NYI     my $self     = shift;
#NYI     my $filename = shift;
#NYI 
#NYI     my $type;
#NYI     my $width;
#NYI     my $height;
#NYI     my $x_dpi = 96;
#NYI     my $y_dpi = 96;
#NYI     my $image_name;
#NYI 
#NYI 
#NYI     ( $image_name ) = fileparse( $filename );
#NYI 
#NYI     # Open the image file and import the data.
#NYI     my $fh = FileHandle->new( $filename );
#NYI     croak "Couldn't import $filename: $!" unless defined $fh;
#NYI     binmode $fh;
#NYI 
#NYI     # Slurp the file into a string and do some size calcs.
#NYI     my $data = do { local $/; <$fh> };
#NYI     my $size = length $data;
#NYI 
#NYI 
#NYI     if ( unpack( 'x A3', $data ) eq 'PNG' ) {
#NYI 
#NYI         # Test for PNGs.
#NYI         ( $type, $width, $height, $x_dpi, $y_dpi ) =
#NYI           $self->_process_png( $data, $filename );
#NYI 
#NYI         $self->{_image_types}->{png} = 1;
#NYI     }
#NYI     elsif ( unpack( 'n', $data ) == 0xFFD8 ) {
#NYI 
#NYI         # Test for JPEG files.
#NYI         ( $type, $width, $height, $x_dpi, $y_dpi ) =
#NYI           $self->_process_jpg( $data, $filename );
#NYI 
#NYI         $self->{_image_types}->{jpeg} = 1;
#NYI     }
#NYI     elsif ( unpack( 'A2', $data ) eq 'BM' ) {
#NYI 
#NYI         # Test for BMPs.
#NYI         ( $type, $width, $height ) = $self->_process_bmp( $data, $filename );
#NYI 
#NYI         $self->{_image_types}->{bmp} = 1;
#NYI     }
#NYI     else {
#NYI         croak "Unsupported image format for file: $filename\n";
#NYI     }
#NYI 
#NYI     push @{ $self->{_images} }, [ $filename, $type ];
#NYI 
#NYI     # Set a default dpi for images with 0 dpi.
#NYI     $x_dpi = 96 if $x_dpi == 0;
#NYI     $y_dpi = 96 if $y_dpi == 0;
#NYI 
#NYI     $fh->close;
#NYI 
#NYI     return ( $type, $width, $height, $image_name, $x_dpi, $y_dpi );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _process_png()
#NYI #
#NYI # Extract width and height information from a PNG file.
#NYI #
#NYI sub _process_png {
#NYI 
#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];
#NYI 
#NYI     my $type   = 'png';
#NYI     my $width  = 0;
#NYI     my $height = 0;
#NYI     my $x_dpi  = 96;
#NYI     my $y_dpi  = 96;
#NYI 
#NYI     my $offset      = 8;
#NYI     my $data_length = length $data;
#NYI 
#NYI     # Search through the image data to read the height and width in the
#NYI     # IHDR element. Also read the DPI in the pHYs element.
#NYI     while ( $offset < $data_length ) {
#NYI 
#NYI         my $length = unpack "N",  substr $data, $offset + 0, 4;
#NYI         my $type   = unpack "A4", substr $data, $offset + 4, 4;
#NYI 
#NYI         if ( $type eq "IHDR" ) {
#NYI             $width  = unpack "N", substr $data, $offset + 8,  4;
#NYI             $height = unpack "N", substr $data, $offset + 12, 4;
#NYI         }
#NYI 
#NYI         if ( $type eq "pHYs" ) {
#NYI             my $x_ppu = unpack "N", substr $data, $offset + 8,  4;
#NYI             my $y_ppu = unpack "N", substr $data, $offset + 12, 4;
#NYI             my $units = unpack "C", substr $data, $offset + 16, 1;
#NYI 
#NYI             if ( $units == 1 ) {
#NYI                 $x_dpi = $x_ppu * 0.0254;
#NYI                 $y_dpi = $y_ppu * 0.0254;
#NYI             }
#NYI         }
#NYI 
#NYI         $offset = $offset + $length + 12;
#NYI 
#NYI         last if $type eq "IEND";
#NYI     }
#NYI 
#NYI     if ( not defined $height ) {
#NYI         croak "$filename: no size data found in png image.\n";
#NYI     }
#NYI 
#NYI     return ( $type, $width, $height, $x_dpi, $y_dpi );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _process_bmp()
#NYI #
#NYI # Extract width and height information from a BMP file.
#NYI #
#NYI # Most of the checks came from old Spredsheet::WriteExcel code.
#NYI #
#NYI sub _process_bmp {
#NYI 
#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];
#NYI     my $type     = 'bmp';
#NYI 
#NYI 
#NYI     # Check that the file is big enough to be a bitmap.
#NYI     if ( length $data <= 0x36 ) {
#NYI         croak "$filename doesn't contain enough data.";
#NYI     }
#NYI 
#NYI 
#NYI     # Read the bitmap width and height. Verify the sizes.
#NYI     my ( $width, $height ) = unpack "x18 V2", $data;
#NYI 
#NYI     if ( $width > 0xFFFF ) {
#NYI         croak "$filename: largest image width $width supported is 65k.";
#NYI     }
#NYI 
#NYI     if ( $height > 0xFFFF ) {
#NYI         croak "$filename: largest image height supported is 65k.";
#NYI     }
#NYI 
#NYI     # Read the bitmap planes and bpp data. Verify them.
#NYI     my ( $planes, $bitcount ) = unpack "x26 v2", $data;
#NYI 
#NYI     if ( $bitcount != 24 ) {
#NYI         croak "$filename isn't a 24bit true color bitmap.";
#NYI     }
#NYI 
#NYI     if ( $planes != 1 ) {
#NYI         croak "$filename: only 1 plane supported in bitmap image.";
#NYI     }
#NYI 
#NYI 
#NYI     # Read the bitmap compression. Verify compression.
#NYI     my $compression = unpack "x30 V", $data;
#NYI 
#NYI     if ( $compression != 0 ) {
#NYI         croak "$filename: compression not supported in bitmap image.";
#NYI     }
#NYI 
#NYI     return ( $type, $width, $height );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _process_jpg()
#NYI #
#NYI # Extract width and height information from a JPEG file.
#NYI #
#NYI sub _process_jpg {
#NYI 
#NYI     my $self     = shift;
#NYI     my $data     = $_[0];
#NYI     my $filename = $_[1];
#NYI     my $type     = 'jpeg';
#NYI     my $x_dpi    = 96;
#NYI     my $y_dpi    = 96;
#NYI     my $width;
#NYI     my $height;
#NYI 
#NYI     my $offset      = 2;
#NYI     my $data_length = length $data;
#NYI 
#NYI     # Search through the image data to read the height and width in the
#NYI     # 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element.
#NYI     while ( $offset < $data_length ) {
#NYI 
#NYI         my $marker = unpack "n", substr $data, $offset + 0, 2;
#NYI         my $length = unpack "n", substr $data, $offset + 2, 2;
#NYI 
#NYI         if ( $marker == 0xFFC0 || $marker == 0xFFC2 ) {
#NYI             $height = unpack "n", substr $data, $offset + 5, 2;
#NYI             $width  = unpack "n", substr $data, $offset + 7, 2;
#NYI         }
#NYI 
#NYI         if ( $marker == 0xFFE0 ) {
#NYI             my $units     = unpack "C", substr $data, $offset + 11, 1;
#NYI             my $x_density = unpack "n", substr $data, $offset + 12, 2;
#NYI             my $y_density = unpack "n", substr $data, $offset + 14, 2;
#NYI 
#NYI             if ( $units == 1 ) {
#NYI                 $x_dpi = $x_density;
#NYI                 $y_dpi = $y_density;
#NYI             }
#NYI 
#NYI             if ( $units == 2 ) {
#NYI                 $x_dpi = $x_density * 2.54;
#NYI                 $y_dpi = $y_density * 2.54;
#NYI             }
#NYI         }
#NYI 
#NYI         $offset = $offset + $length + 2;
#NYI         last if $marker == 0xFFDA;
#NYI     }
#NYI 
#NYI     if ( not defined $height ) {
#NYI         croak "$filename: no size data found in jpeg image.\n";
#NYI     }
#NYI 
#NYI     return ( $type, $width, $height, $x_dpi, $y_dpi );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_sheet_index()
#NYI #
#NYI # Convert a sheet name to its index. Return undef otherwise.
#NYI #
#NYI sub _get_sheet_index {
#NYI 
#NYI     my $self        = shift;
#NYI     my $sheetname   = shift;
#NYI     my $sheet_index = undef;
#NYI 
#NYI     $sheetname =~ s/^'//;
#NYI     $sheetname =~ s/'$//;
#NYI 
#NYI     if ( exists $self->{_sheetnames}->{$sheetname} ) {
#NYI         return $self->{_sheetnames}->{$sheetname}->{_index};
#NYI     }
#NYI     else {
#NYI         return undef;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_optimization()
#NYI #
#NYI # Set the speed/memory optimisation level.
#NYI #
#NYI sub set_optimization {
#NYI 
#NYI     my $self = shift;
#NYI     my $level = defined $_[0] ? $_[0] : 1;
#NYI 
#NYI     croak "set_optimization() must be called before add_worksheet()"
#NYI       if $self->sheets();
#NYI 
#NYI     $self->{_optimization} = $level;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Deprecated methods for backwards compatibility.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI # No longer required by Excel::Writer::XLSX.
#NYI sub compatibility_mode { }
#NYI sub set_codepage       { }
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
#NYI # _write_workbook()
#NYI #
#NYI # Write <workbook> element.
#NYI #
#NYI sub _write_workbook {
#NYI 
#NYI     my $self    = shift;
#NYI     my $schema  = 'http://schemas.openxmlformats.org';
#NYI     my $xmlns   = $schema . '/spreadsheetml/2006/main';
#NYI     my $xmlns_r = $schema . '/officeDocument/2006/relationships';
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns'   => $xmlns,
#NYI         'xmlns:r' => $xmlns_r,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'workbook', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_file_version()
#NYI #
#NYI # Write the <fileVersion> element.
#NYI #
#NYI sub _write_file_version {
#NYI 
#NYI     my $self          = shift;
#NYI     my $app_name      = 'xl';
#NYI     my $last_edited   = 4;
#NYI     my $lowest_edited = 4;
#NYI     my $rup_build     = 4505;
#NYI 
#NYI     my @attributes = (
#NYI         'appName'      => $app_name,
#NYI         'lastEdited'   => $last_edited,
#NYI         'lowestEdited' => $lowest_edited,
#NYI         'rupBuild'     => $rup_build,
#NYI     );
#NYI 
#NYI     if ( $self->{_vba_project} ) {
#NYI         push @attributes, codeName => '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}';
#NYI     }
#NYI 
#NYI     $self->xml_empty_tag( 'fileVersion', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_workbook_pr()
#NYI #
#NYI # Write <workbookPr> element.
#NYI #
#NYI sub _write_workbook_pr {
#NYI 
#NYI     my $self                   = shift;
#NYI     my $date_1904              = $self->{_date_1904};
#NYI     my $show_ink_annotation    = 0;
#NYI     my $auto_compress_pictures = 0;
#NYI     my $default_theme_version  = 124226;
#NYI     my $codename               = $self->{_vba_codename};
#NYI     my @attributes;
#NYI 
#NYI     push @attributes, ( 'codeName' => $codename ) if $codename;
#NYI     push @attributes, ( 'date1904' => 1 )         if $date_1904;
#NYI     push @attributes, ( 'defaultThemeVersion' => $default_theme_version );
#NYI 
#NYI     $self->xml_empty_tag( 'workbookPr', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_book_views()
#NYI #
#NYI # Write <bookViews> element.
#NYI #
#NYI sub _write_book_views {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'bookViews' );
#NYI     $self->_write_workbook_view();
#NYI     $self->xml_end_tag( 'bookViews' );
#NYI }
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_workbook_view()
#NYI #
#NYI # Write <workbookView> element.
#NYI #
#NYI sub _write_workbook_view {
#NYI 
#NYI     my $self          = shift;
#NYI     my $x_window      = $self->{_x_window};
#NYI     my $y_window      = $self->{_y_window};
#NYI     my $window_width  = $self->{_window_width};
#NYI     my $window_height = $self->{_window_height};
#NYI     my $tab_ratio     = $self->{_tab_ratio};
#NYI     my $active_tab    = $self->{_activesheet};
#NYI     my $first_sheet   = $self->{_firstsheet};
#NYI 
#NYI     my @attributes = (
#NYI         'xWindow'      => $x_window,
#NYI         'yWindow'      => $y_window,
#NYI         'windowWidth'  => $window_width,
#NYI         'windowHeight' => $window_height,
#NYI     );
#NYI 
#NYI     # Store the tabRatio attribute when it isn't the default.
#NYI     push @attributes, ( tabRatio => $tab_ratio ) if $tab_ratio != 500;
#NYI 
#NYI     # Store the firstSheet attribute when it isn't the default.
#NYI     push @attributes, ( firstSheet => $first_sheet + 1 ) if $first_sheet > 0;
#NYI 
#NYI     # Store the activeTab attribute when it isn't the first sheet.
#NYI     push @attributes, ( activeTab => $active_tab ) if $active_tab > 0;
#NYI 
#NYI     $self->xml_empty_tag( 'workbookView', @attributes );
#NYI }
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_sheets()
#NYI #
#NYI # Write <sheets> element.
#NYI #
#NYI sub _write_sheets {
#NYI 
#NYI     my $self   = shift;
#NYI     my $id_num = 1;
#NYI 
#NYI     $self->xml_start_tag( 'sheets' );
#NYI 
#NYI     for my $worksheet ( @{ $self->{_worksheets} } ) {
#NYI         $self->_write_sheet( $worksheet->{_name}, $id_num++,
#NYI             $worksheet->{_hidden} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'sheets' );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_sheet()
#NYI #
#NYI # Write <sheet> element.
#NYI #
#NYI sub _write_sheet {
#NYI 
#NYI     my $self     = shift;
#NYI     my $name     = shift;
#NYI     my $sheet_id = shift;
#NYI     my $hidden   = shift;
#NYI     my $r_id     = 'rId' . $sheet_id;
#NYI 
#NYI     my @attributes = (
#NYI         'name'    => $name,
#NYI         'sheetId' => $sheet_id,
#NYI     );
#NYI 
#NYI     push @attributes, ( 'state' => 'hidden' ) if $hidden;
#NYI     push @attributes, ( 'r:id' => $r_id );
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'sheet', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_calc_pr()
#NYI #
#NYI # Write <calcPr> element.
#NYI #
#NYI sub _write_calc_pr {
#NYI 
#NYI     my $self            = shift;
#NYI     my $calc_id         = $self->{_calc_id};
#NYI     my $concurrent_calc = 0;
#NYI 
#NYI     my @attributes = ( calcId => $calc_id );
#NYI 
#NYI     if ( $self->{_calc_mode} eq 'manual' ) {
#NYI         push @attributes, 'calcMode'   => 'manual';
#NYI         push @attributes, 'calcOnSave' => 0;
#NYI     }
#NYI     elsif ( $self->{_calc_mode} eq 'autoNoTable' ) {
#NYI         push @attributes, calcMode => 'autoNoTable';
#NYI     }
#NYI 
#NYI     if ( $self->{_calc_on_load} ) {
#NYI         push @attributes, 'fullCalcOnLoad' => 1;
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'calcPr', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_ext_lst()
#NYI #
#NYI # Write <extLst> element.
#NYI #
#NYI sub _write_ext_lst {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'extLst' );
#NYI     $self->_write_ext();
#NYI     $self->xml_end_tag( 'extLst' );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_ext()
#NYI #
#NYI # Write <ext> element.
#NYI #
#NYI sub _write_ext {
#NYI 
#NYI     my $self     = shift;
#NYI     my $xmlns_mx = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
#NYI     my $uri      = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns:mx' => $xmlns_mx,
#NYI         'uri'      => $uri,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'ext', @attributes );
#NYI     $self->_write_mx_arch_id();
#NYI     $self->xml_end_tag( 'ext' );
#NYI }
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _write_mx_arch_id()
#NYI #
#NYI # Write <mx:ArchID> element.
#NYI #
#NYI sub _write_mx_arch_id {
#NYI 
#NYI     my $self  = shift;
#NYI     my $Flags = 2;
#NYI 
#NYI     my @attributes = ( 'Flags' => $Flags, );
#NYI 
#NYI     $self->xml_empty_tag( 'mx:ArchID', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_defined_names()
#NYI #
#NYI # Write the <definedNames> element.
#NYI #
#NYI sub _write_defined_names {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return unless @{ $self->{_defined_names} };
#NYI 
#NYI     $self->xml_start_tag( 'definedNames' );
#NYI 
#NYI     for my $aref ( @{ $self->{_defined_names} } ) {
#NYI         $self->_write_defined_name( $aref );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'definedNames' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_defined_name()
#NYI #
#NYI # Write the <definedName> element.
#NYI #
#NYI sub _write_defined_name {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     my $name   = $data->[0];
#NYI     my $id     = $data->[1];
#NYI     my $range  = $data->[2];
#NYI     my $hidden = $data->[3];
#NYI 
#NYI     my @attributes = ( 'name' => $name );
#NYI 
#NYI     push @attributes, ( 'localSheetId' => $id ) if $id != -1;
#NYI     push @attributes, ( 'hidden'       => 1 )   if $hidden;
#NYI 
#NYI     $self->xml_data_element( 'definedName', $range, @attributes );
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
#NYI Workbook - A class for writing Excel Workbooks.
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
