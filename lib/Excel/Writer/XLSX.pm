#NYI package Excel::Writer::XLSX;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Excel::Writer::XLSX - Create a new file in the Excel 2007+ XLSX format.
#NYI #
#NYI # Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#NYI #
#NYI # Documentation after __END__
#NYI #
#NYI 
#NYI use 5.008002;
#NYI use strict;
#NYI use warnings;
#NYI use Exporter;
#NYI 
#NYI use strict;
#NYI use Excel::Writer::XLSX::Workbook;
#NYI 
#NYI our @ISA     = qw(Excel::Writer::XLSX::Workbook Exporter);
#NYI our $VERSION = '0.96';
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # new()
#NYI #
#NYI sub new {
#NYI 
#NYI     my $class = shift;
#NYI     my $self  = Excel::Writer::XLSX::Workbook->new( @_ );
#NYI 
#NYI     # Check for file creation failures before re-blessing
#NYI     bless $self, $class if defined $self;
#NYI 
#NYI     return $self;
#NYI }
#NYI 
#NYI 
#NYI 1;
#NYI 
#NYI 
#NYI __END__
#NYI 
#NYI 
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Excel::Writer::XLSX - Create a new file in the Excel 2007+ XLSX format.
#NYI 
#NYI =head1 SYNOPSIS
#NYI 
#NYI To write a string, a formatted string, a number and a formula to the first worksheet in an Excel workbook called perl.xlsx:
#NYI 
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     # Create a new Excel workbook
#NYI     my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );
#NYI 
#NYI     # Add a worksheet
#NYI     $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     #  Add and define a format
#NYI     $format = $workbook->add_format();
#NYI     $format->set_bold();
#NYI     $format->set_color( 'red' );
#NYI     $format->set_align( 'center' );
#NYI 
#NYI     # Write a formatted and unformatted string, row and column notation.
#NYI     $col = $row = 0;
#NYI     $worksheet->write( $row, $col, 'Hi Excel!', $format );
#NYI     $worksheet->write( 1, $col, 'Hi Excel!' );
#NYI 
#NYI     # Write a number and a formula using A1 notation
#NYI     $worksheet->write( 'A3', 1.2345 );
#NYI     $worksheet->write( 'A4', '=SIN(PI()/4)' );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 DESCRIPTION
#NYI 
#NYI The C<Excel::Writer::XLSX> module can be used to create an Excel file in the 2007+ XLSX format.
#NYI 
#NYI The XLSX format is the Office Open XML (OOXML) format used by Excel 2007 and later.
#NYI 
#NYI Multiple worksheets can be added to a workbook and formatting can be applied to cells. Text, numbers, and formulas can be written to the cells.
#NYI 
#NYI This module cannot, as yet, be used to write to an existing Excel XLSX file.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 Excel::Writer::XLSX and Spreadsheet::WriteExcel
#NYI 
#NYI C<Excel::Writer::XLSX> uses the same interface as the L<Spreadsheet::WriteExcel> module which produces an Excel file in binary XLS format.
#NYI 
#NYI Excel::Writer::XLSX supports all of the features of Spreadsheet::WriteExcel and in some cases has more functionality. For more details see L</Compatibility with Spreadsheet::WriteExcel>.
#NYI 
#NYI The main advantage of the XLSX format over the XLS format is that it allows a larger number of rows and columns in a worksheet. The XLSX file format also produces much smaller files than the XLS file format.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 QUICK START
#NYI 
#NYI Excel::Writer::XLSX tries to provide an interface to as many of Excel's features as possible. As a result there is a lot of documentation to accompany the interface and it can be difficult at first glance to see what it important and what is not. So for those of you who prefer to assemble Ikea furniture first and then read the instructions, here are three easy steps:
#NYI 
#NYI 1. Create a new Excel I<workbook> (i.e. file) using C<new()>.
#NYI 
#NYI 2. Add a worksheet to the new workbook using C<add_worksheet()>.
#NYI 
#NYI 3. Write to the worksheet using C<write()>.
#NYI 
#NYI Like this:
#NYI 
#NYI     use Excel::Writer::XLSX;                                   # Step 0
#NYI 
#NYI     my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );    # Step 1
#NYI     $worksheet = $workbook->add_worksheet();                   # Step 2
#NYI     $worksheet->write( 'A1', 'Hi Excel!' );                    # Step 3
#NYI 
#NYI This will create an Excel file called C<perl.xlsx> with a single worksheet and the text C<'Hi Excel!'> in the relevant cell. And that's it. Okay, so there is actually a zeroth step as well, but C<use module> goes without saying. There are many examples that come with the distribution and which you can use to get you started. See L</EXAMPLES>.
#NYI 
#NYI Those of you who read the instructions first and assemble the furniture afterwards will know how to proceed. ;-)
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 WORKBOOK METHODS
#NYI 
#NYI The Excel::Writer::XLSX module provides an object oriented interface to a new Excel workbook. The following methods are available through a new workbook.
#NYI 
#NYI     new()
#NYI     add_worksheet()
#NYI     add_format()
#NYI     add_chart()
#NYI     add_shape()
#NYI     add_vba_project()
#NYI     set_vba_name()
#NYI     close()
#NYI     set_properties()
#NYI     set_custom_property()
#NYI     define_name()
#NYI     set_tempdir()
#NYI     set_custom_color()
#NYI     sheets()
#NYI     get_worksheet_by_name()
#NYI     set_1904()
#NYI     set_optimization()
#NYI     set_calc_mode()
#NYI 
#NYI If you are unfamiliar with object oriented interfaces or the way that they are implemented in Perl have a look at C<perlobj> and C<perltoot> in the main Perl documentation.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 new()
#NYI 
#NYI A new Excel workbook is created using the C<new()> constructor which accepts either a filename or a filehandle as a parameter. The following example creates a new Excel file based on a filename:
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'filename.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI     $worksheet->write( 0, 0, 'Hi Excel!' );
#NYI 
#NYI Here are some other examples of using C<new()> with filenames:
#NYI 
#NYI     my $workbook1 = Excel::Writer::XLSX->new( $filename );
#NYI     my $workbook2 = Excel::Writer::XLSX->new( '/tmp/filename.xlsx' );
#NYI     my $workbook3 = Excel::Writer::XLSX->new( "c:\\tmp\\filename.xlsx" );
#NYI     my $workbook4 = Excel::Writer::XLSX->new( 'c:\tmp\filename.xlsx' );
#NYI 
#NYI The last two examples demonstrates how to create a file on DOS or Windows where it is necessary to either escape the directory separator C<\> or to use single quotes to ensure that it isn't interpolated. For more information see C<perlfaq5: Why can't I use "C:\temp\foo" in DOS paths?>.
#NYI 
#NYI It is recommended that the filename uses the extension C<.xlsx> rather than C<.xls> since the latter causes an Excel warning when used with the XLSX format.
#NYI 
#NYI The C<new()> constructor returns a Excel::Writer::XLSX object that you can use to add worksheets and store data. It should be noted that although C<my> is not specifically required it defines the scope of the new workbook variable and, in the majority of cases, ensures that the workbook is closed properly without explicitly calling the C<close()> method.
#NYI 
#NYI If the file cannot be created, due to file permissions or some other reason,  C<new> will return C<undef>. Therefore, it is good practice to check the return value of C<new> before proceeding. As usual the Perl variable C<$!> will be set if there is a file creation error. You will also see one of the warning messages detailed in L</DIAGNOSTICS>:
#NYI 
#NYI     my $workbook = Excel::Writer::XLSX->new( 'protected.xlsx' );
#NYI     die "Problems creating new Excel file: $!" unless defined $workbook;
#NYI 
#NYI You can also pass a valid filehandle to the C<new()> constructor. For example in a CGI program you could do something like this:
#NYI 
#NYI     binmode( STDOUT );
#NYI     my $workbook = Excel::Writer::XLSX->new( \*STDOUT );
#NYI 
#NYI The requirement for C<binmode()> is explained below.
#NYI 
#NYI See also, the C<cgi.pl> program in the C<examples> directory of the distro.
#NYI 
#NYI In C<mod_perl> programs where you will have to do something like the following:
#NYI 
#NYI     # mod_perl 1
#NYI     ...
#NYI     tie *XLSX, 'Apache';
#NYI     binmode( XLSX );
#NYI     my $workbook = Excel::Writer::XLSX->new( \*XLSX );
#NYI     ...
#NYI 
#NYI     # mod_perl 2
#NYI     ...
#NYI     tie *XLSX => $r;    # Tie to the Apache::RequestRec object
#NYI     binmode( *XLSX );
#NYI     my $workbook = Excel::Writer::XLSX->new( \*XLSX );
#NYI     ...
#NYI 
#NYI See also, the C<mod_perl1.pl> and C<mod_perl2.pl> programs in the C<examples> directory of the distro.
#NYI 
#NYI Filehandles can also be useful if you want to stream an Excel file over a socket or if you want to store an Excel file in a scalar.
#NYI 
#NYI For example here is a way to write an Excel file to a scalar:
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     open my $fh, '>', \my $str or die "Failed to open filehandle: $!";
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( $fh );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     $worksheet->write( 0, 0, 'Hi Excel!' );
#NYI 
#NYI     $workbook->close();
#NYI 
#NYI     # The Excel file in now in $str. Remember to binmode() the output
#NYI     # filehandle before printing it.
#NYI     binmode STDOUT;
#NYI     print $str;
#NYI 
#NYI See also the C<write_to_scalar.pl> and C<filehandle.pl> programs in the C<examples> directory of the distro.
#NYI 
#NYI B<Note about the requirement for> C<binmode()>. An Excel file is comprised of binary data. Therefore, if you are using a filehandle you should ensure that you C<binmode()> it prior to passing it to C<new()>.You should do this regardless of whether you are on a Windows platform or not.
#NYI 
#NYI You don't have to worry about C<binmode()> if you are using filenames instead of filehandles. Excel::Writer::XLSX performs the C<binmode()> internally when it converts the filename to a filehandle. For more information about C<binmode()> see C<perlfunc> and C<perlopentut> in the main Perl documentation.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 add_worksheet( $sheetname )
#NYI 
#NYI At least one worksheet should be added to a new workbook. A worksheet is used to write data into cells:
#NYI 
#NYI     $worksheet1 = $workbook->add_worksheet();               # Sheet1
#NYI     $worksheet2 = $workbook->add_worksheet( 'Foglio2' );    # Foglio2
#NYI     $worksheet3 = $workbook->add_worksheet( 'Data' );       # Data
#NYI     $worksheet4 = $workbook->add_worksheet();               # Sheet4
#NYI 
#NYI If C<$sheetname> is not specified the default Excel convention will be followed, i.e. Sheet1, Sheet2, etc.
#NYI 
#NYI The worksheet name must be a valid Excel worksheet name, i.e. it cannot contain any of the following characters, C<[ ] : * ? / \> and it must be less than 32 characters. In addition, you cannot use the same, case insensitive, C<$sheetname> for more than one worksheet.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 add_format( %properties )
#NYI 
#NYI The C<add_format()> method can be used to create new Format objects which are used to apply formatting to a cell. You can either define the properties at creation time via a hash of property values or later via method calls.
#NYI 
#NYI     $format1 = $workbook->add_format( %props );    # Set properties at creation
#NYI     $format2 = $workbook->add_format();            # Set properties later
#NYI 
#NYI See the L</CELL FORMATTING> section for more details about Format properties and how to set them.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 add_chart( %properties )
#NYI 
#NYI This method is use to create a new chart either as a standalone worksheet (the default) or as an embeddable object that can be inserted into a worksheet via the C<insert_chart()> Worksheet method.
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'column' );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI     type     (required)
#NYI     subtype  (optional)
#NYI     name     (optional)
#NYI     embedded (optional)
#NYI 
#NYI =over
#NYI 
#NYI =item * C<type>
#NYI 
#NYI This is a required parameter. It defines the type of chart that will be created.
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'line' );
#NYI 
#NYI The available types are:
#NYI 
#NYI     area
#NYI     bar
#NYI     column
#NYI     line
#NYI     pie
#NYI     doughnut
#NYI     scatter
#NYI     stock
#NYI 
#NYI =item * C<subtype>
#NYI 
#NYI Used to define a chart subtype where available.
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'bar', subtype => 'stacked' );
#NYI 
#NYI See the L<Excel::Writer::XLSX::Chart> documentation for a list of available chart subtypes.
#NYI 
#NYI =item * C<name>
#NYI 
#NYI Set the name for the chart sheet. The name property is optional and if it isn't supplied will default to C<Chart1 .. n>. The name must be a valid Excel worksheet name. See C<add_worksheet()> for more details on valid sheet names. The C<name> property can be omitted for embedded charts.
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'line', name => 'Results Chart' );
#NYI 
#NYI =item * C<embedded>
#NYI 
#NYI Specifies that the Chart object will be inserted in a worksheet via the C<insert_chart()> Worksheet method. It is an error to try insert a Chart that doesn't have this flag set.
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'line', embedded => 1 );
#NYI 
#NYI     # Configure the chart.
#NYI     ...
#NYI 
#NYI     # Insert the chart into the a worksheet.
#NYI     $worksheet->insert_chart( 'E2', $chart );
#NYI 
#NYI =back
#NYI 
#NYI See Excel::Writer::XLSX::Chart for details on how to configure the chart object once it is created. See also the C<chart_*.pl> programs in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI =head2 add_shape( %properties )
#NYI 
#NYI The C<add_shape()> method can be used to create new shapes that may be inserted into a worksheet.
#NYI 
#NYI You can either define the properties at creation time via a hash of property values or later via method calls.
#NYI 
#NYI     # Set properties at creation.
#NYI     $plus = $workbook->add_shape(
#NYI         type   => 'plus',
#NYI         id     => 3,
#NYI         width  => $pw,
#NYI         height => $ph
#NYI     );
#NYI 
#NYI 
#NYI     # Default rectangle shape. Set properties later.
#NYI     $rect =  $workbook->add_shape();
#NYI 
#NYI See L<Excel::Writer::XLSX::Shape> for details on how to configure the shape object once it is created.
#NYI 
#NYI See also the C<shape*.pl> programs in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI =head2 add_vba_project( 'vbaProject.bin' )
#NYI 
#NYI The C<add_vba_project()> method can be used to add macros or functions to an Excel::Writer::XLSX file using a binary VBA project file that has been extracted from an existing Excel C<xlsm> file.
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'file.xlsm' );
#NYI 
#NYI     $workbook->add_vba_project( './vbaProject.bin' );
#NYI 
#NYI The supplied C<extract_vba> utility can be used to extract the required C<vbaProject.bin> file from an existing Excel file:
#NYI 
#NYI     $ extract_vba file.xlsm
#NYI     Extracted 'vbaProject.bin' successfully
#NYI 
#NYI Macros can be tied to buttons using the worksheet C<insert_button()> method (see the L</WORKSHEET METHODS> section for details):
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro' } );
#NYI 
#NYI Note, Excel uses the file extension C<xlsm> instead of C<xlsx> for files that contain macros. It is advisable to follow the same convention.
#NYI 
#NYI See also the C<macros.pl> example file and the L<WORKING WITH VBA MACROS>.
#NYI 
#NYI 
#NYI 
#NYI =head2 set_vba_name()
#NYI 
#NYI The C<set_vba_name()> method can be used to set the VBA codename for the workbook. This is sometimes required when a C<vbaProject macro> included via C<add_vba_project()> refers to the workbook. The default Excel VBA name of C<ThisWorkbook> is used if a user defined name isn't specified. See also L<WORKING WITH VBA MACROS>.
#NYI 
#NYI 
#NYI =head2 close()
#NYI 
#NYI In general your Excel file will be closed automatically when your program ends or when the Workbook object goes out of scope, however the C<close()> method can be used to explicitly close an Excel file.
#NYI 
#NYI     $workbook->close();
#NYI 
#NYI An explicit C<close()> is required if the file must be closed prior to performing some external action on it such as copying it, reading its size or attaching it to an email.
#NYI 
#NYI In addition, C<close()> may be required to prevent perl's garbage collector from disposing of the Workbook, Worksheet and Format objects in the wrong order. Situations where this can occur are:
#NYI 
#NYI =over 4
#NYI 
#NYI =item *
#NYI 
#NYI If C<my()> was not used to declare the scope of a workbook variable created using C<new()>.
#NYI 
#NYI =item *
#NYI 
#NYI If the C<new()>, C<add_worksheet()> or C<add_format()> methods are called in subroutines.
#NYI 
#NYI =back
#NYI 
#NYI The reason for this is that Excel::Writer::XLSX relies on Perl's C<DESTROY> mechanism to trigger destructor methods in a specific sequence. This may not happen in cases where the Workbook, Worksheet and Format variables are not lexically scoped or where they have different lexical scopes.
#NYI 
#NYI In general, if you create a file with a size of 0 bytes or you fail to create a file you need to call C<close()>.
#NYI 
#NYI The return value of C<close()> is the same as that returned by perl when it closes the file created by C<new()>. This allows you to handle error conditions in the usual way:
#NYI 
#NYI     $workbook->close() or die "Error closing file: $!";
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_size( $width, $height )
#NYI 
#NYI The C<set_size()> method can be used to set the size of a workbook window.
#NYI 
#NYI     $workbook->set_size(1200, 800);
#NYI 
#NYI The Excel window size was used in Excel 2007 to define the width and height of a workbook window within the Multiple Document Interface (MDI). In later versions of Excel for Windows this interface was dropped. This method is currently only useful when setting the window size in Excel for Mac 2011. The units are pixels and the default size is 1073 x 644.
#NYI 
#NYI Note, this doesn't equate exactly to the Excel for Mac pixel size since it is based on the original Excel 2007 for Windows sizing.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_properties()
#NYI 
#NYI The C<set_properties> method can be used to set the document properties of the Excel file created by C<Excel::Writer::XLSX>. These properties are visible when you use the C<< Office Button -> Prepare -> Properties >> option in Excel and are also available to external applications that read or index Windows files.
#NYI 
#NYI The properties should be passed in hash format as follows:
#NYI 
#NYI     $workbook->set_properties(
#NYI         title    => 'This is an example spreadsheet',
#NYI         author   => 'John McNamara',
#NYI         comments => 'Created with Perl and Excel::Writer::XLSX',
#NYI     );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI     title
#NYI     subject
#NYI     author
#NYI     manager
#NYI     company
#NYI     category
#NYI     keywords
#NYI     comments
#NYI     status
#NYI     hyperlink_base
#NYI 
#NYI See also the C<properties.pl> program in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_custom_property( $name, $value, $type)
#NYI 
#NYI The C<set_custom_property> method can be used to set one of more custom document properties not covered by the C<set_properties()> method above. These properties are visible when you use the C<< Office Button -> Prepare -> Properties -> Advanced Properties -> Custom >> option in Excel and are also available to external applications that read or index Windows files.
#NYI 
#NYI The C<set_custom_property> method takes 3 parameters:
#NYI 
#NYI     $workbook-> set_custom_property( $name, $value, $type);
#NYI 
#NYI Where the available types are:
#NYI 
#NYI     text
#NYI     date
#NYI     number
#NYI     bool
#NYI 
#NYI For example:
#NYI 
#NYI     $workbook->set_custom_property( 'Checked by',      'Eve',                  'text'   );
#NYI     $workbook->set_custom_property( 'Date completed',  '2016-12-12T23:00:00Z', 'date'   );
#NYI     $workbook->set_custom_property( 'Document number', '12345' ,               'number' );
#NYI     $workbook->set_custom_property( 'Reference',       '1.2345',               'number' );
#NYI     $workbook->set_custom_property( 'Has review',      1,                      'bool'   );
#NYI     $workbook->set_custom_property( 'Signed off',      0,                      'bool'   );
#NYI     $workbook->set_custom_property( 'Department',      $some_string,           'text'   );
#NYI     $workbook->set_custom_property( 'Scale',           '1.2345678901234',      'number' );
#NYI 
#NYI Dates should by in ISO8601 C<yyyy-mm-ddThh:mm:ss.sssZ> date format in Zulu time, as shown above.
#NYI 
#NYI The C<text> and C<number> types are optional since they can usually be inferred from the data:
#NYI 
#NYI     $workbook->set_custom_property( 'Checked by', 'Eve'    );
#NYI     $workbook->set_custom_property( 'Reference',  '1.2345' );
#NYI 
#NYI 
#NYI The C<$name> and C<$value> parameters are limited to 255 characters by Excel.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 define_name()
#NYI 
#NYI This method is used to defined a name that can be used to represent a value, a single cell or a range of cells in a workbook.
#NYI 
#NYI For example to set a global/workbook name:
#NYI 
#NYI     # Global/workbook names.
#NYI     $workbook->define_name( 'Exchange_rate', '=0.96' );
#NYI     $workbook->define_name( 'Sales',         '=Sheet1!$G$1:$H$10' );
#NYI 
#NYI It is also possible to define a local/worksheet name by prefixing the name with the sheet name using the syntax C<sheetname!definedname>:
#NYI 
#NYI     # Local/worksheet name.
#NYI     $workbook->define_name( 'Sheet2!Sales',  '=Sheet2!$G$1:$G$10' );
#NYI 
#NYI If the sheet name contains spaces or special characters you must enclose it in single quotes like in Excel:
#NYI 
#NYI     $workbook->define_name( "'New Data'!Sales",  '=Sheet2!$G$1:$G$10' );
#NYI 
#NYI See the defined_name.pl program in the examples dir of the distro.
#NYI 
#NYI Refer to the following to see Excel's syntax rules for defined names: L<http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx#BMsyntax_rules_for_names>
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_tempdir()
#NYI 
#NYI C<Excel::Writer::XLSX> stores worksheet data in temporary files prior to assembling the final workbook.
#NYI 
#NYI The C<File::Temp> module is used to create these temporary files. File::Temp uses C<File::Spec> to determine an appropriate location for these files such as C</tmp> or C<c:\windows\temp>. You can find out which directory is used on your system as follows:
#NYI 
#NYI     perl -MFile::Spec -le "print File::Spec->tmpdir()"
#NYI 
#NYI If the default temporary file directory isn't accessible to your application, or doesn't contain enough space, you can specify an alternative location using the C<set_tempdir()> method:
#NYI 
#NYI     $workbook->set_tempdir( '/tmp/writeexcel' );
#NYI     $workbook->set_tempdir( 'c:\windows\temp\writeexcel' );
#NYI 
#NYI The directory for the temporary file must exist, C<set_tempdir()> will not create a new directory.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_custom_color( $index, $red, $green, $blue )
#NYI 
#NYI The method is maintained for backward compatibility with Spreadsheet::WriteExcel. Excel::Writer::XLSX programs don't require this method and colours can be specified using a Html style C<#RRGGBB> value, see L</WORKING WITH COLOURS>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 sheets( 0, 1, ... )
#NYI 
#NYI The C<sheets()> method returns a list, or a sliced list, of the worksheets in a workbook.
#NYI 
#NYI If no arguments are passed the method returns a list of all the worksheets in the workbook. This is useful if you want to repeat an operation on each worksheet:
#NYI 
#NYI     for $worksheet ( $workbook->sheets() ) {
#NYI         print $worksheet->get_name();
#NYI     }
#NYI 
#NYI 
#NYI You can also specify a slice list to return one or more worksheet objects:
#NYI 
#NYI     $worksheet = $workbook->sheets( 0 );
#NYI     $worksheet->write( 'A1', 'Hello' );
#NYI 
#NYI 
#NYI Or since the return value from C<sheets()> is a reference to a worksheet object you can write the above example as:
#NYI 
#NYI     $workbook->sheets( 0 )->write( 'A1', 'Hello' );
#NYI 
#NYI 
#NYI The following example returns the first and last worksheet in a workbook:
#NYI 
#NYI     for $worksheet ( $workbook->sheets( 0, -1 ) ) {
#NYI         # Do something
#NYI     }
#NYI 
#NYI 
#NYI Array slices are explained in the C<perldata> manpage.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 get_worksheet_by_name()
#NYI 
#NYI The C<get_worksheet_by_name()> function return a worksheet or chartsheet object in the workbook using the sheetname:
#NYI 
#NYI     $worksheet = $workbook->get_worksheet_by_name('Sheet1');
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_1904()
#NYI 
#NYI Excel stores dates as real numbers where the integer part stores the number of days since the epoch and the fractional part stores the percentage of the day. The epoch can be either 1900 or 1904. Excel for Windows uses 1900 and Excel for Macintosh uses 1904. However, Excel on either platform will convert automatically between one system and the other.
#NYI 
#NYI Excel::Writer::XLSX stores dates in the 1900 format by default. If you wish to change this you can call the C<set_1904()> workbook method. You can query the current value by calling the C<get_1904()> workbook method. This returns 0 for 1900 and 1 for 1904.
#NYI 
#NYI See also L</DATES AND TIME IN EXCEL> for more information about working with Excel's date system.
#NYI 
#NYI In general you probably won't need to use C<set_1904()>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_optimization()
#NYI 
#NYI The C<set_optimization()> method is used to turn on optimizations in the Excel::Writer::XLSX module. Currently there is only one optimization available and that is to reduce memory usage.
#NYI 
#NYI     $workbook->set_optimization();
#NYI 
#NYI 
#NYI See L</SPEED AND MEMORY USAGE> for more background information.
#NYI 
#NYI Note, that with this optimization turned on a row of data is written and then discarded when a cell in a new row is added via one of the Worksheet C<write_*()> methods. As such data should be written in sequential row order once the optimization is turned on.
#NYI 
#NYI This method must be called before any calls to C<add_worksheet()>.
#NYI 
#NYI 
#NYI 
#NYI =head2 set_calc_mode( $mode )
#NYI 
#NYI Set the calculation mode for formulas in the workbook. This is mainly of use for workbooks with slow formulas where you want to allow the user to calculate them manually.
#NYI 
#NYI The mode parameter can be one of the following strings:
#NYI 
#NYI =over
#NYI 
#NYI =item C<auto>
#NYI 
#NYI The default. Excel will re-calculate formulas when a formula or a value affecting the formula changes.
#NYI 
#NYI =item C<manual>
#NYI 
#NYI Only re-calculate formulas when the user requires it. Generally by pressing F9.
#NYI 
#NYI =item C<auto_except_tables>
#NYI 
#NYI Excel will automatically re-calculate formulas except for tables.
#NYI 
#NYI =back
#NYI 
#NYI =head1 WORKSHEET METHODS
#NYI 
#NYI A new worksheet is created by calling the C<add_worksheet()> method from a workbook object:
#NYI 
#NYI     $worksheet1 = $workbook->add_worksheet();
#NYI     $worksheet2 = $workbook->add_worksheet();
#NYI 
#NYI The following methods are available through a new worksheet:
#NYI 
#NYI     write()
#NYI     write_number()
#NYI     write_string()
#NYI     write_rich_string()
#NYI     keep_leading_zeros()
#NYI     write_blank()
#NYI     write_row()
#NYI     write_col()
#NYI     write_date_time()
#NYI     write_url()
#NYI     write_url_range()
#NYI     write_formula()
#NYI     write_boolean()
#NYI     write_comment()
#NYI     show_comments()
#NYI     set_comments_author()
#NYI     add_write_handler()
#NYI     insert_image()
#NYI     insert_chart()
#NYI     insert_shape()
#NYI     insert_button()
#NYI     data_validation()
#NYI     conditional_formatting()
#NYI     add_sparkline()
#NYI     add_table()
#NYI     get_name()
#NYI     activate()
#NYI     select()
#NYI     hide()
#NYI     set_first_sheet()
#NYI     protect()
#NYI     set_selection()
#NYI     set_row()
#NYI     set_default_row()
#NYI     set_column()
#NYI     outline_settings()
#NYI     freeze_panes()
#NYI     split_panes()
#NYI     merge_range()
#NYI     merge_range_type()
#NYI     set_zoom()
#NYI     right_to_left()
#NYI     hide_zero()
#NYI     set_tab_color()
#NYI     autofilter()
#NYI     filter_column()
#NYI     filter_column_list()
#NYI     set_vba_name()
#NYI 
#NYI 
#NYI 
#NYI =head2 Cell notation
#NYI 
#NYI Excel::Writer::XLSX supports two forms of notation to designate the position of cells: Row-column notation and A1 notation.
#NYI 
#NYI Row-column notation uses a zero based index for both row and column while A1 notation uses the standard Excel alphanumeric sequence of column letter and 1-based row. For example:
#NYI 
#NYI     (0, 0)      # The top left cell in row-column notation.
#NYI     ('A1')      # The top left cell in A1 notation.
#NYI 
#NYI     (1999, 29)  # Row-column notation.
#NYI     ('AD2000')  # The same cell in A1 notation.
#NYI 
#NYI Row-column notation is useful if you are referring to cells programmatically:
#NYI 
#NYI     for my $i ( 0 .. 9 ) {
#NYI         $worksheet->write( $i, 0, 'Hello' );    # Cells A1 to A10
#NYI     }
#NYI 
#NYI A1 notation is useful for setting up a worksheet manually and for working with formulas:
#NYI 
#NYI     $worksheet->write( 'H1', 200 );
#NYI     $worksheet->write( 'H2', '=H1+1' );
#NYI 
#NYI In formulas and applicable methods you can also use the C<A:A> column notation:
#NYI 
#NYI     $worksheet->write( 'A1', '=SUM(B:B)' );
#NYI 
#NYI The C<Excel::Writer::XLSX::Utility> module that is included in the distro contains helper functions for dealing with A1 notation, for example:
#NYI 
#NYI     use Excel::Writer::XLSX::Utility;
#NYI 
#NYI     ( $row, $col ) = xl_cell_to_rowcol( 'C2' );    # (1, 2)
#NYI     $str           = xl_rowcol_to_cell( 1, 2 );    # C2
#NYI 
#NYI For simplicity, the parameter lists for the worksheet method calls in the following sections are given in terms of row-column notation. In all cases it is also possible to use A1 notation.
#NYI 
#NYI Note: in Excel it is also possible to use a R1C1 notation. This is not supported by Excel::Writer::XLSX.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write( $row, $column, $token, $format )
#NYI 
#NYI Excel makes a distinction between data types such as strings, numbers, blanks, formulas and hyperlinks. To simplify the process of writing data the C<write()> method acts as a general alias for several more specific methods:
#NYI 
#NYI     write_string()
#NYI     write_number()
#NYI     write_blank()
#NYI     write_formula()
#NYI     write_url()
#NYI     write_row()
#NYI     write_col()
#NYI 
#NYI The general rule is that if the data looks like a I<something> then a I<something> is written. Here are some examples in both row-column and A1 notation:
#NYI 
#NYI                                                         # Same as:
#NYI     $worksheet->write( 0, 0, 'Hello'                 ); # write_string()
#NYI     $worksheet->write( 1, 0, 'One'                   ); # write_string()
#NYI     $worksheet->write( 2, 0,  2                      ); # write_number()
#NYI     $worksheet->write( 3, 0,  3.00001                ); # write_number()
#NYI     $worksheet->write( 4, 0,  ""                     ); # write_blank()
#NYI     $worksheet->write( 5, 0,  ''                     ); # write_blank()
#NYI     $worksheet->write( 6, 0,  undef                  ); # write_blank()
#NYI     $worksheet->write( 7, 0                          ); # write_blank()
#NYI     $worksheet->write( 8, 0,  'http://www.perl.com/' ); # write_url()
#NYI     $worksheet->write( 'A9',  'ftp://ftp.cpan.org/'  ); # write_url()
#NYI     $worksheet->write( 'A10', 'internal:Sheet1!A1'   ); # write_url()
#NYI     $worksheet->write( 'A11', 'external:c:\foo.xlsx' ); # write_url()
#NYI     $worksheet->write( 'A12', '=A3 + 3*A4'           ); # write_formula()
#NYI     $worksheet->write( 'A13', '=SIN(PI()/4)'         ); # write_formula()
#NYI     $worksheet->write( 'A14', \@array                ); # write_row()
#NYI     $worksheet->write( 'A15', [\@array]              ); # write_col()
#NYI 
#NYI     # And if the keep_leading_zeros property is set:
#NYI     $worksheet->write( 'A16', '2'                    ); # write_number()
#NYI     $worksheet->write( 'A17', '02'                   ); # write_string()
#NYI     $worksheet->write( 'A18', '00002'                ); # write_string()
#NYI 
#NYI     # Write an array formula. Not available in Spreadsheet::WriteExcel.
#NYI     $worksheet->write( 'A19', '{=SUM(A1:B1*A2:B2)}'  ); # write_formula()
#NYI 
#NYI 
#NYI The "looks like" rule is defined by regular expressions:
#NYI 
#NYI C<write_number()> if C<$token> is a number based on the following regex: C<$token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/>.
#NYI 
#NYI C<write_string()> if C<keep_leading_zeros()> is set and C<$token> is an integer with leading zeros based on the following regex: C<$token =~ /^0\d+$/>.
#NYI 
#NYI C<write_blank()> if C<$token> is undef or a blank string: C<undef>, C<""> or C<''>.
#NYI 
#NYI C<write_url()> if C<$token> is a http, https, ftp or mailto URL based on the following regexes: C<$token =~ m|^[fh]tt?ps?://|> or C<$token =~ m|^mailto:|>.
#NYI 
#NYI C<write_url()> if C<$token> is an internal or external sheet reference based on the following regex: C<$token =~ m[^(in|ex)ternal:]>.
#NYI 
#NYI C<write_formula()> if the first character of C<$token> is C<"=">.
#NYI 
#NYI C<write_array_formula()> if the C<$token> matches C</^{=.*}$/>.
#NYI 
#NYI C<write_row()> if C<$token> is an array ref.
#NYI 
#NYI C<write_col()> if C<$token> is an array ref of array refs.
#NYI 
#NYI C<write_string()> if none of the previous conditions apply.
#NYI 
#NYI The C<$format> parameter is optional. It should be a valid Format object, see L</CELL FORMATTING>:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_bold();
#NYI     $format->set_color( 'red' );
#NYI     $format->set_align( 'center' );
#NYI 
#NYI     $worksheet->write( 4, 0, 'Hello', $format );    # Formatted string
#NYI 
#NYI The write() method will ignore empty strings or C<undef> tokens unless a format is also supplied. As such you needn't worry about special handling for empty or C<undef> values in your data. See also the C<write_blank()> method.
#NYI 
#NYI One problem with the C<write()> method is that occasionally data looks like a number but you don't want it treated as a number. For example, zip codes or ID numbers often start with a leading zero. If you write this data as a number then the leading zero(s) will be stripped. You can change this default behaviour by using the C<keep_leading_zeros()> method. While this property is in place any integers with leading zeros will be treated as strings and the zeros will be preserved. See the C<keep_leading_zeros()> section for a full discussion of this issue.
#NYI 
#NYI You can also add your own data handlers to the C<write()> method using C<add_write_handler()>.
#NYI 
#NYI The C<write()> method will also handle Unicode strings in C<UTF-8> format.
#NYI 
#NYI The C<write> methods return:
#NYI 
#NYI     0 for success.
#NYI    -1 for insufficient number of arguments.
#NYI    -2 for row or column out of bounds.
#NYI    -3 for string too long.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_number( $row, $column, $number, $format )
#NYI 
#NYI Write an integer or a float to the cell specified by C<$row> and C<$column>:
#NYI 
#NYI     $worksheet->write_number( 0, 0, 123456 );
#NYI     $worksheet->write_number( 'A2', 2.3451 );
#NYI 
#NYI See the note about L</Cell notation>. The C<$format> parameter is optional.
#NYI 
#NYI In general it is sufficient to use the C<write()> method.
#NYI 
#NYI B<Note>: some versions of Excel 2007 do not display the calculated values of formulas written by Excel::Writer::XLSX. Applying all available Service Packs to Excel should fix this.
#NYI 
#NYI 
#NYI 
#NYI =head2 write_string( $row, $column, $string, $format )
#NYI 
#NYI Write a string to the cell specified by C<$row> and C<$column>:
#NYI 
#NYI     $worksheet->write_string( 0, 0, 'Your text here' );
#NYI     $worksheet->write_string( 'A2', 'or here' );
#NYI 
#NYI The maximum string size is 32767 characters. However the maximum string segment that Excel can display in a cell is 1000. All 32767 characters can be displayed in the formula bar.
#NYI 
#NYI The C<$format> parameter is optional.
#NYI 
#NYI The C<write()> method will also handle strings in C<UTF-8> format. See also the C<unicode_*.pl> programs in the examples directory of the distro.
#NYI 
#NYI In general it is sufficient to use the C<write()> method. However, you may sometimes wish to use the C<write_string()> method to write data that looks like a number but that you don't want treated as a number. For example, zip codes or phone numbers:
#NYI 
#NYI     # Write as a plain string
#NYI     $worksheet->write_string( 'A1', '01209' );
#NYI 
#NYI However, if the user edits this string Excel may convert it back to a number. To get around this you can use the Excel text format C<@>:
#NYI 
#NYI     # Format as a string. Doesn't change to a number when edited
#NYI     my $format1 = $workbook->add_format( num_format => '@' );
#NYI     $worksheet->write_string( 'A2', '01209', $format1 );
#NYI 
#NYI See also the note about L</Cell notation>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_rich_string( $row, $column, $format, $string, ..., $cell_format )
#NYI 
#NYI The C<write_rich_string()> method is used to write strings with multiple formats. For example to write the string "This is B<bold> and this is I<italic>" you would use the following:
#NYI 
#NYI     my $bold   = $workbook->add_format( bold   => 1 );
#NYI     my $italic = $workbook->add_format( italic => 1 );
#NYI 
#NYI     $worksheet->write_rich_string( 'A1',
#NYI         'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );
#NYI 
#NYI The basic rule is to break the string into fragments and put a C<$format> object before the fragment that you want to format. For example:
#NYI 
#NYI     # Unformatted string.
#NYI       'This is an example string'
#NYI 
#NYI     # Break it into fragments.
#NYI       'This is an ', 'example', ' string'
#NYI 
#NYI     # Add formatting before the fragments you want formatted.
#NYI       'This is an ', $format, 'example', ' string'
#NYI 
#NYI     # In Excel::Writer::XLSX.
#NYI     $worksheet->write_rich_string( 'A1',
#NYI         'This is an ', $format, 'example', ' string' );
#NYI 
#NYI String fragments that don't have a format are given a default format. So for example when writing the string "Some B<bold> text" you would use the first example below but it would be equivalent to the second:
#NYI 
#NYI     # With default formatting:
#NYI     my $bold    = $workbook->add_format( bold => 1 );
#NYI 
#NYI     $worksheet->write_rich_string( 'A1',
#NYI         'Some ', $bold, 'bold', ' text' );
#NYI 
#NYI     # Or more explicitly:
#NYI     my $bold    = $workbook->add_format( bold => 1 );
#NYI     my $default = $workbook->add_format();
#NYI 
#NYI     $worksheet->write_rich_string( 'A1',
#NYI         $default, 'Some ', $bold, 'bold', $default, ' text' );
#NYI 
#NYI As with Excel, only the font properties of the format such as font name, style, size, underline, color and effects are applied to the string fragments. Other features such as border, background, text wrap and alignment must be applied to the cell.
#NYI 
#NYI The C<write_rich_string()> method allows you to do this by using the last argument as a cell format (if it is a format object). The following example centers a rich string in the cell:
#NYI 
#NYI     my $bold   = $workbook->add_format( bold  => 1 );
#NYI     my $center = $workbook->add_format( align => 'center' );
#NYI 
#NYI     $worksheet->write_rich_string( 'A5',
#NYI         'Some ', $bold, 'bold text', ' centered', $center );
#NYI 
#NYI See the C<rich_strings.pl> example in the distro for more examples.
#NYI 
#NYI     my $bold   = $workbook->add_format( bold        => 1 );
#NYI     my $italic = $workbook->add_format( italic      => 1 );
#NYI     my $red    = $workbook->add_format( color       => 'red' );
#NYI     my $blue   = $workbook->add_format( color       => 'blue' );
#NYI     my $center = $workbook->add_format( align       => 'center' );
#NYI     my $super  = $workbook->add_format( font_script => 1 );
#NYI 
#NYI 
#NYI     # Write some strings with multiple formats.
#NYI     $worksheet->write_rich_string( 'A1',
#NYI         'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );
#NYI 
#NYI     $worksheet->write_rich_string( 'A3',
#NYI         'This is ', $red, 'red', ' and this is ', $blue, 'blue' );
#NYI 
#NYI     $worksheet->write_rich_string( 'A5',
#NYI         'Some ', $bold, 'bold text', ' centered', $center );
#NYI 
#NYI     $worksheet->write_rich_string( 'A7',
#NYI         $italic, 'j = k', $super, '(n-1)', $center );
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/rich_strings.jpg" width="640" height="420" alt="Output from rich_strings.pl" /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI As with C<write_sting()> the maximum string size is 32767 characters. See also the note about L</Cell notation>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 keep_leading_zeros()
#NYI 
#NYI This method changes the default handling of integers with leading zeros when using the C<write()> method.
#NYI 
#NYI The C<write()> method uses regular expressions to determine what type of data to write to an Excel worksheet. If the data looks like a number it writes a number using C<write_number()>. One problem with this approach is that occasionally data looks like a number but you don't want it treated as a number.
#NYI 
#NYI Zip codes and ID numbers, for example, often start with a leading zero. If you write this data as a number then the leading zero(s) will be stripped. This is the also the default behaviour when you enter data manually in Excel.
#NYI 
#NYI To get around this you can use one of three options. Write a formatted number, write the number as a string or use the C<keep_leading_zeros()> method to change the default behaviour of C<write()>:
#NYI 
#NYI     # Implicitly write a number, the leading zero is removed: 1209
#NYI     $worksheet->write( 'A1', '01209' );
#NYI 
#NYI     # Write a zero padded number using a format: 01209
#NYI     my $format1 = $workbook->add_format( num_format => '00000' );
#NYI     $worksheet->write( 'A2', '01209', $format1 );
#NYI 
#NYI     # Write explicitly as a string: 01209
#NYI     $worksheet->write_string( 'A3', '01209' );
#NYI 
#NYI     # Write implicitly as a string: 01209
#NYI     $worksheet->keep_leading_zeros();
#NYI     $worksheet->write( 'A4', '01209' );
#NYI 
#NYI 
#NYI The above code would generate a worksheet that looked like the following:
#NYI 
#NYI      -----------------------------------------------------------
#NYI     |   |     A     |     B     |     C     |     D     | ...
#NYI      -----------------------------------------------------------
#NYI     | 1 |      1209 |           |           |           | ...
#NYI     | 2 |     01209 |           |           |           | ...
#NYI     | 3 | 01209     |           |           |           | ...
#NYI     | 4 | 01209     |           |           |           | ...
#NYI 
#NYI 
#NYI The examples are on different sides of the cells due to the fact that Excel displays strings with a left justification and numbers with a right justification by default. You can change this by using a format to justify the data, see L</CELL FORMATTING>.
#NYI 
#NYI It should be noted that if the user edits the data in examples C<A3> and C<A4> the strings will revert back to numbers. Again this is Excel's default behaviour. To avoid this you can use the text format C<@>:
#NYI 
#NYI     # Format as a string (01209)
#NYI     my $format2 = $workbook->add_format( num_format => '@' );
#NYI     $worksheet->write_string( 'A5', '01209', $format2 );
#NYI 
#NYI The C<keep_leading_zeros()> property is off by default. The C<keep_leading_zeros()> method takes 0 or 1 as an argument. It defaults to 1 if an argument isn't specified:
#NYI 
#NYI     $worksheet->keep_leading_zeros();       # Set on
#NYI     $worksheet->keep_leading_zeros( 1 );    # Set on
#NYI     $worksheet->keep_leading_zeros( 0 );    # Set off
#NYI 
#NYI See also the C<add_write_handler()> method.
#NYI 
#NYI 
#NYI =head2 write_blank( $row, $column, $format )
#NYI 
#NYI Write a blank cell specified by C<$row> and C<$column>:
#NYI 
#NYI     $worksheet->write_blank( 0, 0, $format );
#NYI 
#NYI This method is used to add formatting to a cell which doesn't contain a string or number value.
#NYI 
#NYI Excel differentiates between an "Empty" cell and a "Blank" cell. An "Empty" cell is a cell which doesn't contain data whilst a "Blank" cell is a cell which doesn't contain data but does contain formatting. Excel stores "Blank" cells but ignores "Empty" cells.
#NYI 
#NYI As such, if you write an empty cell without formatting it is ignored:
#NYI 
#NYI     $worksheet->write( 'A1', undef, $format );    # write_blank()
#NYI     $worksheet->write( 'A2', undef );             # Ignored
#NYI 
#NYI This seemingly uninteresting fact means that you can write arrays of data without special treatment for C<undef> or empty string values.
#NYI 
#NYI See the note about L</Cell notation>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_row( $row, $column, $array_ref, $format )
#NYI 
#NYI The C<write_row()> method can be used to write a 1D or 2D array of data in one go. This is useful for converting the results of a database query into an Excel worksheet. You must pass a reference to the array of data rather than the array itself. The C<write()> method is then called for each element of the data. For example:
#NYI 
#NYI     @array = ( 'awk', 'gawk', 'mawk' );
#NYI     $array_ref = \@array;
#NYI 
#NYI     $worksheet->write_row( 0, 0, $array_ref );
#NYI 
#NYI     # The above example is equivalent to:
#NYI     $worksheet->write( 0, 0, $array[0] );
#NYI     $worksheet->write( 0, 1, $array[1] );
#NYI     $worksheet->write( 0, 2, $array[2] );
#NYI 
#NYI 
#NYI Note: For convenience the C<write()> method behaves in the same way as C<write_row()> if it is passed an array reference. Therefore the following two method calls are equivalent:
#NYI 
#NYI     $worksheet->write_row( 'A1', $array_ref );    # Write a row of data
#NYI     $worksheet->write(     'A1', $array_ref );    # Same thing
#NYI 
#NYI As with all of the write methods the C<$format> parameter is optional. If a format is specified it is applied to all the elements of the data array.
#NYI 
#NYI Array references within the data will be treated as columns. This allows you to write 2D arrays of data in one go. For example:
#NYI 
#NYI     @eec =  (
#NYI                 ['maggie', 'milly', 'molly', 'may'  ],
#NYI                 [13,       14,      15,      16     ],
#NYI                 ['shell',  'star',  'crab',  'stone']
#NYI             );
#NYI 
#NYI     $worksheet->write_row( 'A1', \@eec );
#NYI 
#NYI 
#NYI Would produce a worksheet as follows:
#NYI 
#NYI      -----------------------------------------------------------
#NYI     |   |    A    |    B    |    C    |    D    |    E    | ...
#NYI      -----------------------------------------------------------
#NYI     | 1 | maggie  | 13      | shell   | ...     |  ...    | ...
#NYI     | 2 | milly   | 14      | star    | ...     |  ...    | ...
#NYI     | 3 | molly   | 15      | crab    | ...     |  ...    | ...
#NYI     | 4 | may     | 16      | stone   | ...     |  ...    | ...
#NYI     | 5 | ...     | ...     | ...     | ...     |  ...    | ...
#NYI     | 6 | ...     | ...     | ...     | ...     |  ...    | ...
#NYI 
#NYI 
#NYI To write the data in a row-column order refer to the C<write_col()> method below.
#NYI 
#NYI Any C<undef> values in the data will be ignored unless a format is applied to the data, in which case a formatted blank cell will be written. In either case the appropriate row or column value will still be incremented.
#NYI 
#NYI To find out more about array references refer to C<perlref> and C<perlreftut> in the main Perl documentation. To find out more about 2D arrays or "lists of lists" refer to C<perllol>.
#NYI 
#NYI The C<write_row()> method returns the first error encountered when writing the elements of the data or zero if no errors were encountered. See the return values described for the C<write()> method above.
#NYI 
#NYI See also the C<write_arrays.pl> program in the C<examples> directory of the distro.
#NYI 
#NYI The C<write_row()> method allows the following idiomatic conversion of a text file to an Excel file:
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'file.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     open INPUT, 'file.txt' or die "Couldn't open file: $!";
#NYI 
#NYI     $worksheet->write( $. -1, 0, [split] ) while <INPUT>;
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_col( $row, $column, $array_ref, $format )
#NYI 
#NYI The C<write_col()> method can be used to write a 1D or 2D array of data in one go. This is useful for converting the results of a database query into an Excel worksheet. You must pass a reference to the array of data rather than the array itself. The C<write()> method is then called for each element of the data. For example:
#NYI 
#NYI     @array = ( 'awk', 'gawk', 'mawk' );
#NYI     $array_ref = \@array;
#NYI 
#NYI     $worksheet->write_col( 0, 0, $array_ref );
#NYI 
#NYI     # The above example is equivalent to:
#NYI     $worksheet->write( 0, 0, $array[0] );
#NYI     $worksheet->write( 1, 0, $array[1] );
#NYI     $worksheet->write( 2, 0, $array[2] );
#NYI 
#NYI As with all of the write methods the C<$format> parameter is optional. If a format is specified it is applied to all the elements of the data array.
#NYI 
#NYI Array references within the data will be treated as rows. This allows you to write 2D arrays of data in one go. For example:
#NYI 
#NYI     @eec =  (
#NYI                 ['maggie', 'milly', 'molly', 'may'  ],
#NYI                 [13,       14,      15,      16     ],
#NYI                 ['shell',  'star',  'crab',  'stone']
#NYI             );
#NYI 
#NYI     $worksheet->write_col( 'A1', \@eec );
#NYI 
#NYI 
#NYI Would produce a worksheet as follows:
#NYI 
#NYI      -----------------------------------------------------------
#NYI     |   |    A    |    B    |    C    |    D    |    E    | ...
#NYI      -----------------------------------------------------------
#NYI     | 1 | maggie  | milly   | molly   | may     |  ...    | ...
#NYI     | 2 | 13      | 14      | 15      | 16      |  ...    | ...
#NYI     | 3 | shell   | star    | crab    | stone   |  ...    | ...
#NYI     | 4 | ...     | ...     | ...     | ...     |  ...    | ...
#NYI     | 5 | ...     | ...     | ...     | ...     |  ...    | ...
#NYI     | 6 | ...     | ...     | ...     | ...     |  ...    | ...
#NYI 
#NYI 
#NYI To write the data in a column-row order refer to the C<write_row()> method above.
#NYI 
#NYI Any C<undef> values in the data will be ignored unless a format is applied to the data, in which case a formatted blank cell will be written. In either case the appropriate row or column value will still be incremented.
#NYI 
#NYI As noted above the C<write()> method can be used as a synonym for C<write_row()> and C<write_row()> handles nested array refs as columns. Therefore, the following two method calls are equivalent although the more explicit call to C<write_col()> would be preferable for maintainability:
#NYI 
#NYI     $worksheet->write_col( 'A1', $array_ref     ); # Write a column of data
#NYI     $worksheet->write(     'A1', [ $array_ref ] ); # Same thing
#NYI 
#NYI To find out more about array references refer to C<perlref> and C<perlreftut> in the main Perl documentation. To find out more about 2D arrays or "lists of lists" refer to C<perllol>.
#NYI 
#NYI The C<write_col()> method returns the first error encountered when writing the elements of the data or zero if no errors were encountered. See the return values described for the C<write()> method above.
#NYI 
#NYI See also the C<write_arrays.pl> program in the C<examples> directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_date_time( $row, $col, $date_string, $format )
#NYI 
#NYI The C<write_date_time()> method can be used to write a date or time to the cell specified by C<$row> and C<$column>:
#NYI 
#NYI     $worksheet->write_date_time( 'A1', '2004-05-13T23:20', $date_format );
#NYI 
#NYI The C<$date_string> should be in the following format:
#NYI 
#NYI     yyyy-mm-ddThh:mm:ss.sss
#NYI 
#NYI This conforms to an ISO8601 date but it should be noted that the full range of ISO8601 formats are not supported.
#NYI 
#NYI The following variations on the C<$date_string> parameter are permitted:
#NYI 
#NYI     yyyy-mm-ddThh:mm:ss.sss         # Standard format
#NYI     yyyy-mm-ddT                     # No time
#NYI               Thh:mm:ss.sss         # No date
#NYI     yyyy-mm-ddThh:mm:ss.sssZ        # Additional Z (but not time zones)
#NYI     yyyy-mm-ddThh:mm:ss             # No fractional seconds
#NYI     yyyy-mm-ddThh:mm                # No seconds
#NYI 
#NYI Note that the C<T> is required in all cases.
#NYI 
#NYI A date should always have a C<$format>, otherwise it will appear as a number, see L</DATES AND TIME IN EXCEL> and L</CELL FORMATTING>. Here is a typical example:
#NYI 
#NYI     my $date_format = $workbook->add_format( num_format => 'mm/dd/yy' );
#NYI     $worksheet->write_date_time( 'A1', '2004-05-13T23:20', $date_format );
#NYI 
#NYI Valid dates should be in the range 1900-01-01 to 9999-12-31, for the 1900 epoch and 1904-01-01 to 9999-12-31, for the 1904 epoch. As with Excel, dates outside these ranges will be written as a string.
#NYI 
#NYI See also the date_time.pl program in the C<examples> directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_url( $row, $col, $url, $format, $label )
#NYI 
#NYI Write a hyperlink to a URL in the cell specified by C<$row> and C<$column>. The hyperlink is comprised of two elements: the visible label and the invisible link. The visible label is the same as the link unless an alternative label is specified. The C<$label> parameter is optional. The label is written using the C<write()> method. Therefore it is possible to write strings, numbers or formulas as labels.
#NYI 
#NYI The C<$format> parameter is also optional, however, without a format the link won't look like a link.
#NYI 
#NYI The suggested format is:
#NYI 
#NYI     my $format = $workbook->add_format( color => 'blue', underline => 1 );
#NYI 
#NYI B<Note>, this behaviour is different from Spreadsheet::WriteExcel which provides a default hyperlink format if one isn't specified by the user.
#NYI 
#NYI There are four web style URI's supported: C<http://>, C<https://>, C<ftp://> and C<mailto:>:
#NYI 
#NYI     $worksheet->write_url( 0, 0, 'ftp://www.perl.org/',       $format );
#NYI     $worksheet->write_url( 'A3', 'http://www.perl.com/',      $format );
#NYI     $worksheet->write_url( 'A4', 'mailto:jmcnamara@cpan.org', $format );
#NYI 
#NYI You can display an alternative string using the C<$label> parameter:
#NYI 
#NYI     $worksheet->write_url( 1, 0, 'http://www.perl.com/', $format, 'Perl' );
#NYI 
#NYI If you wish to have some other cell data such as a number or a formula you can overwrite the cell using another call to C<write_*()>:
#NYI 
#NYI     $worksheet->write_url( 'A1', 'http://www.perl.com/' );
#NYI 
#NYI     # Overwrite the URL string with a formula. The cell is still a link.
#NYI     $worksheet->write_formula( 'A1', '=1+1', $format );
#NYI 
#NYI There are two local URIs supported: C<internal:> and C<external:>. These are used for hyperlinks to internal worksheet references or external workbook and worksheet references:
#NYI 
#NYI     $worksheet->write_url( 'A6',  'internal:Sheet2!A1',              $format );
#NYI     $worksheet->write_url( 'A7',  'internal:Sheet2!A1',              $format );
#NYI     $worksheet->write_url( 'A8',  'internal:Sheet2!A1:B2',           $format );
#NYI     $worksheet->write_url( 'A9',  q{internal:'Sales Data'!A1},       $format );
#NYI     $worksheet->write_url( 'A10', 'external:c:\temp\foo.xlsx',       $format );
#NYI     $worksheet->write_url( 'A11', 'external:c:\foo.xlsx#Sheet2!A1',  $format );
#NYI     $worksheet->write_url( 'A12', 'external:..\foo.xlsx',            $format );
#NYI     $worksheet->write_url( 'A13', 'external:..\foo.xlsx#Sheet2!A1',  $format );
#NYI     $worksheet->write_url( 'A13', 'external:\\\\NET\share\foo.xlsx', $format );
#NYI 
#NYI All of the these URI types are recognised by the C<write()> method, see above.
#NYI 
#NYI Worksheet references are typically of the form C<Sheet1!A1>. You can also refer to a worksheet range using the standard Excel notation: C<Sheet1!A1:B2>.
#NYI 
#NYI In external links the workbook and worksheet name must be separated by the C<#> character: C<external:Workbook.xlsx#Sheet1!A1'>.
#NYI 
#NYI You can also link to a named range in the target worksheet. For example say you have a named range called C<my_name> in the workbook C<c:\temp\foo.xlsx> you could link to it as follows:
#NYI 
#NYI     $worksheet->write_url( 'A14', 'external:c:\temp\foo.xlsx#my_name' );
#NYI 
#NYI Excel requires that worksheet names containing spaces or non alphanumeric characters are single quoted as follows C<'Sales Data'!A1>. If you need to do this in a single quoted string then you can either escape the single quotes C<\'> or use the quote operator C<q{}> as described in C<perlop> in the main Perl documentation.
#NYI 
#NYI Links to network files are also supported. MS/Novell Network files normally begin with two back slashes as follows C<\\NETWORK\etc>. In order to generate this in a single or double quoted string you will have to escape the backslashes,  C<'\\\\NETWORK\etc'>.
#NYI 
#NYI If you are using double quote strings then you should be careful to escape anything that looks like a metacharacter. For more information see C<perlfaq5: Why can't I use "C:\temp\foo" in DOS paths?>.
#NYI 
#NYI Finally, you can avoid most of these quoting problems by using forward slashes. These are translated internally to backslashes:
#NYI 
#NYI     $worksheet->write_url( 'A14', "external:c:/temp/foo.xlsx" );
#NYI     $worksheet->write_url( 'A15', 'external://NETWORK/share/foo.xlsx' );
#NYI 
#NYI Note: Excel::Writer::XLSX will escape the following characters in URLs as required by Excel: C<< \s " < > \ [  ] ` ^ { } >> unless the URL already contains C<%xx> style escapes. In which case it is assumed that the URL was escaped correctly by the user and will by passed directly to Excel.
#NYI 
#NYI Excel limits hyperlink links and anchor/locations to 255 characters each.
#NYI 
#NYI See also, the note about L</Cell notation>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_formula( $row, $column, $formula, $format, $value )
#NYI 
#NYI Write a formula or function to the cell specified by C<$row> and C<$column>:
#NYI 
#NYI     $worksheet->write_formula( 0, 0, '=$B$3 + B4' );
#NYI     $worksheet->write_formula( 1, 0, '=SIN(PI()/4)' );
#NYI     $worksheet->write_formula( 2, 0, '=SUM(B1:B5)' );
#NYI     $worksheet->write_formula( 'A4', '=IF(A3>1,"Yes", "No")' );
#NYI     $worksheet->write_formula( 'A5', '=AVERAGE(1, 2, 3, 4)' );
#NYI     $worksheet->write_formula( 'A6', '=DATEVALUE("1-Jan-2001")' );
#NYI 
#NYI Array formulas are also supported:
#NYI 
#NYI     $worksheet->write_formula( 'A7', '{=SUM(A1:B1*A2:B2)}' );
#NYI 
#NYI See also the C<write_array_formula()> method below.
#NYI 
#NYI See the note about L</Cell notation>. For more information about writing Excel formulas see L</FORMULAS AND FUNCTIONS IN EXCEL>
#NYI 
#NYI If required, it is also possible to specify the calculated value of the formula. This is occasionally necessary when working with non-Excel applications that don't calculate the value of the formula. The calculated C<$value> is added at the end of the argument list:
#NYI 
#NYI     $worksheet->write( 'A1', '=2+2', $format, 4 );
#NYI 
#NYI However, this probably isn't something that you will ever need to do. If you do use this feature then do so with care.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_array_formula($first_row, $first_col, $last_row, $last_col, $formula, $format, $value)
#NYI 
#NYI Write an array formula to a cell range. In Excel an array formula is a formula that performs a calculation on a set of values. It can return a single value or a range of values.
#NYI 
#NYI An array formula is indicated by a pair of braces around the formula: C<{=SUM(A1:B1*A2:B2)}>.  If the array formula returns a single value then the C<$first_> and C<$last_> parameters should be the same:
#NYI 
#NYI     $worksheet->write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}');
#NYI 
#NYI It this case however it is easier to just use the C<write_formula()> or C<write()> methods:
#NYI 
#NYI     # Same as above but more concise.
#NYI     $worksheet->write( 'A1', '{=SUM(B1:C1*B2:C2)}' );
#NYI     $worksheet->write_formula( 'A1', '{=SUM(B1:C1*B2:C2)}' );
#NYI 
#NYI For array formulas that return a range of values you must specify the range that the return values will be written to:
#NYI 
#NYI     $worksheet->write_array_formula( 'A1:A3',    '{=TREND(C1:C3,B1:B3)}' );
#NYI     $worksheet->write_array_formula( 0, 0, 2, 0, '{=TREND(C1:C3,B1:B3)}' );
#NYI 
#NYI If required, it is also possible to specify the calculated value of the formula. This is occasionally necessary when working with non-Excel applications that don't calculate the value of the formula. However, using this parameter only writes a single value to the upper left cell in the result array. For a multi-cell array formula where the results are required, the other result values can be specified by using C<write_number()> to write to the appropriate cell:
#NYI 
#NYI     # Specify the result for a single cell range.
#NYI     $worksheet->write_array_formula( 'A1:A3', '{=SUM(B1:C1*B2:C2)}, $format, 2005 );
#NYI 
#NYI     # Specify the results for a multi cell range.
#NYI     $worksheet->write_array_formula( 'A1:A3', '{=TREND(C1:C3,B1:B3)}', $format, 105 );
#NYI     $worksheet->write_number( 'A2', 12, format );
#NYI     $worksheet->write_number( 'A3', 14, format );
#NYI 
#NYI In addition, some early versions of Excel 2007 don't calculate the values of array formulas when they aren't supplied. Installing the latest Office Service Pack should fix this issue.
#NYI 
#NYI See also the C<array_formula.pl> program in the C<examples> directory of the distro.
#NYI 
#NYI Note: Array formulas are not supported by Spreadsheet::WriteExcel.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_boolean( $row, $column, $value, $format )
#NYI 
#NYI Write an Excel boolean value to the cell specified by C<$row> and C<$column>:
#NYI 
#NYI     $worksheet->write_boolean( 'A1', 1          );  # TRUE
#NYI     $worksheet->write_boolean( 'A2', 0          );  # FALSE
#NYI     $worksheet->write_boolean( 'A3', undef      );  # FALSE
#NYI     $worksheet->write_boolean( 'A3', 0, $format );  # FALSE, with format.
#NYI 
#NYI A C<$value> that is true or false using Perl's rules will be written as an Excel boolean C<TRUE> or C<FALSE> value.
#NYI 
#NYI See the note about L</Cell notation>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 store_formula( $formula )
#NYI 
#NYI Deprecated. This is a Spreadsheet::WriteExcel method that is no longer required by Excel::Writer::XLSX. See below.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 repeat_formula( $row, $col, $formula, $format )
#NYI 
#NYI Deprecated. This is a Spreadsheet::WriteExcel method that is no longer required by Excel::Writer::XLSX.
#NYI 
#NYI In Spreadsheet::WriteExcel it was computationally expensive to write formulas since they were parsed by a recursive descent parser. The C<store_formula()> and C<repeat_formula()> methods were used as a way of avoiding the overhead of repeated formulas by reusing a pre-parsed formula.
#NYI 
#NYI In Excel::Writer::XLSX this is no longer necessary since it is just as quick to write a formula as it is to write a string or a number.
#NYI 
#NYI The methods remain for backward compatibility but new Excel::Writer::XLSX programs shouldn't use them.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 write_comment( $row, $column, $string, ... )
#NYI 
#NYI The C<write_comment()> method is used to add a comment to a cell. A cell comment is indicated in Excel by a small red triangle in the upper right-hand corner of the cell. Moving the cursor over the red triangle will reveal the comment.
#NYI 
#NYI The following example shows how to add a comment to a cell:
#NYI 
#NYI     $worksheet->write        ( 2, 2, 'Hello' );
#NYI     $worksheet->write_comment( 2, 2, 'This is a comment.' );
#NYI 
#NYI As usual you can replace the C<$row> and C<$column> parameters with an C<A1> cell reference. See the note about L</Cell notation>.
#NYI 
#NYI     $worksheet->write        ( 'C3', 'Hello');
#NYI     $worksheet->write_comment( 'C3', 'This is a comment.' );
#NYI 
#NYI The C<write_comment()> method will also handle strings in C<UTF-8> format.
#NYI 
#NYI     $worksheet->write_comment( 'C3', "\x{263a}" );       # Smiley
#NYI     $worksheet->write_comment( 'C4', 'Comment ca va?' );
#NYI 
#NYI In addition to the basic 3 argument form of C<write_comment()> you can pass in several optional key/value pairs to control the format of the comment. For example:
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', visible => 1, author => 'Perl' );
#NYI 
#NYI Most of these options are quite specific and in general the default comment behaves will be all that you need. However, should you need greater control over the format of the cell comment the following options are available:
#NYI 
#NYI     author
#NYI     visible
#NYI     x_scale
#NYI     width
#NYI     y_scale
#NYI     height
#NYI     color
#NYI     start_cell
#NYI     start_row
#NYI     start_col
#NYI     x_offset
#NYI     y_offset
#NYI 
#NYI 
#NYI =over 4
#NYI 
#NYI =item Option: author
#NYI 
#NYI This option is used to indicate who is the author of the cell comment. Excel displays the author of the comment in the status bar at the bottom of the worksheet. This is usually of interest in corporate environments where several people might review and provide comments to a workbook.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Atonement', author => 'Ian McEwan' );
#NYI 
#NYI The default author for all cell comments can be set using the C<set_comments_author()> method (see below).
#NYI 
#NYI     $worksheet->set_comments_author( 'Perl' );
#NYI 
#NYI 
#NYI =item Option: visible
#NYI 
#NYI This option is used to make a cell comment visible when the worksheet is opened. The default behaviour in Excel is that comments are initially hidden. However, it is also possible in Excel to make individual or all comments visible. In Excel::Writer::XLSX individual comments can be made visible as follows:
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', visible => 1 );
#NYI 
#NYI It is possible to make all comments in a worksheet visible using the C<show_comments()> worksheet method (see below). Alternatively, if all of the cell comments have been made visible you can hide individual comments:
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', visible => 0 );
#NYI 
#NYI 
#NYI =item Option: x_scale
#NYI 
#NYI This option is used to set the width of the cell comment box as a factor of the default width.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', x_scale => 2 );
#NYI     $worksheet->write_comment( 'C4', 'Hello', x_scale => 4.2 );
#NYI 
#NYI 
#NYI =item Option: width
#NYI 
#NYI This option is used to set the width of the cell comment box explicitly in pixels.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', width => 200 );
#NYI 
#NYI 
#NYI =item Option: y_scale
#NYI 
#NYI This option is used to set the height of the cell comment box as a factor of the default height.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', y_scale => 2 );
#NYI     $worksheet->write_comment( 'C4', 'Hello', y_scale => 4.2 );
#NYI 
#NYI 
#NYI =item Option: height
#NYI 
#NYI This option is used to set the height of the cell comment box explicitly in pixels.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', height => 200 );
#NYI 
#NYI 
#NYI =item Option: color
#NYI 
#NYI This option is used to set the background colour of cell comment box. You can use one of the named colours recognised by Excel::Writer::XLSX or a Html style C<#RRGGBB> colour. See L</WORKING WITH COLOURS>.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', color => 'green' );
#NYI     $worksheet->write_comment( 'C4', 'Hello', color => '#FF6600' ); # Orange
#NYI 
#NYI 
#NYI =item Option: start_cell
#NYI 
#NYI This option is used to set the cell in which the comment will appear. By default Excel displays comments one cell to the right and one cell above the cell to which the comment relates. However, you can change this behaviour if you wish. In the following example the comment which would appear by default in cell C<D2> is moved to C<E2>.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', start_cell => 'E2' );
#NYI 
#NYI 
#NYI =item Option: start_row
#NYI 
#NYI This option is used to set the row in which the comment will appear. See the C<start_cell> option above. The row is zero indexed.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', start_row => 0 );
#NYI 
#NYI 
#NYI =item Option: start_col
#NYI 
#NYI This option is used to set the column in which the comment will appear. See the C<start_cell> option above. The column is zero indexed.
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', start_col => 4 );
#NYI 
#NYI 
#NYI =item Option: x_offset
#NYI 
#NYI This option is used to change the x offset, in pixels, of a comment within a cell:
#NYI 
#NYI     $worksheet->write_comment( 'C3', $comment, x_offset => 30 );
#NYI 
#NYI 
#NYI =item Option: y_offset
#NYI 
#NYI This option is used to change the y offset, in pixels, of a comment within a cell:
#NYI 
#NYI     $worksheet->write_comment('C3', $comment, x_offset => 30);
#NYI 
#NYI 
#NYI =back
#NYI 
#NYI You can apply as many of these options as you require.
#NYI 
#NYI B<Note about using options that adjust the position of the cell comment such as start_cell, start_row, start_col, x_offset and y_offset>: Excel only displays offset cell comments when they are displayed as "visible". Excel does B<not> display hidden cells as moved when you mouse over them.
#NYI 
#NYI B<Note about row height and comments>. If you specify the height of a row that contains a comment then Excel::Writer::XLSX will adjust the height of the comment to maintain the default or user specified dimensions. However, the height of a row can also be adjusted automatically by Excel if the text wrap property is set or large fonts are used in the cell. This means that the height of the row is unknown to the module at run time and thus the comment box is stretched with the row. Use the C<set_row()> method to specify the row height explicitly and avoid this problem.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 show_comments()
#NYI 
#NYI This method is used to make all cell comments visible when a worksheet is opened.
#NYI 
#NYI     $worksheet->show_comments();
#NYI 
#NYI Individual comments can be made visible using the C<visible> parameter of the C<write_comment> method (see above):
#NYI 
#NYI     $worksheet->write_comment( 'C3', 'Hello', visible => 1 );
#NYI 
#NYI If all of the cell comments have been made visible you can hide individual comments as follows:
#NYI 
#NYI     $worksheet->show_comments();
#NYI     $worksheet->write_comment( 'C3', 'Hello', visible => 0 );
#NYI 
#NYI 
#NYI 
#NYI =head2 set_comments_author()
#NYI 
#NYI This method is used to set the default author of all cell comments.
#NYI 
#NYI     $worksheet->set_comments_author( 'Perl' );
#NYI 
#NYI Individual comment authors can be set using the C<author> parameter of the C<write_comment> method (see above).
#NYI 
#NYI The default comment author is an empty string, C<''>, if no author is specified.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 add_write_handler( $re, $code_ref )
#NYI 
#NYI This method is used to extend the Excel::Writer::XLSX write() method to handle user defined data.
#NYI 
#NYI If you refer to the section on C<write()> above you will see that it acts as an alias for several more specific C<write_*> methods. However, it doesn't always act in exactly the way that you would like it to.
#NYI 
#NYI One solution is to filter the input data yourself and call the appropriate C<write_*> method. Another approach is to use the C<add_write_handler()> method to add your own automated behaviour to C<write()>.
#NYI 
#NYI The C<add_write_handler()> method take two arguments, C<$re>, a regular expression to match incoming data and C<$code_ref> a callback function to handle the matched data:
#NYI 
#NYI     $worksheet->add_write_handler( qr/^\d\d\d\d$/, \&my_write );
#NYI 
#NYI (In the these examples the C<qr> operator is used to quote the regular expression strings, see L<perlop> for more details).
#NYI 
#NYI The method is used as follows. say you wished to write 7 digit ID numbers as a string so that any leading zeros were preserved*, you could do something like the following:
#NYI 
#NYI     $worksheet->add_write_handler( qr/^\d{7}$/, \&write_my_id );
#NYI 
#NYI 
#NYI     sub write_my_id {
#NYI         my $worksheet = shift;
#NYI         return $worksheet->write_string( @_ );
#NYI     }
#NYI 
#NYI * You could also use the C<keep_leading_zeros()> method for this.
#NYI 
#NYI Then if you call C<write()> with an appropriate string it will be handled automatically:
#NYI 
#NYI     # Writes 0000000. It would normally be written as a number; 0.
#NYI     $worksheet->write( 'A1', '0000000' );
#NYI 
#NYI The callback function will receive a reference to the calling worksheet and all of the other arguments that were passed to C<write()>. The callback will see an C<@_> argument list that looks like the following:
#NYI 
#NYI     $_[0]   A ref to the calling worksheet. *
#NYI     $_[1]   Zero based row number.
#NYI     $_[2]   Zero based column number.
#NYI     $_[3]   A number or string or token.
#NYI     $_[4]   A format ref if any.
#NYI     $_[5]   Any other arguments.
#NYI     ...
#NYI 
#NYI     *  It is good style to shift this off the list so the @_ is the same
#NYI        as the argument list seen by write().
#NYI 
#NYI Your callback should C<return()> the return value of the C<write_*> method that was called or C<undef> to indicate that you rejected the match and want C<write()> to continue as normal.
#NYI 
#NYI So for example if you wished to apply the previous filter only to ID values that occur in the first column you could modify your callback function as follows:
#NYI 
#NYI 
#NYI     sub write_my_id {
#NYI         my $worksheet = shift;
#NYI         my $col       = $_[1];
#NYI 
#NYI         if ( $col == 0 ) {
#NYI             return $worksheet->write_string( @_ );
#NYI         }
#NYI         else {
#NYI             # Reject the match and return control to write()
#NYI             return undef;
#NYI         }
#NYI     }
#NYI 
#NYI Now, you will get different behaviour for the first column and other columns:
#NYI 
#NYI     $worksheet->write( 'A1', '0000000' );    # Writes 0000000
#NYI     $worksheet->write( 'B1', '0000000' );    # Writes 0
#NYI 
#NYI 
#NYI You may add more than one handler in which case they will be called in the order that they were added.
#NYI 
#NYI Note, the C<add_write_handler()> method is particularly suited for handling dates.
#NYI 
#NYI See the C<write_handler 1-4> programs in the C<examples> directory for further examples.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 insert_image( $row, $col, $filename, $x, $y, $x_scale, $y_scale )
#NYI 
#NYI This method can be used to insert a image into a worksheet. The image can be in PNG, JPEG or BMP format. The C<$x>, C<$y>, C<$x_scale> and C<$y_scale> parameters are optional.
#NYI 
#NYI     $worksheet1->insert_image( 'A1', 'perl.bmp' );
#NYI     $worksheet2->insert_image( 'A1', '../images/perl.bmp' );
#NYI     $worksheet3->insert_image( 'A1', '.c:\images\perl.bmp' );
#NYI 
#NYI The parameters C<$x> and C<$y> can be used to specify an offset from the top left hand corner of the cell specified by C<$row> and C<$col>. The offset values are in pixels.
#NYI 
#NYI     $worksheet1->insert_image('A1', 'perl.bmp', 32, 10);
#NYI 
#NYI The offsets can be greater than the width or height of the underlying cell. This can be occasionally useful if you wish to align two or more images relative to the same cell.
#NYI 
#NYI The parameters C<$x_scale> and C<$y_scale> can be used to scale the inserted image horizontally and vertically:
#NYI 
#NYI     # Scale the inserted image: width x 2.0, height x 0.8
#NYI     $worksheet->insert_image( 'A1', 'perl.bmp', 0, 0, 2, 0.8 );
#NYI 
#NYI Note: you must call C<set_row()> or C<set_column()> before C<insert_image()> if you wish to change the default dimensions of any of the rows or columns that the image occupies. The height of a row can also change if you use a font that is larger than the default. This in turn will affect the scaling of your image. To avoid this you should explicitly set the height of the row using C<set_row()> if it contains a font size that will change the row height.
#NYI 
#NYI BMP images must be 24 bit, true colour, bitmaps. In general it is best to avoid BMP images since they aren't compressed.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 insert_chart( $row, $col, $chart, $x, $y, $x_scale, $y_scale )
#NYI 
#NYI This method can be used to insert a Chart object into a worksheet. The Chart must be created by the C<add_chart()> Workbook method and it must have the C<embedded> option set.
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'line', embedded => 1 );
#NYI 
#NYI     # Configure the chart.
#NYI     ...
#NYI 
#NYI     # Insert the chart into the a worksheet.
#NYI     $worksheet->insert_chart( 'E2', $chart );
#NYI 
#NYI See C<add_chart()> for details on how to create the Chart object and L<Excel::Writer::XLSX::Chart> for details on how to configure it. See also the C<chart_*.pl> programs in the examples directory of the distro.
#NYI 
#NYI The C<$x>, C<$y>, C<$x_scale> and C<$y_scale> parameters are optional.
#NYI 
#NYI The parameters C<$x> and C<$y> can be used to specify an offset from the top left hand corner of the cell specified by C<$row> and C<$col>. The offset values are in pixels.
#NYI 
#NYI     $worksheet1->insert_chart( 'E2', $chart, 3, 3 );
#NYI 
#NYI The parameters C<$x_scale> and C<$y_scale> can be used to scale the inserted chart horizontally and vertically:
#NYI 
#NYI     # Scale the width by 120% and the height by 150%
#NYI     $worksheet->insert_chart( 'E2', $chart, 0, 0, 1.2, 1.5 );
#NYI 
#NYI =head2 insert_shape( $row, $col, $shape, $x, $y, $x_scale, $y_scale )
#NYI 
#NYI This method can be used to insert a Shape object into a worksheet. The Shape must be created by the C<add_shape()> Workbook method.
#NYI 
#NYI     my $shape = $workbook->add_shape( name => 'My Shape', type => 'plus' );
#NYI 
#NYI     # Configure the shape.
#NYI     $shape->set_text('foo');
#NYI     ...
#NYI 
#NYI     # Insert the shape into the a worksheet.
#NYI     $worksheet->insert_shape( 'E2', $shape );
#NYI 
#NYI See C<add_shape()> for details on how to create the Shape object and L<Excel::Writer::XLSX::Shape> for details on how to configure it.
#NYI 
#NYI The C<$x>, C<$y>, C<$x_scale> and C<$y_scale> parameters are optional.
#NYI 
#NYI The parameters C<$x> and C<$y> can be used to specify an offset from the top left hand corner of the cell specified by C<$row> and C<$col>. The offset values are in pixels.
#NYI 
#NYI     $worksheet1->insert_shape( 'E2', $chart, 3, 3 );
#NYI 
#NYI The parameters C<$x_scale> and C<$y_scale> can be used to scale the inserted shape horizontally and vertically:
#NYI 
#NYI     # Scale the width by 120% and the height by 150%
#NYI     $worksheet->insert_shape( 'E2', $shape, 0, 0, 1.2, 1.5 );
#NYI 
#NYI See also the C<shape*.pl> programs in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 insert_button( $row, $col, { %properties })
#NYI 
#NYI The C<insert_button()> method can be used to insert an Excel form button into a worksheet.
#NYI 
#NYI This method is generally only useful when used in conjunction with the Workbook C<add_vba_project()> method to tie the button to a macro from an embedded VBA project:
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'file.xlsm' );
#NYI     ...
#NYI     $workbook->add_vba_project( './vbaProject.bin' );
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro' } );
#NYI 
#NYI The properties of the button that can be set are:
#NYI 
#NYI     macro
#NYI     caption
#NYI     width
#NYI     height
#NYI     x_scale
#NYI     y_scale
#NYI     x_offset
#NYI     y_offset
#NYI 
#NYI 
#NYI =over
#NYI 
#NYI =item Option: macro
#NYI 
#NYI This option is used to set the macro that the button will invoke when the user clicks on it. The macro should be included using the Workbook C<add_vba_project()> method shown above.
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro' } );
#NYI 
#NYI The default macro is C<ButtonX_Click> where X is the button number.
#NYI 
#NYI =item Option: caption
#NYI 
#NYI This option is used to set the caption on the button. The default is C<Button X> where X is the button number.
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro', caption => 'Hello' } );
#NYI 
#NYI =item Option: width
#NYI 
#NYI This option is used to set the width of the button in pixels.
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro', width => 128 } );
#NYI 
#NYI The default button width is 64 pixels which is the width of a default cell.
#NYI 
#NYI =item Option: height
#NYI 
#NYI This option is used to set the height of the button in pixels.
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro', height => 40 } );
#NYI 
#NYI The default button height is 20 pixels which is the height of a default cell.
#NYI 
#NYI =item Option: x_scale
#NYI 
#NYI This option is used to set the width of the button as a factor of the default width.
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro', x_scale => 2.0 );
#NYI 
#NYI =item Option: y_scale
#NYI 
#NYI This option is used to set the height of the button as a factor of the default height.
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro', y_scale => 2.0 );
#NYI 
#NYI 
#NYI =item Option: x_offset
#NYI 
#NYI This option is used to change the x offset, in pixels, of a button within a cell:
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro', x_offset => 2 );
#NYI 
#NYI =item Option: y_offset
#NYI 
#NYI This option is used to change the y offset, in pixels, of a comment within a cell.
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI Note: Button is the only Excel form element that is available in Excel::Writer::XLSX. Form elements represent a lot of work to implement and the underlying VML syntax isn't very much fun.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 data_validation()
#NYI 
#NYI The C<data_validation()> method is used to construct an Excel data validation or to limit the user input to a dropdown list of values.
#NYI 
#NYI 
#NYI     $worksheet->data_validation('B3',
#NYI         {
#NYI             validate => 'integer',
#NYI             criteria => '>',
#NYI             value    => 100,
#NYI         });
#NYI 
#NYI     $worksheet->data_validation('B5:B9',
#NYI         {
#NYI             validate => 'list',
#NYI             value    => ['open', 'high', 'close'],
#NYI         });
#NYI 
#NYI This method contains a lot of parameters and is described in detail in a separate section L</DATA VALIDATION IN EXCEL>.
#NYI 
#NYI 
#NYI See also the C<data_validate.pl> program in the examples directory of the distro
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 conditional_formatting()
#NYI 
#NYI The C<conditional_formatting()> method is used to add formatting to a cell or range of cells based on user defined criteria.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:J10',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => '>=',
#NYI             value    => 50,
#NYI             format   => $format1,
#NYI         }
#NYI     );
#NYI 
#NYI This method contains a lot of parameters and is described in detail in a separate section L<CONDITIONAL FORMATTING IN EXCEL>.
#NYI 
#NYI See also the C<conditional_format.pl> program in the examples directory of the distro
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 add_sparkline()
#NYI 
#NYI The C<add_sparkline()> worksheet method is used to add sparklines to a cell or a range of cells.
#NYI 
#NYI     $worksheet->add_sparkline(
#NYI         {
#NYI             location => 'F2',
#NYI             range    => 'Sheet1!A2:E2',
#NYI             type     => 'column',
#NYI             style    => 12,
#NYI         }
#NYI     );
#NYI 
#NYI This method contains a lot of parameters and is described in detail in a separate section L</SPARKLINES IN EXCEL>.
#NYI 
#NYI See also the C<sparklines1.pl> and C<sparklines2.pl> example programs in the C<examples> directory of the distro.
#NYI 
#NYI B<Note:> Sparklines are a feature of Excel 2010+ only. You can write them to an XLSX file that can be read by Excel 2007 but they won't be displayed.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 add_table()
#NYI 
#NYI The C<add_table()> method is used to group a range of cells into an Excel Table.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { ... } );
#NYI 
#NYI This method contains a lot of parameters and is described in detail in a separate section L<TABLES IN EXCEL>.
#NYI 
#NYI See also the C<tables.pl> program in the examples directory of the distro
#NYI 
#NYI 
#NYI 
#NYI =head2 get_name()
#NYI 
#NYI The C<get_name()> method is used to retrieve the name of a worksheet. For example:
#NYI 
#NYI     for my $sheet ( $workbook->sheets() ) {
#NYI         print $sheet->get_name();
#NYI     }
#NYI 
#NYI For reasons related to the design of Excel::Writer::XLSX and to the internals of Excel there is no C<set_name()> method. The only way to set the worksheet name is via the C<add_worksheet()> method.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 activate()
#NYI 
#NYI The C<activate()> method is used to specify which worksheet is initially visible in a multi-sheet workbook:
#NYI 
#NYI     $worksheet1 = $workbook->add_worksheet( 'To' );
#NYI     $worksheet2 = $workbook->add_worksheet( 'the' );
#NYI     $worksheet3 = $workbook->add_worksheet( 'wind' );
#NYI 
#NYI     $worksheet3->activate();
#NYI 
#NYI This is similar to the Excel VBA activate method. More than one worksheet can be selected via the C<select()> method, see below, however only one worksheet can be active.
#NYI 
#NYI The default active worksheet is the first worksheet.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 select()
#NYI 
#NYI The C<select()> method is used to indicate that a worksheet is selected in a multi-sheet workbook:
#NYI 
#NYI     $worksheet1->activate();
#NYI     $worksheet2->select();
#NYI     $worksheet3->select();
#NYI 
#NYI A selected worksheet has its tab highlighted. Selecting worksheets is a way of grouping them together so that, for example, several worksheets could be printed in one go. A worksheet that has been activated via the C<activate()> method will also appear as selected.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 hide()
#NYI 
#NYI The C<hide()> method is used to hide a worksheet:
#NYI 
#NYI     $worksheet2->hide();
#NYI 
#NYI You may wish to hide a worksheet in order to avoid confusing a user with intermediate data or calculations.
#NYI 
#NYI A hidden worksheet can not be activated or selected so this method is mutually exclusive with the C<activate()> and C<select()> methods. In addition, since the first worksheet will default to being the active worksheet, you cannot hide the first worksheet without activating another sheet:
#NYI 
#NYI     $worksheet2->activate();
#NYI     $worksheet1->hide();
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_first_sheet()
#NYI 
#NYI The C<activate()> method determines which worksheet is initially selected. However, if there are a large number of worksheets the selected worksheet may not appear on the screen. To avoid this you can select which is the leftmost visible worksheet using C<set_first_sheet()>:
#NYI 
#NYI     for ( 1 .. 20 ) {
#NYI         $workbook->add_worksheet;
#NYI     }
#NYI 
#NYI     $worksheet21 = $workbook->add_worksheet();
#NYI     $worksheet22 = $workbook->add_worksheet();
#NYI 
#NYI     $worksheet21->set_first_sheet();
#NYI     $worksheet22->activate();
#NYI 
#NYI This method is not required very often. The default value is the first worksheet.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 protect( $password, \%options )
#NYI 
#NYI The C<protect()> method is used to protect a worksheet from modification:
#NYI 
#NYI     $worksheet->protect();
#NYI 
#NYI The C<protect()> method also has the effect of enabling a cell's C<locked> and C<hidden> properties if they have been set. A I<locked> cell cannot be edited and this property is on by default for all cells. A I<hidden> cell will display the results of a formula but not the formula itself.
#NYI 
#NYI See the C<protection.pl> program in the examples directory of the distro for an illustrative example and the C<set_locked> and C<set_hidden> format methods in L</CELL FORMATTING>.
#NYI 
#NYI You can optionally add a password to the worksheet protection:
#NYI 
#NYI     $worksheet->protect( 'drowssap' );
#NYI 
#NYI Passing the empty string C<''> is the same as turning on protection without a password.
#NYI 
#NYI Note, the worksheet level password in Excel provides very weak protection. It does not encrypt your data and is very easy to deactivate. Full workbook encryption is not supported by C<Excel::Writer::XLSX> since it requires a completely different file format and would take several man months to implement.
#NYI 
#NYI You can specify which worksheet elements you wish to protect by passing a hash_ref with any or all of the following keys:
#NYI 
#NYI     # Default shown.
#NYI     %options = (
#NYI         objects               => 0,
#NYI         scenarios             => 0,
#NYI         format_cells          => 0,
#NYI         format_columns        => 0,
#NYI         format_rows           => 0,
#NYI         insert_columns        => 0,
#NYI         insert_rows           => 0,
#NYI         insert_hyperlinks     => 0,
#NYI         delete_columns        => 0,
#NYI         delete_rows           => 0,
#NYI         select_locked_cells   => 1,
#NYI         sort                  => 0,
#NYI         autofilter            => 0,
#NYI         pivot_tables          => 0,
#NYI         select_unlocked_cells => 1,
#NYI     );
#NYI 
#NYI The default boolean values are shown above. Individual elements can be protected as follows:
#NYI 
#NYI     $worksheet->protect( 'drowssap', { insert_rows => 1 } );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_selection( $first_row, $first_col, $last_row, $last_col )
#NYI 
#NYI This method can be used to specify which cell or cells are selected in a worksheet. The most common requirement is to select a single cell, in which case C<$last_row> and C<$last_col> can be omitted. The active cell within a selected range is determined by the order in which C<$first> and C<$last> are specified. It is also possible to specify a cell or a range using A1 notation. See the note about L</Cell notation>.
#NYI 
#NYI Examples:
#NYI 
#NYI     $worksheet1->set_selection( 3, 3 );          # 1. Cell D4.
#NYI     $worksheet2->set_selection( 3, 3, 6, 6 );    # 2. Cells D4 to G7.
#NYI     $worksheet3->set_selection( 6, 6, 3, 3 );    # 3. Cells G7 to D4.
#NYI     $worksheet4->set_selection( 'D4' );          # Same as 1.
#NYI     $worksheet5->set_selection( 'D4:G7' );       # Same as 2.
#NYI     $worksheet6->set_selection( 'G7:D4' );       # Same as 3.
#NYI 
#NYI The default cell selections is (0, 0), 'A1'.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_row( $row, $height, $format, $hidden, $level, $collapsed )
#NYI 
#NYI This method can be used to change the default properties of a row. All parameters apart from C<$row> are optional.
#NYI 
#NYI The most common use for this method is to change the height of a row:
#NYI 
#NYI     $worksheet->set_row( 0, 20 );    # Row 1 height set to 20
#NYI 
#NYI If you wish to set the format without changing the height you can pass C<undef> as the height parameter:
#NYI 
#NYI     $worksheet->set_row( 0, undef, $format );
#NYI 
#NYI The C<$format> parameter will be applied to any cells in the row that don't have a format. For example
#NYI 
#NYI     $worksheet->set_row( 0, undef, $format1 );    # Set the format for row 1
#NYI     $worksheet->write( 'A1', 'Hello' );           # Defaults to $format1
#NYI     $worksheet->write( 'B1', 'Hello', $format2 ); # Keeps $format2
#NYI 
#NYI If you wish to define a row format in this way you should call the method before any calls to C<write()>. Calling it afterwards will overwrite any format that was previously specified.
#NYI 
#NYI The C<$hidden> parameter should be set to 1 if you wish to hide a row. This can be used, for example, to hide intermediary steps in a complicated calculation:
#NYI 
#NYI     $worksheet->set_row( 0, 20,    $format, 1 );
#NYI     $worksheet->set_row( 1, undef, undef,   1 );
#NYI 
#NYI The C<$level> parameter is used to set the outline level of the row. Outlines are described in L</OUTLINES AND GROUPING IN EXCEL>. Adjacent rows with the same outline level are grouped together into a single outline.
#NYI 
#NYI The following example sets an outline level of 1 for rows 1 and 2 (zero-indexed):
#NYI 
#NYI     $worksheet->set_row( 1, undef, undef, 0, 1 );
#NYI     $worksheet->set_row( 2, undef, undef, 0, 1 );
#NYI 
#NYI The C<$hidden> parameter can also be used to hide collapsed outlined rows when used in conjunction with the C<$level> parameter.
#NYI 
#NYI     $worksheet->set_row( 1, undef, undef, 1, 1 );
#NYI     $worksheet->set_row( 2, undef, undef, 1, 1 );
#NYI 
#NYI For collapsed outlines you should also indicate which row has the collapsed C<+> symbol using the optional C<$collapsed> parameter.
#NYI 
#NYI     $worksheet->set_row( 3, undef, undef, 0, 0, 1 );
#NYI 
#NYI For a more complete example see the C<outline.pl> and C<outline_collapsed.pl> programs in the examples directory of the distro.
#NYI 
#NYI Excel allows up to 7 outline levels. Therefore the C<$level> parameter should be in the range C<0 E<lt>= $level E<lt>= 7>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )
#NYI 
#NYI This method can be used to change the default properties of a single column or a range of columns. All parameters apart from C<$first_col> and C<$last_col> are optional.
#NYI 
#NYI If C<set_column()> is applied to a single column the value of C<$first_col> and C<$last_col> should be the same. In the case where C<$last_col> is zero it is set to the same value as C<$first_col>.
#NYI 
#NYI It is also possible, and generally clearer, to specify a column range using the form of A1 notation used for columns. See the note about L</Cell notation>.
#NYI 
#NYI Examples:
#NYI 
#NYI     $worksheet->set_column( 0, 0, 20 );    # Column  A   width set to 20
#NYI     $worksheet->set_column( 1, 3, 30 );    # Columns B-D width set to 30
#NYI     $worksheet->set_column( 'E:E', 20 );   # Column  E   width set to 20
#NYI     $worksheet->set_column( 'F:H', 30 );   # Columns F-H width set to 30
#NYI 
#NYI The width corresponds to the column width value that is specified in Excel. It is approximately equal to the length of a string in the default font of Calibri 11. Unfortunately, there is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel.
#NYI 
#NYI As usual the C<$format> parameter is optional, for additional information, see L</CELL FORMATTING>. If you wish to set the format without changing the width you can pass C<undef> as the width parameter:
#NYI 
#NYI     $worksheet->set_column( 0, 0, undef, $format );
#NYI 
#NYI The C<$format> parameter will be applied to any cells in the column that don't have a format. For example
#NYI 
#NYI     $worksheet->set_column( 'A:A', undef, $format1 );    # Set format for col 1
#NYI     $worksheet->write( 'A1', 'Hello' );                  # Defaults to $format1
#NYI     $worksheet->write( 'A2', 'Hello', $format2 );        # Keeps $format2
#NYI 
#NYI If you wish to define a column format in this way you should call the method before any calls to C<write()>. If you call it afterwards it won't have any effect.
#NYI 
#NYI A default row format takes precedence over a default column format
#NYI 
#NYI     $worksheet->set_row( 0, undef, $format1 );           # Set format for row 1
#NYI     $worksheet->set_column( 'A:A', undef, $format2 );    # Set format for col 1
#NYI     $worksheet->write( 'A1', 'Hello' );                  # Defaults to $format1
#NYI     $worksheet->write( 'A2', 'Hello' );                  # Defaults to $format2
#NYI 
#NYI The C<$hidden> parameter should be set to 1 if you wish to hide a column. This can be used, for example, to hide intermediary steps in a complicated calculation:
#NYI 
#NYI     $worksheet->set_column( 'D:D', 20,    $format, 1 );
#NYI     $worksheet->set_column( 'E:E', undef, undef,   1 );
#NYI 
#NYI The C<$level> parameter is used to set the outline level of the column. Outlines are described in L</OUTLINES AND GROUPING IN EXCEL>. Adjacent columns with the same outline level are grouped together into a single outline.
#NYI 
#NYI The following example sets an outline level of 1 for columns B to G:
#NYI 
#NYI     $worksheet->set_column( 'B:G', undef, undef, 0, 1 );
#NYI 
#NYI The C<$hidden> parameter can also be used to hide collapsed outlined columns when used in conjunction with the C<$level> parameter.
#NYI 
#NYI     $worksheet->set_column( 'B:G', undef, undef, 1, 1 );
#NYI 
#NYI For collapsed outlines you should also indicate which row has the collapsed C<+> symbol using the optional C<$collapsed> parameter.
#NYI 
#NYI     $worksheet->set_column( 'H:H', undef, undef, 0, 0, 1 );
#NYI 
#NYI For a more complete example see the C<outline.pl> and C<outline_collapsed.pl> programs in the examples directory of the distro.
#NYI 
#NYI Excel allows up to 7 outline levels. Therefore the C<$level> parameter should be in the range C<0 E<lt>= $level E<lt>= 7>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_default_row( $height, $hide_unused_rows )
#NYI 
#NYI The C<set_default_row()> method is used to set the limited number of default row properties allowed by Excel. These are the default height and the option to hide unused rows.
#NYI 
#NYI     $worksheet->set_default_row( 24 );  # Set the default row height to 24.
#NYI 
#NYI The option to hide unused rows is used by Excel as an optimisation so that the user can hide a large number of rows without generating a very large file with an entry for each hidden row.
#NYI 
#NYI     $worksheet->set_default_row( undef, 1 );
#NYI 
#NYI See the C<hide_row_col.pl> example program.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 outline_settings( $visible, $symbols_below, $symbols_right, $auto_style )
#NYI 
#NYI The C<outline_settings()> method is used to control the appearance of outlines in Excel. Outlines are described in L</OUTLINES AND GROUPING IN EXCEL>.
#NYI 
#NYI The C<$visible> parameter is used to control whether or not outlines are visible. Setting this parameter to 0 will cause all outlines on the worksheet to be hidden. They can be unhidden in Excel by means of the "Show Outline Symbols" command button. The default setting is 1 for visible outlines.
#NYI 
#NYI     $worksheet->outline_settings( 0 );
#NYI 
#NYI The C<$symbols_below> parameter is used to control whether the row outline symbol will appear above or below the outline level bar. The default setting is 1 for symbols to appear below the outline level bar.
#NYI 
#NYI The C<$symbols_right> parameter is used to control whether the column outline symbol will appear to the left or the right of the outline level bar. The default setting is 1 for symbols to appear to the right of the outline level bar.
#NYI 
#NYI The C<$auto_style> parameter is used to control whether the automatic outline generator in Excel uses automatic styles when creating an outline. This has no effect on a file generated by C<Excel::Writer::XLSX> but it does have an effect on how the worksheet behaves after it is created. The default setting is 0 for "Automatic Styles" to be turned off.
#NYI 
#NYI The default settings for all of these parameters correspond to Excel's default parameters.
#NYI 
#NYI 
#NYI The worksheet parameters controlled by C<outline_settings()> are rarely used.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 freeze_panes( $row, $col, $top_row, $left_col )
#NYI 
#NYI This method can be used to divide a worksheet into horizontal or vertical regions known as panes and to also "freeze" these panes so that the splitter bars are not visible. This is the same as the C<Window-E<gt>Freeze Panes> menu command in Excel
#NYI 
#NYI The parameters C<$row> and C<$col> are used to specify the location of the split. It should be noted that the split is specified at the top or left of a cell and that the method uses zero based indexing. Therefore to freeze the first row of a worksheet it is necessary to specify the split at row 2 (which is 1 as the zero-based index). This might lead you to think that you are using a 1 based index but this is not the case.
#NYI 
#NYI You can set one of the C<$row> and C<$col> parameters as zero if you do not want either a vertical or horizontal split.
#NYI 
#NYI Examples:
#NYI 
#NYI     $worksheet->freeze_panes( 1, 0 );    # Freeze the first row
#NYI     $worksheet->freeze_panes( 'A2' );    # Same using A1 notation
#NYI     $worksheet->freeze_panes( 0, 1 );    # Freeze the first column
#NYI     $worksheet->freeze_panes( 'B1' );    # Same using A1 notation
#NYI     $worksheet->freeze_panes( 1, 2 );    # Freeze first row and first 2 columns
#NYI     $worksheet->freeze_panes( 'C2' );    # Same using A1 notation
#NYI 
#NYI The parameters C<$top_row> and C<$left_col> are optional. They are used to specify the top-most or left-most visible row or column in the scrolling region of the panes. For example to freeze the first row and to have the scrolling region begin at row twenty:
#NYI 
#NYI     $worksheet->freeze_panes( 1, 0, 20, 0 );
#NYI 
#NYI You cannot use A1 notation for the C<$top_row> and C<$left_col> parameters.
#NYI 
#NYI 
#NYI See also the C<panes.pl> program in the C<examples> directory of the distribution.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 split_panes( $y, $x, $top_row, $left_col )
#NYI 
#NYI 
#NYI This method can be used to divide a worksheet into horizontal or vertical regions known as panes. This method is different from the C<freeze_panes()> method in that the splits between the panes will be visible to the user and each pane will have its own scroll bars.
#NYI 
#NYI The parameters C<$y> and C<$x> are used to specify the vertical and horizontal position of the split. The units for C<$y> and C<$x> are the same as those used by Excel to specify row height and column width. However, the vertical and horizontal units are different from each other. Therefore you must specify the C<$y> and C<$x> parameters in terms of the row heights and column widths that you have set or the default values which are C<15> for a row and C<8.43> for a column.
#NYI 
#NYI You can set one of the C<$y> and C<$x> parameters as zero if you do not want either a vertical or horizontal split. The parameters C<$top_row> and C<$left_col> are optional. They are used to specify the top-most or left-most visible row or column in the bottom-right pane.
#NYI 
#NYI Example:
#NYI 
#NYI     $worksheet->split_panes( 15, 0,   );    # First row
#NYI     $worksheet->split_panes( 0,  8.43 );    # First column
#NYI     $worksheet->split_panes( 15, 8.43 );    # First row and column
#NYI 
#NYI You cannot use A1 notation with this method.
#NYI 
#NYI See also the C<freeze_panes()> method and the C<panes.pl> program in the C<examples> directory of the distribution.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 merge_range( $first_row, $first_col, $last_row, $last_col, $token, $format )
#NYI 
#NYI The C<merge_range()> method allows you to merge cells that contain other types of alignment in addition to the merging:
#NYI 
#NYI     my $format = $workbook->add_format(
#NYI         border => 6,
#NYI         valign => 'vcenter',
#NYI         align  => 'center',
#NYI     );
#NYI 
#NYI     $worksheet->merge_range( 'B3:D4', 'Vertical and horizontal', $format );
#NYI 
#NYI C<merge_range()> writes its C<$token> argument using the worksheet C<write()> method. Therefore it will handle numbers, strings, formulas or urls as required. If you need to specify the required C<write_*()> method use the C<merge_range_type()> method, see below.
#NYI 
#NYI The full possibilities of this method are shown in the C<merge3.pl> to C<merge6.pl> programs in the C<examples> directory of the distribution.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 merge_range_type( $type, $first_row, $first_col, $last_row, $last_col, ... )
#NYI 
#NYI The C<merge_range()> method, see above, uses C<write()> to insert the required data into to a merged range. However, there may be times where this isn't what you require so as an alternative the C<merge_range_type ()> method allows you to specify the type of data you wish to write. For example:
#NYI 
#NYI     $worksheet->merge_range_type( 'number',  'B2:C2', 123,    $format1 );
#NYI     $worksheet->merge_range_type( 'string',  'B4:C4', 'foo',  $format2 );
#NYI     $worksheet->merge_range_type( 'formula', 'B6:C6', '=1+2', $format3 );
#NYI 
#NYI The C<$type> must be one of the following, which corresponds to a C<write_*()> method:
#NYI 
#NYI     'number'
#NYI     'string'
#NYI     'formula'
#NYI     'array_formula'
#NYI     'blank'
#NYI     'rich_string'
#NYI     'date_time'
#NYI     'url'
#NYI 
#NYI Any arguments after the range should be whatever the appropriate method accepts:
#NYI 
#NYI     $worksheet->merge_range_type( 'rich_string', 'B8:C8',
#NYI                                   'This is ', $bold, 'bold', $format4 );
#NYI 
#NYI Note, you must always pass a C<$format> object as an argument, even if it is a default format.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_zoom( $scale )
#NYI 
#NYI Set the worksheet zoom factor in the range C<10 E<lt>= $scale E<lt>= 400>:
#NYI 
#NYI     $worksheet1->set_zoom( 50 );
#NYI     $worksheet2->set_zoom( 75 );
#NYI     $worksheet3->set_zoom( 300 );
#NYI     $worksheet4->set_zoom( 400 );
#NYI 
#NYI The default zoom factor is 100. You cannot zoom to "Selection" because it is calculated by Excel at run-time.
#NYI 
#NYI Note, C<set_zoom()> does not affect the scale of the printed page. For that you should use C<set_print_scale()>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 right_to_left()
#NYI 
#NYI The C<right_to_left()> method is used to change the default direction of the worksheet from left-to-right, with the A1 cell in the top left, to right-to-left, with the A1 cell in the top right.
#NYI 
#NYI     $worksheet->right_to_left();
#NYI 
#NYI This is useful when creating Arabic, Hebrew or other near or far eastern worksheets that use right-to-left as the default direction.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 hide_zero()
#NYI 
#NYI The C<hide_zero()> method is used to hide any zero values that appear in cells.
#NYI 
#NYI     $worksheet->hide_zero();
#NYI 
#NYI In Excel this option is found under Tools->Options->View.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_tab_color()
#NYI 
#NYI The C<set_tab_color()> method is used to change the colour of the worksheet tab. You can use one of the standard colour names provided by the Format object or a Html style C<#RRGGBB> colour. See L</WORKING WITH COLOURS>.
#NYI 
#NYI     $worksheet1->set_tab_color( 'red' );
#NYI     $worksheet2->set_tab_color( '#FF6600' );
#NYI 
#NYI See the C<tab_colors.pl> program in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 autofilter( $first_row, $first_col, $last_row, $last_col )
#NYI 
#NYI This method allows an autofilter to be added to a worksheet. An autofilter is a way of adding drop down lists to the headers of a 2D range of worksheet data. This allows users to filter the data based on simple criteria so that some data is shown and some is hidden.
#NYI 
#NYI To add an autofilter to a worksheet:
#NYI 
#NYI     $worksheet->autofilter( 0, 0, 10, 3 );
#NYI     $worksheet->autofilter( 'A1:D11' );    # Same as above in A1 notation.
#NYI 
#NYI Filter conditions can be applied using the C<filter_column()> or C<filter_column_list()> method.
#NYI 
#NYI See the C<autofilter.pl> program in the examples directory of the distro for a more detailed example.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 filter_column( $column, $expression )
#NYI 
#NYI The C<filter_column> method can be used to filter columns in a autofilter range based on simple conditions.
#NYI 
#NYI B<NOTE:> It isn't sufficient to just specify the filter condition. You must also hide any rows that don't match the filter condition. Rows are hidden using the C<set_row()> C<visible> parameter. C<Excel::Writer::XLSX> cannot do this automatically since it isn't part of the file format. See the C<autofilter.pl> program in the examples directory of the distro for an example.
#NYI 
#NYI The conditions for the filter are specified using simple expressions:
#NYI 
#NYI     $worksheet->filter_column( 'A', 'x > 2000' );
#NYI     $worksheet->filter_column( 'B', 'x > 2000 and x < 5000' );
#NYI 
#NYI The C<$column> parameter can either be a zero indexed column number or a string column name.
#NYI 
#NYI The following operators are available:
#NYI 
#NYI     Operator        Synonyms
#NYI        ==           =   eq  =~
#NYI        !=           <>  ne  !=
#NYI        >
#NYI        <
#NYI        >=
#NYI        <=
#NYI 
#NYI        and          &&
#NYI        or           ||
#NYI 
#NYI The operator synonyms are just syntactic sugar to make you more comfortable using the expressions. It is important to remember that the expressions will be interpreted by Excel and not by perl.
#NYI 
#NYI An expression can comprise a single statement or two statements separated by the C<and> and C<or> operators. For example:
#NYI 
#NYI     'x <  2000'
#NYI     'x >  2000'
#NYI     'x == 2000'
#NYI     'x >  2000 and x <  5000'
#NYI     'x == 2000 or  x == 5000'
#NYI 
#NYI Filtering of blank or non-blank data can be achieved by using a value of C<Blanks> or C<NonBlanks> in the expression:
#NYI 
#NYI     'x == Blanks'
#NYI     'x == NonBlanks'
#NYI 
#NYI Excel also allows some simple string matching operations:
#NYI 
#NYI     'x =~ b*'   # begins with b
#NYI     'x !~ b*'   # doesn't begin with b
#NYI     'x =~ *b'   # ends with b
#NYI     'x !~ *b'   # doesn't end with b
#NYI     'x =~ *b*'  # contains b
#NYI     'x !~ *b*'  # doesn't contains b
#NYI 
#NYI You can also use C<*> to match any character or number and C<?> to match any single character or number. No other regular expression quantifier is supported by Excel's filters. Excel's regular expression characters can be escaped using C<~>.
#NYI 
#NYI The placeholder variable C<x> in the above examples can be replaced by any simple string. The actual placeholder name is ignored internally so the following are all equivalent:
#NYI 
#NYI     'x     < 2000'
#NYI     'col   < 2000'
#NYI     'Price < 2000'
#NYI 
#NYI Also, note that a filter condition can only be applied to a column in a range specified by the C<autofilter()> Worksheet method.
#NYI 
#NYI See the C<autofilter.pl> program in the examples directory of the distro for a more detailed example.
#NYI 
#NYI B<Note> L<Spreadsheet::WriteExcel> supports Top 10 style filters. These aren't currently supported by Excel::Writer::XLSX but may be added later.
#NYI 
#NYI 
#NYI =head2 filter_column_list( $column, @matches )
#NYI 
#NYI Prior to Excel 2007 it was only possible to have either 1 or 2 filter conditions such as the ones shown above in the C<filter_column> method.
#NYI 
#NYI Excel 2007 introduced a new list style filter where it is possible to specify 1 or more 'or' style criteria. For example if your column contained data for the first six months the initial data would be displayed as all selected as shown on the left. Then if you selected 'March', 'April' and 'May' they would be displayed as shown on the right.
#NYI 
#NYI     No criteria selected      Some criteria selected.
#NYI 
#NYI     [/] (Select all)          [X] (Select all)
#NYI     [/] January               [ ] January
#NYI     [/] February              [ ] February
#NYI     [/] March                 [/] March
#NYI     [/] April                 [/] April
#NYI     [/] May                   [/] May
#NYI     [/] June                  [ ] June
#NYI 
#NYI The C<filter_column_list()> method can be used to represent these types of filters:
#NYI 
#NYI     $worksheet->filter_column_list( 'A', 'March', 'April', 'May' );
#NYI 
#NYI The C<$column> parameter can either be a zero indexed column number or a string column name.
#NYI 
#NYI One or more criteria can be selected:
#NYI 
#NYI     $worksheet->filter_column_list( 0, 'March' );
#NYI     $worksheet->filter_column_list( 1, 100, 110, 120, 130 );
#NYI 
#NYI B<NOTE:> It isn't sufficient to just specify the filter condition. You must also hide any rows that don't match the filter condition. Rows are hidden using the C<set_row()> C<visible> parameter. C<Excel::Writer::XLSX> cannot do this automatically since it isn't part of the file format. See the C<autofilter.pl> program in the examples directory of the distro for an example.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 convert_date_time( $date_string )
#NYI 
#NYI The C<convert_date_time()> method is used internally by the C<write_date_time()> method to convert date strings to a number that represents an Excel date and time.
#NYI 
#NYI It is exposed as a public method for utility purposes.
#NYI 
#NYI The C<$date_string> format is detailed in the C<write_date_time()> method.
#NYI 
#NYI =head2 Worksheet set_vba_name()
#NYI 
#NYI The Worksheet C<set_vba_name()> method can be used to set the VBA codename for the
#NYI worksheet (there is a similar method for the workbook VBA name). This is sometimes required when a C<vbaProject> macro included via C<add_vba_project()> refers to the worksheet. The default Excel VBA name of C<Sheet1>, etc., is used if a user defined name isn't specified.
#NYI 
#NYI See also L<WORKING WITH VBA MACROS>.
#NYI 
#NYI 
#NYI 
#NYI =head1 PAGE SET-UP METHODS
#NYI 
#NYI Page set-up methods affect the way that a worksheet looks when it is printed. They control features such as page headers and footers and margins. These methods are really just standard worksheet methods. They are documented here in a separate section for the sake of clarity.
#NYI 
#NYI The following methods are available for page set-up:
#NYI 
#NYI     set_landscape()
#NYI     set_portrait()
#NYI     set_page_view()
#NYI     set_paper()
#NYI     center_horizontally()
#NYI     center_vertically()
#NYI     set_margins()
#NYI     set_header()
#NYI     set_footer()
#NYI     repeat_rows()
#NYI     repeat_columns()
#NYI     hide_gridlines()
#NYI     print_row_col_headers()
#NYI     print_area()
#NYI     print_across()
#NYI     fit_to_pages()
#NYI     set_start_page()
#NYI     set_print_scale()
#NYI     print_black_and_white()
#NYI     set_h_pagebreaks()
#NYI     set_v_pagebreaks()
#NYI 
#NYI A common requirement when working with Excel::Writer::XLSX is to apply the same page set-up features to all of the worksheets in a workbook. To do this you can use the C<sheets()> method of the C<workbook> class to access the array of worksheets in a workbook:
#NYI 
#NYI     for $worksheet ( $workbook->sheets() ) {
#NYI         $worksheet->set_landscape();
#NYI     }
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_landscape()
#NYI 
#NYI This method is used to set the orientation of a worksheet's printed page to landscape:
#NYI 
#NYI     $worksheet->set_landscape();    # Landscape mode
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_portrait()
#NYI 
#NYI This method is used to set the orientation of a worksheet's printed page to portrait. The default worksheet orientation is portrait, so you won't generally need to call this method.
#NYI 
#NYI     $worksheet->set_portrait();    # Portrait mode
#NYI 
#NYI 
#NYI 
#NYI =head2 set_page_view()
#NYI 
#NYI This method is used to display the worksheet in "Page View/Layout" mode.
#NYI 
#NYI     $worksheet->set_page_view();
#NYI 
#NYI 
#NYI 
#NYI =head2 set_paper( $index )
#NYI 
#NYI This method is used to set the paper format for the printed output of a worksheet. The following paper styles are available:
#NYI 
#NYI     Index   Paper format            Paper size
#NYI     =====   ============            ==========
#NYI       0     Printer default         -
#NYI       1     Letter                  8 1/2 x 11 in
#NYI       2     Letter Small            8 1/2 x 11 in
#NYI       3     Tabloid                 11 x 17 in
#NYI       4     Ledger                  17 x 11 in
#NYI       5     Legal                   8 1/2 x 14 in
#NYI       6     Statement               5 1/2 x 8 1/2 in
#NYI       7     Executive               7 1/4 x 10 1/2 in
#NYI       8     A3                      297 x 420 mm
#NYI       9     A4                      210 x 297 mm
#NYI      10     A4 Small                210 x 297 mm
#NYI      11     A5                      148 x 210 mm
#NYI      12     B4                      250 x 354 mm
#NYI      13     B5                      182 x 257 mm
#NYI      14     Folio                   8 1/2 x 13 in
#NYI      15     Quarto                  215 x 275 mm
#NYI      16     -                       10x14 in
#NYI      17     -                       11x17 in
#NYI      18     Note                    8 1/2 x 11 in
#NYI      19     Envelope  9             3 7/8 x 8 7/8
#NYI      20     Envelope 10             4 1/8 x 9 1/2
#NYI      21     Envelope 11             4 1/2 x 10 3/8
#NYI      22     Envelope 12             4 3/4 x 11
#NYI      23     Envelope 14             5 x 11 1/2
#NYI      24     C size sheet            -
#NYI      25     D size sheet            -
#NYI      26     E size sheet            -
#NYI      27     Envelope DL             110 x 220 mm
#NYI      28     Envelope C3             324 x 458 mm
#NYI      29     Envelope C4             229 x 324 mm
#NYI      30     Envelope C5             162 x 229 mm
#NYI      31     Envelope C6             114 x 162 mm
#NYI      32     Envelope C65            114 x 229 mm
#NYI      33     Envelope B4             250 x 353 mm
#NYI      34     Envelope B5             176 x 250 mm
#NYI      35     Envelope B6             176 x 125 mm
#NYI      36     Envelope                110 x 230 mm
#NYI      37     Monarch                 3.875 x 7.5 in
#NYI      38     Envelope                3 5/8 x 6 1/2 in
#NYI      39     Fanfold                 14 7/8 x 11 in
#NYI      40     German Std Fanfold      8 1/2 x 12 in
#NYI      41     German Legal Fanfold    8 1/2 x 13 in
#NYI 
#NYI 
#NYI Note, it is likely that not all of these paper types will be available to the end user since it will depend on the paper formats that the user's printer supports. Therefore, it is best to stick to standard paper types.
#NYI 
#NYI     $worksheet->set_paper( 1 );    # US Letter
#NYI     $worksheet->set_paper( 9 );    # A4
#NYI 
#NYI If you do not specify a paper type the worksheet will print using the printer's default paper.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 center_horizontally()
#NYI 
#NYI Center the worksheet data horizontally between the margins on the printed page:
#NYI 
#NYI     $worksheet->center_horizontally();
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 center_vertically()
#NYI 
#NYI Center the worksheet data vertically between the margins on the printed page:
#NYI 
#NYI     $worksheet->center_vertically();
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_margins( $inches )
#NYI 
#NYI There are several methods available for setting the worksheet margins on the printed page:
#NYI 
#NYI     set_margins()        # Set all margins to the same value
#NYI     set_margins_LR()     # Set left and right margins to the same value
#NYI     set_margins_TB()     # Set top and bottom margins to the same value
#NYI     set_margin_left();   # Set left margin
#NYI     set_margin_right();  # Set right margin
#NYI     set_margin_top();    # Set top margin
#NYI     set_margin_bottom(); # Set bottom margin
#NYI 
#NYI All of these methods take a distance in inches as a parameter. Note: 1 inch = 25.4mm. C<;-)> The default left and right margin is 0.7 inch. The default top and bottom margin is 0.75 inch. Note, these defaults are different from the defaults used in the binary file format by Spreadsheet::WriteExcel.
#NYI 
#NYI 
#NYI 
#NYI =head2 set_header( $string, $margin )
#NYI 
#NYI Headers and footers are generated using a C<$string> which is a combination of plain text and control characters. The C<$margin> parameter is optional.
#NYI 
#NYI The available control character are:
#NYI 
#NYI     Control             Category            Description
#NYI     =======             ========            ===========
#NYI     &L                  Justification       Left
#NYI     &C                                      Center
#NYI     &R                                      Right
#NYI 
#NYI     &P                  Information         Page number
#NYI     &N                                      Total number of pages
#NYI     &D                                      Date
#NYI     &T                                      Time
#NYI     &F                                      File name
#NYI     &A                                      Worksheet name
#NYI     &Z                                      Workbook path
#NYI 
#NYI     &fontsize           Font                Font size
#NYI     &"font,style"                           Font name and style
#NYI     &U                                      Single underline
#NYI     &E                                      Double underline
#NYI     &S                                      Strikethrough
#NYI     &X                                      Superscript
#NYI     &Y                                      Subscript
#NYI 
#NYI     &[Picture]          Images              Image placeholder
#NYI     &G                                      Same as &[Picture]
#NYI 
#NYI     &&                  Miscellaneous       Literal ampersand &
#NYI 
#NYI 
#NYI Text in headers and footers can be justified (aligned) to the left, center and right by prefixing the text with the control characters C<&L>, C<&C> and C<&R>.
#NYI 
#NYI For example (with ASCII art representation of the results):
#NYI 
#NYI     $worksheet->set_header('&LHello');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     | Hello                                                         |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI     $worksheet->set_header('&CHello');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     |                          Hello                                |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI     $worksheet->set_header('&RHello');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     |                                                         Hello |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI For simple text, if you do not specify any justification the text will be centred. However, you must prefix the text with C<&C> if you specify a font name or any other formatting:
#NYI 
#NYI     $worksheet->set_header('Hello');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     |                          Hello                                |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI You can have text in each of the justification regions:
#NYI 
#NYI     $worksheet->set_header('&LCiao&CBello&RCielo');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     | Ciao                     Bello                          Cielo |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI The information control characters act as variables that Excel will update as the workbook or worksheet changes. Times and dates are in the users default format:
#NYI 
#NYI     $worksheet->set_header('&CPage &P of &N');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     |                        Page 1 of 6                            |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI     $worksheet->set_header('&CUpdated at &T');
#NYI 
#NYI      ---------------------------------------------------------------
#NYI     |                                                               |
#NYI     |                    Updated at 12:30 PM                        |
#NYI     |                                                               |
#NYI 
#NYI 
#NYI Images can be inserted using the options shown below. Each image must have a placeholder in header string using the C<&[Picture]> or C<&G> control characters:
#NYI 
#NYI     $worksheet->set_header( '&L&G', 0.3, { image_left => 'logo.jpg' });
#NYI 
#NYI 
#NYI 
#NYI You can specify the font size of a section of the text by prefixing it with the control character C<&n> where C<n> is the font size:
#NYI 
#NYI     $worksheet1->set_header( '&C&30Hello Big' );
#NYI     $worksheet2->set_header( '&C&10Hello Small' );
#NYI 
#NYI You can specify the font of a section of the text by prefixing it with the control sequence C<&"font,style"> where C<fontname> is a font name such as "Courier New" or "Times New Roman" and C<style> is one of the standard Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
#NYI 
#NYI     $worksheet1->set_header( '&C&"Courier New,Italic"Hello' );
#NYI     $worksheet2->set_header( '&C&"Courier New,Bold Italic"Hello' );
#NYI     $worksheet3->set_header( '&C&"Times New Roman,Regular"Hello' );
#NYI 
#NYI It is possible to combine all of these features together to create sophisticated headers and footers. As an aid to setting up complicated headers and footers you can record a page set-up as a macro in Excel and look at the format strings that VBA produces. Remember however that VBA uses two double quotes C<""> to indicate a single double quote. For the last example above the equivalent VBA code looks like this:
#NYI 
#NYI     .LeftHeader   = ""
#NYI     .CenterHeader = "&""Times New Roman,Regular""Hello"
#NYI     .RightHeader  = ""
#NYI 
#NYI 
#NYI To include a single literal ampersand C<&> in a header or footer you should use a double ampersand C<&&>:
#NYI 
#NYI     $worksheet1->set_header('&CCuriouser && Curiouser - Attorneys at Law');
#NYI 
#NYI As stated above the margin parameter is optional. As with the other margins the value should be in inches. The default header and footer margin is 0.3 inch. Note, the default margin is different from the default used in the binary file format by Spreadsheet::WriteExcel. The header and footer margin size can be set as follows:
#NYI 
#NYI     $worksheet->set_header( '&CHello', 0.75 );
#NYI 
#NYI The header and footer margins are independent of the top and bottom margins.
#NYI 
#NYI The available options are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<image_left> The path to the image. Requires a C<&G> or C<&[Picture]> placeholder.
#NYI 
#NYI =item * C<image_center> Same as above.
#NYI 
#NYI =item * C<image_right> Same as above.
#NYI 
#NYI =item * C<scale_with_doc> Scale header with document. Defaults to true.
#NYI 
#NYI =item * C<align_with_margins> Align header to margins. Defaults to true.
#NYI 
#NYI =back
#NYI 
#NYI The image options must have an accompanying C<&[Picture]> or C<&G> control
#NYI character in the header string:
#NYI 
#NYI     $worksheet->set_header(
#NYI         '&L&[Picture]&C&[Picture]&R&[Picture]',
#NYI         undef, # If you don't want to change the margin.
#NYI         {
#NYI             image_left   => 'red.jpg',
#NYI             image_center => 'blue.jpg',
#NYI             image_right  => 'yellow.jpg'
#NYI         }
#NYI       );
#NYI 
#NYI 
#NYI Note, the header or footer string must be less than 255 characters. Strings longer than this will not be written and a warning will be generated.
#NYI 
#NYI The C<set_header()> method can also handle Unicode strings in C<UTF-8> format.
#NYI 
#NYI     $worksheet->set_header( "&C\x{263a}" )
#NYI 
#NYI 
#NYI See, also the C<headers.pl> program in the C<examples> directory of the distribution.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_footer( $string, $margin )
#NYI 
#NYI The syntax of the C<set_footer()> method is the same as C<set_header()>,  see above.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 repeat_rows( $first_row, $last_row )
#NYI 
#NYI Set the number of rows to repeat at the top of each printed page.
#NYI 
#NYI For large Excel documents it is often desirable to have the first row or rows of the worksheet print out at the top of each page. This can be achieved by using the C<repeat_rows()> method. The parameters C<$first_row> and C<$last_row> are zero based. The C<$last_row> parameter is optional if you only wish to specify one row:
#NYI 
#NYI     $worksheet1->repeat_rows( 0 );    # Repeat the first row
#NYI     $worksheet2->repeat_rows( 0, 1 ); # Repeat the first two rows
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 repeat_columns( $first_col, $last_col )
#NYI 
#NYI Set the columns to repeat at the left hand side of each printed page.
#NYI 
#NYI For large Excel documents it is often desirable to have the first column or columns of the worksheet print out at the left hand side of each page. This can be achieved by using the C<repeat_columns()> method. The parameters C<$first_column> and C<$last_column> are zero based. The C<$last_column> parameter is optional if you only wish to specify one column. You can also specify the columns using A1 column notation, see the note about L</Cell notation>.
#NYI 
#NYI     $worksheet1->repeat_columns( 0 );        # Repeat the first column
#NYI     $worksheet2->repeat_columns( 0, 1 );     # Repeat the first two columns
#NYI     $worksheet3->repeat_columns( 'A:A' );    # Repeat the first column
#NYI     $worksheet4->repeat_columns( 'A:B' );    # Repeat the first two columns
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 hide_gridlines( $option )
#NYI 
#NYI This method is used to hide the gridlines on the screen and printed page. Gridlines are the lines that divide the cells on a worksheet. Screen and printed gridlines are turned on by default in an Excel worksheet. If you have defined your own cell borders you may wish to hide the default gridlines.
#NYI 
#NYI     $worksheet->hide_gridlines();
#NYI 
#NYI The following values of C<$option> are valid:
#NYI 
#NYI     0 : Don't hide gridlines
#NYI     1 : Hide printed gridlines only
#NYI     2 : Hide screen and printed gridlines
#NYI 
#NYI If you don't supply an argument or use C<undef> the default option is 1, i.e. only the printed gridlines are hidden.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 print_row_col_headers()
#NYI 
#NYI Set the option to print the row and column headers on the printed page.
#NYI 
#NYI An Excel worksheet looks something like the following;
#NYI 
#NYI      ------------------------------------------
#NYI     |   |   A   |   B   |   C   |   D   |  ...
#NYI      ------------------------------------------
#NYI     | 1 |       |       |       |       |  ...
#NYI     | 2 |       |       |       |       |  ...
#NYI     | 3 |       |       |       |       |  ...
#NYI     | 4 |       |       |       |       |  ...
#NYI     |...|  ...  |  ...  |  ...  |  ...  |  ...
#NYI 
#NYI The headers are the letters and numbers at the top and the left of the worksheet. Since these headers serve mainly as a indication of position on the worksheet they generally do not appear on the printed page. If you wish to have them printed you can use the C<print_row_col_headers()> method :
#NYI 
#NYI     $worksheet->print_row_col_headers();
#NYI 
#NYI Do not confuse these headers with page headers as described in the C<set_header()> section above.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 print_area( $first_row, $first_col, $last_row, $last_col )
#NYI 
#NYI This method is used to specify the area of the worksheet that will be printed. All four parameters must be specified. You can also use A1 notation, see the note about L</Cell notation>.
#NYI 
#NYI 
#NYI     $worksheet1->print_area( 'A1:H20' );    # Cells A1 to H20
#NYI     $worksheet2->print_area( 0, 0, 19, 7 ); # The same
#NYI     $worksheet2->print_area( 'A:H' );       # Columns A to H if rows have data
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 print_across()
#NYI 
#NYI The C<print_across> method is used to change the default print direction. This is referred to by Excel as the sheet "page order".
#NYI 
#NYI     $worksheet->print_across();
#NYI 
#NYI The default page order is shown below for a worksheet that extends over 4 pages. The order is called "down then across":
#NYI 
#NYI     [1] [3]
#NYI     [2] [4]
#NYI 
#NYI However, by using the C<print_across> method the print order will be changed to "across then down":
#NYI 
#NYI     [1] [2]
#NYI     [3] [4]
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 fit_to_pages( $width, $height )
#NYI 
#NYI The C<fit_to_pages()> method is used to fit the printed area to a specific number of pages both vertically and horizontally. If the printed area exceeds the specified number of pages it will be scaled down to fit. This guarantees that the printed area will always appear on the specified number of pages even if the page size or margins change.
#NYI 
#NYI     $worksheet1->fit_to_pages( 1, 1 );    # Fit to 1x1 pages
#NYI     $worksheet2->fit_to_pages( 2, 1 );    # Fit to 2x1 pages
#NYI     $worksheet3->fit_to_pages( 1, 2 );    # Fit to 1x2 pages
#NYI 
#NYI The print area can be defined using the C<print_area()> method as described above.
#NYI 
#NYI A common requirement is to fit the printed output to I<n> pages wide but have the height be as long as necessary. To achieve this set the C<$height> to zero:
#NYI 
#NYI     $worksheet1->fit_to_pages( 1, 0 );    # 1 page wide and as long as necessary
#NYI 
#NYI Note that although it is valid to use both C<fit_to_pages()> and C<set_print_scale()> on the same worksheet only one of these options can be active at a time. The last method call made will set the active option.
#NYI 
#NYI Note that C<fit_to_pages()> will override any manual page breaks that are defined in the worksheet.
#NYI 
#NYI Note: When using C<fit_to_pages()> it may also be required to set the printer paper size using C<set_paper()> or else Excel will default to "US Letter".
#NYI 
#NYI 
#NYI =head2 set_start_page( $start_page )
#NYI 
#NYI The C<set_start_page()> method is used to set the number of the starting page when the worksheet is printed out. The default value is 1.
#NYI 
#NYI     $worksheet->set_start_page( 2 );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_print_scale( $scale )
#NYI 
#NYI Set the scale factor of the printed page. Scale factors in the range C<10 E<lt>= $scale E<lt>= 400> are valid:
#NYI 
#NYI     $worksheet1->set_print_scale( 50 );
#NYI     $worksheet2->set_print_scale( 75 );
#NYI     $worksheet3->set_print_scale( 300 );
#NYI     $worksheet4->set_print_scale( 400 );
#NYI 
#NYI The default scale factor is 100. Note, C<set_print_scale()> does not affect the scale of the visible page in Excel. For that you should use C<set_zoom()>.
#NYI 
#NYI Note also that although it is valid to use both C<fit_to_pages()> and C<set_print_scale()> on the same worksheet only one of these options can be active at a time. The last method call made will set the active option.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 print_black_and_white()
#NYI 
#NYI Set the option to print the worksheet in black and white:
#NYI 
#NYI     $worksheet->print_black_and_white();
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_h_pagebreaks( @breaks )
#NYI 
#NYI Add horizontal page breaks to a worksheet. A page break causes all the data that follows it to be printed on the next page. Horizontal page breaks act between rows. To create a page break between rows 20 and 21 you must specify the break at row 21. However in zero index notation this is actually row 20. So you can pretend for a small while that you are using 1 index notation:
#NYI 
#NYI     $worksheet1->set_h_pagebreaks( 20 );    # Break between row 20 and 21
#NYI 
#NYI The C<set_h_pagebreaks()> method will accept a list of page breaks and you can call it more than once:
#NYI 
#NYI     $worksheet2->set_h_pagebreaks( 20,  40,  60,  80,  100 );    # Add breaks
#NYI     $worksheet2->set_h_pagebreaks( 120, 140, 160, 180, 200 );    # Add some more
#NYI 
#NYI Note: If you specify the "fit to page" option via the C<fit_to_pages()> method it will override all manual page breaks.
#NYI 
#NYI There is a silent limitation of about 1000 horizontal page breaks per worksheet in line with an Excel internal limitation.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_v_pagebreaks( @breaks )
#NYI 
#NYI Add vertical page breaks to a worksheet. A page break causes all the data that follows it to be printed on the next page. Vertical page breaks act between columns. To create a page break between columns 20 and 21 you must specify the break at column 21. However in zero index notation this is actually column 20. So you can pretend for a small while that you are using 1 index notation:
#NYI 
#NYI     $worksheet1->set_v_pagebreaks(20); # Break between column 20 and 21
#NYI 
#NYI The C<set_v_pagebreaks()> method will accept a list of page breaks and you can call it more than once:
#NYI 
#NYI     $worksheet2->set_v_pagebreaks( 20,  40,  60,  80,  100 );    # Add breaks
#NYI     $worksheet2->set_v_pagebreaks( 120, 140, 160, 180, 200 );    # Add some more
#NYI 
#NYI Note: If you specify the "fit to page" option via the C<fit_to_pages()> method it will override all manual page breaks.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 CELL FORMATTING
#NYI 
#NYI This section describes the methods and properties that are available for formatting cells in Excel. The properties of a cell that can be formatted include: fonts, colours, patterns, borders, alignment and number formatting.
#NYI 
#NYI 
#NYI =head2 Creating and using a Format object
#NYI 
#NYI Cell formatting is defined through a Format object. Format objects are created by calling the workbook C<add_format()> method as follows:
#NYI 
#NYI     my $format1 = $workbook->add_format();            # Set properties later
#NYI     my $format2 = $workbook->add_format( %props );    # Set at creation
#NYI 
#NYI The format object holds all the formatting properties that can be applied to a cell, a row or a column. The process of setting these properties is discussed in the next section.
#NYI 
#NYI Once a Format object has been constructed and its properties have been set it can be passed as an argument to the worksheet C<write> methods as follows:
#NYI 
#NYI     $worksheet->write( 0, 0, 'One', $format );
#NYI     $worksheet->write_string( 1, 0, 'Two', $format );
#NYI     $worksheet->write_number( 2, 0, 3, $format );
#NYI     $worksheet->write_blank( 3, 0, $format );
#NYI 
#NYI Formats can also be passed to the worksheet C<set_row()> and C<set_column()> methods to define the default property for a row or column.
#NYI 
#NYI     $worksheet->set_row( 0, 15, $format );
#NYI     $worksheet->set_column( 0, 0, 15, $format );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Format methods and Format properties
#NYI 
#NYI The following table shows the Excel format categories, the formatting properties that can be applied and the equivalent object method:
#NYI 
#NYI 
#NYI     Category   Description       Property        Method Name
#NYI     --------   -----------       --------        -----------
#NYI     Font       Font type         font            set_font()
#NYI                Font size         size            set_size()
#NYI                Font color        color           set_color()
#NYI                Bold              bold            set_bold()
#NYI                Italic            italic          set_italic()
#NYI                Underline         underline       set_underline()
#NYI                Strikeout         font_strikeout  set_font_strikeout()
#NYI                Super/Subscript   font_script     set_font_script()
#NYI                Outline           font_outline    set_font_outline()
#NYI                Shadow            font_shadow     set_font_shadow()
#NYI 
#NYI     Number     Numeric format    num_format      set_num_format()
#NYI 
#NYI     Protection Lock cells        locked          set_locked()
#NYI                Hide formulas     hidden          set_hidden()
#NYI 
#NYI     Alignment  Horizontal align  align           set_align()
#NYI                Vertical align    valign          set_align()
#NYI                Rotation          rotation        set_rotation()
#NYI                Text wrap         text_wrap       set_text_wrap()
#NYI                Justify last      text_justlast   set_text_justlast()
#NYI                Center across     center_across   set_center_across()
#NYI                Indentation       indent          set_indent()
#NYI                Shrink to fit     shrink          set_shrink()
#NYI 
#NYI     Pattern    Cell pattern      pattern         set_pattern()
#NYI                Background color  bg_color        set_bg_color()
#NYI                Foreground color  fg_color        set_fg_color()
#NYI 
#NYI     Border     Cell border       border          set_border()
#NYI                Bottom border     bottom          set_bottom()
#NYI                Top border        top             set_top()
#NYI                Left border       left            set_left()
#NYI                Right border      right           set_right()
#NYI                Border color      border_color    set_border_color()
#NYI                Bottom color      bottom_color    set_bottom_color()
#NYI                Top color         top_color       set_top_color()
#NYI                Left color        left_color      set_left_color()
#NYI                Right color       right_color     set_right_color()
#NYI                Diagonal type     diag_type       set_diag_type()
#NYI                Diagonal border   diag_border     set_diag_border()
#NYI                Diagonal color    diag_color      set_diag_color()
#NYI 
#NYI There are two ways of setting Format properties: by using the object method interface or by setting the property directly. For example, a typical use of the method interface would be as follows:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_bold();
#NYI     $format->set_color( 'red' );
#NYI 
#NYI By comparison the properties can be set directly by passing a hash of properties to the Format constructor:
#NYI 
#NYI     my $format = $workbook->add_format( bold => 1, color => 'red' );
#NYI 
#NYI or after the Format has been constructed by means of the C<set_format_properties()> method as follows:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_format_properties( bold => 1, color => 'red' );
#NYI 
#NYI You can also store the properties in one or more named hashes and pass them to the required method:
#NYI 
#NYI     my %font = (
#NYI         font  => 'Calibri',
#NYI         size  => 12,
#NYI         color => 'blue',
#NYI         bold  => 1,
#NYI     );
#NYI 
#NYI     my %shading = (
#NYI         bg_color => 'green',
#NYI         pattern  => 1,
#NYI     );
#NYI 
#NYI 
#NYI     my $format1 = $workbook->add_format( %font );            # Font only
#NYI     my $format2 = $workbook->add_format( %font, %shading );  # Font and shading
#NYI 
#NYI 
#NYI The provision of two ways of setting properties might lead you to wonder which is the best way. The method mechanism may be better if you prefer setting properties via method calls (which the author did when the code was first written) otherwise passing properties to the constructor has proved to be a little more flexible and self documenting in practice. An additional advantage of working with property hashes is that it allows you to share formatting between workbook objects as shown in the example above.
#NYI 
#NYI The Perl/Tk style of adding properties is also supported:
#NYI 
#NYI     my %font = (
#NYI         -font  => 'Calibri',
#NYI         -size  => 12,
#NYI         -color => 'blue',
#NYI         -bold  => 1,
#NYI     );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Working with formats
#NYI 
#NYI The default format is Calibri 11 with all other properties off.
#NYI 
#NYI Each unique format in Excel::Writer::XLSX must have a corresponding Format object. It isn't possible to use a Format with a write() method and then redefine the Format for use at a later stage. This is because a Format is applied to a cell not in its current state but in its final state. Consider the following example:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_bold();
#NYI     $format->set_color( 'red' );
#NYI     $worksheet->write( 'A1', 'Cell A1', $format );
#NYI     $format->set_color( 'green' );
#NYI     $worksheet->write( 'B1', 'Cell B1', $format );
#NYI 
#NYI Cell A1 is assigned the Format C<$format> which is initially set to the colour red. However, the colour is subsequently set to green. When Excel displays Cell A1 it will display the final state of the Format which in this case will be the colour green.
#NYI 
#NYI In general a method call without an argument will turn a property on, for example:
#NYI 
#NYI     my $format1 = $workbook->add_format();
#NYI     $format1->set_bold();       # Turns bold on
#NYI     $format1->set_bold( 1 );    # Also turns bold on
#NYI     $format1->set_bold( 0 );    # Turns bold off
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 FORMAT METHODS
#NYI 
#NYI The Format object methods are described in more detail in the following sections. In addition, there is a Perl program called C<formats.pl> in the C<examples> directory of the WriteExcel distribution. This program creates an Excel workbook called C<formats.xlsx> which contains examples of almost all the format types.
#NYI 
#NYI The following Format methods are available:
#NYI 
#NYI     set_font()
#NYI     set_size()
#NYI     set_color()
#NYI     set_bold()
#NYI     set_italic()
#NYI     set_underline()
#NYI     set_font_strikeout()
#NYI     set_font_script()
#NYI     set_font_outline()
#NYI     set_font_shadow()
#NYI     set_num_format()
#NYI     set_locked()
#NYI     set_hidden()
#NYI     set_align()
#NYI     set_rotation()
#NYI     set_text_wrap()
#NYI     set_text_justlast()
#NYI     set_center_across()
#NYI     set_indent()
#NYI     set_shrink()
#NYI     set_pattern()
#NYI     set_bg_color()
#NYI     set_fg_color()
#NYI     set_border()
#NYI     set_bottom()
#NYI     set_top()
#NYI     set_left()
#NYI     set_right()
#NYI     set_border_color()
#NYI     set_bottom_color()
#NYI     set_top_color()
#NYI     set_left_color()
#NYI     set_right_color()
#NYI     set_diag_type()
#NYI     set_diag_border()
#NYI     set_diag_color()
#NYI 
#NYI 
#NYI The above methods can also be applied directly as properties. For example C<< $format->set_bold() >> is equivalent to C<< $workbook->add_format(bold => 1) >>.
#NYI 
#NYI 
#NYI =head2 set_format_properties( %properties )
#NYI 
#NYI The properties of an existing Format object can be also be set by means of C<set_format_properties()>:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_format_properties( bold => 1, color => 'red' );
#NYI 
#NYI However, this method is here mainly for legacy reasons. It is preferable to set the properties in the format constructor:
#NYI 
#NYI     my $format = $workbook->add_format( bold => 1, color => 'red' );
#NYI 
#NYI 
#NYI =head2 set_font( $fontname )
#NYI 
#NYI     Default state:      Font is Calibri
#NYI     Default action:     None
#NYI     Valid args:         Any valid font name
#NYI 
#NYI Specify the font used:
#NYI 
#NYI     $format->set_font('Times New Roman');
#NYI 
#NYI Excel can only display fonts that are installed on the system that it is running on. Therefore it is best to use the fonts that come as standard such as 'Calibri', 'Times New Roman' and 'Courier New'. See also the Fonts worksheet created by formats.pl
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_size()
#NYI 
#NYI     Default state:      Font size is 10
#NYI     Default action:     Set font size to 1
#NYI     Valid args:         Integer values from 1 to as big as your screen.
#NYI 
#NYI 
#NYI Set the font size. Excel adjusts the height of a row to accommodate the largest font size in the row. You can also explicitly specify the height of a row using the set_row() worksheet method.
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_size( 30 );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_color()
#NYI 
#NYI     Default state:      Excels default color, usually black
#NYI     Default action:     Set the default color
#NYI     Valid args:         Integers from 8..63 or the following strings:
#NYI                         'black'
#NYI                         'blue'
#NYI                         'brown'
#NYI                         'cyan'
#NYI                         'gray'
#NYI                         'green'
#NYI                         'lime'
#NYI                         'magenta'
#NYI                         'navy'
#NYI                         'orange'
#NYI                         'pink'
#NYI                         'purple'
#NYI                         'red'
#NYI                         'silver'
#NYI                         'white'
#NYI                         'yellow'
#NYI 
#NYI Set the font colour. The C<set_color()> method is used as follows:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_color( 'red' );
#NYI     $worksheet->write( 0, 0, 'wheelbarrow', $format );
#NYI 
#NYI Note: The C<set_color()> method is used to set the colour of the font in a cell. To set the colour of a cell use the C<set_bg_color()> and C<set_pattern()> methods.
#NYI 
#NYI For additional examples see the 'Named colors' and 'Standard colors' worksheets created by formats.pl in the examples directory.
#NYI 
#NYI See also L</WORKING WITH COLOURS>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_bold()
#NYI 
#NYI     Default state:      bold is off
#NYI     Default action:     Turn bold on
#NYI     Valid args:         0, 1
#NYI 
#NYI Set the bold property of the font:
#NYI 
#NYI     $format->set_bold();  # Turn bold on
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_italic()
#NYI 
#NYI     Default state:      Italic is off
#NYI     Default action:     Turn italic on
#NYI     Valid args:         0, 1
#NYI 
#NYI Set the italic property of the font:
#NYI 
#NYI     $format->set_italic();  # Turn italic on
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_underline()
#NYI 
#NYI     Default state:      Underline is off
#NYI     Default action:     Turn on single underline
#NYI     Valid args:         0  = No underline
#NYI                         1  = Single underline
#NYI                         2  = Double underline
#NYI                         33 = Single accounting underline
#NYI                         34 = Double accounting underline
#NYI 
#NYI Set the underline property of the font.
#NYI 
#NYI     $format->set_underline();   # Single underline
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_font_strikeout()
#NYI 
#NYI     Default state:      Strikeout is off
#NYI     Default action:     Turn strikeout on
#NYI     Valid args:         0, 1
#NYI 
#NYI Set the strikeout property of the font.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_font_script()
#NYI 
#NYI     Default state:      Super/Subscript is off
#NYI     Default action:     Turn Superscript on
#NYI     Valid args:         0  = Normal
#NYI                         1  = Superscript
#NYI                         2  = Subscript
#NYI 
#NYI Set the superscript/subscript property of the font.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_font_outline()
#NYI 
#NYI     Default state:      Outline is off
#NYI     Default action:     Turn outline on
#NYI     Valid args:         0, 1
#NYI 
#NYI Macintosh only.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_font_shadow()
#NYI 
#NYI     Default state:      Shadow is off
#NYI     Default action:     Turn shadow on
#NYI     Valid args:         0, 1
#NYI 
#NYI Macintosh only.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_num_format()
#NYI 
#NYI     Default state:      General format
#NYI     Default action:     Format index 1
#NYI     Valid args:         See the following table
#NYI 
#NYI This method is used to define the numerical format of a number in Excel. It controls whether a number is displayed as an integer, a floating point number, a date, a currency value or some other user defined format.
#NYI 
#NYI The numerical format of a cell can be specified by using a format string or an index to one of Excel's built-in formats:
#NYI 
#NYI     my $format1 = $workbook->add_format();
#NYI     my $format2 = $workbook->add_format();
#NYI     $format1->set_num_format( 'd mmm yyyy' );    # Format string
#NYI     $format2->set_num_format( 0x0f );            # Format index
#NYI 
#NYI     $worksheet->write( 0, 0, 36892.521, $format1 );    # 1 Jan 2001
#NYI     $worksheet->write( 0, 0, 36892.521, $format2 );    # 1-Jan-01
#NYI 
#NYI 
#NYI Using format strings you can define very sophisticated formatting of numbers.
#NYI 
#NYI     $format01->set_num_format( '0.000' );
#NYI     $worksheet->write( 0, 0, 3.1415926, $format01 );    # 3.142
#NYI 
#NYI     $format02->set_num_format( '#,##0' );
#NYI     $worksheet->write( 1, 0, 1234.56, $format02 );      # 1,235
#NYI 
#NYI     $format03->set_num_format( '#,##0.00' );
#NYI     $worksheet->write( 2, 0, 1234.56, $format03 );      # 1,234.56
#NYI 
#NYI     $format04->set_num_format( '$0.00' );
#NYI     $worksheet->write( 3, 0, 49.99, $format04 );        # $49.99
#NYI 
#NYI     # Note you can use other currency symbols such as the pound or yen as well.
#NYI     # Other currencies may require the use of Unicode.
#NYI 
#NYI     $format07->set_num_format( 'mm/dd/yy' );
#NYI     $worksheet->write( 6, 0, 36892.521, $format07 );    # 01/01/01
#NYI 
#NYI     $format08->set_num_format( 'mmm d yyyy' );
#NYI     $worksheet->write( 7, 0, 36892.521, $format08 );    # Jan 1 2001
#NYI 
#NYI     $format09->set_num_format( 'd mmmm yyyy' );
#NYI     $worksheet->write( 8, 0, 36892.521, $format09 );    # 1 January 2001
#NYI 
#NYI     $format10->set_num_format( 'dd/mm/yyyy hh:mm AM/PM' );
#NYI     $worksheet->write( 9, 0, 36892.521, $format10 );    # 01/01/2001 12:30 AM
#NYI 
#NYI     $format11->set_num_format( '0 "dollar and" .00 "cents"' );
#NYI     $worksheet->write( 10, 0, 1.87, $format11 );        # 1 dollar and .87 cents
#NYI 
#NYI     # Conditional numerical formatting.
#NYI     $format12->set_num_format( '[Green]General;[Red]-General;General' );
#NYI     $worksheet->write( 11, 0, 123, $format12 );         # > 0 Green
#NYI     $worksheet->write( 12, 0, -45, $format12 );         # < 0 Red
#NYI     $worksheet->write( 13, 0, 0,   $format12 );         # = 0 Default colour
#NYI 
#NYI     # Zip code
#NYI     $format13->set_num_format( '00000' );
#NYI     $worksheet->write( 14, 0, '01209', $format13 );
#NYI 
#NYI 
#NYI The number system used for dates is described in L</DATES AND TIME IN EXCEL>.
#NYI 
#NYI The colour format should have one of the following values:
#NYI 
#NYI     [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]
#NYI 
#NYI Alternatively you can specify the colour based on a colour index as follows: C<[Color n]>, where n is a standard Excel colour index - 7. See the 'Standard colors' worksheet created by formats.pl.
#NYI 
#NYI For more information refer to the documentation on formatting in the C<docs> directory of the Excel::Writer::XLSX distro, the Excel on-line help or L<http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx>.
#NYI 
#NYI You should ensure that the format string is valid in Excel prior to using it in WriteExcel.
#NYI 
#NYI Excel's built-in formats are shown in the following table:
#NYI 
#NYI     Index   Index   Format String
#NYI     0       0x00    General
#NYI     1       0x01    0
#NYI     2       0x02    0.00
#NYI     3       0x03    #,##0
#NYI     4       0x04    #,##0.00
#NYI     5       0x05    ($#,##0_);($#,##0)
#NYI     6       0x06    ($#,##0_);[Red]($#,##0)
#NYI     7       0x07    ($#,##0.00_);($#,##0.00)
#NYI     8       0x08    ($#,##0.00_);[Red]($#,##0.00)
#NYI     9       0x09    0%
#NYI     10      0x0a    0.00%
#NYI     11      0x0b    0.00E+00
#NYI     12      0x0c    # ?/?
#NYI     13      0x0d    # ??/??
#NYI     14      0x0e    m/d/yy
#NYI     15      0x0f    d-mmm-yy
#NYI     16      0x10    d-mmm
#NYI     17      0x11    mmm-yy
#NYI     18      0x12    h:mm AM/PM
#NYI     19      0x13    h:mm:ss AM/PM
#NYI     20      0x14    h:mm
#NYI     21      0x15    h:mm:ss
#NYI     22      0x16    m/d/yy h:mm
#NYI     ..      ....    ...........
#NYI     37      0x25    (#,##0_);(#,##0)
#NYI     38      0x26    (#,##0_);[Red](#,##0)
#NYI     39      0x27    (#,##0.00_);(#,##0.00)
#NYI     40      0x28    (#,##0.00_);[Red](#,##0.00)
#NYI     41      0x29    _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
#NYI     42      0x2a    _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
#NYI     43      0x2b    _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
#NYI     44      0x2c    _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
#NYI     45      0x2d    mm:ss
#NYI     46      0x2e    [h]:mm:ss
#NYI     47      0x2f    mm:ss.0
#NYI     48      0x30    ##0.0E+0
#NYI     49      0x31    @
#NYI 
#NYI 
#NYI For examples of these formatting codes see the 'Numerical formats' worksheet created by formats.pl. See also the number_formats1.html and the number_formats2.html documents in the C<docs> directory of the distro.
#NYI 
#NYI Note 1. Numeric formats 23 to 36 are not documented by Microsoft and may differ in international versions.
#NYI 
#NYI Note 2. The dollar sign appears as the defined local currency symbol.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_locked()
#NYI 
#NYI     Default state:      Cell locking is on
#NYI     Default action:     Turn locking on
#NYI     Valid args:         0, 1
#NYI 
#NYI This property can be used to prevent modification of a cells contents. Following Excel's convention, cell locking is turned on by default. However, it only has an effect if the worksheet has been protected, see the worksheet C<protect()> method.
#NYI 
#NYI     my $locked = $workbook->add_format();
#NYI     $locked->set_locked( 1 );    # A non-op
#NYI 
#NYI     my $unlocked = $workbook->add_format();
#NYI     $locked->set_locked( 0 );
#NYI 
#NYI     # Enable worksheet protection
#NYI     $worksheet->protect();
#NYI 
#NYI     # This cell cannot be edited.
#NYI     $worksheet->write( 'A1', '=1+2', $locked );
#NYI 
#NYI     # This cell can be edited.
#NYI     $worksheet->write( 'A2', '=1+2', $unlocked );
#NYI 
#NYI Note: This offers weak protection even with a password, see the note in relation to the C<protect()> method.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_hidden()
#NYI 
#NYI     Default state:      Formula hiding is off
#NYI     Default action:     Turn hiding on
#NYI     Valid args:         0, 1
#NYI 
#NYI This property is used to hide a formula while still displaying its result. This is generally used to hide complex calculations from end users who are only interested in the result. It only has an effect if the worksheet has been protected, see the worksheet C<protect()> method.
#NYI 
#NYI     my $hidden = $workbook->add_format();
#NYI     $hidden->set_hidden();
#NYI 
#NYI     # Enable worksheet protection
#NYI     $worksheet->protect();
#NYI 
#NYI     # The formula in this cell isn't visible
#NYI     $worksheet->write( 'A1', '=1+2', $hidden );
#NYI 
#NYI 
#NYI Note: This offers weak protection even with a password, see the note in relation to the C<protect()> method.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_align()
#NYI 
#NYI     Default state:      Alignment is off
#NYI     Default action:     Left alignment
#NYI     Valid args:         'left'              Horizontal
#NYI                         'center'
#NYI                         'right'
#NYI                         'fill'
#NYI                         'justify'
#NYI                         'center_across'
#NYI 
#NYI                         'top'               Vertical
#NYI                         'vcenter'
#NYI                         'bottom'
#NYI                         'vjustify'
#NYI 
#NYI This method is used to set the horizontal and vertical text alignment within a cell. Vertical and horizontal alignments can be combined. The method is used as follows:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_align( 'center' );
#NYI     $format->set_align( 'vcenter' );
#NYI     $worksheet->set_row( 0, 30 );
#NYI     $worksheet->write( 0, 0, 'X', $format );
#NYI 
#NYI Text can be aligned across two or more adjacent cells using the C<center_across> property. However, for genuine merged cells it is better to use the C<merge_range()> worksheet method.
#NYI 
#NYI The C<vjustify> (vertical justify) option can be used to provide automatic text wrapping in a cell. The height of the cell will be adjusted to accommodate the wrapped text. To specify where the text wraps use the C<set_text_wrap()> method.
#NYI 
#NYI 
#NYI For further examples see the 'Alignment' worksheet created by formats.pl.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_center_across()
#NYI 
#NYI     Default state:      Center across selection is off
#NYI     Default action:     Turn center across on
#NYI     Valid args:         1
#NYI 
#NYI Text can be aligned across two or more adjacent cells using the C<set_center_across()> method. This is an alias for the C<set_align('center_across')> method call.
#NYI 
#NYI Only one cell should contain the text, the other cells should be blank:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_center_across();
#NYI 
#NYI     $worksheet->write( 1, 1, 'Center across selection', $format );
#NYI     $worksheet->write_blank( 1, 2, $format );
#NYI 
#NYI See also the C<merge1.pl> to C<merge6.pl> programs in the C<examples> directory and the C<merge_range()> method.
#NYI 
#NYI 
#NYI 
#NYI =head2 set_text_wrap()
#NYI 
#NYI     Default state:      Text wrap is off
#NYI     Default action:     Turn text wrap on
#NYI     Valid args:         0, 1
#NYI 
#NYI 
#NYI Here is an example using the text wrap property, the escape character C<\n> is used to indicate the end of line:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_text_wrap();
#NYI     $worksheet->write( 0, 0, "It's\na bum\nwrap", $format );
#NYI 
#NYI Excel will adjust the height of the row to accommodate the wrapped text. A similar effect can be obtained without newlines using the C<set_align('vjustify')> method. See the C<textwrap.pl> program in the C<examples> directory.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_rotation()
#NYI 
#NYI     Default state:      Text rotation is off
#NYI     Default action:     None
#NYI     Valid args:         Integers in the range -90 to 90 and 270
#NYI 
#NYI Set the rotation of the text in a cell. The rotation can be any angle in the range -90 to 90 degrees.
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_rotation( 30 );
#NYI     $worksheet->write( 0, 0, 'This text is rotated', $format );
#NYI 
#NYI 
#NYI The angle 270 is also supported. This indicates text where the letters run from top to bottom.
#NYI 
#NYI 
#NYI 
#NYI =head2 set_indent()
#NYI 
#NYI     Default state:      Text indentation is off
#NYI     Default action:     Indent text 1 level
#NYI     Valid args:         Positive integers
#NYI 
#NYI 
#NYI This method can be used to indent text. The argument, which should be an integer, is taken as the level of indentation:
#NYI 
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_indent( 2 );
#NYI     $worksheet->write( 0, 0, 'This text is indented', $format );
#NYI 
#NYI 
#NYI Indentation is a horizontal alignment property. It will override any other horizontal properties but it can be used in conjunction with vertical properties.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_shrink()
#NYI 
#NYI     Default state:      Text shrinking is off
#NYI     Default action:     Turn "shrink to fit" on
#NYI     Valid args:         1
#NYI 
#NYI 
#NYI This method can be used to shrink text so that it fits in a cell.
#NYI 
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_shrink();
#NYI     $worksheet->write( 0, 0, 'Honey, I shrunk the text!', $format );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_text_justlast()
#NYI 
#NYI     Default state:      Justify last is off
#NYI     Default action:     Turn justify last on
#NYI     Valid args:         0, 1
#NYI 
#NYI 
#NYI Only applies to Far Eastern versions of Excel.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_pattern()
#NYI 
#NYI     Default state:      Pattern is off
#NYI     Default action:     Solid fill is on
#NYI     Valid args:         0 .. 18
#NYI 
#NYI Set the background pattern of a cell.
#NYI 
#NYI Examples of the available patterns are shown in the 'Patterns' worksheet created by formats.pl. However, it is unlikely that you will ever need anything other than Pattern 1 which is a solid fill of the background color.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_bg_color()
#NYI 
#NYI     Default state:      Color is off
#NYI     Default action:     Solid fill.
#NYI     Valid args:         See set_color()
#NYI 
#NYI The C<set_bg_color()> method can be used to set the background colour of a pattern. Patterns are defined via the C<set_pattern()> method. If a pattern hasn't been defined then a solid fill pattern is used as the default.
#NYI 
#NYI Here is an example of how to set up a solid fill in a cell:
#NYI 
#NYI     my $format = $workbook->add_format();
#NYI 
#NYI     $format->set_pattern();    # This is optional when using a solid fill
#NYI 
#NYI     $format->set_bg_color( 'green' );
#NYI     $worksheet->write( 'A1', 'Ray', $format );
#NYI 
#NYI For further examples see the 'Patterns' worksheet created by formats.pl.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_fg_color()
#NYI 
#NYI     Default state:      Color is off
#NYI     Default action:     Solid fill.
#NYI     Valid args:         See set_color()
#NYI 
#NYI 
#NYI The C<set_fg_color()> method can be used to set the foreground colour of a pattern.
#NYI 
#NYI For further examples see the 'Patterns' worksheet created by formats.pl.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_border()
#NYI 
#NYI     Also applies to:    set_bottom()
#NYI                         set_top()
#NYI                         set_left()
#NYI                         set_right()
#NYI 
#NYI     Default state:      Border is off
#NYI     Default action:     Set border type 1
#NYI     Valid args:         0-13, See below.
#NYI 
#NYI A cell border is comprised of a border on the bottom, top, left and right. These can be set to the same value using C<set_border()> or individually using the relevant method calls shown above.
#NYI 
#NYI The following shows the border styles sorted by Excel::Writer::XLSX index number:
#NYI 
#NYI     Index   Name            Weight   Style
#NYI     =====   =============   ======   ===========
#NYI     0       None            0
#NYI     1       Continuous      1        -----------
#NYI     2       Continuous      2        -----------
#NYI     3       Dash            1        - - - - - -
#NYI     4       Dot             1        . . . . . .
#NYI     5       Continuous      3        -----------
#NYI     6       Double          3        ===========
#NYI     7       Continuous      0        -----------
#NYI     8       Dash            2        - - - - - -
#NYI     9       Dash Dot        1        - . - . - .
#NYI     10      Dash Dot        2        - . - . - .
#NYI     11      Dash Dot Dot    1        - . . - . .
#NYI     12      Dash Dot Dot    2        - . . - . .
#NYI     13      SlantDash Dot   2        / - . / - .
#NYI 
#NYI 
#NYI The following shows the borders sorted by style:
#NYI 
#NYI     Name            Weight   Style         Index
#NYI     =============   ======   ===========   =====
#NYI     Continuous      0        -----------   7
#NYI     Continuous      1        -----------   1
#NYI     Continuous      2        -----------   2
#NYI     Continuous      3        -----------   5
#NYI     Dash            1        - - - - - -   3
#NYI     Dash            2        - - - - - -   8
#NYI     Dash Dot        1        - . - . - .   9
#NYI     Dash Dot        2        - . - . - .   10
#NYI     Dash Dot Dot    1        - . . - . .   11
#NYI     Dash Dot Dot    2        - . . - . .   12
#NYI     Dot             1        . . . . . .   4
#NYI     Double          3        ===========   6
#NYI     None            0                      0
#NYI     SlantDash Dot   2        / - . / - .   13
#NYI 
#NYI 
#NYI The following shows the borders in the order shown in the Excel Dialog.
#NYI 
#NYI     Index   Style             Index   Style
#NYI     =====   =====             =====   =====
#NYI     0       None              12      - . . - . .
#NYI     7       -----------       13      / - . / - .
#NYI     4       . . . . . .       10      - . - . - .
#NYI     11      - . . - . .       8       - - - - - -
#NYI     9       - . - . - .       2       -----------
#NYI     3       - - - - - -       5       -----------
#NYI     1       -----------       6       ===========
#NYI 
#NYI 
#NYI Examples of the available border styles are shown in the 'Borders' worksheet created by formats.pl.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_border_color()
#NYI 
#NYI     Also applies to:    set_bottom_color()
#NYI                         set_top_color()
#NYI                         set_left_color()
#NYI                         set_right_color()
#NYI 
#NYI     Default state:      Color is off
#NYI     Default action:     Undefined
#NYI     Valid args:         See set_color()
#NYI 
#NYI 
#NYI Set the colour of the cell borders. A cell border is comprised of a border on the bottom, top, left and right. These can be set to the same colour using C<set_border_color()> or individually using the relevant method calls shown above. Examples of the border styles and colours are shown in the 'Borders' worksheet created by formats.pl.
#NYI 
#NYI 
#NYI =head2 set_diag_type()
#NYI 
#NYI     Default state:      Diagonal border is off.
#NYI     Default action:     None.
#NYI     Valid args:         1-3, See below.
#NYI 
#NYI Set the diagonal border type for the cell. Three types of diagonal borders are available in Excel:
#NYI 
#NYI    1: From bottom left to top right.
#NYI    2: From top left to bottom right.
#NYI    3: Same as 1 and 2 combined.
#NYI 
#NYI For example:
#NYI 
#NYI     $format->set_diag_type( 3 );
#NYI 
#NYI 
#NYI 
#NYI =head2 set_diag_border()
#NYI 
#NYI     Default state:      Border is off
#NYI     Default action:     Set border type 1
#NYI     Valid args:         0-13, See below.
#NYI 
#NYI Set the diagonal border style. Same as the parameter to C<set_border()> above.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 set_diag_color()
#NYI 
#NYI     Default state:      Color is off
#NYI     Default action:     Undefined
#NYI     Valid args:         See set_color()
#NYI 
#NYI 
#NYI Set the colour of the diagonal cell border:
#NYI 
#NYI     $format->set_diag_type( 3 );
#NYI     $format->set_diag_border( 7 );
#NYI     $format->set_diag_color( 'red' );
#NYI 
#NYI 
#NYI 
#NYI =head2 copy( $format )
#NYI 
#NYI This method is used to copy all of the properties from one Format object to another:
#NYI 
#NYI     my $lorry1 = $workbook->add_format();
#NYI     $lorry1->set_bold();
#NYI     $lorry1->set_italic();
#NYI     $lorry1->set_color( 'red' );    # lorry1 is bold, italic and red
#NYI 
#NYI     my $lorry2 = $workbook->add_format();
#NYI     $lorry2->copy( $lorry1 );
#NYI     $lorry2->set_color( 'yellow' );    # lorry2 is bold, italic and yellow
#NYI 
#NYI The C<copy()> method is only useful if you are using the method interface to Format properties. It generally isn't required if you are setting Format properties directly using hashes.
#NYI 
#NYI 
#NYI Note: this is not a copy constructor, both objects must exist prior to copying.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 UNICODE IN EXCEL
#NYI 
#NYI The following is a brief introduction to handling Unicode in C<Excel::Writer::XLSX>.
#NYI 
#NYI I<For a more general introduction to Unicode handling in Perl see> L<perlunitut> and L<perluniintro>.
#NYI 
#NYI Excel::Writer::XLSX writer differs from Spreadsheet::WriteExcel in that it only handles Unicode data in C<UTF-8> format and doesn't try to handle legacy UTF-16 Excel formats.
#NYI 
#NYI If the data is in C<UTF-8> format then Excel::Writer::XLSX will handle it automatically.
#NYI 
#NYI If you are dealing with non-ASCII characters that aren't in C<UTF-8> then perl provides useful tools in the guise of the C<Encode> module to help you to convert to the required format. For example:
#NYI 
#NYI     use Encode 'decode';
#NYI 
#NYI     my $string = 'some string with koi8-r characters';
#NYI        $string = decode('koi8-r', $string); # koi8-r to utf8
#NYI 
#NYI Alternatively you can read data from an encoded file and convert it to C<UTF-8> as you read it in:
#NYI 
#NYI 
#NYI     my $file = 'unicode_koi8r.txt';
#NYI     open FH, '<:encoding(koi8-r)', $file or die "Couldn't open $file: $!\n";
#NYI 
#NYI     my $row = 0;
#NYI     while ( <FH> ) {
#NYI         # Data read in is now in utf8 format.
#NYI         chomp;
#NYI         $worksheet->write( $row++, 0, $_ );
#NYI     }
#NYI 
#NYI These methodologies are explained in more detail in L<perlunitut>, L<perluniintro> and L<perlunicode>.
#NYI 
#NYI If the program contains UTF-8 text then you will also need to add C<use utf8> to the includes:
#NYI 
#NYI     use utf8;
#NYI 
#NYI     ...
#NYI 
#NYI     $worksheet->write( 'A1', 'Some UTF-8 string' );
#NYI 
#NYI 
#NYI See also the C<unicode_*.pl> programs in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 WORKING WITH COLOURS
#NYI 
#NYI Throughout Excel::Writer::XLSX colours can be specified using a Html style C<#RRGGBB> value. For example with a Format object:
#NYI 
#NYI     $format->set_font_color( '#FF0000' );
#NYI 
#NYI For backward compatibility a limited number of color names are supported:
#NYI 
#NYI     $format->set_font_color( 'red' );
#NYI 
#NYI The color names supported are:
#NYI 
#NYI     black
#NYI     blue
#NYI     brown
#NYI     cyan
#NYI     gray
#NYI     green
#NYI     lime
#NYI     magenta
#NYI     navy
#NYI     orange
#NYI     pink
#NYI     purple
#NYI     red
#NYI     silver
#NYI     white
#NYI     yellow
#NYI 
#NYI See also C<colors.pl> in the C<examples> directory.
#NYI 
#NYI 
#NYI =head1 DATES AND TIME IN EXCEL
#NYI 
#NYI There are two important things to understand about dates and times in Excel:
#NYI 
#NYI =over 4
#NYI 
#NYI =item 1 A date/time in Excel is a real number plus an Excel number format.
#NYI 
#NYI =item 2 Excel::Writer::XLSX doesn't automatically convert date/time strings in C<write()> to an Excel date/time.
#NYI 
#NYI =back
#NYI 
#NYI These two points are explained in more detail below along with some suggestions on how to convert times and dates to the required format.
#NYI 
#NYI 
#NYI =head2 An Excel date/time is a number plus a format
#NYI 
#NYI If you write a date string with C<write()> then all you will get is a string:
#NYI 
#NYI     $worksheet->write( 'A1', '02/03/04' );   # !! Writes a string not a date. !!
#NYI 
#NYI Dates and times in Excel are represented by real numbers, for example "Jan 1 2001 12:30 AM" is represented by the number 36892.521.
#NYI 
#NYI The integer part of the number stores the number of days since the epoch and the fractional part stores the percentage of the day.
#NYI 
#NYI A date or time in Excel is just like any other number. To have the number display as a date you must apply an Excel number format to it. Here are some examples.
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'date_examples.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     $worksheet->set_column( 'A:A', 30 );    # For extra visibility.
#NYI 
#NYI     my $number = 39506.5;
#NYI 
#NYI     $worksheet->write( 'A1', $number );             #   39506.5
#NYI 
#NYI     my $format2 = $workbook->add_format( num_format => 'dd/mm/yy' );
#NYI     $worksheet->write( 'A2', $number, $format2 );    #  28/02/08
#NYI 
#NYI     my $format3 = $workbook->add_format( num_format => 'mm/dd/yy' );
#NYI     $worksheet->write( 'A3', $number, $format3 );    #  02/28/08
#NYI 
#NYI     my $format4 = $workbook->add_format( num_format => 'd-m-yyyy' );
#NYI     $worksheet->write( 'A4', $number, $format4 );    #  28-2-2008
#NYI 
#NYI     my $format5 = $workbook->add_format( num_format => 'dd/mm/yy hh:mm' );
#NYI     $worksheet->write( 'A5', $number, $format5 );    #  28/02/08 12:00
#NYI 
#NYI     my $format6 = $workbook->add_format( num_format => 'd mmm yyyy' );
#NYI     $worksheet->write( 'A6', $number, $format6 );    # 28 Feb 2008
#NYI 
#NYI     my $format7 = $workbook->add_format( num_format => 'mmm d yyyy hh:mm AM/PM' );
#NYI     $worksheet->write('A7', $number , $format7);     #  Feb 28 2008 12:00 PM
#NYI 
#NYI 
#NYI =head2 Excel::Writer::XLSX doesn't automatically convert date/time strings
#NYI 
#NYI Excel::Writer::XLSX doesn't automatically convert input date strings into Excel's formatted date numbers due to the large number of possible date formats and also due to the possibility of misinterpretation.
#NYI 
#NYI For example, does C<02/03/04> mean March 2 2004, February 3 2004 or even March 4 2002.
#NYI 
#NYI Therefore, in order to handle dates you will have to convert them to numbers and apply an Excel format. Some methods for converting dates are listed in the next section.
#NYI 
#NYI The most direct way is to convert your dates to the ISO8601 C<yyyy-mm-ddThh:mm:ss.sss> date format and use the C<write_date_time()> worksheet method:
#NYI 
#NYI     $worksheet->write_date_time( 'A2', '2001-01-01T12:20', $format );
#NYI 
#NYI See the C<write_date_time()> section of the documentation for more details.
#NYI 
#NYI A general methodology for handling date strings with C<write_date_time()> is:
#NYI 
#NYI     1. Identify incoming date/time strings with a regex.
#NYI     2. Extract the component parts of the date/time using the same regex.
#NYI     3. Convert the date/time to the ISO8601 format.
#NYI     4. Write the date/time using write_date_time() and a number format.
#NYI 
#NYI Here is an example:
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'example.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # Set the default format for dates.
#NYI     my $date_format = $workbook->add_format( num_format => 'mmm d yyyy' );
#NYI 
#NYI     # Increase column width to improve visibility of data.
#NYI     $worksheet->set_column( 'A:C', 20 );
#NYI 
#NYI     # Simulate reading from a data source.
#NYI     my $row = 0;
#NYI 
#NYI     while ( <DATA> ) {
#NYI         chomp;
#NYI 
#NYI         my $col  = 0;
#NYI         my @data = split ' ';
#NYI 
#NYI         for my $item ( @data ) {
#NYI 
#NYI             # Match dates in the following formats: d/m/yy, d/m/yyyy
#NYI             if ( $item =~ qr[^(\d{1,2})/(\d{1,2})/(\d{4})$] ) {
#NYI 
#NYI                 # Change to the date format required by write_date_time().
#NYI                 my $date = sprintf "%4d-%02d-%02dT", $3, $2, $1;
#NYI 
#NYI                 $worksheet->write_date_time( $row, $col++, $date,
#NYI                     $date_format );
#NYI             }
#NYI             else {
#NYI 
#NYI                 # Just plain data
#NYI                 $worksheet->write( $row, $col++, $item );
#NYI             }
#NYI         }
#NYI         $row++;
#NYI     }
#NYI 
#NYI     __DATA__
#NYI     Item    Cost    Date
#NYI     Book    10      1/9/2007
#NYI     Beer    4       12/9/2007
#NYI     Bed     500     5/10/2007
#NYI 
#NYI For a slightly more advanced solution you can modify the C<write()> method to handle date formats of your choice via the C<add_write_handler()> method. See the C<add_write_handler()> section of the docs and the write_handler3.pl and write_handler4.pl programs in the examples directory of the distro.
#NYI 
#NYI 
#NYI =head2 Converting dates and times to an Excel date or time
#NYI 
#NYI The C<write_date_time()> method above is just one way of handling dates and times.
#NYI 
#NYI You can also use the C<convert_date_time()> worksheet method to convert from an ISO8601 style date string to an Excel date and time number.
#NYI 
#NYI The L<Excel::Writer::XLSX::Utility> module which is included in the distro has date/time handling functions:
#NYI 
#NYI     use Excel::Writer::XLSX::Utility;
#NYI 
#NYI     $date           = xl_date_list(2002, 1, 1);         # 37257
#NYI     $date           = xl_parse_date("11 July 1997");    # 35622
#NYI     $time           = xl_parse_time('3:21:36 PM');      # 0.64
#NYI     $date           = xl_decode_date_EU("13 May 2002"); # 37389
#NYI 
#NYI Note: some of these functions require additional CPAN modules.
#NYI 
#NYI For date conversions using the CPAN C<DateTime> framework see L<DateTime::Format::Excel> L<http://search.cpan.org/search?dist=DateTime-Format-Excel>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 OUTLINES AND GROUPING IN EXCEL
#NYI 
#NYI 
#NYI Excel allows you to group rows or columns so that they can be hidden or displayed with a single mouse click. This feature is referred to as outlines.
#NYI 
#NYI Outlines can reduce complex data down to a few salient sub-totals or summaries.
#NYI 
#NYI This feature is best viewed in Excel but the following is an ASCII representation of what a worksheet with three outlines might look like. Rows 3-4 and rows 7-8 are grouped at level 2. Rows 2-9 are grouped at level 1. The lines at the left hand side are called outline level bars.
#NYI 
#NYI 
#NYI             ------------------------------------------
#NYI      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
#NYI             ------------------------------------------
#NYI       _    | 1 |   A   |       |       |       |  ...
#NYI      |  _  | 2 |   B   |       |       |       |  ...
#NYI      | |   | 3 |  (C)  |       |       |       |  ...
#NYI      | |   | 4 |  (D)  |       |       |       |  ...
#NYI      | -   | 5 |   E   |       |       |       |  ...
#NYI      |  _  | 6 |   F   |       |       |       |  ...
#NYI      | |   | 7 |  (G)  |       |       |       |  ...
#NYI      | |   | 8 |  (H)  |       |       |       |  ...
#NYI      | -   | 9 |   I   |       |       |       |  ...
#NYI      -     | . |  ...  |  ...  |  ...  |  ...  |  ...
#NYI 
#NYI 
#NYI Clicking the minus sign on each of the level 2 outlines will collapse and hide the data as shown in the next figure. The minus sign changes to a plus sign to indicate that the data in the outline is hidden.
#NYI 
#NYI             ------------------------------------------
#NYI      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
#NYI             ------------------------------------------
#NYI       _    | 1 |   A   |       |       |       |  ...
#NYI      |     | 2 |   B   |       |       |       |  ...
#NYI      | +   | 5 |   E   |       |       |       |  ...
#NYI      |     | 6 |   F   |       |       |       |  ...
#NYI      | +   | 9 |   I   |       |       |       |  ...
#NYI      -     | . |  ...  |  ...  |  ...  |  ...  |  ...
#NYI 
#NYI 
#NYI Clicking on the minus sign on the level 1 outline will collapse the remaining rows as follows:
#NYI 
#NYI             ------------------------------------------
#NYI      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
#NYI             ------------------------------------------
#NYI            | 1 |   A   |       |       |       |  ...
#NYI      +     | . |  ...  |  ...  |  ...  |  ...  |  ...
#NYI 
#NYI 
#NYI Grouping in C<Excel::Writer::XLSX> is achieved by setting the outline level via the C<set_row()> and C<set_column()> worksheet methods:
#NYI 
#NYI     set_row( $row, $height, $format, $hidden, $level, $collapsed )
#NYI     set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )
#NYI 
#NYI The following example sets an outline level of 1 for rows 1 and 2 (zero-indexed) and columns B to G. The parameters C<$height> and C<$XF> are assigned default values since they are undefined:
#NYI 
#NYI     $worksheet->set_row( 1, undef, undef, 0, 1 );
#NYI     $worksheet->set_row( 2, undef, undef, 0, 1 );
#NYI     $worksheet->set_column( 'B:G', undef, undef, 0, 1 );
#NYI 
#NYI Excel allows up to 7 outline levels. Therefore the C<$level> parameter should be in the range C<0 E<lt>= $level E<lt>= 7>.
#NYI 
#NYI Rows and columns can be collapsed by setting the C<$hidden> flag for the hidden rows/columns and setting the C<$collapsed> flag for the row/column that has the collapsed C<+> symbol:
#NYI 
#NYI     $worksheet->set_row( 1, undef, undef, 1, 1 );
#NYI     $worksheet->set_row( 2, undef, undef, 1, 1 );
#NYI     $worksheet->set_row( 3, undef, undef, 0, 0, 1 );          # Collapsed flag.
#NYI 
#NYI     $worksheet->set_column( 'B:G', undef, undef, 1, 1 );
#NYI     $worksheet->set_column( 'H:H', undef, undef, 0, 0, 1 );   # Collapsed flag.
#NYI 
#NYI Note: Setting the C<$collapsed> flag is particularly important for compatibility with OpenOffice.org and Gnumeric.
#NYI 
#NYI For a more complete example see the C<outline.pl> and C<outline_collapsed.pl> programs in the examples directory of the distro.
#NYI 
#NYI Some additional outline properties can be set via the C<outline_settings()> worksheet method, see above.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 DATA VALIDATION IN EXCEL
#NYI 
#NYI Data validation is a feature of Excel which allows you to restrict the data that a users enters in a cell and to display help and warning messages. It also allows you to restrict input to values in a drop down list.
#NYI 
#NYI A typical use case might be to restrict data in a cell to integer values in a certain range, to provide a help message to indicate the required value and to issue a warning if the input data doesn't meet the stated criteria. In Excel::Writer::XLSX we could do that as follows:
#NYI 
#NYI     $worksheet->data_validation('B3',
#NYI         {
#NYI             validate        => 'integer',
#NYI             criteria        => 'between',
#NYI             minimum         => 1,
#NYI             maximum         => 100,
#NYI             input_title     => 'Input an integer:',
#NYI             input_message   => 'Between 1 and 100',
#NYI             error_message   => 'Sorry, try again.',
#NYI         });
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/validation_example.jpg" alt="The output from the above example"/></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI For more information on data validation see the following Microsoft support article "Description and examples of data validation in Excel": L<http://support.microsoft.com/kb/211485>.
#NYI 
#NYI The following sections describe how to use the C<data_validation()> method and its various options.
#NYI 
#NYI 
#NYI =head2 data_validation( $row, $col, { parameter => 'value', ... } )
#NYI 
#NYI The C<data_validation()> method is used to construct an Excel data validation.
#NYI 
#NYI It can be applied to a single cell or a range of cells. You can pass 3 parameters such as C<($row, $col, {...})> or 5 parameters such as C<($first_row, $first_col, $last_row, $last_col, {...})>. You can also use C<A1> style notation. For example:
#NYI 
#NYI     $worksheet->data_validation( 0, 0,       {...} );
#NYI     $worksheet->data_validation( 0, 0, 4, 1, {...} );
#NYI 
#NYI     # Which are the same as:
#NYI 
#NYI     $worksheet->data_validation( 'A1',       {...} );
#NYI     $worksheet->data_validation( 'A1:B5',    {...} );
#NYI 
#NYI See also the note about L</Cell notation> for more information.
#NYI 
#NYI 
#NYI The last parameter in C<data_validation()> must be a hash ref containing the parameters that describe the type and style of the data validation. The allowable parameters are:
#NYI 
#NYI     validate
#NYI     criteria
#NYI     value | minimum | source
#NYI     maximum
#NYI     ignore_blank
#NYI     dropdown
#NYI 
#NYI     input_title
#NYI     input_message
#NYI     show_input
#NYI 
#NYI     error_title
#NYI     error_message
#NYI     error_type
#NYI     show_error
#NYI 
#NYI These parameters are explained in the following sections. Most of the parameters are optional, however, you will generally require the three main options C<validate>, C<criteria> and C<value>.
#NYI 
#NYI     $worksheet->data_validation('B3',
#NYI         {
#NYI             validate => 'integer',
#NYI             criteria => '>',
#NYI             value    => 100,
#NYI         });
#NYI 
#NYI The C<data_validation> method returns:
#NYI 
#NYI      0 for success.
#NYI     -1 for insufficient number of arguments.
#NYI     -2 for row or column out of bounds.
#NYI     -3 for incorrect parameter or value.
#NYI 
#NYI 
#NYI =head2 validate
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<validate> parameter is used to set the type of data that you wish to validate. It is always required and it has no default value. Allowable values are:
#NYI 
#NYI     any
#NYI     integer
#NYI     decimal
#NYI     list
#NYI     date
#NYI     time
#NYI     length
#NYI     custom
#NYI 
#NYI =over
#NYI 
#NYI =item * B<any> is used to specify that the type of data is unrestricted. This is useful to display an input message without restricting the data that can be entered.
#NYI 
#NYI =item * B<integer> restricts the cell to integer values. Excel refers to this as 'whole number'.
#NYI 
#NYI     validate => 'integer',
#NYI     criteria => '>',
#NYI     value    => 100,
#NYI 
#NYI =item * B<decimal> restricts the cell to decimal values.
#NYI 
#NYI     validate => 'decimal',
#NYI     criteria => '>',
#NYI     value    => 38.6,
#NYI 
#NYI =item * B<list> restricts the cell to a set of user specified values. These can be passed in an array ref or as a cell range (named ranges aren't currently supported):
#NYI 
#NYI     validate => 'list',
#NYI     value    => ['open', 'high', 'close'],
#NYI     # Or like this:
#NYI     value    => 'B1:B3',
#NYI 
#NYI Excel requires that range references are only to cells on the same worksheet.
#NYI 
#NYI =item * B<date> restricts the cell to date values. Dates in Excel are expressed as integer values but you can also pass an ISO8601 style string as used in C<write_date_time()>. See also L</DATES AND TIME IN EXCEL> for more information about working with Excel's dates.
#NYI 
#NYI     validate => 'date',
#NYI     criteria => '>',
#NYI     value    => 39653, # 24 July 2008
#NYI     # Or like this:
#NYI     value    => '2008-07-24T',
#NYI 
#NYI =item * B<time> restricts the cell to time values. Times in Excel are expressed as decimal values but you can also pass an ISO8601 style string as used in C<write_date_time()>. See also L</DATES AND TIME IN EXCEL> for more information about working with Excel's times.
#NYI 
#NYI     validate => 'time',
#NYI     criteria => '>',
#NYI     value    => 0.5, # Noon
#NYI     # Or like this:
#NYI     value    => 'T12:00:00',
#NYI 
#NYI =item * B<length> restricts the cell data based on an integer string length. Excel refers to this as 'Text length'.
#NYI 
#NYI     validate => 'length',
#NYI     criteria => '>',
#NYI     value    => 10,
#NYI 
#NYI =item * B<custom> restricts the cell based on an external Excel formula that returns a C<TRUE/FALSE> value.
#NYI 
#NYI     validate => 'custom',
#NYI     value    => '=IF(A10>B10,TRUE,FALSE)',
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI =head2 criteria
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<criteria> parameter is used to set the criteria by which the data in the cell is validated. It is almost always required except for the C<list> and C<custom> validate options. It has no default value. Allowable values are:
#NYI 
#NYI     'between'
#NYI     'not between'
#NYI     'equal to'                  |  '=='  |  '='
#NYI     'not equal to'              |  '!='  |  '<>'
#NYI     'greater than'              |  '>'
#NYI     'less than'                 |  '<'
#NYI     'greater than or equal to'  |  '>='
#NYI     'less than or equal to'     |  '<='
#NYI 
#NYI You can either use Excel's textual description strings, in the first column above, or the more common symbolic alternatives. The following are equivalent:
#NYI 
#NYI     validate => 'integer',
#NYI     criteria => 'greater than',
#NYI     value    => 100,
#NYI 
#NYI     validate => 'integer',
#NYI     criteria => '>',
#NYI     value    => 100,
#NYI 
#NYI The C<list> and C<custom> validate options don't require a C<criteria>. If you specify one it will be ignored.
#NYI 
#NYI     validate => 'list',
#NYI     value    => ['open', 'high', 'close'],
#NYI 
#NYI     validate => 'custom',
#NYI     value    => '=IF(A10>B10,TRUE,FALSE)',
#NYI 
#NYI 
#NYI =head2 value | minimum | source
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<value> parameter is used to set the limiting value to which the C<criteria> is applied. It is always required and it has no default value. You can also use the synonyms C<minimum> or C<source> to make the validation a little clearer and closer to Excel's description of the parameter:
#NYI 
#NYI     # Use 'value'
#NYI     validate => 'integer',
#NYI     criteria => '>',
#NYI     value    => 100,
#NYI 
#NYI     # Use 'minimum'
#NYI     validate => 'integer',
#NYI     criteria => 'between',
#NYI     minimum  => 1,
#NYI     maximum  => 100,
#NYI 
#NYI     # Use 'source'
#NYI     validate => 'list',
#NYI     source   => '$B$1:$B$3',
#NYI 
#NYI 
#NYI =head2 maximum
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<maximum> parameter is used to set the upper limiting value when the C<criteria> is either C<'between'> or C<'not between'>:
#NYI 
#NYI     validate => 'integer',
#NYI     criteria => 'between',
#NYI     minimum  => 1,
#NYI     maximum  => 100,
#NYI 
#NYI 
#NYI =head2 ignore_blank
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<ignore_blank> parameter is used to toggle on and off the 'Ignore blank' option in the Excel data validation dialog. When the option is on the data validation is not applied to blank data in the cell. It is on by default.
#NYI 
#NYI     ignore_blank => 0,  # Turn the option off
#NYI 
#NYI 
#NYI =head2 dropdown
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<dropdown> parameter is used to toggle on and off the 'In-cell dropdown' option in the Excel data validation dialog. When the option is on a dropdown list will be shown for C<list> validations. It is on by default.
#NYI 
#NYI     dropdown => 0,      # Turn the option off
#NYI 
#NYI 
#NYI =head2 input_title
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<input_title> parameter is used to set the title of the input message that is displayed when a cell is entered. It has no default value and is only displayed if the input message is displayed. See the C<input_message> parameter below.
#NYI 
#NYI     input_title   => 'This is the input title',
#NYI 
#NYI The maximum title length is 32 characters.
#NYI 
#NYI 
#NYI =head2 input_message
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<input_message> parameter is used to set the input message that is displayed when a cell is entered. It has no default value.
#NYI 
#NYI     validate      => 'integer',
#NYI     criteria      => 'between',
#NYI     minimum       => 1,
#NYI     maximum       => 100,
#NYI     input_title   => 'Enter the applied discount:',
#NYI     input_message => 'between 1 and 100',
#NYI 
#NYI The message can be split over several lines using newlines, C<"\n"> in double quoted strings.
#NYI 
#NYI     input_message => "This is\na test.",
#NYI 
#NYI The maximum message length is 255 characters.
#NYI 
#NYI 
#NYI =head2 show_input
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<show_input> parameter is used to toggle on and off the 'Show input message when cell is selected' option in the Excel data validation dialog. When the option is off an input message is not displayed even if it has been set using C<input_message>. It is on by default.
#NYI 
#NYI     show_input => 0,      # Turn the option off
#NYI 
#NYI 
#NYI =head2 error_title
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<error_title> parameter is used to set the title of the error message that is displayed when the data validation criteria is not met. The default error title is 'Microsoft Excel'.
#NYI 
#NYI     error_title   => 'Input value is not valid',
#NYI 
#NYI The maximum title length is 32 characters.
#NYI 
#NYI 
#NYI =head2 error_message
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<error_message> parameter is used to set the error message that is displayed when a cell is entered. The default error message is "The value you entered is not valid.\nA user has restricted values that can be entered into the cell.".
#NYI 
#NYI     validate      => 'integer',
#NYI     criteria      => 'between',
#NYI     minimum       => 1,
#NYI     maximum       => 100,
#NYI     error_title   => 'Input value is not valid',
#NYI     error_message => 'It should be an integer between 1 and 100',
#NYI 
#NYI The message can be split over several lines using newlines, C<"\n"> in double quoted strings.
#NYI 
#NYI     input_message => "This is\na test.",
#NYI 
#NYI The maximum message length is 255 characters.
#NYI 
#NYI 
#NYI =head2 error_type
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<error_type> parameter is used to specify the type of error dialog that is displayed. There are 3 options:
#NYI 
#NYI     'stop'
#NYI     'warning'
#NYI     'information'
#NYI 
#NYI The default is C<'stop'>.
#NYI 
#NYI 
#NYI =head2 show_error
#NYI 
#NYI This parameter is passed in a hash ref to C<data_validation()>.
#NYI 
#NYI The C<show_error> parameter is used to toggle on and off the 'Show error alert after invalid data is entered' option in the Excel data validation dialog. When the option is off an error message is not displayed even if it has been set using C<error_message>. It is on by default.
#NYI 
#NYI     show_error => 0,      # Turn the option off
#NYI 
#NYI =head2 Data Validation Examples
#NYI 
#NYI Example 1. Limiting input to an integer greater than a fixed value.
#NYI 
#NYI     $worksheet->data_validation('A1',
#NYI         {
#NYI             validate        => 'integer',
#NYI             criteria        => '>',
#NYI             value           => 0,
#NYI         });
#NYI 
#NYI Example 2. Limiting input to an integer greater than a fixed value where the value is referenced from a cell.
#NYI 
#NYI     $worksheet->data_validation('A2',
#NYI         {
#NYI             validate        => 'integer',
#NYI             criteria        => '>',
#NYI             value           => '=E3',
#NYI         });
#NYI 
#NYI Example 3. Limiting input to a decimal in a fixed range.
#NYI 
#NYI     $worksheet->data_validation('A3',
#NYI         {
#NYI             validate        => 'decimal',
#NYI             criteria        => 'between',
#NYI             minimum         => 0.1,
#NYI             maximum         => 0.5,
#NYI         });
#NYI 
#NYI Example 4. Limiting input to a value in a dropdown list.
#NYI 
#NYI     $worksheet->data_validation('A4',
#NYI         {
#NYI             validate        => 'list',
#NYI             source          => ['open', 'high', 'close'],
#NYI         });
#NYI 
#NYI Example 5. Limiting input to a value in a dropdown list where the list is specified as a cell range.
#NYI 
#NYI     $worksheet->data_validation('A5',
#NYI         {
#NYI             validate        => 'list',
#NYI             source          => '=$E$4:$G$4',
#NYI         });
#NYI 
#NYI Example 6. Limiting input to a date in a fixed range.
#NYI 
#NYI     $worksheet->data_validation('A6',
#NYI         {
#NYI             validate        => 'date',
#NYI             criteria        => 'between',
#NYI             minimum         => '2008-01-01T',
#NYI             maximum         => '2008-12-12T',
#NYI         });
#NYI 
#NYI Example 7. Displaying a message when the cell is selected.
#NYI 
#NYI     $worksheet->data_validation('A7',
#NYI         {
#NYI             validate      => 'integer',
#NYI             criteria      => 'between',
#NYI             minimum       => 1,
#NYI             maximum       => 100,
#NYI             input_title   => 'Enter an integer:',
#NYI             input_message => 'between 1 and 100',
#NYI         });
#NYI 
#NYI See also the C<data_validate.pl> program in the examples directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 CONDITIONAL FORMATTING IN EXCEL
#NYI 
#NYI Conditional formatting is a feature of Excel which allows you to apply a format to a cell or a range of cells based on a certain criteria.
#NYI 
#NYI For example the following criteria is used to highlight cells >= 50 in red in the C<conditional_format.pl> example from the distro:
#NYI 
#NYI     # Write a conditional format over a range.
#NYI     $worksheet->conditional_formatting( 'B3:K12',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => '>=',
#NYI             value    => 50,
#NYI             format   => $format1,
#NYI         }
#NYI     );
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/conditional_example.jpg" alt="The output from the above example"/></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI 
#NYI =head2 conditional_formatting( $row, $col, { parameter => 'value', ... } )
#NYI 
#NYI The C<conditional_formatting()> method is used to apply formatting  based on user defined criteria to an Excel::Writer::XLSX file.
#NYI 
#NYI It can be applied to a single cell or a range of cells. You can pass 3 parameters such as C<($row, $col, {...})> or 5 parameters such as C<($first_row, $first_col, $last_row, $last_col, {...})>. You can also use C<A1> style notation. For example:
#NYI 
#NYI     $worksheet->conditional_formatting( 0, 0,       {...} );
#NYI     $worksheet->conditional_formatting( 0, 0, 4, 1, {...} );
#NYI 
#NYI     # Which are the same as:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1',       {...} );
#NYI     $worksheet->conditional_formatting( 'A1:B5',    {...} );
#NYI 
#NYI See also the note about L</Cell notation> for more information.
#NYI 
#NYI Using C<A1> style notation is also possible to specify non-contiguous ranges, separated by a comma. For example:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:D5,A8:D12', {...} );
#NYI 
#NYI The last parameter in C<conditional_formatting()> must be a hash ref containing the parameters that describe the type and style of the data validation. The main parameters are:
#NYI 
#NYI     type
#NYI     format
#NYI     criteria
#NYI     value
#NYI     minimum
#NYI     maximum
#NYI 
#NYI Other, less commonly used parameters are:
#NYI 
#NYI     min_type
#NYI     mid_type
#NYI     max_type
#NYI     min_value
#NYI     mid_value
#NYI     max_value
#NYI     min_color
#NYI     mid_color
#NYI     max_color
#NYI     bar_color
#NYI     stop_if_true
#NYI     icon_style
#NYI     icons
#NYI     reverse_icons
#NYI     icons_only
#NYI 
#NYI Additional parameters which are used for specific conditional format types are shown in the relevant sections below.
#NYI 
#NYI =head2 type
#NYI 
#NYI This parameter is passed in a hash ref to C<conditional_formatting()>.
#NYI 
#NYI The C<type> parameter is used to set the type of conditional formatting that you wish to apply. It is always required and it has no default value. Allowable C<type> values and their associated parameters are:
#NYI 
#NYI     Type            Parameters
#NYI     ====            ==========
#NYI     cell            criteria
#NYI                     value
#NYI                     minimum
#NYI                     maximum
#NYI 
#NYI     date            criteria
#NYI                     value
#NYI                     minimum
#NYI                     maximum
#NYI 
#NYI     time_period     criteria
#NYI 
#NYI     text            criteria
#NYI                     value
#NYI 
#NYI     average         criteria
#NYI 
#NYI     duplicate       (none)
#NYI 
#NYI     unique          (none)
#NYI 
#NYI     top             criteria
#NYI                     value
#NYI 
#NYI     bottom          criteria
#NYI                     value
#NYI 
#NYI     blanks          (none)
#NYI 
#NYI     no_blanks       (none)
#NYI 
#NYI     errors          (none)
#NYI 
#NYI     no_errors       (none)
#NYI 
#NYI     2_color_scale   (none)
#NYI 
#NYI     3_color_scale   (none)
#NYI 
#NYI     data_bar        (none)
#NYI 
#NYI     formula         criteria
#NYI 
#NYI     icon_set        icon_style
#NYI                     reverse_icons
#NYI                     icons
#NYI                     icons_only
#NYI 
#NYI 
#NYI All conditional formatting types, apart from C<icon_set> have a C<format> parameter, see below.
#NYI 
#NYI =head2 type => 'cell'
#NYI 
#NYI This is the most common conditional formatting type. It is used when a format is applied to a cell based on a simple criterion. For example:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => 'greater than',
#NYI             value    => 5,
#NYI             format   => $red_format,
#NYI         }
#NYI     );
#NYI 
#NYI Or, using the C<between> criteria:
#NYI 
#NYI     $worksheet->conditional_formatting( 'C1:C4',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => 'between',
#NYI             minimum  => 20,
#NYI             maximum  => 30,
#NYI             format   => $green_format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 criteria
#NYI 
#NYI The C<criteria> parameter is used to set the criteria by which the cell data will be evaluated. It has no default value. The most common criteria as applied to C<< { type => 'cell' } >> are:
#NYI 
#NYI     'between'
#NYI     'not between'
#NYI     'equal to'                  |  '=='  |  '='
#NYI     'not equal to'              |  '!='  |  '<>'
#NYI     'greater than'              |  '>'
#NYI     'less than'                 |  '<'
#NYI     'greater than or equal to'  |  '>='
#NYI     'less than or equal to'     |  '<='
#NYI 
#NYI You can either use Excel's textual description strings, in the first column above, or the more common symbolic alternatives.
#NYI 
#NYI Additional criteria which are specific to other conditional format types are shown in the relevant sections below.
#NYI 
#NYI 
#NYI =head2 value
#NYI 
#NYI The C<value> is generally used along with the C<criteria> parameter to set the rule by which the cell data  will be evaluated.
#NYI 
#NYI     type     => 'cell',
#NYI     criteria => '>',
#NYI     value    => 5
#NYI     format   => $format,
#NYI 
#NYI The C<value> property can also be an cell reference.
#NYI 
#NYI     type     => 'cell',
#NYI     criteria => '>',
#NYI     value    => '$C$1',
#NYI     format   => $format,
#NYI 
#NYI 
#NYI =head2 format
#NYI 
#NYI The C<format> parameter is used to specify the format that will be applied to the cell when the conditional formatting criterion is met. The format is created using the C<add_format()> method in the same way as cell formats:
#NYI 
#NYI     $format = $workbook->add_format( bold => 1, italic => 1 );
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => '>',
#NYI             value    => 5
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI The conditional format follows the same rules as in Excel: it is superimposed over the existing cell format and not all font and border properties can be modified. Font properties that can't be modified are font name, font size, superscript and subscript. The border property that cannot be modified is diagonal borders.
#NYI 
#NYI Excel specifies some default formats to be used with conditional formatting. You can replicate them using the following Excel::Writer::XLSX formats:
#NYI 
#NYI     # Light red fill with dark red text.
#NYI 
#NYI     my $format1 = $workbook->add_format(
#NYI         bg_color => '#FFC7CE',
#NYI         color    => '#9C0006',
#NYI     );
#NYI 
#NYI     # Light yellow fill with dark yellow text.
#NYI 
#NYI     my $format2 = $workbook->add_format(
#NYI         bg_color => '#FFEB9C',
#NYI         color    => '#9C6500',
#NYI     );
#NYI 
#NYI     # Green fill with dark green text.
#NYI 
#NYI     my $format3 = $workbook->add_format(
#NYI         bg_color => '#C6EFCE',
#NYI         color    => '#006100',
#NYI     );
#NYI 
#NYI 
#NYI =head2 minimum
#NYI 
#NYI The C<minimum> parameter is used to set the lower limiting value when the C<criteria> is either C<'between'> or C<'not between'>:
#NYI 
#NYI     validate => 'integer',
#NYI     criteria => 'between',
#NYI     minimum  => 1,
#NYI     maximum  => 100,
#NYI 
#NYI 
#NYI =head2 maximum
#NYI 
#NYI The C<maximum> parameter is used to set the upper limiting value when the C<criteria> is either C<'between'> or C<'not between'>. See the previous example.
#NYI 
#NYI 
#NYI =head2 type => 'date'
#NYI 
#NYI The C<date> type is the same as the C<cell> type and uses the same criteria and values. However it allows the C<value>, C<minimum> and C<maximum> properties to be specified in the ISO8601 C<yyyy-mm-ddThh:mm:ss.sss> date format which is detailed in the C<write_date_time()> method.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'date',
#NYI             criteria => 'greater than',
#NYI             value    => '2011-01-01T',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'time_period'
#NYI 
#NYI The C<time_period> type is used to specify Excel's "Dates Occurring" style conditional format.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'time_period',
#NYI             criteria => 'yesterday',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI The period is set in the C<criteria> and can have one of the following values:
#NYI 
#NYI         criteria => 'yesterday',
#NYI         criteria => 'today',
#NYI         criteria => 'last 7 days',
#NYI         criteria => 'last week',
#NYI         criteria => 'this week',
#NYI         criteria => 'next week',
#NYI         criteria => 'last month',
#NYI         criteria => 'this month',
#NYI         criteria => 'next month'
#NYI 
#NYI 
#NYI =head2 type => 'text'
#NYI 
#NYI The C<text> type is used to specify Excel's "Specific Text" style conditional format. It is used to do simple string matching using the C<criteria> and C<value> parameters:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'text',
#NYI             criteria => 'containing',
#NYI             value    => 'foo',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI The C<criteria> can have one of the following values:
#NYI 
#NYI     criteria => 'containing',
#NYI     criteria => 'not containing',
#NYI     criteria => 'begins with',
#NYI     criteria => 'ends with',
#NYI 
#NYI The C<value> parameter should be a string or single character.
#NYI 
#NYI 
#NYI =head2 type => 'average'
#NYI 
#NYI The C<average> type is used to specify Excel's "Average" style conditional format.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'average',
#NYI             criteria => 'above',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI The type of average for the conditional format range is specified by the C<criteria>:
#NYI 
#NYI     criteria => 'above',
#NYI     criteria => 'below',
#NYI     criteria => 'equal or above',
#NYI     criteria => 'equal or below',
#NYI     criteria => '1 std dev above',
#NYI     criteria => '1 std dev below',
#NYI     criteria => '2 std dev above',
#NYI     criteria => '2 std dev below',
#NYI     criteria => '3 std dev above',
#NYI     criteria => '3 std dev below',
#NYI 
#NYI 
#NYI 
#NYI =head2 type => 'duplicate'
#NYI 
#NYI The C<duplicate> type is used to highlight duplicate cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'duplicate',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'unique'
#NYI 
#NYI The C<unique> type is used to highlight unique cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'unique',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'top'
#NYI 
#NYI The C<top> type is used to specify the top C<n> values by number or percentage in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'top',
#NYI             value    => 10,
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI The C<criteria> can be used to indicate that a percentage condition is required:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'top',
#NYI             value    => 10,
#NYI             criteria => '%',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'bottom'
#NYI 
#NYI The C<bottom> type is used to specify the bottom C<n> values by number or percentage in a range.
#NYI 
#NYI It takes the same parameters as C<top>, see above.
#NYI 
#NYI 
#NYI =head2 type => 'blanks'
#NYI 
#NYI The C<blanks> type is used to highlight blank cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'blanks',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'no_blanks'
#NYI 
#NYI The C<no_blanks> type is used to highlight non blank cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'no_blanks',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'errors'
#NYI 
#NYI The C<errors> type is used to highlight error cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'errors',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => 'no_errors'
#NYI 
#NYI The C<no_errors> type is used to highlight non error cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'no_errors',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =head2 type => '2_color_scale'
#NYI 
#NYI The C<2_color_scale> type is used to specify Excel's "2 Color Scale" style conditional format.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A12',
#NYI         {
#NYI             type  => '2_color_scale',
#NYI         }
#NYI     );
#NYI 
#NYI This conditional type can be modified with C<min_type>, C<max_type>, C<min_value>, C<max_value>, C<min_color> and C<max_color>, see below.
#NYI 
#NYI 
#NYI =head2 type => '3_color_scale'
#NYI 
#NYI The C<3_color_scale> type is used to specify Excel's "3 Color Scale" style conditional format.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A12',
#NYI         {
#NYI             type  => '3_color_scale',
#NYI         }
#NYI     );
#NYI 
#NYI This conditional type can be modified with C<min_type>, C<mid_type>, C<max_type>, C<min_value>, C<mid_value>, C<max_value>, C<min_color>, C<mid_color> and C<max_color>, see below.
#NYI 
#NYI 
#NYI =head2 type => 'data_bar'
#NYI 
#NYI The C<data_bar> type is used to specify Excel's "Data Bar" style conditional format.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A12',
#NYI         {
#NYI             type  => 'data_bar',
#NYI         }
#NYI     );
#NYI 
#NYI This conditional type can be modified with C<min_type>, C<max_type>, C<min_value>, C<max_value> and C<bar_color>, see below.
#NYI 
#NYI 
#NYI 
#NYI =head2 type => 'formula'
#NYI 
#NYI The C<formula> type is used to specify a conditional format based on a user defined formula:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A4',
#NYI         {
#NYI             type     => 'formula',
#NYI             criteria => '=$A$1 > 5',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI The formula is specified in the C<criteria>.
#NYI 
#NYI 
#NYI 
#NYI =head2 type => 'icon_set'
#NYI 
#NYI The C<icon_set> type is used to specify a conditional format with a set of icons such as traffic lights or arrows:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:C1',
#NYI         {
#NYI             type         => 'icon_set',
#NYI             icon_style   => '3_traffic_lights',
#NYI         }
#NYI     );
#NYI 
#NYI The icon set style is specified by the C<icon_style> parameter. Valid options are:
#NYI 
#NYI     3_arrows
#NYI     3_arrows_gray
#NYI     3_flags
#NYI     3_signs
#NYI     3_symbols
#NYI     3_symbols_circled
#NYI     3_traffic_lights
#NYI     3_traffic_lights_rimmed
#NYI 
#NYI     4_arrows
#NYI     4_arrows_gray
#NYI     4_ratings
#NYI     4_red_to_black
#NYI     4_traffic_lights
#NYI 
#NYI     5_arrows
#NYI     5_arrows_gray
#NYI     5_quarters
#NYI     5_ratings
#NYI 
#NYI The criteria, type and value of each icon can be specified using the C<icon> array of hash refs with optional C<criteria>, C<type> and C<value> parameters:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:D1',
#NYI         {
#NYI             type         => 'icon_set',
#NYI             icon_style   => '4_red_to_black',
#NYI             icons        => [ {criteria => '>',  type => 'number',     value => 90},
#NYI                               {criteria => '>=', type => 'percentile', value => 50},
#NYI                               {criteria => '>',  type => 'percent',    value => 25},
#NYI                             ],
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI The C<icons criteria> parameter should be either C<< >= >> or C<< > >>. The default C<criteria> is C<< >= >>.
#NYI 
#NYI The C<icons type> parameter should be one of the following values:
#NYI 
#NYI     number
#NYI     percentile
#NYI     percent
#NYI     formula
#NYI 
#NYI The default C<type> is C<percent>.
#NYI 
#NYI The C<icons value> parameter can be a value or formula:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:D1',
#NYI         {
#NYI             type         => 'icon_set',
#NYI             icon_style   => '4_red_to_black',
#NYI             icons        => [ {value => 90},
#NYI                               {value => 50},
#NYI                               {value => 25},
#NYI                             ],
#NYI         }
#NYI     );
#NYI 
#NYI Note: The C<icons> parameters should start with the highest value and with each subsequent one being lower. The default C<value> is C<(n * 100) / number_of_icons>. The lowest number icon in an icon set has properties defined by Excel. Therefore in a C<n> icon set, there is no C<n-1> hash of parameters.
#NYI 
#NYI The order of the icons can be reversed using the C<reverse_icons> parameter:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:C1',
#NYI         {
#NYI             type          => 'icon_set',
#NYI             icon_style    => '3_arrows',
#NYI             reverse_icons => 1,
#NYI         }
#NYI     );
#NYI 
#NYI The icons can be displayed without the cell value using the C<icons_only> parameter:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:C1',
#NYI         {
#NYI             type         => 'icon_set',
#NYI             icon_style   => '3_flags',
#NYI             icons_only   => 1,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 min_type, mid_type, max_type
#NYI 
#NYI The C<min_type> and C<max_type> properties are available when the conditional formatting type is C<2_color_scale>, C<3_color_scale> or C<data_bar>. The C<mid_type> is available for C<3_color_scale>. The properties are used as follows:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A12',
#NYI         {
#NYI             type      => '2_color_scale',
#NYI             min_type  => 'percent',
#NYI             max_type  => 'percent',
#NYI         }
#NYI     );
#NYI 
#NYI The available min/mid/max types are:
#NYI 
#NYI     min        (for min_type only)
#NYI     num
#NYI     percent
#NYI     percentile
#NYI     formula
#NYI     max        (for max_type only)
#NYI 
#NYI 
#NYI =head2 min_value, mid_value, max_value
#NYI 
#NYI The C<min_value> and C<max_value> properties are available when the conditional formatting type is C<2_color_scale>, C<3_color_scale> or C<data_bar>. The C<mid_value> is available for C<3_color_scale>. The properties are used as follows:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A12',
#NYI         {
#NYI             type       => '2_color_scale',
#NYI             min_value  => 10,
#NYI             max_value  => 90,
#NYI         }
#NYI     );
#NYI 
#NYI =head2 min_color, mid_color,  max_color, bar_color
#NYI 
#NYI The C<min_color> and C<max_color> properties are available when the conditional formatting type is C<2_color_scale>, C<3_color_scale> or C<data_bar>. The C<mid_color> is available for C<3_color_scale>. The properties are used as follows:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:A12',
#NYI         {
#NYI             type      => '2_color_scale',
#NYI             min_color => "#C5D9F1",
#NYI             max_color => "#538ED5",
#NYI         }
#NYI     );
#NYI 
#NYI The color can be specifies as an Excel::Writer::XLSX color index or, more usefully, as a HTML style RGB hex number, as shown above.
#NYI 
#NYI 
#NYI =head2 stop_if_true
#NYI 
#NYI The C<stop_if_true> parameter, if set to a true value, will enable the "stop if true" feature on the conditional formatting rule, so that subsequent rules are not examined for any cell on which the conditions for this rule are met.
#NYI 
#NYI 
#NYI =head2 Conditional Formatting Examples
#NYI 
#NYI Example 1. Highlight cells greater than an integer value.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => 'greater than',
#NYI             value    => 5,
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 2. Highlight cells greater than a value in a reference cell.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => 'greater than',
#NYI             value    => '$H$1',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 3. Highlight cells greater than a certain date:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'date',
#NYI             criteria => 'greater than',
#NYI             value    => '2011-01-01T',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 4. Highlight cells with a date in the last seven days:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'time_period',
#NYI             criteria => 'last 7 days',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 5. Highlight cells with strings starting with the letter C<b>:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'text',
#NYI             criteria => 'begins with',
#NYI             value    => 'b',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 6. Highlight cells that are 1 std deviation above the average for the range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'average',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 7. Highlight duplicate cells in a range:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'duplicate',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 8. Highlight unique cells in a range.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'unique',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 9. Highlight the top 10 cells.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'top',
#NYI             value    => 10,
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI Example 10. Highlight blank cells.
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:F10',
#NYI         {
#NYI             type     => 'blanks',
#NYI             format   => $format,
#NYI         }
#NYI     );
#NYI 
#NYI Example 11. Set traffic light icons in 3 cells:
#NYI 
#NYI     $worksheet->conditional_formatting( 'A1:C1',
#NYI         {
#NYI             type         => 'icon_set',
#NYI             icon_style   => '3_traffic_lights',
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI See also the C<conditional_format.pl> example program in C<EXAMPLES>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 SPARKLINES IN EXCEL
#NYI 
#NYI Sparklines are a feature of Excel 2010+ which allows you to add small charts to worksheet cells. These are useful for showing visual trends in data in a compact format.
#NYI 
#NYI In Excel::Writer::XLSX Sparklines can be added to cells using the C<add_sparkline()> worksheet method:
#NYI 
#NYI     $worksheet->add_sparkline(
#NYI         {
#NYI             location => 'F2',
#NYI             range    => 'Sheet1!A2:E2',
#NYI             type     => 'column',
#NYI             style    => 12,
#NYI         }
#NYI     );
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/sparklines1.jpg" alt="Sparklines example."/></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI B<Note:> Sparklines are a feature of Excel 2010+ only. You can write them to an XLSX file that can be read by Excel 2007 but they won't be displayed.
#NYI 
#NYI 
#NYI =head2 add_sparkline( { parameter => 'value', ... } )
#NYI 
#NYI The C<add_sparkline()> worksheet method is used to add sparklines to a cell or a range of cells.
#NYI 
#NYI The parameters to C<add_sparkline()> must be passed in a hash ref. The main sparkline parameters are:
#NYI 
#NYI     location        (required)
#NYI     range           (required)
#NYI     type
#NYI     style
#NYI 
#NYI     markers
#NYI     negative_points
#NYI     axis
#NYI     reverse
#NYI 
#NYI Other, less commonly used parameters are:
#NYI 
#NYI     high_point
#NYI     low_point
#NYI     first_point
#NYI     last_point
#NYI     max
#NYI     min
#NYI     empty_cells
#NYI     show_hidden
#NYI     date_axis
#NYI     weight
#NYI 
#NYI     series_color
#NYI     negative_color
#NYI     markers_color
#NYI     first_color
#NYI     last_color
#NYI     high_color
#NYI     low_color
#NYI 
#NYI These parameters are explained in the sections below:
#NYI 
#NYI =head2 location
#NYI 
#NYI This is the cell where the sparkline will be displayed:
#NYI 
#NYI     location => 'F1'
#NYI 
#NYI The C<location> should be a single cell. (For multiple cells see L<Grouped Sparklines> below).
#NYI 
#NYI To specify the location in row-column notation use the C<xl_rowcol_to_cell()> function from the L<Excel::Writer::XLSX::Utility> module.
#NYI 
#NYI     use Excel::Writer::XLSX::Utility ':rowcol';
#NYI     ...
#NYI     location => xl_rowcol_to_cell( 0, 5 ), # F1
#NYI 
#NYI 
#NYI =head2 range
#NYI 
#NYI This specifies the cell data range that the sparkline will plot:
#NYI 
#NYI     $worksheet->add_sparkline(
#NYI         {
#NYI             location => 'F1',
#NYI             range    => 'A1:E1',
#NYI         }
#NYI     );
#NYI 
#NYI The C<range> should be a 2D array. (For 3D arrays of cells see L<Grouped Sparklines> below).
#NYI 
#NYI If C<range> is not on the same worksheet you can specify its location using the usual Excel notation:
#NYI 
#NYI             range => 'Sheet1!A1:E1',
#NYI 
#NYI If the worksheet contains spaces or special characters you should quote the worksheet name in the same way that Excel does:
#NYI 
#NYI             range => q('Monthly Data'!A1:E1),
#NYI 
#NYI To specify the location in row-column notation use the C<xl_range()> or C<xl_range_formula()> functions from the L<Excel::Writer::XLSX::Utility> module.
#NYI 
#NYI     use Excel::Writer::XLSX::Utility ':rowcol';
#NYI     ...
#NYI     range => xl_range( 1, 1,  0, 4 ),                   # 'A1:E1'
#NYI     range => xl_range_formula( 'Sheet1', 0, 0,  0, 4 ), # 'Sheet1!A2:E2'
#NYI 
#NYI =head2 type
#NYI 
#NYI Specifies the type of sparkline. There are 3 available sparkline types:
#NYI 
#NYI     line    (default)
#NYI     column
#NYI     win_loss
#NYI 
#NYI For example:
#NYI 
#NYI     {
#NYI         location => 'F1',
#NYI         range    => 'A1:E1',
#NYI         type     => 'column',
#NYI     }
#NYI 
#NYI 
#NYI =head2 style
#NYI 
#NYI Excel provides 36 built-in Sparkline styles in 6 groups of 6. The C<style> parameter can be used to replicate these and should be a corresponding number from 1 .. 36.
#NYI 
#NYI     {
#NYI         location => 'A14',
#NYI         range    => 'Sheet2!A2:J2',
#NYI         style    => 3,
#NYI     }
#NYI 
#NYI The style number starts in the top left of the style grid and runs left to right. The default style is 1. It is possible to override colour elements of the sparklines using the C<*_color> parameters below.
#NYI 
#NYI =head2 markers
#NYI 
#NYI Turn on the markers for C<line> style sparklines.
#NYI 
#NYI     {
#NYI         location => 'A6',
#NYI         range    => 'Sheet2!A1:J1',
#NYI         markers  => 1,
#NYI     }
#NYI 
#NYI Markers aren't shown in Excel for C<column> and C<win_loss> sparklines.
#NYI 
#NYI =head2 negative_points
#NYI 
#NYI Highlight negative values in a sparkline range. This is usually required with C<win_loss> sparklines.
#NYI 
#NYI     {
#NYI         location        => 'A21',
#NYI         range           => 'Sheet2!A3:J3',
#NYI         type            => 'win_loss',
#NYI         negative_points => 1,
#NYI     }
#NYI 
#NYI =head2 axis
#NYI 
#NYI Display a horizontal axis in the sparkline:
#NYI 
#NYI     {
#NYI         location => 'A10',
#NYI         range    => 'Sheet2!A1:J1',
#NYI         axis     => 1,
#NYI     }
#NYI 
#NYI =head2 reverse
#NYI 
#NYI Plot the data from right-to-left instead of the default left-to-right:
#NYI 
#NYI     {
#NYI         location => 'A24',
#NYI         range    => 'Sheet2!A4:J4',
#NYI         type     => 'column',
#NYI         reverse  => 1,
#NYI     }
#NYI 
#NYI =head2 weight
#NYI 
#NYI Adjust the default line weight (thickness) for C<line> style sparklines.
#NYI 
#NYI      weight => 0.25,
#NYI 
#NYI The weight value should be one of the following values allowed by Excel:
#NYI 
#NYI     0.25  0.5   0.75
#NYI     1     1.25
#NYI     2.25
#NYI     3
#NYI     4.25
#NYI     6
#NYI 
#NYI =head2 high_point, low_point, first_point, last_point
#NYI 
#NYI Highlight points in a sparkline range.
#NYI 
#NYI         high_point  => 1,
#NYI         low_point   => 1,
#NYI         first_point => 1,
#NYI         last_point  => 1,
#NYI 
#NYI 
#NYI =head2 max, min
#NYI 
#NYI Specify the maximum and minimum vertical axis values:
#NYI 
#NYI         max         => 0.5,
#NYI         min         => -0.5,
#NYI 
#NYI As a special case you can set the maximum and minimum to be for a group of sparklines rather than one:
#NYI 
#NYI         max         => 'group',
#NYI 
#NYI See L<Grouped Sparklines> below.
#NYI 
#NYI =head2 empty_cells
#NYI 
#NYI Define how empty cells are handled in a sparkline.
#NYI 
#NYI     empty_cells => 'zero',
#NYI 
#NYI The available options are:
#NYI 
#NYI     gaps   : show empty cells as gaps (the default).
#NYI     zero   : plot empty cells as 0.
#NYI     connect: Connect points with a line ("line" type  sparklines only).
#NYI 
#NYI =head2 show_hidden
#NYI 
#NYI Plot data in hidden rows and columns:
#NYI 
#NYI     show_hidden => 1,
#NYI 
#NYI Note, this option is off by default.
#NYI 
#NYI =head2 date_axis
#NYI 
#NYI Specify an alternative date axis for the sparkline. This is useful if the data being plotted isn't at fixed width intervals:
#NYI 
#NYI     {
#NYI         location  => 'F3',
#NYI         range     => 'A3:E3',
#NYI         date_axis => 'A4:E4',
#NYI     }
#NYI 
#NYI The number of cells in the date range should correspond to the number of cells in the data range.
#NYI 
#NYI 
#NYI =head2 series_color
#NYI 
#NYI It is possible to override the colour of a sparkline style using the following parameters:
#NYI 
#NYI     series_color
#NYI     negative_color
#NYI     markers_color
#NYI     first_color
#NYI     last_color
#NYI     high_color
#NYI     low_color
#NYI 
#NYI The color should be specified as a HTML style C<#rrggbb> hex value:
#NYI 
#NYI     {
#NYI         location     => 'A18',
#NYI         range        => 'Sheet2!A2:J2',
#NYI         type         => 'column',
#NYI         series_color => '#E965E0',
#NYI     }
#NYI 
#NYI =head2 Grouped Sparklines
#NYI 
#NYI The C<add_sparkline()> worksheet method can be used multiple times to write as many sparklines as are required in a worksheet.
#NYI 
#NYI However, it is sometimes necessary to group contiguous sparklines so that changes that are applied to one are applied to all. In Excel this is achieved by selecting a 3D range of cells for the data C<range> and a 2D range of cells for the C<location>.
#NYI 
#NYI In Excel::Writer::XLSX, you can simulate this by passing an array refs of values to C<location> and C<range>:
#NYI 
#NYI     {
#NYI         location => [ 'A27',          'A28',          'A29'          ],
#NYI         range    => [ 'Sheet2!A5:J5', 'Sheet2!A6:J6', 'Sheet2!A7:J7' ],
#NYI         markers  => 1,
#NYI     }
#NYI 
#NYI =head2 Sparkline examples
#NYI 
#NYI See the C<sparklines1.pl> and C<sparklines2.pl> example programs in the C<examples> directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 TABLES IN EXCEL
#NYI 
#NYI Tables in Excel are a way of grouping a range of cells into a single entity that has common formatting or that can be referenced from formulas. Tables can have column headers, autofilters, total rows, column formulas and default formatting.
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/tables.jpg" width="640" height="420" alt="Output from tables.pl" /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI For more information see "An Overview of Excel Tables" L<http://office.microsoft.com/en-us/excel-help/overview-of-excel-tables-HA010048546.aspx>.
#NYI 
#NYI Note, tables don't work in Excel::Writer::XLSX when C<set_optimization()> mode in on.
#NYI 
#NYI 
#NYI =head2 add_table( $row1, $col1, $row2, $col2, { parameter => 'value', ... })
#NYI 
#NYI Tables are added to a worksheet using the C<add_table()> method:
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { %parameters } );
#NYI 
#NYI The data range can be specified in 'A1' or 'row/col' notation (see also the note about L</Cell notation> for more information):
#NYI 
#NYI 
#NYI     $worksheet->add_table( 'B3:F7' );
#NYI     # Same as:
#NYI     $worksheet->add_table(  2, 1, 6, 5 );
#NYI 
#NYI The last parameter in C<add_table()> should be a hash ref containing the parameters that describe the table options and data. The available parameters are:
#NYI 
#NYI         data
#NYI         autofilter
#NYI         header_row
#NYI         banded_columns
#NYI         banded_rows
#NYI         first_column
#NYI         last_column
#NYI         style
#NYI         total_row
#NYI         columns
#NYI         name
#NYI 
#NYI The table parameters are detailed below. There are no required parameters and the hash ref isn't required if no options are specified.
#NYI 
#NYI 
#NYI 
#NYI =head2 data
#NYI 
#NYI The C<data> parameter can be used to specify the data in the cells of the table.
#NYI 
#NYI     my $data = [
#NYI         [ 'Apples',  10000, 5000, 8000, 6000 ],
#NYI         [ 'Pears',   2000,  3000, 4000, 5000 ],
#NYI         [ 'Bananas', 6000,  6000, 6500, 6000 ],
#NYI         [ 'Oranges', 500,   300,  200,  700 ],
#NYI 
#NYI     ];
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { data => $data } );
#NYI 
#NYI Table data can also be written separately, as an array or individual cells.
#NYI 
#NYI     # These two statements are the same as the single statement above.
#NYI     $worksheet->add_table( 'B3:F7' );
#NYI     $worksheet->write_col( 'B4', $data );
#NYI 
#NYI Writing the cell data separately is occasionally required when you need to control the C<write_*()> method used to populate the cells or if you wish to tweak the cell formatting.
#NYI 
#NYI The C<data> structure should be an array ref of array refs holding row data as shown above.
#NYI 
#NYI =head2 header_row
#NYI 
#NYI The C<header_row> parameter can be used to turn on or off the header row in the table. It is on by default.
#NYI 
#NYI     $worksheet->add_table( 'B4:F7', { header_row => 0 } ); # Turn header off.
#NYI 
#NYI The header row will contain default captions such as C<Column 1>, C<Column 2>,  etc. These captions can be overridden using the C<columns> parameter below.
#NYI 
#NYI 
#NYI =head2 autofilter
#NYI 
#NYI The C<autofilter> parameter can be used to turn on or off the autofilter in the header row. It is on by default.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { autofilter => 0 } ); # Turn autofilter off.
#NYI 
#NYI The C<autofilter> is only shown if the C<header_row> is on. Filters within the table are not supported.
#NYI 
#NYI 
#NYI =head2 banded_rows
#NYI 
#NYI The C<banded_rows> parameter can be used to used to create rows of alternating colour in the table. It is on by default.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { banded_rows => 0 } );
#NYI 
#NYI 
#NYI =head2 banded_columns
#NYI 
#NYI The C<banded_columns> parameter can be used to used to create columns of alternating colour in the table. It is off by default.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { banded_columns => 1 } );
#NYI 
#NYI 
#NYI =head2 first_column
#NYI 
#NYI The C<first_column> parameter can be used to highlight the first column of the table. The type of highlighting will depend on the C<style> of the table. It may be bold text or a different colour. It is off by default.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { first_column => 1 } );
#NYI 
#NYI 
#NYI =head2 last_column
#NYI 
#NYI The C<last_column> parameter can be used to highlight the last column of the table. The type of highlighting will depend on the C<style> of the table. It may be bold text or a different colour. It is off by default.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { last_column => 1 } );
#NYI 
#NYI 
#NYI =head2 style
#NYI 
#NYI The C<style> parameter can be used to set the style of the table. Standard Excel table format names should be used (with matching capitalisation):
#NYI 
#NYI     $worksheet11->add_table(
#NYI         'B3:F7',
#NYI         {
#NYI             data      => $data,
#NYI             style     => 'Table Style Light 11',
#NYI         }
#NYI     );
#NYI 
#NYI The default table style is 'Table Style Medium 9'.
#NYI 
#NYI 
#NYI =head2 name
#NYI 
#NYI By default tables are named C<Table1>, C<Table2>, etc. The C<name> parameter can be used to set the name of the table:
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { name => 'SalesData' } );
#NYI 
#NYI If you override the table name you must ensure that it doesn't clash with an existing table name and that it follows Excel's requirements for table names L<http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx#BMsyntax_rules_for_names>.
#NYI 
#NYI If you need to know the name of the table, for example to use it in a formula, you can get it as follows:
#NYI 
#NYI     my $table      = $worksheet2->add_table( 'B3:F7' );
#NYI     my $table_name = $table->{_name};
#NYI 
#NYI 
#NYI =head2 total_row
#NYI 
#NYI The C<total_row> parameter can be used to turn on the total row in the last row of a table. It is distinguished from the other rows by a different formatting and also with dropdown C<SUBTOTAL> functions.
#NYI 
#NYI     $worksheet->add_table( 'B3:F7', { total_row => 1 } );
#NYI 
#NYI The default total row doesn't have any captions or functions. These must by specified via the C<columns> parameter below.
#NYI 
#NYI =head2 columns
#NYI 
#NYI The C<columns> parameter can be used to set properties for columns within the table.
#NYI 
#NYI The sub-properties that can be set are:
#NYI 
#NYI     header
#NYI     formula
#NYI     total_string
#NYI     total_function
#NYI     total_value
#NYI     format
#NYI     header_format
#NYI 
#NYI The column data must be specified as an array ref of hash refs. For example to override the default 'Column n' style table headers:
#NYI 
#NYI     $worksheet->add_table(
#NYI         'B3:F7',
#NYI         {
#NYI             data    => $data,
#NYI             columns => [
#NYI                 { header => 'Product' },
#NYI                 { header => 'Quarter 1' },
#NYI                 { header => 'Quarter 2' },
#NYI                 { header => 'Quarter 3' },
#NYI                 { header => 'Quarter 4' },
#NYI             ]
#NYI         }
#NYI     );
#NYI 
#NYI If you don't wish to specify properties for a specific column you pass an empty hash ref and the defaults will be applied:
#NYI 
#NYI             ...
#NYI             columns => [
#NYI                 { header => 'Product' },
#NYI                 { header => 'Quarter 1' },
#NYI                 { },                        # Defaults to 'Column 3'.
#NYI                 { header => 'Quarter 3' },
#NYI                 { header => 'Quarter 4' },
#NYI             ]
#NYI             ...
#NYI 
#NYI 
#NYI Column formulas can by applied using the C<formula> column property:
#NYI 
#NYI     $worksheet8->add_table(
#NYI         'B3:G7',
#NYI         {
#NYI             data    => $data,
#NYI             columns => [
#NYI                 { header => 'Product' },
#NYI                 { header => 'Quarter 1' },
#NYI                 { header => 'Quarter 2' },
#NYI                 { header => 'Quarter 3' },
#NYI                 { header => 'Quarter 4' },
#NYI                 {
#NYI                     header  => 'Year',
#NYI                     formula => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])'
#NYI                 },
#NYI             ]
#NYI         }
#NYI     );
#NYI 
#NYI The Excel 2007 C<[#This Row]> and Excel 2010 C<@> structural references are supported within the formula.
#NYI 
#NYI As stated above the C<total_row> table parameter turns on the "Total" row in the table but it doesn't populate it with any defaults. Total captions and functions must be specified via the C<columns> property and the C<total_string>, C<total_function> and C<total_value> sub properties:
#NYI 
#NYI     $worksheet10->add_table(
#NYI         'B3:F8',
#NYI         {
#NYI             data      => $data,
#NYI             total_row => 1,
#NYI             columns   => [
#NYI                 { header => 'Product',   total_string   => 'Totals' },
#NYI                 { header => 'Quarter 1', total_function => 'sum' },
#NYI                 { header => 'Quarter 2', total_function => 'sum' },
#NYI                 { header => 'Quarter 3', total_function => 'sum' },
#NYI                 { header => 'Quarter 4', total_function => 'sum' },
#NYI             ]
#NYI         }
#NYI     );
#NYI 
#NYI The supported totals row C<SUBTOTAL> functions are:
#NYI 
#NYI         average
#NYI         count_nums
#NYI         count
#NYI         max
#NYI         min
#NYI         std_dev
#NYI         sum
#NYI         var
#NYI 
#NYI User defined functions or formulas aren't supported.
#NYI 
#NYI It is also possible to set a calculated value for the C<total_function> using the C<total_value> sub property. This is only necessary when creating workbooks for applications that cannot calculate the value of formulas automatically. This is similar to setting the C<value> optional property in C<write_formula()>:
#NYI 
#NYI     $worksheet10->add_table(
#NYI         'B3:F8',
#NYI         {
#NYI             data      => $data,
#NYI             total_row => 1,
#NYI             columns   => [
#NYI                 { total_string   => 'Totals' },
#NYI                 { total_function => 'sum', total_value => 100 },
#NYI                 { total_function => 'sum', total_value => 200 },
#NYI                 { total_function => 'sum', total_value => 100 },
#NYI                 { total_function => 'sum', total_value => 400 },
#NYI             ]
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI 
#NYI 
#NYI Formatting can also be applied to columns, to the column data using C<format> and to the header using C<header_format>:
#NYI 
#NYI     my $currency_format = $workbook->add_format( num_format => '$#,##0' );
#NYI 
#NYI     $worksheet->add_table(
#NYI         'B3:D8',
#NYI         {
#NYI             data      => $data,
#NYI             total_row => 1,
#NYI             columns   => [
#NYI                 { header => 'Product', total_string => 'Totals' },
#NYI                 {
#NYI                     header         => 'Quarter 1',
#NYI                     total_function => 'sum',
#NYI                     format         => $currency_format,
#NYI                 },
#NYI                 {
#NYI                     header         => 'Quarter 2',
#NYI                     header_format  => $bold,
#NYI                     total_function => 'sum',
#NYI                     format         => $currency_format,
#NYI                 },
#NYI             ]
#NYI         }
#NYI     );
#NYI 
#NYI Standard Excel::Writer::XLSX format objects can be used. However, they should be limited to numerical formats for the columns and simple formatting like text wrap for the headers. Overriding other table formatting may produce inconsistent results.
#NYI 
#NYI 
#NYI 
#NYI =head1 FORMULAS AND FUNCTIONS IN EXCEL
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Introduction
#NYI 
#NYI The following is a brief introduction to formulas and functions in Excel and Excel::Writer::XLSX.
#NYI 
#NYI A formula is a string that begins with an equals sign:
#NYI 
#NYI     '=A1+B1'
#NYI     '=AVERAGE(1, 2, 3)'
#NYI 
#NYI The formula can contain numbers, strings, boolean values, cell references, cell ranges and functions. Named ranges are not supported. Formulas should be written as they appear in Excel, that is cells and functions must be in uppercase.
#NYI 
#NYI Cells in Excel are referenced using the A1 notation system where the column is designated by a letter and the row by a number. Columns range from A to XFD i.e. 0 to 16384, rows range from 1 to 1048576. The C<Excel::Writer::XLSX::Utility> module that is included in the distro contains helper functions for dealing with A1 notation, for example:
#NYI 
#NYI     use Excel::Writer::XLSX::Utility;
#NYI 
#NYI     ( $row, $col ) = xl_cell_to_rowcol( 'C2' );    # (1, 2)
#NYI     $str = xl_rowcol_to_cell( 1, 2 );              # C2
#NYI 
#NYI The Excel C<$> notation in cell references is also supported. This allows you to specify whether a row or column is relative or absolute. This only has an effect if the cell is copied. The following examples show relative and absolute values.
#NYI 
#NYI     '=A1'   # Column and row are relative
#NYI     '=$A1'  # Column is absolute and row is relative
#NYI     '=A$1'  # Column is relative and row is absolute
#NYI     '=$A$1' # Column and row are absolute
#NYI 
#NYI Formulas can also refer to cells in other worksheets of the current workbook. For example:
#NYI 
#NYI     '=Sheet2!A1'
#NYI     '=Sheet2!A1:A5'
#NYI     '=Sheet2:Sheet3!A1'
#NYI     '=Sheet2:Sheet3!A1:A5'
#NYI     q{='Test Data'!A1}
#NYI     q{='Test Data1:Test Data2'!A1}
#NYI 
#NYI The sheet reference and the cell reference are separated by C<!> the exclamation mark symbol. If worksheet names contain spaces, commas or parentheses then Excel requires that the name is enclosed in single quotes as shown in the last two examples above. In order to avoid using a lot of escape characters you can use the quote operator C<q{}> to protect the quotes. See C<perlop> in the main Perl documentation. Only valid sheet names that have been added using the C<add_worksheet()> method can be used in formulas. You cannot reference external workbooks.
#NYI 
#NYI 
#NYI The following table lists the operators that are available in Excel's formulas. The majority of the operators are the same as Perl's, differences are indicated:
#NYI 
#NYI     Arithmetic operators:
#NYI     =====================
#NYI     Operator  Meaning                   Example
#NYI        +      Addition                  1+2
#NYI        -      Subtraction               2-1
#NYI        *      Multiplication            2*3
#NYI        /      Division                  1/4
#NYI        ^      Exponentiation            2^3      # Equivalent to **
#NYI        -      Unary minus               -(1+2)
#NYI        %      Percent (Not modulus)     13%
#NYI 
#NYI 
#NYI     Comparison operators:
#NYI     =====================
#NYI     Operator  Meaning                   Example
#NYI         =     Equal to                  A1 =  B1 # Equivalent to ==
#NYI         <>    Not equal to              A1 <> B1 # Equivalent to !=
#NYI         >     Greater than              A1 >  B1
#NYI         <     Less than                 A1 <  B1
#NYI         >=    Greater than or equal to  A1 >= B1
#NYI         <=    Less than or equal to     A1 <= B1
#NYI 
#NYI 
#NYI     String operator:
#NYI     ================
#NYI     Operator  Meaning                   Example
#NYI         &     Concatenation             "Hello " & "World!" # [1]
#NYI 
#NYI 
#NYI     Reference operators:
#NYI     ====================
#NYI     Operator  Meaning                   Example
#NYI         :     Range operator            A1:A4               # [2]
#NYI         ,     Union operator            SUM(1, 2+2, B3)     # [3]
#NYI 
#NYI 
#NYI     Notes:
#NYI     [1]: Equivalent to "Hello " . "World!" in Perl.
#NYI     [2]: This range is equivalent to cells A1, A2, A3 and A4.
#NYI     [3]: The comma behaves like the list separator in Perl.
#NYI 
#NYI The range and comma operators can have different symbols in non-English versions of Excel, see below.
#NYI 
#NYI For a general introduction to Excel's formulas and an explanation of the syntax of the function refer to the Excel help files or the following: L<http://office.microsoft.com/en-us/assistance/CH062528031033.aspx>.
#NYI 
#NYI In most cases a formula in Excel can be used directly in the C<write_formula> method. However, there are a few potential issues and differences that the user should be aware of. These are explained in the following sections.
#NYI 
#NYI 
#NYI =head2 Non US Excel functions and syntax
#NYI 
#NYI 
#NYI Excel stores formulas in the format of the US English version, regardless of the language or locale of the end-user's version of Excel. Therefore all formula function names written using Excel::Writer::XLSX must be in English:
#NYI 
#NYI     worksheet->write_formula('A1', '=SUM(1, 2, 3)');   # OK
#NYI     worksheet->write_formula('A2', '=SOMME(1, 2, 3)'); # French. Error on load.
#NYI 
#NYI Also, formulas must be written with the US style separator/range operator which is a comma (not semi-colon). Therefore a formula with multiple values should be written as follows:
#NYI 
#NYI     worksheet->write_formula('A1', '=SUM(1, 2, 3)'); # OK
#NYI     worksheet->write_formula('A2', '=SUM(1; 2; 3)'); # Semi-colon. Error on load.
#NYI 
#NYI If you have a non-English version of Excel you can use the following multi-lingual Formula Translator (L<http://en.excel-translator.de/language/>) to help you convert the formula. It can also replace semi-colons with commas.
#NYI 
#NYI 
#NYI =head2 Formulas added in Excel 2010 and later
#NYI 
#NYI Excel 2010 and later added functions which weren't defined in the original file specification. These functions are referred to by Microsoft as I<future> functions. Examples of these functions are C<ACOT>, C<CHISQ.DIST.RT> , C<CONFIDENCE.NORM>, C<STDEV.P>, C<STDEV.S> and C<WORKDAY.INTL>.
#NYI 
#NYI When written using C<write_formula()> these functions need to be fully qualified with a C<_xlfn.> (or other) prefix as they are shown the list below. For example:
#NYI 
#NYI     worksheet->write_formula('A1', '=_xlfn.STDEV.S(B1:B10)')
#NYI 
#NYI They will appear without the prefix in Excel.
#NYI 
#NYI The following list is taken from the MS XLSX extensions documentation on future functions: L<http://msdn.microsoft.com/en-us/library/dd907480%28v=office.12%29.aspx>:
#NYI 
#NYI     _xlfn.ACOT
#NYI     _xlfn.ACOTH
#NYI     _xlfn.AGGREGATE
#NYI     _xlfn.ARABIC
#NYI     _xlfn.BASE
#NYI     _xlfn.BETA.DIST
#NYI     _xlfn.BETA.INV
#NYI     _xlfn.BINOM.DIST
#NYI     _xlfn.BINOM.DIST.RANGE
#NYI     _xlfn.BINOM.INV
#NYI     _xlfn.BITAND
#NYI     _xlfn.BITLSHIFT
#NYI     _xlfn.BITOR
#NYI     _xlfn.BITRSHIFT
#NYI     _xlfn.BITXOR
#NYI     _xlfn.CEILING.MATH
#NYI     _xlfn.CEILING.PRECISE
#NYI     _xlfn.CHISQ.DIST
#NYI     _xlfn.CHISQ.DIST.RT
#NYI     _xlfn.CHISQ.INV
#NYI     _xlfn.CHISQ.INV.RT
#NYI     _xlfn.CHISQ.TEST
#NYI     _xlfn.COMBINA
#NYI     _xlfn.CONFIDENCE.NORM
#NYI     _xlfn.CONFIDENCE.T
#NYI     _xlfn.COT
#NYI     _xlfn.COTH
#NYI     _xlfn.COVARIANCE.P
#NYI     _xlfn.COVARIANCE.S
#NYI     _xlfn.CSC
#NYI     _xlfn.CSCH
#NYI     _xlfn.DAYS
#NYI     _xlfn.DECIMAL
#NYI     ECMA.CEILING
#NYI     _xlfn.ERF.PRECISE
#NYI     _xlfn.ERFC.PRECISE
#NYI     _xlfn.EXPON.DIST
#NYI     _xlfn.F.DIST
#NYI     _xlfn.F.DIST.RT
#NYI     _xlfn.F.INV
#NYI     _xlfn.F.INV.RT
#NYI     _xlfn.F.TEST
#NYI     _xlfn.FILTERXML
#NYI     _xlfn.FLOOR.MATH
#NYI     _xlfn.FLOOR.PRECISE
#NYI     _xlfn.FORECAST.ETS
#NYI     _xlfn.FORECAST.ETS.CONFINT
#NYI     _xlfn.FORECAST.ETS.SEASONALITY
#NYI     _xlfn.FORECAST.ETS.STAT
#NYI     _xlfn.FORECAST.LINEAR
#NYI     _xlfn.FORMULATEXT
#NYI     _xlfn.GAMMA
#NYI     _xlfn.GAMMA.DIST
#NYI     _xlfn.GAMMA.INV
#NYI     _xlfn.GAMMALN.PRECISE
#NYI     _xlfn.GAUSS
#NYI     _xlfn.HYPGEOM.DIST
#NYI     _xlfn.IFNA
#NYI     _xlfn.IMCOSH
#NYI     _xlfn.IMCOT
#NYI     _xlfn.IMCSC
#NYI     _xlfn.IMCSCH
#NYI     _xlfn.IMSEC
#NYI     _xlfn.IMSECH
#NYI     _xlfn.IMSINH
#NYI     _xlfn.IMTAN
#NYI     _xlfn.ISFORMULA
#NYI     ISO.CEILING
#NYI     _xlfn.ISOWEEKNUM
#NYI     _xlfn.LOGNORM.DIST
#NYI     _xlfn.LOGNORM.INV
#NYI     _xlfn.MODE.MULT
#NYI     _xlfn.MODE.SNGL
#NYI     _xlfn.MUNIT
#NYI     _xlfn.NEGBINOM.DIST
#NYI     NETWORKDAYS.INTL
#NYI     _xlfn.NORM.DIST
#NYI     _xlfn.NORM.INV
#NYI     _xlfn.NORM.S.DIST
#NYI     _xlfn.NORM.S.INV
#NYI     _xlfn.NUMBERVALUE
#NYI     _xlfn.PDURATION
#NYI     _xlfn.PERCENTILE.EXC
#NYI     _xlfn.PERCENTILE.INC
#NYI     _xlfn.PERCENTRANK.EXC
#NYI     _xlfn.PERCENTRANK.INC
#NYI     _xlfn.PERMUTATIONA
#NYI     _xlfn.PHI
#NYI     _xlfn.POISSON.DIST
#NYI     _xlfn.QUARTILE.EXC
#NYI     _xlfn.QUARTILE.INC
#NYI     _xlfn.QUERYSTRING
#NYI     _xlfn.RANK.AVG
#NYI     _xlfn.RANK.EQ
#NYI     _xlfn.RRI
#NYI     _xlfn.SEC
#NYI     _xlfn.SECH
#NYI     _xlfn.SHEET
#NYI     _xlfn.SHEETS
#NYI     _xlfn.SKEW.P
#NYI     _xlfn.STDEV.P
#NYI     _xlfn.STDEV.S
#NYI     _xlfn.T.DIST
#NYI     _xlfn.T.DIST.2T
#NYI     _xlfn.T.DIST.RT
#NYI     _xlfn.T.INV
#NYI     _xlfn.T.INV.2T
#NYI     _xlfn.T.TEST
#NYI     _xlfn.UNICHAR
#NYI     _xlfn.UNICODE
#NYI     _xlfn.VAR.P
#NYI     _xlfn.VAR.S
#NYI     _xlfn.WEBSERVICE
#NYI     _xlfn.WEIBULL.DIST
#NYI     WORKDAY.INTL
#NYI     _xlfn.XOR
#NYI     _xlfn.Z.TEST
#NYI 
#NYI 
#NYI =head2 Using Tables in Formulas
#NYI 
#NYI Worksheet tables can be added with Excel::Writer::XLSX using the C<add_table()> method:
#NYI 
#NYI     worksheet->add_table('B3:F7', {options});
#NYI 
#NYI By default tables are named C<Table1>, C<Table2>, etc., in the order that they are added. However it can also be set by the user using the C<name> parameter:
#NYI 
#NYI     worksheet->add_table('B3:F7', {'name': 'SalesData'});
#NYI 
#NYI If you need to know the name of the table, for example to use it in a formula,
#NYI you can get it as follows:
#NYI 
#NYI     table = worksheet->add_table('B3:F7');
#NYI     table_name = table->{_name};
#NYI 
#NYI When used in a formula a table name such as C<TableX> should be referred to as C<TableX[]> (like a Perl array):
#NYI 
#NYI     worksheet->write_formula('A5', '=VLOOKUP("Sales", Table1[], 2, FALSE');
#NYI 
#NYI 
#NYI =head2 Dealing with #NAME? errors
#NYI 
#NYI If there is an error in the syntax of a formula it is usually displayed in
#NYI Excel as C<#NAME?>. If you encounter an error like this you can debug it as
#NYI follows:
#NYI 
#NYI =over
#NYI 
#NYI =item 1. Ensure the formula is valid in Excel by copying and pasting it into a cell. Note, this should be done in Excel and not other applications such as OpenOffice or LibreOffice since they may have slightly different syntax.
#NYI 
#NYI =item 2. Ensure the formula is using comma separators instead of semi-colons, see L<Non US Excel functions and syntax> above.
#NYI 
#NYI =item 3. Ensure the formula is in English, see L<Non US Excel functions and syntax> above.
#NYI 
#NYI =item 4. Ensure that the formula doesn't contain an Excel 2010+ future function as listed in L<Formulas added in Excel 2010 and later> above. If it does then ensure that the correct prefix is used.
#NYI 
#NYI =back
#NYI 
#NYI Finally if you have completed all the previous steps and still get a C<#NAME?> error you can examine a valid Excel file to see what the correct syntax should be. To do this you should create a valid formula in Excel and save the file. You can then examine the XML in the unzipped file.
#NYI 
#NYI The following shows how to do that using Linux C<unzip> and libxml's xmllint
#NYI L<http://xmlsoft.org/xmllint.html> to format the XML for clarity:
#NYI 
#NYI     $ unzip myfile.xlsx -d myfile
#NYI     $ xmllint --format myfile/xl/worksheets/sheet1.xml | grep '<f>'
#NYI 
#NYI             <f>SUM(1, 2, 3)</f>
#NYI 
#NYI 
#NYI =head2 Formula Results
#NYI 
#NYI Excel::Writer::XLSX doesn't calculate the result of a formula and instead stores the value 0 as the formula result. It then sets a global flag in the XLSX file to say that all formulas and functions should be recalculated when the file is opened.
#NYI 
#NYI This is the method recommended in the Excel documentation and in general it works fine with spreadsheet applications. However, applications that don't have a facility to calculate formulas will only display the 0 results. Examples of such applications are Excel Viewer, PDF Converters, and some mobile device applications.
#NYI 
#NYI If required, it is also possible to specify the calculated result of the
#NYI formula using the optional last C<value> parameter in C<write_formula>:
#NYI 
#NYI     worksheet->write_formula('A1', '=2+2', num_format, 4);
#NYI 
#NYI The C<value> parameter can be a number, a string, a boolean sting (C<'TRUE'> or C<'FALSE'>) or one of the following Excel error codes:
#NYI 
#NYI     #DIV/0!
#NYI     #N/A
#NYI     #NAME?
#NYI     #NULL!
#NYI     #NUM!
#NYI     #REF!
#NYI     #VALUE!
#NYI 
#NYI It is also possible to specify the calculated result of an array formula created with C<write_array_formula>:
#NYI 
#NYI     # Specify the result for a single cell range.
#NYI     worksheet->write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}', format, 2005);
#NYI 
#NYI However, using this parameter only writes a single value to the upper left cell in the result array. For a multi-cell array formula where the results are required, the other result values can be specified by using C<write_number()> to write to the appropriate cell:
#NYI 
#NYI     # Specify the results for a multi cell range.
#NYI     worksheet->write_array_formula('A1:A3', '{=TREND(C1:C3,B1:B3)}', format, 15);
#NYI     worksheet->write_number('A2', 12, format);
#NYI     worksheet->write_number('A3', 14, format);
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 WORKING WITH VBA MACROS
#NYI 
#NYI An Excel C<xlsm> file is exactly the same as a C<xlsx> file except that is includes an additional C<vbaProject.bin> file which contains functions and/or macros. Excel uses a different extension to differentiate between the two file formats since files containing macros are usually subject to additional security checks.
#NYI 
#NYI The C<vbaProject.bin> file is a binary OLE COM container. This was the format used in older C<xls> versions of Excel prior to Excel 2007. Unlike all of the other components of an xlsx/xlsm file the data isn't stored in XML format. Instead the functions and macros as stored as pre-parsed binary format. As such it wouldn't be feasible to define macros and create a C<vbaProject.bin> file from scratch (at least not in the remaining lifespan and interest levels of the author).
#NYI 
#NYI Instead a workaround is used to extract C<vbaProject.bin> files from existing xlsm files and then add these to Excel::Writer::XLSX files.
#NYI 
#NYI 
#NYI =head2 The extract_vba utility
#NYI 
#NYI The C<extract_vba> utility is used to extract the C<vbaProject.bin> binary from an Excel 2007+ xlsm file. The utility is included in the Excel::Writer::XLSX bin directory and is also installed as a standalone executable file:
#NYI 
#NYI     $ extract_vba macro_file.xlsm
#NYI     Extracted: vbaProject.bin
#NYI 
#NYI 
#NYI =head2 Adding the VBA macros to a Excel::Writer::XLSX file
#NYI 
#NYI Once the C<vbaProject.bin> file has been extracted it can be added to the Excel::Writer::XLSX workbook using the C<add_vba_project()> method:
#NYI 
#NYI     $workbook->add_vba_project( './vbaProject.bin' );
#NYI 
#NYI If the VBA file contains functions you can then refer to them in calculations using C<write_formula>:
#NYI 
#NYI     $worksheet->write_formula( 'A1', '=MyMortgageCalc(200000, 25)' );
#NYI 
#NYI Excel files that contain functions and macros should use an C<xlsm> extension or else Excel will complain and possibly not open the file:
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'file.xlsm' );
#NYI 
#NYI It is also possible to assign a macro to a button that is inserted into a
#NYI worksheet using the C<insert_button()> method:
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'file.xlsm' );
#NYI     ...
#NYI     $workbook->add_vba_project( './vbaProject.bin' );
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'my_macro' } );
#NYI 
#NYI 
#NYI It may be necessary to specify a more explicit macro name prefixed by the workbook VBA name as follows:
#NYI 
#NYI     $worksheet->insert_button( 'C2', { macro => 'ThisWorkbook.my_macro' } );
#NYI 
#NYI See the C<macros.pl> from the examples directory for a working example.
#NYI 
#NYI Note: Button is the only VBA Control supported by Excel::Writer::XLSX. Due to the large effort in implementation (1+ man months) it is unlikely that any other form elements will be added in the future.
#NYI 
#NYI 
#NYI =head2 Setting the VBA codenames
#NYI 
#NYI VBA macros generally refer to workbook and worksheet objects. If the VBA codenames aren't specified then Excel::Writer::XLSX will use the Excel defaults of C<ThisWorkbook> and C<Sheet1>, C<Sheet2> etc.
#NYI 
#NYI If the macro uses other codenames you can set them using the workbook and worksheet C<set_vba_name()> methods as follows:
#NYI 
#NYI       $workbook->set_vba_name( 'MyWorkbook' );
#NYI       $worksheet->set_vba_name( 'MySheet' );
#NYI 
#NYI You can find the names that are used in the VBA editor or by unzipping the C<xlsm> file and grepping the files. The following shows how to do that using libxml's xmllint L<http://xmlsoft.org/xmllint.html> to format the XML for clarity:
#NYI 
#NYI     $ unzip myfile.xlsm -d myfile
#NYI     $ xmllint --format `find myfile -name "*.xml" | xargs` | grep "Pr.*codeName"
#NYI 
#NYI       <workbookPr codeName="MyWorkbook" defaultThemeVersion="124226"/>
#NYI       <sheetPr codeName="MySheet"/>
#NYI 
#NYI 
#NYI Note: This step is particularly important for macros created with non-English versions of Excel.
#NYI 
#NYI 
#NYI 
#NYI =head2 What to do if it doesn't work
#NYI 
#NYI This feature should be considered experimental and there is no guarantee that it will work in all cases. Some effort may be required and some knowledge of VBA will certainly help. If things don't work out here are some things to try:
#NYI 
#NYI =over
#NYI 
#NYI =item *
#NYI 
#NYI Start with a simple macro file, ensure that it works and then add complexity.
#NYI 
#NYI =item *
#NYI 
#NYI Try to extract the macros from an Excel 2007 file. The method should work with macros from later versions (it was also tested with Excel 2010 macros). However there may be features in the macro files of more recent version of Excel that aren't backward compatible.
#NYI 
#NYI =item *
#NYI 
#NYI Check the code names that macros use to refer to the workbook and worksheets (see the previous section above). In general VBA uses a code name of C<ThisWorkbook> to refer to the current workbook and the sheet name (such as C<Sheet1>) to refer to the worksheets. These are the defaults used by Excel::Writer::XLSX. If the macro uses other names then you can specify these using the workbook and worksheet C<set_vba_name()> methods:
#NYI 
#NYI       $workbook>set_vba_name( 'MyWorkbook' );
#NYI       $worksheet->set_vba_name( 'MySheet' );
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI =head1 EXAMPLES
#NYI 
#NYI See L<Excel::Writer::XLSX::Examples> for a full list of examples.
#NYI 
#NYI 
#NYI =head2 Example 1
#NYI 
#NYI The following example shows some of the basic features of Excel::Writer::XLSX.
#NYI 
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     # Create a new workbook called simple.xlsx and add a worksheet
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'simple.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # The general syntax is write($row, $column, $token). Note that row and
#NYI     # column are zero indexed
#NYI 
#NYI     # Write some text
#NYI     $worksheet->write( 0, 0, 'Hi Excel!' );
#NYI 
#NYI 
#NYI     # Write some numbers
#NYI     $worksheet->write( 2, 0, 3 );
#NYI     $worksheet->write( 3, 0, 3.00000 );
#NYI     $worksheet->write( 4, 0, 3.00001 );
#NYI     $worksheet->write( 5, 0, 3.14159 );
#NYI 
#NYI 
#NYI     # Write some formulas
#NYI     $worksheet->write( 7, 0, '=A3 + A6' );
#NYI     $worksheet->write( 8, 0, '=IF(A5>3,"Yes", "No")' );
#NYI 
#NYI 
#NYI     # Write a hyperlink
#NYI     my $hyperlink_format = $workbook->add_format(
#NYI         color     => 'blue',
#NYI         underline => 1,
#NYI     );
#NYI 
#NYI     $worksheet->write( 10, 0, 'http://www.perl.com/', $hyperlink_format );
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/a_simple.jpg" width="640" height="420" alt="Output from a_simple.pl" /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Example 2
#NYI 
#NYI The following is a general example which demonstrates some features of working with multiple worksheets.
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     # Create a new Excel workbook
#NYI     my $workbook = Excel::Writer::XLSX->new( 'regions.xlsx' );
#NYI 
#NYI     # Add some worksheets
#NYI     my $north = $workbook->add_worksheet( 'North' );
#NYI     my $south = $workbook->add_worksheet( 'South' );
#NYI     my $east  = $workbook->add_worksheet( 'East' );
#NYI     my $west  = $workbook->add_worksheet( 'West' );
#NYI 
#NYI     # Add a Format
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_bold();
#NYI     $format->set_color( 'blue' );
#NYI 
#NYI     # Add a caption to each worksheet
#NYI     for my $worksheet ( $workbook->sheets() ) {
#NYI         $worksheet->write( 0, 0, 'Sales', $format );
#NYI     }
#NYI 
#NYI     # Write some data
#NYI     $north->write( 0, 1, 200000 );
#NYI     $south->write( 0, 1, 100000 );
#NYI     $east->write( 0, 1, 150000 );
#NYI     $west->write( 0, 1, 100000 );
#NYI 
#NYI     # Set the active worksheet
#NYI     $south->activate();
#NYI 
#NYI     # Set the width of the first column
#NYI     $south->set_column( 0, 0, 20 );
#NYI 
#NYI     # Set the active cell
#NYI     $south->set_selection( 0, 1 );
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/regions.jpg" width="640" height="420" alt="Output from regions.pl" /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Example 3
#NYI 
#NYI Example of how to add conditional formatting to an Excel::Writer::XLSX file. The example below highlights cells that have a value greater than or equal to 50 in red and cells below that value in green.
#NYI 
#NYI     #!/usr/bin/perl
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'conditional_format.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI 
#NYI     # This example below highlights cells that have a value greater than or
#NYI     # equal to 50 in red and cells below that value in green.
#NYI 
#NYI     # Light red fill with dark red text.
#NYI     my $format1 = $workbook->add_format(
#NYI         bg_color => '#FFC7CE',
#NYI         color    => '#9C0006',
#NYI 
#NYI     );
#NYI 
#NYI     # Green fill with dark green text.
#NYI     my $format2 = $workbook->add_format(
#NYI         bg_color => '#C6EFCE',
#NYI         color    => '#006100',
#NYI 
#NYI     );
#NYI 
#NYI     # Some sample data to run the conditional formatting against.
#NYI     my $data = [
#NYI         [ 34, 72,  38, 30, 75, 48, 75, 66, 84, 86 ],
#NYI         [ 6,  24,  1,  84, 54, 62, 60, 3,  26, 59 ],
#NYI         [ 28, 79,  97, 13, 85, 93, 93, 22, 5,  14 ],
#NYI         [ 27, 71,  40, 17, 18, 79, 90, 93, 29, 47 ],
#NYI         [ 88, 25,  33, 23, 67, 1,  59, 79, 47, 36 ],
#NYI         [ 24, 100, 20, 88, 29, 33, 38, 54, 54, 88 ],
#NYI         [ 6,  57,  88, 28, 10, 26, 37, 7,  41, 48 ],
#NYI         [ 52, 78,  1,  96, 26, 45, 47, 33, 96, 36 ],
#NYI         [ 60, 54,  81, 66, 81, 90, 80, 93, 12, 55 ],
#NYI         [ 70, 5,   46, 14, 71, 19, 66, 36, 41, 21 ],
#NYI     ];
#NYI 
#NYI     my $caption = 'Cells with values >= 50 are in light red. '
#NYI       . 'Values < 50 are in light green';
#NYI 
#NYI     # Write the data.
#NYI     $worksheet->write( 'A1', $caption );
#NYI     $worksheet->write_col( 'B3', $data );
#NYI 
#NYI     # Write a conditional format over a range.
#NYI     $worksheet->conditional_formatting( 'B3:K12',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => '>=',
#NYI             value    => 50,
#NYI             format   => $format1,
#NYI         }
#NYI     );
#NYI 
#NYI     # Write another conditional format over the same range.
#NYI     $worksheet->conditional_formatting( 'B3:K12',
#NYI         {
#NYI             type     => 'cell',
#NYI             criteria => '<',
#NYI             value    => 50,
#NYI             format   => $format2,
#NYI         }
#NYI     );
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/conditional_format.jpg" width="640" height="420" alt="Output from conditional_format.pl" /></center></p>
#NYI 
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Example 4
#NYI 
#NYI The following is a simple example of using functions.
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     # Create a new workbook and add a worksheet
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'stats.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet( 'Test data' );
#NYI 
#NYI     # Set the column width for columns 1
#NYI     $worksheet->set_column( 0, 0, 20 );
#NYI 
#NYI 
#NYI     # Create a format for the headings
#NYI     my $format = $workbook->add_format();
#NYI     $format->set_bold();
#NYI 
#NYI 
#NYI     # Write the sample data
#NYI     $worksheet->write( 0, 0, 'Sample', $format );
#NYI     $worksheet->write( 0, 1, 1 );
#NYI     $worksheet->write( 0, 2, 2 );
#NYI     $worksheet->write( 0, 3, 3 );
#NYI     $worksheet->write( 0, 4, 4 );
#NYI     $worksheet->write( 0, 5, 5 );
#NYI     $worksheet->write( 0, 6, 6 );
#NYI     $worksheet->write( 0, 7, 7 );
#NYI     $worksheet->write( 0, 8, 8 );
#NYI 
#NYI     $worksheet->write( 1, 0, 'Length', $format );
#NYI     $worksheet->write( 1, 1, 25.4 );
#NYI     $worksheet->write( 1, 2, 25.4 );
#NYI     $worksheet->write( 1, 3, 24.8 );
#NYI     $worksheet->write( 1, 4, 25.0 );
#NYI     $worksheet->write( 1, 5, 25.3 );
#NYI     $worksheet->write( 1, 6, 24.9 );
#NYI     $worksheet->write( 1, 7, 25.2 );
#NYI     $worksheet->write( 1, 8, 24.8 );
#NYI 
#NYI     # Write some statistical functions
#NYI     $worksheet->write( 4, 0, 'Count', $format );
#NYI     $worksheet->write( 4, 1, '=COUNT(B1:I1)' );
#NYI 
#NYI     $worksheet->write( 5, 0, 'Sum', $format );
#NYI     $worksheet->write( 5, 1, '=SUM(B2:I2)' );
#NYI 
#NYI     $worksheet->write( 6, 0, 'Average', $format );
#NYI     $worksheet->write( 6, 1, '=AVERAGE(B2:I2)' );
#NYI 
#NYI     $worksheet->write( 7, 0, 'Min', $format );
#NYI     $worksheet->write( 7, 1, '=MIN(B2:I2)' );
#NYI 
#NYI     $worksheet->write( 8, 0, 'Max', $format );
#NYI     $worksheet->write( 8, 1, '=MAX(B2:I2)' );
#NYI 
#NYI     $worksheet->write( 9, 0, 'Standard Deviation', $format );
#NYI     $worksheet->write( 9, 1, '=STDEV(B2:I2)' );
#NYI 
#NYI     $worksheet->write( 10, 0, 'Kurtosis', $format );
#NYI     $worksheet->write( 10, 1, '=KURT(B2:I2)' );
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/stats.jpg" width="640" height="420" alt="Output from stats.pl" /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Example 5
#NYI 
#NYI The following example converts a tab separated file called C<tab.txt> into an Excel file called C<tab.xlsx>.
#NYI 
#NYI     #!/usr/bin/perl -w
#NYI 
#NYI     use strict;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     open( TABFILE, 'tab.txt' ) or die "tab.txt: $!";
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'tab.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # Row and column are zero indexed
#NYI     my $row = 0;
#NYI 
#NYI     while ( <TABFILE> ) {
#NYI         chomp;
#NYI 
#NYI         # Split on single tab
#NYI         my @fields = split( '\t', $_ );
#NYI 
#NYI         my $col = 0;
#NYI         for my $token ( @fields ) {
#NYI             $worksheet->write( $row, $col, $token );
#NYI             $col++;
#NYI         }
#NYI         $row++;
#NYI     }
#NYI 
#NYI 
#NYI NOTE: This is a simple conversion program for illustrative purposes only. For converting a CSV or Tab separated or any other type of delimited text file to Excel I recommend the more rigorous csv2xls program that is part of H.Merijn Brand's L<Text::CSV_XS> module distro.
#NYI 
#NYI See the examples/csv2xls link here: L<http://search.cpan.org/~hmbrand/Text-CSV_XS/MANIFEST>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head2 Additional Examples
#NYI 
#NYI The following is a description of the example files that are provided
#NYI in the standard Excel::Writer::XLSX distribution. They demonstrate the
#NYI different features and options of the module. See L<Excel::Writer::XLSX::Examples> for more details.
#NYI 
#NYI     Getting started
#NYI     ===============
#NYI     a_simple.pl             A simple demo of some of the features.
#NYI     bug_report.pl           A template for submitting bug reports.
#NYI     demo.pl                 A demo of some of the available features.
#NYI     formats.pl              All the available formatting on several worksheets.
#NYI     regions.pl              A simple example of multiple worksheets.
#NYI     stats.pl                Basic formulas and functions.
#NYI 
#NYI 
#NYI     Intermediate
#NYI     ============
#NYI     autofilter.pl           Examples of worksheet autofilters.
#NYI     array_formula.pl        Examples of how to write array formulas.
#NYI     cgi.pl                  A simple CGI program.
#NYI     chart_area.pl           A demo of area style charts.
#NYI     chart_bar.pl            A demo of bar (vertical histogram) style charts.
#NYI     chart_column.pl         A demo of column (histogram) style charts.
#NYI     chart_line.pl           A demo of line style charts.
#NYI     chart_pie.pl            A demo of pie style charts.
#NYI     chart_doughnut.pl       A demo of doughnut style charts.
#NYI     chart_radar.pl          A demo of radar style charts.
#NYI     chart_scatter.pl        A demo of scatter style charts.
#NYI     chart_secondary_axis.pl A demo of a line chart with a secondary axis.
#NYI     chart_combined.pl       A demo of a combined column and line chart.
#NYI     chart_pareto.pl         A demo of a combined Pareto chart.
#NYI     chart_stock.pl          A demo of stock style charts.
#NYI     chart_data_table.pl     A demo of a chart with a data table on the axis.
#NYI     chart_data_tools.pl     A demo of charts with data highlighting options.
#NYI     chart_clustered.pl      A demo of a chart with a clustered axis.
#NYI     chart_styles.pl         A demo of the available chart styles.
#NYI     colors.pl               A demo of the colour palette and named colours.
#NYI     comments1.pl            Add comments to worksheet cells.
#NYI     comments2.pl            Add comments with advanced options.
#NYI     conditional_format.pl   Add conditional formats to a range of cells.
#NYI     data_validate.pl        An example of data validation and dropdown lists.
#NYI     date_time.pl            Write dates and times with write_date_time().
#NYI     defined_name.pl         Example of how to create defined names.
#NYI     diag_border.pl          A simple example of diagonal cell borders.
#NYI     filehandle.pl           Examples of working with filehandles.
#NYI     headers.pl              Examples of worksheet headers and footers.
#NYI     hide_row_col.pl         Example of hiding rows and columns.
#NYI     hide_sheet.pl           Simple example of hiding a worksheet.
#NYI     hyperlink1.pl           Shows how to create web hyperlinks.
#NYI     hyperlink2.pl           Examples of internal and external hyperlinks.
#NYI     indent.pl               An example of cell indentation.
#NYI     macros.pl               An example of adding macros from an existing file.
#NYI     merge1.pl               A simple example of cell merging.
#NYI     merge2.pl               A simple example of cell merging with formatting.
#NYI     merge3.pl               Add hyperlinks to merged cells.
#NYI     merge4.pl               An advanced example of merging with formatting.
#NYI     merge5.pl               An advanced example of merging with formatting.
#NYI     merge6.pl               An example of merging with Unicode strings.
#NYI     mod_perl1.pl            A simple mod_perl 1 program.
#NYI     mod_perl2.pl            A simple mod_perl 2 program.
#NYI     panes.pl                An examples of how to create panes.
#NYI     outline.pl              An example of outlines and grouping.
#NYI     outline_collapsed.pl    An example of collapsed outlines.
#NYI     protection.pl           Example of cell locking and formula hiding.
#NYI     rich_strings.pl         Example of strings with multiple formats.
#NYI     right_to_left.pl        Change default sheet direction to right to left.
#NYI     sales.pl                An example of a simple sales spreadsheet.
#NYI     shape1.pl               Insert shapes in worksheet.
#NYI     shape2.pl               Insert shapes in worksheet. With properties.
#NYI     shape3.pl               Insert shapes in worksheet. Scaled.
#NYI     shape4.pl               Insert shapes in worksheet. With modification.
#NYI     shape5.pl               Insert shapes in worksheet. With connections.
#NYI     shape6.pl               Insert shapes in worksheet. With connections.
#NYI     shape7.pl               Insert shapes in worksheet. One to many connections.
#NYI     shape8.pl               Insert shapes in worksheet. One to many connections.
#NYI     shape_all.pl            Demo of all the available shape and connector types.
#NYI     sparklines1.pl          Simple sparklines demo.
#NYI     sparklines2.pl          Sparklines demo showing formatting options.
#NYI     stats_ext.pl            Same as stats.pl with external references.
#NYI     stocks.pl               Demonstrates conditional formatting.
#NYI     tab_colors.pl           Example of how to set worksheet tab colours.
#NYI     tables.pl               Add Excel tables to a worksheet.
#NYI     write_handler1.pl       Example of extending the write() method. Step 1.
#NYI     write_handler2.pl       Example of extending the write() method. Step 2.
#NYI     write_handler3.pl       Example of extending the write() method. Step 3.
#NYI     write_handler4.pl       Example of extending the write() method. Step 4.
#NYI     write_to_scalar.pl      Example of writing an Excel file to a Perl scalar.
#NYI 
#NYI     Unicode
#NYI     =======
#NYI     unicode_2022_jp.pl      Japanese: ISO-2022-JP.
#NYI     unicode_8859_11.pl      Thai:     ISO-8859_11.
#NYI     unicode_8859_7.pl       Greek:    ISO-8859_7.
#NYI     unicode_big5.pl         Chinese:  BIG5.
#NYI     unicode_cp1251.pl       Russian:  CP1251.
#NYI     unicode_cp1256.pl       Arabic:   CP1256.
#NYI     unicode_cyrillic.pl     Russian:  Cyrillic.
#NYI     unicode_koi8r.pl        Russian:  KOI8-R.
#NYI     unicode_polish_utf8.pl  Polish :  UTF8.
#NYI     unicode_shift_jis.pl    Japanese: Shift JIS.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 LIMITATIONS
#NYI 
#NYI The following limits are imposed by Excel 2007+:
#NYI 
#NYI     Description                             Limit
#NYI     --------------------------------------  ------
#NYI     Maximum number of chars in a string     32,767
#NYI     Maximum number of columns               16,384
#NYI     Maximum number of rows                  1,048,576
#NYI     Maximum chars in a sheet name           31
#NYI     Maximum chars in a header/footer        254
#NYI 
#NYI     Maximum characters in hyperlink url     255
#NYI     Maximum characters in hyperlink anchor  255
#NYI     Maximum number of unique hyperlinks*    65,530
#NYI 
#NYI * Per worksheet. Excel allows a greater number of non-unique hyperlinks if they are contiguous and can be grouped into a single range. This will be supported in a later version of Excel::Writer::XLSX if possible.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 Compatibility with Spreadsheet::WriteExcel
#NYI 
#NYI The C<Excel::Writer::XLSX> module is a drop-in replacement for C<Spreadsheet::WriteExcel>.
#NYI 
#NYI It supports all of the features of Spreadsheet::WriteExcel with some minor differences noted below.
#NYI 
#NYI     Workbook Methods            Support
#NYI     ================            ======
#NYI     new()                       Yes
#NYI     add_worksheet()             Yes
#NYI     add_format()                Yes
#NYI     add_chart()                 Yes
#NYI     add_shape()                 Yes. Not in Spreadsheet::WriteExcel.
#NYI     add_vba_project()           Yes. Not in Spreadsheet::WriteExcel.
#NYI     close()                     Yes
#NYI     set_properties()            Yes
#NYI     define_name()               Yes
#NYI     set_tempdir()               Yes
#NYI     set_custom_color()          Yes
#NYI     sheets()                    Yes
#NYI     set_1904()                  Yes
#NYI     set_optimization()          Yes. Not required in Spreadsheet::WriteExcel.
#NYI     add_chart_ext()             Not supported. Not required in Excel::Writer::XLSX.
#NYI     compatibility_mode()        Deprecated. Not required in Excel::Writer::XLSX.
#NYI     set_codepage()              Deprecated. Not required in Excel::Writer::XLSX.
#NYI 
#NYI 
#NYI     Worksheet Methods           Support
#NYI     =================           =======
#NYI     write()                     Yes
#NYI     write_number()              Yes
#NYI     write_string()              Yes
#NYI     write_rich_string()         Yes. Not in Spreadsheet::WriteExcel.
#NYI     write_blank()               Yes
#NYI     write_row()                 Yes
#NYI     write_col()                 Yes
#NYI     write_date_time()           Yes
#NYI     write_url()                 Yes
#NYI     write_formula()             Yes
#NYI     write_array_formula()       Yes. Not in Spreadsheet::WriteExcel.
#NYI     keep_leading_zeros()        Yes
#NYI     write_comment()             Yes
#NYI     show_comments()             Yes
#NYI     set_comments_author()       Yes
#NYI     add_write_handler()         Yes
#NYI     insert_image()              Yes.
#NYI     insert_chart()              Yes
#NYI     insert_shape()              Yes. Not in Spreadsheet::WriteExcel.
#NYI     insert_button()             Yes. Not in Spreadsheet::WriteExcel.
#NYI     data_validation()           Yes
#NYI     conditional_formatting()    Yes. Not in Spreadsheet::WriteExcel.
#NYI     add_sparkline()             Yes. Not in Spreadsheet::WriteExcel.
#NYI     add_table()                 Yes. Not in Spreadsheet::WriteExcel.
#NYI     get_name()                  Yes
#NYI     activate()                  Yes
#NYI     select()                    Yes
#NYI     hide()                      Yes
#NYI     set_first_sheet()           Yes
#NYI     protect()                   Yes
#NYI     set_selection()             Yes
#NYI     set_row()                   Yes.
#NYI     set_column()                Yes.
#NYI     set_default_row()           Yes. Not in Spreadsheet::WriteExcel.
#NYI     outline_settings()          Yes
#NYI     freeze_panes()              Yes
#NYI     split_panes()               Yes
#NYI     merge_range()               Yes
#NYI     merge_range_type()          Yes. Not in Spreadsheet::WriteExcel.
#NYI     set_zoom()                  Yes
#NYI     right_to_left()             Yes
#NYI     hide_zero()                 Yes
#NYI     set_tab_color()             Yes
#NYI     autofilter()                Yes
#NYI     filter_column()             Yes
#NYI     filter_column_list()        Yes. Not in Spreadsheet::WriteExcel.
#NYI     write_utf16be_string()      Deprecated. Use Perl utf8 strings instead.
#NYI     write_utf16le_string()      Deprecated. Use Perl utf8 strings instead.
#NYI     store_formula()             Deprecated. See docs.
#NYI     repeat_formula()            Deprecated. See docs.
#NYI     write_url_range()           Not supported. Not required in Excel::Writer::XLSX.
#NYI 
#NYI     Page Set-up Methods         Support
#NYI     ===================         =======
#NYI     set_landscape()             Yes
#NYI     set_portrait()              Yes
#NYI     set_page_view()             Yes
#NYI     set_paper()                 Yes
#NYI     center_horizontally()       Yes
#NYI     center_vertically()         Yes
#NYI     set_margins()               Yes
#NYI     set_header()                Yes
#NYI     set_footer()                Yes
#NYI     repeat_rows()               Yes
#NYI     repeat_columns()            Yes
#NYI     hide_gridlines()            Yes
#NYI     print_row_col_headers()     Yes
#NYI     print_area()                Yes
#NYI     print_across()              Yes
#NYI     fit_to_pages()              Yes
#NYI     set_start_page()            Yes
#NYI     set_print_scale()           Yes
#NYI     set_h_pagebreaks()          Yes
#NYI     set_v_pagebreaks()          Yes
#NYI 
#NYI     Format Methods              Support
#NYI     ==============              =======
#NYI     set_font()                  Yes
#NYI     set_size()                  Yes
#NYI     set_color()                 Yes
#NYI     set_bold()                  Yes
#NYI     set_italic()                Yes
#NYI     set_underline()             Yes
#NYI     set_font_strikeout()        Yes
#NYI     set_font_script()           Yes
#NYI     set_font_outline()          Yes
#NYI     set_font_shadow()           Yes
#NYI     set_num_format()            Yes
#NYI     set_locked()                Yes
#NYI     set_hidden()                Yes
#NYI     set_align()                 Yes
#NYI     set_rotation()              Yes
#NYI     set_text_wrap()             Yes
#NYI     set_text_justlast()         Yes
#NYI     set_center_across()         Yes
#NYI     set_indent()                Yes
#NYI     set_shrink()                Yes
#NYI     set_pattern()               Yes
#NYI     set_bg_color()              Yes
#NYI     set_fg_color()              Yes
#NYI     set_border()                Yes
#NYI     set_bottom()                Yes
#NYI     set_top()                   Yes
#NYI     set_left()                  Yes
#NYI     set_right()                 Yes
#NYI     set_border_color()          Yes
#NYI     set_bottom_color()          Yes
#NYI     set_top_color()             Yes
#NYI     set_left_color()            Yes
#NYI     set_right_color()           Yes
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 REQUIREMENTS
#NYI 
#NYI L<http://search.cpan.org/search?dist=Archive-Zip/>.
#NYI 
#NYI Perl 5.8.2.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 SPEED AND MEMORY USAGE
#NYI 
#NYI C<Spreadsheet::WriteExcel> was written to optimise speed and reduce memory usage. However, these design goals meant that it wasn't easy to implement features that many users requested such as writing formatting and data separately.
#NYI 
#NYI As a result C<Excel::Writer::XLSX> takes a different design approach and holds a lot more data in memory so that it is functionally more flexible.
#NYI 
#NYI The effect of this is that Excel::Writer::XLSX is about 30% slower than Spreadsheet::WriteExcel and uses 5 times more memory.
#NYI 
#NYI In addition the extended row and column ranges in Excel 2007+ mean that it is possible to run out of memory creating large files. This was almost never an issue with Spreadsheet::WriteExcel.
#NYI 
#NYI This memory usage can be reduced almost completely by using the Workbook C<set_optimization()> method:
#NYI 
#NYI     $workbook->set_optimization();
#NYI 
#NYI This also gives an increase in performance to within 1-10% of Spreadsheet::WriteExcel, see below.
#NYI 
#NYI The trade-off is that you won't be able to take advantage of any new features that manipulate cell data after it is written. One such feature is Tables.
#NYI 
#NYI 
#NYI =head2 Performance figures
#NYI 
#NYI The performance figures below show execution speed and memory usage for 60 columns x N rows for a 50/50 mixture of strings and numbers. Percentage speeds are relative to Spreadsheet::WriteExcel.
#NYI 
#NYI     Excel::Writer::XLSX
#NYI          Rows  Time (s)    Memory (bytes)  Rel. Time
#NYI           400      0.66         6,586,254       129%
#NYI           800      1.26        13,099,422       125%
#NYI          1600      2.55        26,126,361       123%
#NYI          3200      5.16        52,211,284       125%
#NYI          6400     10.47       104,401,428       128%
#NYI         12800     21.48       208,784,519       131%
#NYI         25600     43.90       417,700,746       126%
#NYI         51200     88.52       835,900,298       126%
#NYI 
#NYI     Excel::Writer::XLSX + set_optimisation()
#NYI          Rows  Time (s)    Memory (bytes)  Rel. Time
#NYI           400      0.70            63,059       135%
#NYI           800      1.10            63,059       110%
#NYI          1600      2.30            63,062       111%
#NYI          3200      4.44            63,062       107%
#NYI          6400      8.91            63,062       109%
#NYI         12800     17.69            63,065       108%
#NYI         25600     35.15            63,065       101%
#NYI         51200     70.67            63,065       101%
#NYI 
#NYI     Spreadsheet::WriteExcel
#NYI          Rows  Time (s)    Memory (bytes)
#NYI           400      0.51         1,265,583
#NYI           800      1.01         2,424,855
#NYI          1600      2.07         4,743,400
#NYI          3200      4.14         9,411,139
#NYI          6400      8.20        18,766,915
#NYI         12800     16.39        37,478,468
#NYI         25600     34.72        75,044,423
#NYI         51200     70.21       150,543,431
#NYI 
#NYI 
#NYI =head1 DOWNLOADING
#NYI 
#NYI The latest version of this module is always available at: L<http://search.cpan.org/search?dist=Excel-Writer-XLSX/>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 INSTALLATION
#NYI 
#NYI The module can be installed using the standard Perl procedure:
#NYI 
#NYI             perl Makefile.PL
#NYI             make
#NYI             make test
#NYI             make install    # You may need to be sudo/root
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 DIAGNOSTICS
#NYI 
#NYI 
#NYI =over 4
#NYI 
#NYI =item Filename required by Excel::Writer::XLSX->new()
#NYI 
#NYI A filename must be given in the constructor.
#NYI 
#NYI =item Can't open filename. It may be in use or protected.
#NYI 
#NYI The file cannot be opened for writing. The directory that you are writing to may be protected or the file may be in use by another program.
#NYI 
#NYI 
#NYI =item Can't call method "XXX" on an undefined value at someprogram.pl.
#NYI 
#NYI On Windows this is usually caused by the file that you are trying to create clashing with a version that is already open and locked by Excel.
#NYI 
#NYI =item The file you are trying to open 'file.xls' is in a different format than specified by the file extension.
#NYI 
#NYI This warning occurs when you create an XLSX file but give it an xls extension.
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 WRITING EXCEL FILES
#NYI 
#NYI Depending on your requirements, background and general sensibilities you may prefer one of the following methods of getting data into Excel:
#NYI 
#NYI =over 4
#NYI 
#NYI =item * Spreadsheet::WriteExcel
#NYI 
#NYI This module is the precursor to Excel::Writer::XLSX and uses the same interface. It produces files in the Excel Biff xls format that was used in Excel versions 97-2003. These files can still be read by Excel 2007 but have some limitations in relation to the number of rows and columns that the format supports.
#NYI 
#NYI L<Spreadsheet::WriteExcel>.
#NYI 
#NYI =item * Win32::OLE module and office automation
#NYI 
#NYI This requires a Windows platform and an installed copy of Excel. This is the most powerful and complete method for interfacing with Excel.
#NYI 
#NYI L<Win32::OLE>
#NYI 
#NYI =item * CSV, comma separated variables or text
#NYI 
#NYI Excel will open and automatically convert files with a C<csv> extension.
#NYI 
#NYI To create CSV files refer to the L<Text::CSV_XS> module.
#NYI 
#NYI 
#NYI =item * DBI with DBD::ADO or DBD::ODBC
#NYI 
#NYI Excel files contain an internal index table that allows them to act like a database file. Using one of the standard Perl database modules you can connect to an Excel file as a database.
#NYI 
#NYI 
#NYI =back
#NYI 
#NYI For other Perl-Excel modules try the following search: L<http://search.cpan.org/search?mode=module&query=excel>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 READING EXCEL FILES
#NYI 
#NYI To read data from Excel files try:
#NYI 
#NYI =over 4
#NYI 
#NYI =item * Spreadsheet::XLSX
#NYI 
#NYI A module for reading formatted or unformatted data form XLSX files.
#NYI 
#NYI L<Spreadsheet::XLSX>
#NYI 
#NYI =item * SimpleXlsx
#NYI 
#NYI A lightweight module for reading data from XLSX files.
#NYI 
#NYI L<SimpleXlsx>
#NYI 
#NYI =item * Spreadsheet::ParseExcel
#NYI 
#NYI This module can read  data from an Excel XLS file but it doesn't support the XLSX format.
#NYI 
#NYI L<Spreadsheet::ParseExcel>
#NYI 
#NYI =item * Win32::OLE module and office automation (reading)
#NYI 
#NYI See above.
#NYI 
#NYI =item * DBI with DBD::ADO or DBD::ODBC.
#NYI 
#NYI See above.
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI For other Perl-Excel modules try the following search: L<http://search.cpan.org/search?mode=module&query=excel>.
#NYI 
#NYI 
#NYI =head1 BUGS
#NYI 
#NYI =over
#NYI 
#NYI =item * Memory usage is very high for large worksheets.
#NYI 
#NYI If you run out of memory creating large worksheets use the C<set_optimization()> method. See L</SPEED AND MEMORY USAGE> for more information.
#NYI 
#NYI =item * Perl packaging programs can't find chart modules.
#NYI 
#NYI When using Excel::Writer::XLSX charts with Perl packagers such as PAR or Cava you should explicitly include the chart that you are trying to create in your C<use> statements. This isn't a bug as such but it might help someone from banging their head off a wall:
#NYI 
#NYI     ...
#NYI     use Excel::Writer::XLSX;
#NYI     use Excel::Writer::XLSX::Chart::Column;
#NYI     ...
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI If you wish to submit a bug report run the C<bug_report.pl> program in the C<examples> directory of the distro.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI The bug tracker is on Github: L<https://github.com/jmcnamara/excel-writer-xlsx/issues>.
#NYI 
#NYI 
#NYI =head1 TO DO
#NYI 
#NYI The roadmap is as follows:
#NYI 
#NYI =over 4
#NYI 
#NYI =item * New separated data/formatting API to allow cells to be formatted after data is added.
#NYI 
#NYI =item * More charting features.
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 REPOSITORY
#NYI 
#NYI The Excel::Writer::XLSX source code in host on github: L<http://github.com/jmcnamara/excel-writer-xlsx>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 MAILING LIST
#NYI 
#NYI There is a Google group for discussing and asking questions about Excel::Writer::XLSX. This is a good place to search to see if your question has been asked before:  L<http://groups.google.com/group/spreadsheet-writeexcel>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 DONATIONS and SPONSORSHIP
#NYI 
#NYI If you'd care to donate to the Excel::Writer::XLSX project or sponsor a new feature, you can do so via PayPal: L<http://tinyurl.com/7ayes>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 SEE ALSO
#NYI 
#NYI Spreadsheet::WriteExcel: L<http://search.cpan.org/dist/Spreadsheet-WriteExcel>.
#NYI 
#NYI Spreadsheet::ParseExcel: L<http://search.cpan.org/dist/Spreadsheet-ParseExcel>.
#NYI 
#NYI Spreadsheet::XLSX: L<http://search.cpan.org/dist/Spreadsheet-XLSX>.
#NYI 
#NYI 
#NYI 
#NYI =head1 ACKNOWLEDGEMENTS
#NYI 
#NYI 
#NYI The following people contributed to the debugging, testing or enhancement of Excel::Writer::XLSX:
#NYI 
#NYI Rob Messer of IntelliSurvey gave me the initial prompt to port Spreadsheet::WriteExcel to the XLSX format. IntelliSurvey (L<http://www.intellisurvey.com>) also sponsored large files optimisations and the charting feature.
#NYI 
#NYI Bariatric Advantage (L<http://www.bariatricadvantage.com>) sponsored work on chart formatting.
#NYI 
#NYI Eric Johnson provided the ability to use secondary axes with charts.  Thanks to Foxtons (L<http://foxtons.co.uk>) for sponsoring this work.
#NYI 
#NYI BuildFax (L<http://www.buildfax.com>) sponsored the Tables feature and the Chart point formatting feature.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 DISCLAIMER OF WARRANTY
#NYI 
#NYI Because this software is licensed free of charge, there is no warranty for the software, to the extent permitted by applicable law. Except when otherwise stated in writing the copyright holders and/or other parties provide the software "as is" without warranty of any kind, either expressed or implied, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and performance of the software is with you. Should the software prove defective, you assume the cost of all necessary servicing, repair, or correction.
#NYI 
#NYI In no event unless required by applicable law or agreed to in writing will any copyright holder, or any other party who may modify and/or redistribute the software as permitted by the above licence, be liable to you for damages, including any general, special, incidental, or consequential damages arising out of the use or inability to use the software (including but not limited to loss of data or data being rendered inaccurate or losses sustained by you or third parties or a failure of the software to operate with any other software), even if such holder or other party has been advised of the possibility of such damages.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 LICENSE
#NYI 
#NYI Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 AUTHOR
#NYI 
#NYI John McNamara jmcnamara@cpan.org
#NYI 
#NYI     Wilderness for miles, eyes so mild and wise
#NYI     Oasis child, born and so wild
#NYI     Don't I know you better than the rest
#NYI     All deception, all deception from you
#NYI 
#NYI     Any way you run, you run before us
#NYI     Black and white horse arching among us
#NYI     Any way you run, you run before us
#NYI     Black and white horse arching among us
#NYI 
#NYI       -- Beach House
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 COPYRIGHT
#NYI 
#NYI Copyright MM-MMXVII, John McNamara.
#NYI 
#NYI All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
