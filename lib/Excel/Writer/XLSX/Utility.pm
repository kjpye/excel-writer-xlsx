#NYI package Excel::Writer::XLSX::Utility;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Utility - Helper functions for Excel::Writer::XLSX.
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
#NYI use Exporter;
#NYI use warnings;
#NYI use autouse 'Date::Calc'  => qw(Delta_DHMS Decode_Date_EU Decode_Date_US);
#NYI use autouse 'Date::Manip' => qw(ParseDate Date_Init);
#NYI 
#NYI our $VERSION = '0.96';
#NYI 
#NYI # Row and column functions
#NYI my @rowcol = qw(
#NYI   xl_rowcol_to_cell
#NYI   xl_cell_to_rowcol
#NYI   xl_col_to_name
#NYI   xl_range
#NYI   xl_range_formula
#NYI   xl_inc_row
#NYI   xl_dec_row
#NYI   xl_inc_col
#NYI   xl_dec_col
#NYI );
#NYI 
#NYI # Date and Time functions
#NYI my @dates = qw(
#NYI   xl_date_list
#NYI   xl_date_1904
#NYI   xl_parse_time
#NYI   xl_parse_date
#NYI   xl_parse_date_init
#NYI   xl_decode_date_EU
#NYI   xl_decode_date_US
#NYI );
#NYI 
#NYI our @ISA         = qw(Exporter);
#NYI our @EXPORT_OK   = ();
#NYI our @EXPORT      = ( @rowcol, @dates, 'quote_sheetname' );
#NYI our %EXPORT_TAGS = (
#NYI     rowcol => \@rowcol,
#NYI     dates  => \@dates
#NYI );
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_rowcol_to_cell($row, $col, $row_absolute, $col_absolute)
#NYI #
#NYI sub xl_rowcol_to_cell {
#NYI 
#NYI     my $row     = $_[0] + 1;          # Change from 0-indexed to 1 indexed.
#NYI     my $col     = $_[1];
#NYI     my $row_abs = $_[2] ? '$' : '';
#NYI     my $col_abs = $_[3] ? '$' : '';
#NYI 
#NYI 
#NYI     my $col_str = xl_col_to_name( $col, $col_abs );
#NYI 
#NYI     return $col_str . $row_abs . $row;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_cell_to_rowcol($string)
#NYI #
#NYI # Returns: ($row, $col, $row_absolute, $col_absolute)
#NYI #
#NYI # The $row_absolute and $col_absolute parameters aren't documented because they
#NYI # mainly used internally and aren't very useful to the user.
#NYI #
#NYI sub xl_cell_to_rowcol {
#NYI 
#NYI     my $cell = shift;
#NYI 
#NYI     return ( 0, 0, 0, 0 ) unless $cell;
#NYI 
#NYI     $cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/;
#NYI 
#NYI     my $col_abs = $1 eq "" ? 0 : 1;
#NYI     my $col     = $2;
#NYI     my $row_abs = $3 eq "" ? 0 : 1;
#NYI     my $row     = $4;
#NYI 
#NYI     # Convert base26 column string to number
#NYI     # All your Base are belong to us.
#NYI     my @chars = split //, $col;
#NYI     my $expn = 0;
#NYI     $col = 0;
#NYI 
#NYI     while ( @chars ) {
#NYI         my $char = pop( @chars );    # LS char first
#NYI         $col += ( ord( $char ) - ord( 'A' ) + 1 ) * ( 26**$expn );
#NYI         $expn++;
#NYI     }
#NYI 
#NYI     # Convert 1-index to zero-index
#NYI     $row--;
#NYI     $col--;
#NYI 
#NYI     return $row, $col, $row_abs, $col_abs;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_col_to_name($col, $col_absolute)
#NYI #
#NYI sub xl_col_to_name {
#NYI 
#NYI     my $col     = $_[0];
#NYI     my $col_abs = $_[1] ? '$' : '';
#NYI     my $col_str = '';
#NYI 
#NYI     # Change from 0-indexed to 1 indexed.
#NYI     $col++;
#NYI 
#NYI     while ( $col ) {
#NYI 
#NYI         # Set remainder from 1 .. 26
#NYI         my $remainder = $col % 26 || 26;
#NYI 
#NYI         # Convert the $remainder to a character. C-ishly.
#NYI         my $col_letter = chr( ord( 'A' ) + $remainder - 1 );
#NYI 
#NYI         # Accumulate the column letters, right to left.
#NYI         $col_str = $col_letter . $col_str;
#NYI 
#NYI         # Get the next order of magnitude.
#NYI         $col = int( ( $col - 1 ) / 26 );
#NYI     }
#NYI 
#NYI     return $col_abs . $col_str;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_range($row_1, $row_2, $col_1, $col_2, $row_abs_1, $row_abs_2, $col_abs_1, $col_abs_2)
#NYI #
#NYI sub xl_range {
#NYI 
#NYI     my ( $row_1,     $row_2,     $col_1,     $col_2 )     = @_[ 0 .. 3 ];
#NYI     my ( $row_abs_1, $row_abs_2, $col_abs_1, $col_abs_2 ) = @_[ 4 .. 7 ];
#NYI 
#NYI     my $range1 = xl_rowcol_to_cell( $row_1, $col_1, $row_abs_1, $col_abs_1 );
#NYI     my $range2 = xl_rowcol_to_cell( $row_2, $col_2, $row_abs_2, $col_abs_2 );
#NYI 
#NYI     return $range1 . ':' . $range2;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_range_formula($sheetname, $row_1, $row_2, $col_1, $col_2)
#NYI #
#NYI sub xl_range_formula {
#NYI 
#NYI     my ( $sheetname, $row_1, $row_2, $col_1, $col_2 ) = @_;
#NYI 
#NYI     $sheetname = quote_sheetname( $sheetname );
#NYI 
#NYI     my $range = xl_range( $row_1, $row_2, $col_1, $col_2, 1, 1, 1, 1 );
#NYI 
#NYI     return '=' . $sheetname . '!' . $range
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # quote_sheetname()
#NYI #
#NYI # Sheetnames used in references should be quoted if they contain any spaces,
#NYI # special characters or if they look like something that isn't a sheet name.
#NYI #
#NYI sub quote_sheetname {
#NYI 
#NYI     my $sheetname = $_[0];
#NYI 
#NYI     # Use Excel's conventions and quote the sheet name if it contains any
#NYI     # non-word character or if it isn't already quoted.
#NYI     if ( $sheetname =~ /\W/ && $sheetname !~ /^'/ ) {
#NYI         # Double quote any single quotes.
#NYI         $sheetname =~ s/'/''/g;
#NYI         $sheetname = q(') . $sheetname . q(');
#NYI     }
#NYI 
#NYI     return $sheetname;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_inc_row($string)
#NYI #
#NYI sub xl_inc_row {
#NYI 
#NYI     my $cell = shift;
#NYI     my ( $row, $col, $row_abs, $col_abs ) = xl_cell_to_rowcol( $cell );
#NYI 
#NYI     return xl_rowcol_to_cell( ++$row, $col, $row_abs, $col_abs );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_dec_row($string)
#NYI #
#NYI # Decrements the row number of an Excel cell reference in A1 notation.
#NYI # For example C4 to C3
#NYI #
#NYI # Returns: a cell reference string.
#NYI #
#NYI sub xl_dec_row {
#NYI 
#NYI     my $cell = shift;
#NYI     my ( $row, $col, $row_abs, $col_abs ) = xl_cell_to_rowcol( $cell );
#NYI 
#NYI     return xl_rowcol_to_cell( --$row, $col, $row_abs, $col_abs );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_inc_col($string)
#NYI #
#NYI # Increments the column number of an Excel cell reference in A1 notation.
#NYI # For example C3 to D3
#NYI #
#NYI # Returns: a cell reference string.
#NYI #
#NYI sub xl_inc_col {
#NYI 
#NYI     my $cell = shift;
#NYI     my ( $row, $col, $row_abs, $col_abs ) = xl_cell_to_rowcol( $cell );
#NYI 
#NYI     return xl_rowcol_to_cell( $row, ++$col, $row_abs, $col_abs );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_dec_col($string)
#NYI #
#NYI sub xl_dec_col {
#NYI 
#NYI     my $cell = shift;
#NYI     my ( $row, $col, $row_abs, $col_abs ) = xl_cell_to_rowcol( $cell );
#NYI 
#NYI     return xl_rowcol_to_cell( $row, --$col, $row_abs, $col_abs );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_date_list($years, $months, $days, $hours, $minutes, $seconds)
#NYI #
#NYI sub xl_date_list {
#NYI 
#NYI     return undef unless @_;
#NYI 
#NYI     my $years   = $_[0];
#NYI     my $months  = $_[1] || 1;
#NYI     my $days    = $_[2] || 1;
#NYI     my $hours   = $_[3] || 0;
#NYI     my $minutes = $_[4] || 0;
#NYI     my $seconds = $_[5] || 0;
#NYI 
#NYI     my @date = ( $years, $months, $days, $hours, $minutes, $seconds );
#NYI     my @epoch = ( 1899, 12, 31, 0, 0, 0 );
#NYI 
#NYI     ( $days, $hours, $minutes, $seconds ) = Delta_DHMS( @epoch, @date );
#NYI 
#NYI     my $date =
#NYI       $days + ( $hours * 3600 + $minutes * 60 + $seconds ) / ( 24 * 60 * 60 );
#NYI 
#NYI     # Add a day for Excel's missing leap day in 1900
#NYI     $date++ if ( $date > 59 );
#NYI 
#NYI     return $date;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_parse_time($string)
#NYI #
#NYI sub xl_parse_time {
#NYI 
#NYI     my $time = shift;
#NYI 
#NYI     if ( $time =~ /(\d+):(\d\d):?((?:\d\d)(?:\.\d+)?)?(?:\s+)?(am|pm)?/i ) {
#NYI 
#NYI         my $hours    = $1;
#NYI         my $minutes  = $2;
#NYI         my $seconds  = $3 || 0;
#NYI         my $meridian = lc( $4 || '' );
#NYI 
#NYI         # Normalise midnight and midday
#NYI         $hours = 0 if ( $hours == 12 && $meridian ne '' );
#NYI 
#NYI         # Add 12 hours to the pm times. Note: 12.00 pm has been set to 0.00.
#NYI         $hours += 12 if $meridian eq 'pm';
#NYI 
#NYI         # Calculate the time as a fraction of 24 hours in seconds
#NYI         return ( $hours * 3600 + $minutes * 60 + $seconds ) / ( 24 * 60 * 60 );
#NYI 
#NYI     }
#NYI     else {
#NYI         return undef;    # Not a valid time string
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_parse_date($string)
#NYI #
#NYI sub xl_parse_date {
#NYI 
#NYI     my $date = ParseDate( $_[0] );
#NYI 
#NYI     # Unpack the return value from ParseDate()
#NYI     my ( $years, $months, $days, $hours, undef, $minutes, undef, $seconds ) =
#NYI       unpack( "A4     A2      A2     A2      C        A2      C       A2",
#NYI         $date );
#NYI 
#NYI     # Convert to Excel date
#NYI     return xl_date_list( $years, $months, $days, $hours, $minutes, $seconds );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_parse_date_init("variable=value", ...)
#NYI #
#NYI sub xl_parse_date_init {
#NYI 
#NYI     Date_Init( @_ );    # How lazy is that.
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_decode_date_EU($string)
#NYI #
#NYI sub xl_decode_date_EU {
#NYI 
#NYI     return undef unless @_;
#NYI 
#NYI     my $date = shift;
#NYI     my @date;
#NYI     my $time = 0;
#NYI 
#NYI     # Remove and decode the time portion of the string
#NYI     if ( $date =~ s/(\d+:\d\d:?(\d\d(\.\d+)?)?(\s+)?(am|pm)?)//i ) {
#NYI         $time = xl_parse_time( $1 );
#NYI     }
#NYI 
#NYI     # Return if the string is now blank, i.e. it contained a time only.
#NYI     return $time if $date =~ /^\s*$/;
#NYI 
#NYI     # Decode the date portion of the string
#NYI     @date = Decode_Date_EU( $date );
#NYI     return undef unless @date;
#NYI 
#NYI     return xl_date_list( @date ) + $time;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_decode_date_US($string)
#NYI #
#NYI sub xl_decode_date_US {
#NYI 
#NYI     return undef unless @_;
#NYI 
#NYI     my $date = shift;
#NYI     my @date;
#NYI     my $time = 0;
#NYI 
#NYI     # Remove and decode the time portion of the string
#NYI     if ( $date =~ s/(\d+:\d\d:?(\d\d(\.\d+)?)?(\s+)?(am|pm)?)//i ) {
#NYI         $time = xl_parse_time( $1 );
#NYI     }
#NYI 
#NYI     # Return if the string is now blank, i.e. it contained a time only.
#NYI     return $time if $date =~ /^\s*$/;
#NYI 
#NYI     # Decode the date portion of the string
#NYI     @date = Decode_Date_US( $date );
#NYI     return undef unless @date;
#NYI 
#NYI     return xl_date_list( @date ) + $time;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # xl_decode_date_US($string)
#NYI #
#NYI sub xl_date_1904 {
#NYI 
#NYI     my $date = $_[0] || 0;
#NYI 
#NYI     if ( $date < 1462 ) {
#NYI 
#NYI         # before 1904
#NYI         $date = 0;
#NYI     }
#NYI     else {
#NYI         $date -= 1462;
#NYI     }
#NYI 
#NYI     return $date;
#NYI }
#NYI 
#NYI 
#NYI 1;
#NYI 
#NYI 
#NYI __END__
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Utility - Helper functions for L<Excel::Writer::XLSX>.
#NYI 
#NYI =head1 SYNOPSIS
#NYI 
#NYI Functions to help with some common tasks when using L<Excel::Writer::XLSX>.
#NYI 
#NYI These functions mainly relate to dealing with rows and columns in A1 notation and to handling dates and times.
#NYI 
#NYI     use Excel::Writer::XLSX::Utility;                     # Import everything
#NYI 
#NYI     ($row, $col)    = xl_cell_to_rowcol( 'C2' );          # (1, 2)
#NYI     $str            = xl_rowcol_to_cell( 1, 2 );          # C2
#NYI     $str            = xl_col_to_name( 702 );              # AAA
#NYI     $str            = xl_inc_col( 'Z1'  );                # AA1
#NYI     $str            = xl_dec_col( 'AA1' );                # Z1
#NYI 
#NYI     $date           = xl_date_list(2002, 1, 1);           # 37257
#NYI     $date           = xl_parse_date( '11 July 1997' );    # 35622
#NYI     $time           = xl_parse_time( '3:21:36 PM' );      # 0.64
#NYI     $date           = xl_decode_date_EU( '13 May 2002' ); # 37389
#NYI 
#NYI =head1 DESCRIPTION
#NYI 
#NYI This module provides a set of functions to help with some common tasks encountered when using the L<Excel::Writer::XLSX> module. The two main categories of function are:
#NYI 
#NYI Row and column functions: these are used to deal with Excel's A1 representation of cells. The functions in this category are:
#NYI 
#NYI     xl_rowcol_to_cell
#NYI     xl_cell_to_rowcol
#NYI     xl_col_to_name
#NYI     xl_range
#NYI     xl_range_formula
#NYI     xl_inc_row
#NYI     xl_dec_row
#NYI     xl_inc_col
#NYI     xl_dec_col
#NYI 
#NYI Date and Time functions: these are used to convert dates and times to the numeric format used by Excel. The functions in this category are:
#NYI 
#NYI     xl_date_list
#NYI     xl_date_1904
#NYI     xl_parse_time
#NYI     xl_parse_date
#NYI     xl_parse_date_init
#NYI     xl_decode_date_EU
#NYI     xl_decode_date_US
#NYI 
#NYI All of these functions are exported by default. However, you can use import lists if you wish to limit the functions that are imported:
#NYI 
#NYI     use Excel::Writer::XLSX::Utility;                  # Import everything
#NYI     use Excel::Writer::XLSX::Utility qw(xl_date_list); # xl_date_list only
#NYI     use Excel::Writer::XLSX::Utility qw(:rowcol);      # Row/col functions
#NYI     use Excel::Writer::XLSX::Utility qw(:dates);       # Date functions
#NYI 
#NYI =head1 ROW AND COLUMN FUNCTIONS
#NYI 
#NYI L<Excel::Writer::XLSX> supports two forms of notation to designate the position of cells: Row-column notation and A1 notation.
#NYI 
#NYI Row-column notation uses a zero based index for both row and column while A1 notation uses the standard Excel alphanumeric sequence of column letter and 1-based row. Columns range from A to XFD, i.e. 0 to 16,383, rows range from 0 to 1,048,575 in Excel 2007+. For example:
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
#NYI     $worksheet->write( 'H2', '=H7+1' );
#NYI 
#NYI The functions in the following sections can be used for dealing with A1 notation, for example:
#NYI 
#NYI     ( $row, $col ) = xl_cell_to_rowcol('C2');    # (1, 2)
#NYI     $str           = xl_rowcol_to_cell( 1, 2 );  # C2
#NYI 
#NYI 
#NYI Cell references in Excel can be either relative or absolute. Absolute references are prefixed by the dollar symbol as shown below:
#NYI 
#NYI     A1      # Column and row are relative
#NYI     $A1     # Column is absolute and row is relative
#NYI     A$1     # Column is relative and row is absolute
#NYI     $A$1    # Column and row are absolute
#NYI 
#NYI An absolute reference only makes a difference if the cell is copied. Refer to the Excel documentation for further details. All of the following functions support absolute references.
#NYI 
#NYI =head2 xl_rowcol_to_cell($row, $col, $row_absolute, $col_absolute)
#NYI 
#NYI     Parameters: $row:           Integer
#NYI                 $col:           Integer
#NYI                 $row_absolute:  Boolean (1/0) [optional, default is 0]
#NYI                 $col_absolute:  Boolean (1/0) [optional, default is 0]
#NYI 
#NYI     Returns:    A string in A1 cell notation
#NYI 
#NYI 
#NYI This function converts a zero based row and column cell reference to a A1 style string:
#NYI 
#NYI     $str = xl_rowcol_to_cell( 0, 0 );    # A1
#NYI     $str = xl_rowcol_to_cell( 0, 1 );    # B1
#NYI     $str = xl_rowcol_to_cell( 1, 0 );    # A2
#NYI 
#NYI 
#NYI The optional parameters C<$row_absolute> and C<$col_absolute> can be used to indicate if the row or column is absolute:
#NYI 
#NYI     $str = xl_rowcol_to_cell( 0, 0, 0, 1 );    # $A1
#NYI     $str = xl_rowcol_to_cell( 0, 0, 1, 0 );    # A$1
#NYI     $str = xl_rowcol_to_cell( 0, 0, 1, 1 );    # $A$1
#NYI 
#NYI See above for an explanation of absolute cell references.
#NYI 
#NYI =head2 xl_cell_to_rowcol($string)
#NYI 
#NYI 
#NYI     Parameters: $string         String in A1 format
#NYI 
#NYI     Returns:    List            ($row, $col)
#NYI 
#NYI This function converts an Excel cell reference in A1 notation to a zero based row and column. The function will also handle Excel's absolute, C<$>, cell notation.
#NYI 
#NYI     my ( $row, $col ) = xl_cell_to_rowcol('A1');      # (0, 0)
#NYI     my ( $row, $col ) = xl_cell_to_rowcol('B1');      # (0, 1)
#NYI     my ( $row, $col ) = xl_cell_to_rowcol('C2');      # (1, 2)
#NYI     my ( $row, $col ) = xl_cell_to_rowcol('$C2');     # (1, 2)
#NYI     my ( $row, $col ) = xl_cell_to_rowcol('C$2');     # (1, 2)
#NYI     my ( $row, $col ) = xl_cell_to_rowcol('$C$2');    # (1, 2)
#NYI 
#NYI =head2 xl_col_to_name($col, $col_absolute)
#NYI 
#NYI     Parameters: $col:           Integer
#NYI                 $col_absolute:  Boolean (1/0) [optional, default is 0]
#NYI 
#NYI     Returns:    A column string name.
#NYI 
#NYI 
#NYI This function converts a zero based column reference to a string:
#NYI 
#NYI     $str = xl_col_to_name(0);      # A
#NYI     $str = xl_col_to_name(1);      # B
#NYI     $str = xl_col_to_name(702);    # AAA
#NYI 
#NYI 
#NYI The optional parameter C<$col_absolute> can be used to indicate if the column is absolute:
#NYI 
#NYI     $str = xl_col_to_name( 0, 0 );    # A
#NYI     $str = xl_col_to_name( 0, 1 );    # $A
#NYI     $str = xl_col_to_name( 1, 1 );    # $B
#NYI 
#NYI =head2 xl_range($row_1, $row_2, $col_1, $col_2, $row_abs_1, $row_abs_2, $col_abs_1, $col_abs_2)
#NYI 
#NYI     Parameters: $sheetname      String
#NYI                 $row_1:         Integer
#NYI                 $row_2:         Integer
#NYI                 $col_1:         Integer
#NYI                 $col_2:         Integer
#NYI                 $row_abs_1:     Boolean (1/0) [optional, default is 0]
#NYI                 $row_abs_2:     Boolean (1/0) [optional, default is 0]
#NYI                 $col_abs_1:     Boolean (1/0) [optional, default is 0]
#NYI                 $col_abs_2:     Boolean (1/0) [optional, default is 0]
#NYI 
#NYI     Returns:    A worksheet range formula as a string.
#NYI 
#NYI This function converts zero based row and column cell references to an A1 style range string:
#NYI 
#NYI     my $str = xl_range( 0, 9, 0, 0 );          # A1:A10
#NYI     my $str = xl_range( 1, 8, 2, 2 );          # C2:C9
#NYI     my $str = xl_range( 0, 3, 0, 4 );          # A1:E4
#NYI     my $str = xl_range( 0, 3, 0, 4, 1 );       # A$1:E4
#NYI     my $str = xl_range( 0, 3, 0, 4, 1, 1 );    # A$1:E$4
#NYI 
#NYI =head2 xl_range_formula($sheetname, $row_1, $row_2, $col_1, $col_2)
#NYI 
#NYI     Parameters: $sheetname      String
#NYI                 $row_1:         Integer
#NYI                 $row_2:         Integer
#NYI                 $col_1:         Integer
#NYI                 $col_2:         Integer
#NYI 
#NYI     Returns:    A worksheet range formula as a string.
#NYI 
#NYI This function converts zero based row and column cell references to an A1 style formula string:
#NYI 
#NYI     my $str = xl_range_formula( 'Sheet1', 0, 9,  0, 0 ); # =Sheet1!$A$1:$A$10
#NYI     my $str = xl_range_formula( 'Sheet2', 6, 65, 1, 1 ); # =Sheet2!$B$7:$B$66
#NYI     my $str = xl_range_formula( 'New data', 1, 8, 2, 2 );# ='New data'!$C$2:$C$9
#NYI 
#NYI This is useful for setting ranges in Chart objects:
#NYI 
#NYI     $chart->add_series(
#NYI         categories => xl_range_formula( 'Sheet1', 1, 9, 0, 0 ),
#NYI         values     => xl_range_formula( 'Sheet1', 1, 9, 1, 1 ),
#NYI     );
#NYI 
#NYI     # Which is the same as:
#NYI 
#NYI     $chart->add_series(
#NYI         categories => '=Sheet1!$A$2:$A$10',
#NYI         values     => '=Sheet1!$B$2:$B$10',
#NYI     );
#NYI 
#NYI =head2 xl_inc_row($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a string in A1 format
#NYI 
#NYI     Returns:    Incremented string in A1 format
#NYI 
#NYI This functions takes a cell reference string in A1 notation and increments the row. The function will also handle Excel's absolute, C<$>, cell notation:
#NYI 
#NYI     my $str = xl_inc_row( 'A1' );      # A2
#NYI     my $str = xl_inc_row( 'B$2' );     # B$3
#NYI     my $str = xl_inc_row( '$C3' );     # $C4
#NYI     my $str = xl_inc_row( '$D$4' );    # $D$5
#NYI 
#NYI =head2 xl_dec_row($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a string in A1 format
#NYI 
#NYI     Returns:    Decremented string in A1 format
#NYI 
#NYI This functions takes a cell reference string in A1 notation and decrements the row. The function will also handle Excel's absolute, C<$>, cell notation:
#NYI 
#NYI     my $str = xl_dec_row( 'A2' );      # A1
#NYI     my $str = xl_dec_row( 'B$3' );     # B$2
#NYI     my $str = xl_dec_row( '$C4' );     # $C3
#NYI     my $str = xl_dec_row( '$D$5' );    # $D$4
#NYI 
#NYI =head2 xl_inc_col($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a string in A1 format
#NYI 
#NYI     Returns:    Incremented string in A1 format
#NYI 
#NYI This functions takes a cell reference string in A1 notation and increments the column. The function will also handle Excel's absolute, C<$>, cell notation:
#NYI 
#NYI     my $str = xl_inc_col( 'A1' );      # B1
#NYI     my $str = xl_inc_col( 'Z1' );      # AA1
#NYI     my $str = xl_inc_col( '$B1' );     # $C1
#NYI     my $str = xl_inc_col( '$D$5' );    # $E$5
#NYI 
#NYI =head2 xl_dec_col($string)
#NYI 
#NYI     Parameters: $string, a string in A1 format
#NYI 
#NYI     Returns:    Decremented string in A1 format
#NYI 
#NYI This functions takes a cell reference string in A1 notation and decrements the column. The function will also handle Excel's absolute, C<$>, cell notation:
#NYI 
#NYI     my $str = xl_dec_col( 'B1' );      # A1
#NYI     my $str = xl_dec_col( 'AA1' );     # Z1
#NYI     my $str = xl_dec_col( '$C1' );     # $B1
#NYI     my $str = xl_dec_col( '$E$5' );    # $D$5
#NYI 
#NYI =head1 TIME AND DATE FUNCTIONS
#NYI 
#NYI Dates and times in Excel are represented by real numbers, for example "Jan 1 2001 12:30 AM" is represented by the number 36892.521.
#NYI 
#NYI The integer part of the number stores the number of days since the epoch and the fractional part stores the percentage of the day in seconds.
#NYI 
#NYI A date or time in Excel is like any other number. To display the number as a date you must apply a number format to it: Refer to the C<set_num_format()> method in the Excel::Writer::XLSX documentation:
#NYI 
#NYI     $date = xl_date_list( 2001, 1, 1, 12, 30 );
#NYI     $format->set_num_format( 'mmm d yyyy hh:mm AM/PM' );
#NYI     $worksheet->write( 'A1', $date, $format );    # Jan 1 2001 12:30 AM
#NYI 
#NYI The date handling functions below are supplied for historical reasons. In the current version of the module it is easier to just use the C<write_date_time()> function to write dates or times. See the DATES AND TIME IN EXCEL section of the main L<Excel::Writer::XLSX> documentation for details.
#NYI 
#NYI In addition to using the functions below you must install the L<Date::Manip> and L<Date::Calc> modules. See L<REQUIREMENTS> and the individual requirements of each functions.
#NYI 
#NYI For a C<DateTime.pm> solution see the L<DateTime::Format::Excel> module.
#NYI 
#NYI =head2 xl_date_list($years, $months, $days, $hours, $minutes, $seconds)
#NYI 
#NYI 
#NYI     Parameters: $years:         Integer
#NYI                 $months:        Integer [optional, default is 1]
#NYI                 $days:          Integer [optional, default is 1]
#NYI                 $hours:         Integer [optional, default is 0]
#NYI                 $minutes:       Integer [optional, default is 0]
#NYI                 $seconds:       Float   [optional, default is 0]
#NYI 
#NYI     Returns:    A number that represents an Excel date
#NYI                 or undef for an invalid date.
#NYI 
#NYI     Requires:   Date::Calc
#NYI 
#NYI This function converts an array of data into a number that represents an Excel date. All of the parameters are optional except for C<$years>.
#NYI 
#NYI     $date1 = xl_date_list( 2002, 1, 2 );                # 2 Jan 2002
#NYI     $date2 = xl_date_list( 2002, 1, 2, 12 );            # 2 Jan 2002 12:00 pm
#NYI     $date3 = xl_date_list( 2002, 1, 2, 12, 30 );        # 2 Jan 2002 12:30 pm
#NYI     $date4 = xl_date_list( 2002, 1, 2, 12, 30, 45 );    # 2 Jan 2002 12:30:45 pm
#NYI 
#NYI This function can be used in conjunction with functions that parse date and time strings. In fact it is used in most of the following functions.
#NYI 
#NYI =head2 xl_parse_time($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a textual representation of a time
#NYI 
#NYI     Returns:    A number that represents an Excel time
#NYI                 or undef for an invalid time.
#NYI 
#NYI This function converts a time string into a number that represents an Excel time. The following time formats are valid:
#NYI 
#NYI     hh:mm       [AM|PM]
#NYI     hh:mm       [AM|PM]
#NYI     hh:mm:ss    [AM|PM]
#NYI     hh:mm:ss.ss [AM|PM]
#NYI 
#NYI 
#NYI The meridian, AM or PM, is optional and case insensitive. A 24 hour time is assumed if the meridian is omitted.
#NYI 
#NYI     $time1 = xl_parse_time( '12:18' );
#NYI     $time2 = xl_parse_time( '12:18:14' );
#NYI     $time3 = xl_parse_time( '12:18:14 AM' );
#NYI     $time4 = xl_parse_time( '1:18:14 AM' );
#NYI 
#NYI Time in Excel is expressed as a fraction of the day in seconds. Therefore you can calculate an Excel time as follows:
#NYI 
#NYI     $time = ( $hours * 3600 + $minutes * 60 + $seconds ) / ( 24 * 60 * 60 );
#NYI 
#NYI =head2 xl_parse_date($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a textual representation of a date and time
#NYI 
#NYI     Returns:    A number that represents an Excel date
#NYI                 or undef for an invalid date.
#NYI 
#NYI     Requires:   Date::Manip and Date::Calc
#NYI 
#NYI This function converts a date and time string into a number that represents an Excel date.
#NYI 
#NYI The parsing is performed using the C<ParseDate()> function of the L<Date::Manip> module. Refer to the C<Date::Manip> documentation for further information about the date and time formats that can be parsed. In order to use this function you will probably have to initialise some C<Date::Manip> variables via the C<xl_parse_date_init()> function, see below.
#NYI 
#NYI     xl_parse_date_init( "TZ=GMT", "DateFormat=non-US" );
#NYI 
#NYI     $date1 = xl_parse_date( "11/7/97" );
#NYI     $date2 = xl_parse_date( "Friday 11 July 1997" );
#NYI     $date3 = xl_parse_date( "10:30 AM Friday 11 July 1997" );
#NYI     $date4 = xl_parse_date( "Today" );
#NYI     $date5 = xl_parse_date( "Yesterday" );
#NYI 
#NYI Note, if you parse a string that represents a time but not a date this function will add the current date. If you want the time without the date you can do something like the following:
#NYI 
#NYI     $time  = xl_parse_date( "10:30 AM" );
#NYI     $time -= int( $time );
#NYI 
#NYI =head2 xl_parse_date_init("variable=value", ...)
#NYI 
#NYI 
#NYI     Parameters: A list of Date::Manip variable strings
#NYI 
#NYI     Returns:    A list of all the Date::Manip strings
#NYI 
#NYI     Requires:   Date::Manip
#NYI 
#NYI This function is used to initialise variables required by the L<Date::Manip> module. You should call this function before calling C<xl_parse_date()>. It need only be called once.
#NYI 
#NYI This function is a thin wrapper for the C<Date::Manip::Date_Init()> function. You can use C<Date_Init()>  directly if you wish. Refer to the C<Date::Manip> documentation for further information.
#NYI 
#NYI     xl_parse_date_init( "TZ=MST", "DateFormat=US" );
#NYI     $date1 = xl_parse_date( "11/7/97" );    # November 7th 1997
#NYI 
#NYI     xl_parse_date_init( "TZ=GMT", "DateFormat=non-US" );
#NYI     $date1 = xl_parse_date( "11/7/97" );    # July 11th 1997
#NYI 
#NYI =head2 xl_decode_date_EU($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a textual representation of a date and time
#NYI 
#NYI     Returns:    A number that represents an Excel date
#NYI                 or undef for an invalid date.
#NYI 
#NYI     Requires:   Date::Calc
#NYI 
#NYI This function converts a date and time string into a number that represents an Excel date.
#NYI 
#NYI The date parsing is performed using the C<Decode_Date_EU()> function of the L<Date::Calc> module. Refer to the C<Date::Calc> documentation for further information about the date formats that can be parsed. Also note the following from the C<Date::Calc> documentation:
#NYI 
#NYI "If the year is given as one or two digits only (i.e., if the year is less than 100), it is mapped to the window 1970 -2069 as follows:"
#NYI 
#NYI      0 <= $year <  70  ==>  $year += 2000;
#NYI     70 <= $year < 100  ==>  $year += 1900;
#NYI 
#NYI The time portion of the string is parsed using the C<xl_parse_time()> function described above.
#NYI 
#NYI Note: the EU in the function name means that a European date format is assumed if it is not clear from the string. See the first example below.
#NYI 
#NYI     $date1 = xl_decode_date_EU( "11/7/97" );                    #11 July 1997
#NYI     $date2 = xl_decode_date_EU( "Sat 12 Sept 1998" );
#NYI     $date3 = xl_decode_date_EU( "4:30 AM Sat 12 Sept 1998" );
#NYI 
#NYI =head2 xl_decode_date_US($string)
#NYI 
#NYI 
#NYI     Parameters: $string, a textual representation of a date and time
#NYI 
#NYI     Returns:    A number that represents an Excel date
#NYI                 or undef for an invalid date.
#NYI 
#NYI     Requires:   Date::Calc
#NYI 
#NYI This function converts a date and time string into a number that represents an Excel date.
#NYI 
#NYI The date parsing is performed using the C<Decode_Date_US()> function of the L<Date::Calc> module. Refer to the C<Date::Calc> documentation for further information about the date formats that can be parsed. Also note the following from the C<Date::Calc> documentation:
#NYI 
#NYI "If the year is given as one or two digits only (i.e., if the year is less than 100), it is mapped to the window 1970 -2069 as follows:"
#NYI 
#NYI      0 <= $year <  70  ==>  $year += 2000;
#NYI     70 <= $year < 100  ==>  $year += 1900;
#NYI 
#NYI The time portion of the string is parsed using the C<xl_parse_time()> function described above.
#NYI 
#NYI Note: the US in the function name means that an American date format is assumed if it is not clear from the string. See the first example below.
#NYI 
#NYI     $date1 = xl_decode_date_US( "11/7/97" );                 # 7 November 1997
#NYI     $date2 = xl_decode_date_US( "Sept 12 Saturday 1998" );
#NYI     $date3 = xl_decode_date_US( "4:30 AM Sept 12 Sat 1998" );
#NYI 
#NYI =head2 xl_date_1904($date)
#NYI 
#NYI 
#NYI     Parameters: $date, an Excel date with a 1900 epoch
#NYI 
#NYI     Returns:    an Excel date with a 1904 epoch or zero if
#NYI                 the $date is before 1904
#NYI 
#NYI 
#NYI This function converts an Excel date based on the 1900 epoch into a date based on the 1904 epoch.
#NYI 
#NYI     $date1 = xl_date_list( 2002, 1, 13 );    # 13 Jan 2002, 1900 epoch
#NYI     $date2 = xl_date_1904( $date1 );         # 13 Jan 2002, 1904 epoch
#NYI 
#NYI See also the C<set_1904()> workbook method in the L<Excel::Writer::XLSX> documentation.
#NYI 
#NYI =head1 REQUIREMENTS
#NYI 
#NYI The date and time functions require functions from the L<Date::Manip> and L<Date::Calc> modules. The required functions are "autoused" from these modules so that you do not have to install them unless you wish to use the date and time routines. Therefore it is possible to use the row and column functions without having C<Date::Manip> and C<Date::Calc> installed.
#NYI 
#NYI For more information about "autousing" refer to the documentation on the C<autouse> pragma.
#NYI 
#NYI =head1 BUGS
#NYI 
#NYI When using the autoused functions from C<Date::Manip> and C<Date::Calc> on Perl 5.6.0 with C<-w> you will get a warning like this:
#NYI 
#NYI     "Subroutine xxx redefined ..."
#NYI 
#NYI The current workaround for this is to put C<use warnings;> near the beginning of your program.
#NYI 
#NYI =head1 AUTHOR
#NYI 
#NYI John McNamara jmcnamara@cpan.org
#NYI 
#NYI =head1 COPYRIGHT
#NYI 
#NYI Copyright MM-MMXVII, John McNamara.
#NYI 
#NYI All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
#NYI 
