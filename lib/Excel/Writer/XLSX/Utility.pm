use v6.c+;

unit module Excel::Writer::XLSX::Utility;

###############################################################################
#
# Utility - Helper functions for Excel::Writer::XLSX.
#
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

#NYI use strict;
#NYI use Exporter;
#NYI use warnings;
#NYI use autouse 'Date::Calc'  => qw(Delta_DHMS Decode_Date_EU Decode_Date_US);
#NYI use autouse 'Date::Manip' => qw(ParseDate Date_Init);

#NYI our $VERSION = '0.96';

# Row and column functions
my @rowcol = <
  xl-rowcol-to-cell
  xl-cell-to-rowcol
  xl-col-to-name
  xl-range
  xl-range-formula
  xl-inc-row
  xl-dec-row
  xl-inc-col
  xl-dec-col
>;

# Date and Time functions
my @dates = <
  xl-date-list
  xl-date_1904
  xl-parse-time
  xl-parse-date
  xl-parse-date-init
  xl-decode-date-EU
  xl-decode-date-US
>;

#NYI our @ISA         = qw(Exporter);
#NYI our @EXPORT_OK   = ();
#NYI our @EXPORT      = ( @rowcol, @dates, 'quote-sheetname' );
#NYI our %EXPORT_TAGS = (
#NYI     rowcol => \@rowcol,
#NYI     dates  => \@dates
#NYI );


sub croak($string) {
  die $string;
}

sub carp($string) {
  warn $string;
}

###############################################################################
#
# xl-rowcol-to-cell($row, $col, $row-absolute, $col-absolute)
#
sub xl-rowcol-to-cell($row, $col, $row-abs, $col-abs) is export {

    $row++;          # Change from 0-indexed to 1 indexed.
    $row-abs = $row-abs ?? '$' !! '';
    $col-abs = $col-abs ?? '$' !! '';


    my $col-str = xl-col-to-name( $col, $col-abs );

    return $col-str ~ $row-abs ~ $row;
}


###############################################################################
#
# xl-cell-to-rowcol($string)
#
# Returns: ($row, $col, $row-absolute, $col-absolute)
#
# The $row-absolute and $col-absolute parameters aren't documented because they
# mainly used internally and aren't very useful to the user.
#
sub xl-cell-to-rowcol($cell) is export {
    return ( 0, 0, 0, 0 ) unless $cell;

    $cell ~~ /
              (\$?)
              (<[A..Z]> ** 1..3)
              (\$?)
              (\d+)
             /;

    my $col-abs = $0 eq "" ?? 0 !! 1;
    my $col     = $2;
    my $row-abs = $2 eq "" ?? 0 !! 1;
    my $row     = $3;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars = $col.comb;
    my $expn = 0;
    $col = 0;

    while @chars.elems {
        my $char = @chars.pop;    # LS char first
        $col += ( $char.ord - 'A'.ord + 1 ) * ( 26 ** $expn );
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col, $row-abs, $col-abs;
}


###############################################################################
#
# xl-col-to-name($col, $col-absolute)
#
sub xl-col-to-name($col, $col-abs) is export {

    $col-abs = $col-abs ?? '$' !! '';
    my $col-str = '';

    # Change from 0-indexed to 1 indexed.
    $col++;

    while $col {

        # Set remainder from 1 .. 26
        my $remainder = $col % 26 || 26;

        # Convert the $remainder to a character. C-ishly.
        my $col-letter = chr( 'A'.ord + $remainder - 1 );

        # Accumulate the column letters, right to left.
        $col-str = $col-letter ~ $col-str;

        # Get the next order of magnitude.
        $col = ( ( $col - 1 ) / 26 ).int;
    }

    return $col-abs ~ $col-str;
}


###############################################################################
#
# xl-range($row-1, $row-2, $col-1, $col-2, $row-abs-1, $row-abs-2, $col-abs-1, $col-abs-2)
#
sub xl-range($row1,     $row2,     $col1,     $col2,
             $row-abs1, $row-abs2, $col-abs1, $col-abs2) is export {
    my $range1 = xl-rowcol-to-cell( $row1, $col1, $row-abs1, $col-abs1 );
    my $range2 = xl-rowcol-to-cell( $row2, $col2, $row-abs2, $col-abs2 );

    return $range1 ~ ':' ~ $range2;
}


###############################################################################
#
# xl-range-formula($sheetname, $row-1, $row-2, $col-1, $col-2)
#
sub xl-range-formula($sheetname, $row1, $row2, $col1, $col2) is export {
    $sheetname = quote-sheetname( $sheetname );

    my $range = xl-range( $row1, $row2, $col1, $col2, 1, 1, 1, 1 );

    return '=' ~ $sheetname ~ '!' ~ $range
}


###############################################################################
#
# quote-sheetname()
#
# Sheetnames used in references should be quoted if they contain any spaces,
# special characters or if they look like something that isn't a sheet name.
#
sub quote-sheetname($sheetname) is export {
    # Use Excel's conventions and quote the sheet name if it contains any
    # non-word character or if it isn't already quoted.
    if $sheetname ~~ /\W/ && $sheetname !~~ /^\'/ {
        # Double quote any single quotes.
        $sheetname ~~ s:g/\'/''/;
        $sheetname = "'" ~ $sheetname ~ "'";
    }

    return $sheetname;
}


###############################################################################
#
# xl-inc-row($string)
#
sub xl-inc-row($cell) is export {
    my ( $row, $col, $row-abs, $col-abs ) = xl-cell-to-rowcol( $cell );

    return xl-rowcol-to-cell( $row + 1, $col, $row-abs, $col-abs );
}


###############################################################################
#
# xl-dec-row($string)
#
# Decrements the row number of an Excel cell reference in A1 notation.
# For example C4 to C3
#
# Returns: a cell reference string.
#
sub xl-dec-row($cell) is export {
    my ( $row, $col, $row-abs, $col-abs ) = xl-cell-to-rowcol( $cell );

    return xl-rowcol-to-cell( $row - 1, $col, $row-abs, $col-abs );
}


###############################################################################
#
# xl-inc-col($string)
#
# Increments the column number of an Excel cell reference in A1 notation.
# For example C3 to D3
#
# Returns: a cell reference string.
#
sub xl-inc-col($cell) is export {
    my ( $row, $col, $row-abs, $col-abs ) = xl-cell-to-rowcol( $cell );

    return xl-rowcol-to-cell( $row, $col+1, $row-abs, $col-abs );
}


###############################################################################
#
# xl-dec-col($string)
#
sub xl-dec-col($cell) is export {
    my ( $row, $col, $row-abs, $col-abs ) = xl-cell-to-rowcol( $cell );

    return xl-rowcol-to-cell( $row, $col - 1, $row-abs, $col-abs );
}


###############################################################################
#
# xl-date-list($years, $months, $days, $hours, $minutes, $seconds)
#
sub xl-date-list($years,
                 $months = 1,
                 $days = 1,
                 $hours = 0,
                 $minutes = 0,
                 $seconds = 0) is export {

    return Nil unless $years.defined;

    my $datetime = DateTime.new( $years, $months, $days, $hours, $minutes, $seconds );
    my $epoch = DateTime.new( 1899, 12, 31, 0, 0, 0 );

    
    my $date = ($datetime - $epoch) / (24 * 60 * 60);

    # Add a day for Excel's missing leap day in 1900
    $date++ if $date > 59;

    return $date;
}


###############################################################################
#
# xl-parse-time($string)
#
sub xl-parse-time($time) is export {
    if $time ~~ /:i
                 (\d+)
                 ':'
                 (\d\d)
                 ':'?
                 (
                   [\d\d]
                   [\.\d+]?
                 )?
                 \s*
                 (am|pm)?
                / {

        my $hours    = $0;
        my $minutes  = $1;
        my $seconds  = $2 || 0;
        my $meridian = lc( $3 || '' );

        # Normalise midnight and midday
        $hours = 0 if ( $hours == 12 && $meridian ne '' );

        # Add 12 hours to the pm times. Note: 12.00 pm has been set to 0.00.
        $hours += 12 if $meridian eq 'pm';

        # Calculate the time as a fraction of 24 hours in seconds
        return ( $hours * 3600 + $minutes * 60 + $seconds ) / ( 24 * 60 * 60 );

    }
    else {
        return Nil;    # Not a valid time string
    }
}


###############################################################################
#
# xl-parse-date($string)
#
#NYI sub xl-parse-date($rawdate) is export {

#NYI     my $date = ParseDate( $rawdate );

#NYI     # Unpack the return value from ParseDate()
#NYI     $date ~~ /(....)(..)(..)(..).(..).(..)/;
#NYI     my ( $years, $months, $days, $hours, $minutes, $seconds ) = ($0, $1, $2, $3, $4, $5);
   

#NYI     # Convert to Excel date
#NYI     return xl-date-list( $years, $months, $days, $hours, $minutes, $seconds );
#NYI }


###############################################################################
#
# xl-parse-date-init("variable=value", ...)
#
#NYI sub xl-parse-date-init(*@args) is export {

#NYI     Date-Init( @args );    # How lazy is that.
#NYI }


###############################################################################
#
# xl-decode-date-EU($string)
#
sub xl-decode-date-EU($date) is export {

    return Nil unless $date.defined;

    my @date;
    my $time = 0;

    # Remove and decode the time portion of the string
    if $date ~~ s:i/
                    (\d+ ':' \d\d ':'? [\d\d[\.\d+]?]?\s*[am|pm]?)
                  // {
        $time = xl-parse-time( $0 );
    }

    # Return if the string is now blank, i.e. it contained a time only.
    return $time if $date ~~ /^\s*$/;

    # Decode the date portion of the string
    @date = Decode-Date-EU( $date );
    return Nil unless @date;

    return xl-date-list( @date ) + $time;
}


###############################################################################
#
# xl-decode-date-US($string)
#
sub xl-decode-date-US($date) is export {

    return Nil unless $date.defined;

    my @date;
    my $time = 0;

    # Remove and decode the time portion of the string
    if $date ~~ s:i/(\d+:\d\d:?[\d\d[\.\d+]?]?\s*[am|pm]?)// {
        $time = xl-parse-time( $0 );
    }

    # Return if the string is now blank, i.e. it contained a time only.
    return $time if $date ~~ /^\s*$/;

    # Decode the date portion of the string
    @date = Decode-Date-US( $date );
    return Nil unless @date;

    return xl-date-list( @date ) + $time;
}


###############################################################################
#
# xl-decode-date-US($string)
#
sub xl-date_1904($date = 0) is export {
    if $date < 1462 {
        # before 1904
        $date = 0;
    }
    else {
        $date -= 1462;
    }

    return $date;
}

# Functions to emulate stuff from other modules in the Perl 5 version
#
# The date routines are less strict in what they will accept than the orinials they are emulating

sub Decode-Date-EU($date) {
  Decode-Date($date, 'EU');
}

sub Decode-Date-US($date) {
  Decode-Date($date, 'US');
}

my %month = (
  'jan' =>  1,
  'feb' =>  2,
  'mar' =>  3,
  'apr' =>  4,
  'may' =>  5,
  'jun' =>  6,
  'jul' =>  7,
  'aug' =>  8,
  'sep' =>  9,
  'oct' => 10,
  'nov' => 11,
  'dec' => 12,
);

sub fix-year($year) {
  return $year if $year < 0;
  return $year if $year >= 100;
  my $year-now = Date(now).year;
  my $century = ($year-now / 100).int;
  $year += $century;
  $year += 100 if $year-now - $year > 50;
  $year -= 100 if $year-now - $year < -50;
}

sub Decode-Date($date, $bias) {
  my $year;
  my $month;
  my $day;

  if $date ~~ /:i (\d+) \W+ (jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec) \D+ (\d+)/ {
    $year = $2;
    $day = $0;
    $month = %month{$1.lc};
    return (fix-year($year), $month, $day);
  }
  if $date ~~ / (\d+) \D+ (\d+) \D+ (\d+) / {
    $day = $0;
    $month = $1;
    if    ($day <= 12 and $month > 12)
       or ($bias eq 'US') {
      ($day, $month) = ($month, $day);
    }
    $year = $2;
    return (fix-year($year), $month, $day);
  }
  if $date ~~ / (\d+) / {
    my $sub = $0;
    given $sub.chars {
      when 3 {
        $sub ~~ /(.)(.)(.)/;
        $day = $0;
        $month = $0;
        $year = $2;
      }
      when 4 {
        $sub ~~ /(.)(.)(..)/;
        $day = $0;
        $month = $0;
        $year = $2;
      }
      when 5 {
        $sub ~~ /(.)(..)(..)/;
        $day = $0;
        $month = $0;
        $year = $2;
      }
      when 6 {
        $sub ~~ /(..)(..)(..)/;
        $day = $0;
        $month = $0;
        $year = $2;
      }
      when 7 {
        $sub ~~ /(.)(..)(....)/;
        $day = $0;
        $month = $0;
        $year = $2;
      }
      when 8 {
        $sub ~~ /(..)(..)(..)/;
        $day = $0;
        $month = $0;
        $year = $2;
      }
    }
    return (fix-year($year), $month, $day);
  }
}

=begin pod
=head1 NAME

Utility - Helper functions for L<Excel::Writer::XLSX>.

=head1 SYNOPSIS

Functions to help with some common tasks when using L<Excel::Writer::XLSX>.

These functions mainly relate to dealing with rows and columns in A1 notation and to handling dates and times.

    use Excel::Writer::XLSX::Utility;                     # Import everything

    ($row, $col)    = xl-cell-to-rowcol( 'C2' );          # (1, 2)
    $str            = xl-rowcol-to-cell( 1, 2 );          # C2
    $str            = xl-col-to-name( 702 );              # AAA
    $str            = xl-inc-col( 'Z1'  );                # AA1
    $str            = xl-dec-col( 'AA1' );                # Z1

    $date           = xl-date-list(2002, 1, 1);           # 37257
    $date           = xl-parse-date( '11 July 1997' );    # 35622
    $time           = xl-parse-time( '3:21:36 PM' );      # 0.64
    $date           = xl-decode-date-EU( '13 May 2002' ); # 37389

=head1 DESCRIPTION

This module provides a set of functions to help with some common tasks encountered when using the L<Excel::Writer::XLSX> module. The two main categories of function are:

Row and column functions: these are used to deal with Excel's A1 representation of cells. The functions in this category are:

    xl-rowcol-to-cell
    xl-cell-to-rowcol
    xl-col-to-name
    xl-range
    xl-range-formula
    xl-inc-row
    xl-dec-row
    xl-inc-col
    xl-dec-col

Date and Time functions: these are used to convert dates and times to the numeric format used by Excel. The functions in this category are:

    xl-date-list
    xl-date_1904
    xl-parse-time
    xl-parse-date
    xl-parse-date-init
    xl-decode-date-EU
    xl-decode-date-US

All of these functions are exported by default. However, you can use import lists if you wish to limit the functions that are imported:

    use Excel::Writer::XLSX::Utility;                  # Import everything
    use Excel::Writer::XLSX::Utility qw(xl-date-list); # xl-date-list only
    use Excel::Writer::XLSX::Utility qw(:rowcol);      # Row/col functions
    use Excel::Writer::XLSX::Utility qw(:dates);       # Date functions

=head1 ROW AND COLUMN FUNCTIONS

L<Excel::Writer::XLSX> supports two forms of notation to designate the position of cells: Row-column notation and A1 notation.

Row-column notation uses a zero based index for both row and column while A1 notation uses the standard Excel alphanumeric sequence of column letter and 1-based row. Columns range from A to XFD, i.e. 0 to 16,383, rows range from 0 to 1,048,575 in Excel 2007+. For example:

    (0, 0)      # The top left cell in row-column notation.
    ('A1')      # The top left cell in A1 notation.

    (1999, 29)  # Row-column notation.
    ('AD2000')  # The same cell in A1 notation.

Row-column notation is useful if you are referring to cells programmatically:

    for my $i ( 0 .. 9 ) {
        $worksheet->write( $i, 0, 'Hello' );    # Cells A1 to A10
    }

A1 notation is useful for setting up a worksheet manually and for working with formulas:

    $worksheet->write( 'H1', 200 );
    $worksheet->write( 'H2', '=H7+1' );

The functions in the following sections can be used for dealing with A1 notation, for example:

    ( $row, $col ) = xl-cell-to-rowcol('C2');    # (1, 2)
    $str           = xl-rowcol-to-cell( 1, 2 );  # C2


Cell references in Excel can be either relative or absolute. Absolute references are prefixed by the dollar symbol as shown below:

    A1      # Column and row are relative
    $A1     # Column is absolute and row is relative
    A$1     # Column is relative and row is absolute
    $A$1    # Column and row are absolute

An absolute reference only makes a difference if the cell is copied. Refer to the Excel documentation for further details. All of the following functions support absolute references.

=head2 xl-rowcol-to-cell($row, $col, $row-absolute, $col-absolute)

    Parameters: $row:           Integer
                $col:           Integer
                $row-absolute:  Boolean (1/0) [optional, default is 0]
                $col-absolute:  Boolean (1/0) [optional, default is 0]

    Returns:    A string in A1 cell notation


This function converts a zero based row and column cell reference to a A1 style string:

    $str = xl-rowcol-to-cell( 0, 0 );    # A1
    $str = xl-rowcol-to-cell( 0, 1 );    # B1
    $str = xl-rowcol-to-cell( 1, 0 );    # A2


The optional parameters C<$row-absolute> and C<$col-absolute> can be used to indicate if the row or column is absolute:

    $str = xl-rowcol-to-cell( 0, 0, 0, 1 );    # $A1
    $str = xl-rowcol-to-cell( 0, 0, 1, 0 );    # A$1
    $str = xl-rowcol-to-cell( 0, 0, 1, 1 );    # $A$1

See above for an explanation of absolute cell references.

=head2 xl-cell-to-rowcol($string)


    Parameters: $string         String in A1 format

    Returns:    List            ($row, $col)

This function converts an Excel cell reference in A1 notation to a zero based row and column. The function will also handle Excel's absolute, C<$>, cell notation.

    my ( $row, $col ) = xl-cell-to-rowcol('A1');      # (0, 0)
    my ( $row, $col ) = xl-cell-to-rowcol('B1');      # (0, 1)
    my ( $row, $col ) = xl-cell-to-rowcol('C2');      # (1, 2)
    my ( $row, $col ) = xl-cell-to-rowcol('$C2');     # (1, 2)
    my ( $row, $col ) = xl-cell-to-rowcol('C$2');     # (1, 2)
    my ( $row, $col ) = xl-cell-to-rowcol('$C$2');    # (1, 2)

=head2 xl-col-to-name($col, $col-absolute)

    Parameters: $col:           Integer
                $col-absolute:  Boolean (1/0) [optional, default is 0]

    Returns:    A column string name.


This function converts a zero based column reference to a string:

    $str = xl-col-to-name(0);      # A
    $str = xl-col-to-name(1);      # B
    $str = xl-col-to-name(702);    # AAA


The optional parameter C<$col-absolute> can be used to indicate if the column is absolute:

    $str = xl-col-to-name( 0, 0 );    # A
    $str = xl-col-to-name( 0, 1 );    # $A
    $str = xl-col-to-name( 1, 1 );    # $B

=head2 xl-range($row-1, $row-2, $col-1, $col-2, $row-abs-1, $row-abs-2, $col-abs-1, $col-abs-2)

    Parameters: $sheetname      String
                $row-1:         Integer
                $row-2:         Integer
                $col-1:         Integer
                $col-2:         Integer
                $row-abs-1:     Boolean (1/0) [optional, default is 0]
                $row-abs-2:     Boolean (1/0) [optional, default is 0]
                $col-abs-1:     Boolean (1/0) [optional, default is 0]
                $col-abs-2:     Boolean (1/0) [optional, default is 0]

    Returns:    A worksheet range formula as a string.

This function converts zero based row and column cell references to an A1 style range string:

    my $str = xl-range( 0, 9, 0, 0 );          # A1:A10
    my $str = xl-range( 1, 8, 2, 2 );          # C2:C9
    my $str = xl-range( 0, 3, 0, 4 );          # A1:E4
    my $str = xl-range( 0, 3, 0, 4, 1 );       # A$1:E4
    my $str = xl-range( 0, 3, 0, 4, 1, 1 );    # A$1:E$4

=head2 xl-range-formula($sheetname, $row-1, $row-2, $col-1, $col-2)

    Parameters: $sheetname      String
                $row-1:         Integer
                $row-2:         Integer
                $col-1:         Integer
                $col-2:         Integer

    Returns:    A worksheet range formula as a string.

This function converts zero based row and column cell references to an A1 style formula string:

    my $str = xl-range-formula( 'Sheet1', 0, 9,  0, 0 ); # =Sheet1!$A$1:$A$10
    my $str = xl-range-formula( 'Sheet2', 6, 65, 1, 1 ); # =Sheet2!$B$7:$B$66
    my $str = xl-range-formula( 'New data', 1, 8, 2, 2 );# ='New data'!$C$2:$C$9

This is useful for setting ranges in Chart objects:

    $chart->add-series(
        categories => xl-range-formula( 'Sheet1', 1, 9, 0, 0 ),
        values     => xl-range-formula( 'Sheet1', 1, 9, 1, 1 ),
    );

    # Which is the same as:

    $chart->add-series(
        categories => '=Sheet1!$A$2:$A$10',
        values     => '=Sheet1!$B$2:$B$10',
    );

=head2 xl-inc-row($string)


    Parameters: $string, a string in A1 format

    Returns:    Incremented string in A1 format

This functions takes a cell reference string in A1 notation and increments the row. The function will also handle Excel's absolute, C<$>, cell notation:

    my $str = xl-inc-row( 'A1' );      # A2
    my $str = xl-inc-row( 'B$2' );     # B$3
    my $str = xl-inc-row( '$C3' );     # $C4
    my $str = xl-inc-row( '$D$4' );    # $D$5

=head2 xl-dec-row($string)


    Parameters: $string, a string in A1 format

    Returns:    Decremented string in A1 format

This functions takes a cell reference string in A1 notation and decrements the row. The function will also handle Excel's absolute, C<$>, cell notation:

    my $str = xl-dec-row( 'A2' );      # A1
    my $str = xl-dec-row( 'B$3' );     # B$2
    my $str = xl-dec-row( '$C4' );     # $C3
    my $str = xl-dec-row( '$D$5' );    # $D$4

=head2 xl-inc-col($string)


    Parameters: $string, a string in A1 format

    Returns:    Incremented string in A1 format

This functions takes a cell reference string in A1 notation and increments the column. The function will also handle Excel's absolute, C<$>, cell notation:

    my $str = xl-inc-col( 'A1' );      # B1
    my $str = xl-inc-col( 'Z1' );      # AA1
    my $str = xl-inc-col( '$B1' );     # $C1
    my $str = xl-inc-col( '$D$5' );    # $E$5

=head2 xl-dec-col($string)

    Parameters: $string, a string in A1 format

    Returns:    Decremented string in A1 format

This functions takes a cell reference string in A1 notation and decrements the column. The function will also handle Excel's absolute, C<$>, cell notation:

    my $str = xl-dec-col( 'B1' );      # A1
    my $str = xl-dec-col( 'AA1' );     # Z1
    my $str = xl-dec-col( '$C1' );     # $B1
    my $str = xl-dec-col( '$E$5' );    # $D$5

=head1 TIME AND DATE FUNCTIONS

Dates and times in Excel are represented by real numbers, for example "Jan 1 2001 12:30 AM" is represented by the number 36892.521.

The integer part of the number stores the number of days since the epoch and the fractional part stores the percentage of the day in seconds.

A date or time in Excel is like any other number. To display the number as a date you must apply a number format to it: Refer to the C<set-num-format()> method in the Excel::Writer::XLSX documentation:

    $date = xl-date-list( 2001, 1, 1, 12, 30 );
    $format->set-num-format( 'mmm d yyyy hh:mm AM/PM' );
    $worksheet->write( 'A1', $date, $format );    # Jan 1 2001 12:30 AM

The date handling functions below are supplied for historical reasons. In the current version of the module it is easier to just use the C<write-date-time()> function to write dates or times. See the DATES AND TIME IN EXCEL section of the main L<Excel::Writer::XLSX> documentation for details.

In addition to using the functions below you must install the L<Date::Manip> and L<Date::Calc> modules. See L<REQUIREMENTS> and the individual requirements of each functions.

For a C<DateTime.pm> solution see the L<DateTime::Format::Excel> module.

=head2 xl-date-list($years, $months, $days, $hours, $minutes, $seconds)


    Parameters: $years:         Integer
                $months:        Integer [optional, default is 1]
                $days:          Integer [optional, default is 1]
                $hours:         Integer [optional, default is 0]
                $minutes:       Integer [optional, default is 0]
                $seconds:       Float   [optional, default is 0]

    Returns:    A number that represents an Excel date
                or undef for an invalid date.

    Requires:   Date::Calc

This function converts an array of data into a number that represents an Excel date. All of the parameters are optional except for C<$years>.

    $date1 = xl-date-list( 2002, 1, 2 );                # 2 Jan 2002
    $date2 = xl-date-list( 2002, 1, 2, 12 );            # 2 Jan 2002 12:00 pm
    $date3 = xl-date-list( 2002, 1, 2, 12, 30 );        # 2 Jan 2002 12:30 pm
    $date4 = xl-date-list( 2002, 1, 2, 12, 30, 45 );    # 2 Jan 2002 12:30:45 pm

This function can be used in conjunction with functions that parse date and time strings. In fact it is used in most of the following functions.

=head2 xl-parse-time($string)


    Parameters: $string, a textual representation of a time

    Returns:    A number that represents an Excel time
                or undef for an invalid time.

This function converts a time string into a number that represents an Excel time. The following time formats are valid:

    hh:mm       [AM|PM]
    hh:mm       [AM|PM]
    hh:mm:ss    [AM|PM]
    hh:mm:ss.ss [AM|PM]


The meridian, AM or PM, is optional and case insensitive. A 24 hour time is assumed if the meridian is omitted.

    $time1 = xl-parse-time( '12:18' );
    $time2 = xl-parse-time( '12:18:14' );
    $time3 = xl-parse-time( '12:18:14 AM' );
    $time4 = xl-parse-time( '1:18:14 AM' );

Time in Excel is expressed as a fraction of the day in seconds. Therefore you can calculate an Excel time as follows:

    $time = ( $hours * 3600 + $minutes * 60 + $seconds ) / ( 24 * 60 * 60 );

=head2 xl-parse-date($string)


    Parameters: $string, a textual representation of a date and time

    Returns:    A number that represents an Excel date
                or undef for an invalid date.

    Requires:   Date::Manip and Date::Calc

This function converts a date and time string into a number that represents an Excel date.

The parsing is performed using the C<ParseDate()> function of the L<Date::Manip> module. Refer to the C<Date::Manip> documentation for further information about the date and time formats that can be parsed. In order to use this function you will probably have to initialise some C<Date::Manip> variables via the C<xl-parse-date-init()> function, see below.

    xl-parse-date-init( "TZ=GMT", "DateFormat=non-US" );

    $date1 = xl-parse-date( "11/7/97" );
    $date2 = xl-parse-date( "Friday 11 July 1997" );
    $date3 = xl-parse-date( "10:30 AM Friday 11 July 1997" );
    $date4 = xl-parse-date( "Today" );
    $date5 = xl-parse-date( "Yesterday" );

Note, if you parse a string that represents a time but not a date this function will add the current date. If you want the time without the date you can do something like the following:

    $time  = xl-parse-date( "10:30 AM" );
    $time -= int( $time );

=head2 xl-parse-date-init("variable=value", ...)


    Parameters: A list of Date::Manip variable strings

    Returns:    A list of all the Date::Manip strings

    Requires:   Date::Manip

This function is used to initialise variables required by the L<Date::Manip> module. You should call this function before calling C<xl-parse-date()>. It need only be called once.

This function is a thin wrapper for the C<Date::Manip::Date-Init()> function. You can use C<Date-Init()>  directly if you wish. Refer to the C<Date::Manip> documentation for further information.

    xl-parse-date-init( "TZ=MST", "DateFormat=US" );
    $date1 = xl-parse-date( "11/7/97" );    # November 7th 1997

    xl-parse-date-init( "TZ=GMT", "DateFormat=non-US" );
    $date1 = xl-parse-date( "11/7/97" );    # July 11th 1997

=head2 xl-decode-date-EU($string)


    Parameters: $string, a textual representation of a date and time

    Returns:    A number that represents an Excel date
                or undef for an invalid date.

    Requires:   Date::Calc

This function converts a date and time string into a number that represents an Excel date.

The date parsing is performed using the C<Decode-Date-EU()> function of the L<Date::Calc> module. Refer to the C<Date::Calc> documentation for further information about the date formats that can be parsed. Also note the following from the C<Date::Calc> documentation:

"If the year is given as one or two digits only (i.e., if the year is less than 100), it is mapped to the window 1970 -2069 as follows:"

     0 <= $year <  70  ==>  $year += 2000;
    70 <= $year < 100  ==>  $year += 1900;

The time portion of the string is parsed using the C<xl-parse-time()> function described above.

Note: the EU in the function name means that a European date format is assumed if it is not clear from the string. See the first example below.

    $date1 = xl-decode-date-EU( "11/7/97" );                    #11 July 1997
    $date2 = xl-decode-date-EU( "Sat 12 Sept 1998" );
    $date3 = xl-decode-date-EU( "4:30 AM Sat 12 Sept 1998" );

=head2 xl-decode-date-US($string)


    Parameters: $string, a textual representation of a date and time

    Returns:    A number that represents an Excel date
                or undef for an invalid date.

    Requires:   Date::Calc

This function converts a date and time string into a number that represents an Excel date.

The date parsing is performed using the C<Decode-Date-US()> function of the L<Date::Calc> module. Refer to the C<Date::Calc> documentation for further information about the date formats that can be parsed. Also note the following from the C<Date::Calc> documentation:

"If the year is given as one or two digits only (i.e., if the year is less than 100), it is mapped to the window 1970 -2069 as follows:"

     0 <= $year <  70  ==>  $year += 2000;
    70 <= $year < 100  ==>  $year += 1900;

The time portion of the string is parsed using the C<xl-parse-time()> function described above.

Note: the US in the function name means that an American date format is assumed if it is not clear from the string. See the first example below.

    $date1 = xl-decode-date-US( "11/7/97" );                 # 7 November 1997
    $date2 = xl-decode-date-US( "Sept 12 Saturday 1998" );
    $date3 = xl-decode-date-US( "4:30 AM Sept 12 Sat 1998" );

=head2 xl-date_1904($date)


    Parameters: $date, an Excel date with a 1900 epoch

    Returns:    an Excel date with a 1904 epoch or zero if
                the $date is before 1904


This function converts an Excel date based on the 1900 epoch into a date based on the 1904 epoch.

    $date1 = xl-date-list( 2002, 1, 13 );    # 13 Jan 2002, 1900 epoch
    $date2 = xl-date_1904( $date1 );         # 13 Jan 2002, 1904 epoch

See also the C<set-1904()> workbook method in the L<Excel::Writer::XLSX> documentation.

=head1 REQUIREMENTS

The date and time functions require functions from the L<Date::Manip> and L<Date::Calc> modules. The required functions are "autoused" from these modules so that you do not have to install them unless you wish to use the date and time routines. Therefore it is possible to use the row and column functions without having C<Date::Manip> and C<Date::Calc> installed.

For more information about "autousing" refer to the documentation on the C<autouse> pragma.

=head1 BUGS

When using the autoused functions from C<Date::Manip> and C<Date::Calc> on Perl 5.6.0 with C<-w> you will get a warning like this:

    "Subroutine xxx redefined ..."

The current workaround for this is to put C<use warnings;> near the beginning of your program.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMXVII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=end pod
