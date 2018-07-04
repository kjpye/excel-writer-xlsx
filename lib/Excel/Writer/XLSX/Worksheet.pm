unit class Excel::Writer::XLSX::Worksheet;

###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use v6.c;
use File::Temp;
use Excel::Writer::XLSX::Format;
#use Excel::Writer::XLSX::Drawing;
use Excel::Writer::XLSX::Package::XMLwriter;
use Excel::Writer::XLSX::Utility;

#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';

    my $rowmax = 1_048_576;
    my $colmax = 16_384;
    my $strmax = 32767;

    has $.name;
    has $.index;
    has $!activesheet;
    has $!firstsheet;
    has $!str_total;
    has $!str_unique;
    has $!str_table;
    has $!date_1904;
    has $!palette;
    has $!optimization    = 0;
    has $!tempdir;
    has $!excel2003_style;

    has @!ext_sheets    = [];
    has $!fileclosed    = 0;
    has $!excel_version = 2007;

    has $!xls_rowmax = $rowmax;
    has $!xls_colmax = $colmax;
    has $!xls_strmax = $strmax;
    has $!dim_rowmin = Nil;
    has $!dim_rowmax = Nil;
    has $!dim_colmin = Nil;
    has $!dim_colmax = Nil;

    has %!colinfo    = {};
    has @!selections = [];
    has $.hidden     = 0;
    has $!active     = 0;
    has $!tab_color  = 0;

    has @!panes       = [];
    has $!active_pane = 3;
    has $!selected    = 0;

    has $!page_setup_changed = 0;
    has $!paper_size         = 0;
    has $!orientation        = 1;

    has $!print_options_changed = 0;
    has $!hcenter               = 0;
    has $!vcenter               = 0;
    has $!print_gridlines       = 0;
    has $!screen_gridlines      = 1;
    has $!print_headers         = 0;
    has $!page_view             = 0;

    has $!header_footer_changed = 0;
    has $!header                = '';
    has $!footer                = '';
    has $!header_footer_aligns  = 1;
    has $!header_footer_scales  = 1;
    has @!header_images         = [];
    has @!footer_images         = [];

    has $!margin_left   = 0.7;
    has $!margin_right  = 0.7;
    has $!margin_top    = 0.75;
    has $!margin_bottom = 0.75;
    has $!margin_header = 0.3;
    has $!margin_footer = 0.3;

    has $!repeat_rows = '';
    has $!repeat_cols = '';
    has $!print_area  = '';

    has $!page_order     = 0;
    has $!black_white    = 0;
    has $!draft_quality  = 0;
    has $!print_comments = 0;
    has $!page_start     = 0;

    has $!fit_page   = 0;
    has $!fit_width  = 0;
    has $!fit_height = 0;

    has @!hbreaks = [];
    has @!vbreaks = [];

    has $!protect  = 0;
    has $!password = Nil;

    has %!set_cols = {};
    has %!set_rows = {};

    has $!zoom              = 100;
    has $!zoom_scale_normal = 1;
    has $!print_scale       = 100;
    has $!right_to_left     = 0;
    has $!show_zeros        = 1;
    has $!leading_zeros     = 0;

    has $!outline_row_level = 0;
    has $!outline_col_level = 0;
    has $!outline_style     = 0;
    has $!outline_below     = 1;
    has $!outline_right     = 1;
    has $!outline_on        = 1;
    has $!outline_changed   = 0;

    has $!original_row_height = 15;
    has $!default_row_height  = 15;
    has $!default_row_pixels  = 20;
    has $!default_col_pixels  = 64;
    has $!default_row_zeroed  = 0;

    has %!names = {};

    has @!write_match = [];


    has %!table = {};
    has @!merge = [];

    has $!has_vml             = 0;
    has $!has_header_vml      = 0;
    has $!has_comments        = 0;
    has %!comments            = {};
    has @!comments_array      = [];
    has $!comments_author     = '';
    has $!comments_visible    = 0;
    has $!vml_shape_id        = 1024;
    has @!buttons_array       = [];
    has @!header_images_array = [];

    has $!autofilter   = '';
    has $!filter_on    = 0;
    has @!filter_range = [];
    has %!filter_cols  = {};

    has %!col_sizes        = {};
    has %!row_sizes        = {};
    has %!col_formats      = {};
    has $!col_size_changed = 0;
    has $!row_size_changed = 0;

    has $!last_shape_id          = 1;
    has $!rel_count              = 0;
    has $!hlink_count            = 0;
    has @!hlink_refs             = [];
    has @!external_hyper_links   = [];
    has @!external_drawing_links = [];
    has @!external_comment_links = [];
    has @!external_vml_links     = [];
    has @!external_table_links   = [];
    has @!drawing_links          = [];
    has @!vml_drawing_links      = [];
    has @!charts                 = [];
    has @!images                 = [];
    has @!tables                 = [];
    has @!sparklines             = [];
    has @!shapes                 = [];
    has %!shape_hash             = {};
    has $!has_shapes             = 0;
    has $!drawing                = 0;

    has $!horizontal_dpi = 0;
    has $!vertical_dpi   = 0;

    has $!rstring      = '';
    has $!previous_row = 0;

###############################################################################
#
# Public and private API methods.
#
###############################################################################


###############################################################################
#
# new()
#
# Constructor.
#
#NYI sub new {
#NYI 
#NYI     my $class  = shift;
#NYI     my $fh     = shift;
#NYI     my $self   = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );
#NYI 
#NYI     $self->{_name}            = $_[0];
#NYI     $self->{_index}           = $_[1];
#NYI     $self->{_activesheet}     = $_[2];
#NYI     $self->{_firstsheet}      = $_[3];
#NYI     $self->{_str_total}       = $_[4];
#NYI     $self->{_str_unique}      = $_[5];
#NYI     $self->{_str_table}       = $_[6];
#NYI     $self->{_date_1904}       = $_[7];
#NYI     $self->{_palette}         = $_[8];
#NYI     $self->{_optimization}    = $_[9] || 0;
#NYI     $self->{_tempdir}         = $_[10];
#NYI     $self->{_excel2003_style} = $_[11];

#NYI     if ( $self->{_optimization} == 1 ) {
#NYI         my $fh = tempfile( DIR => $self->{_tempdir} );
#NYI         binmode $fh, ':utf8';
#NYI 
#NYI         $self->{_cell_data_fh} = $fh;
#NYI         $self->{_fh}           = $fh;
#NYI     }
#NYI 
#NYI     $self->{_validations}  = [];
#NYI     $self->{_cond_formats} = {};
#NYI     $self->{_dxf_priority} = 1;
#NYI 
#NYI     if ( $self->{_excel2003_style} ) {
#NYI         $self->{_original_row_height}  = 12.75;
#NYI         $self->{_default_row_height}   = 12.75;
#NYI         $self->{_default_row_pixels}   = 17;
#NYI         $self->{_margin_left}          = 0.75;
#NYI         $self->{_margin_right}         = 0.75;
#NYI         $self->{_margin_top}           = 1;
#NYI         $self->{_margin_bottom}        = 1;
#NYI         $self->{_margin_header}        = 0.5;
#NYI         $self->{_margin_footer}        = 0.5;
#NYI         $self->{_header_footer_aligns} = 0;
#NYI     }
#NYI 
#NYI     bless $self, $class;
#NYI     return $self;
#NYI }
#NYI 
###############################################################################
#
# set_xml_writer()
#
# Over-ridden to ensure that write_single_row() is called for the final row
# when optimisation mode is on.
#
method set_xml_writer($filename) {

    if $!optimization == 1 {
        self.write_single_row();
    }

    self.SUPER::set_xml_writer( $filename ); # TODO
}


###############################################################################
#
# assemble_xml_file()
#
# Assemble and write the XML file.
#
method assemble_xml_file {

    self.xml_declaration();

    # Write the root worksheet element.
    self.write_worksheet();

    # Write the worksheet properties.
    self.write_sheet_pr();

    # Write the worksheet dimensions.
    self.write_dimension();

    # Write the sheet view properties.
    self.write_sheet_views();

    # Write the sheet format properties.
    self.write_sheet_format_pr();

    # Write the sheet column info.
    self.write_cols();

    # Write the worksheet data such as rows columns and cells.
    if $!optimization == 0 {
        self.write_sheet_data();
    }
    else {
        self.write_optimized_sheet_data();
    }

    # Write the sheetProtection element.
    self.write_sheet_protection();

    # Write the worksheet calculation properties.
    #$self->_write_sheet_calc_pr();

    # Write the worksheet phonetic properties.
    if $!excel2003_style {
        self.write_phonetic_pr();
    }

    # Write the autoFilter element.
    self.write_auto_filter();

    # Write the mergeCells element.
    self.write_merge_cells();

    # Write the conditional formats.
    self.write_conditional_formats();

    # Write the dataValidations element.
    self.write_data_validations();

    # Write the hyperlink element.
    self.write_hyperlinks();

    # Write the printOptions element.
    self.write_print_options();

    # Write the worksheet page_margins.
    self.write_page_margins();

    # Write the worksheet page setup.
    self.write_page_setup();

    # Write the headerFooter element.
    self.write_header_footer();

    # Write the rowBreaks element.
    self.write_row_breaks();

    # Write the colBreaks element.
    self.write_col_breaks();

    # Write the drawing element.
    self.write_drawings();

    # Write the legacyDrawing element.
    self.write_legacy_drawing();

    # Write the legacyDrawingHF element.
    self.write_legacy_drawing_hf();

    # Write the tableParts element.
    self.write_table_parts();

    # Write the extLst and sparklines.
    self.write_ext_sparklines();

    # Close the worksheet tag.
    self.xml_end_tag( 'worksheet' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# _close()
#
# Write the worksheet elements.
#
#NYI sub _close {
#NYI 
#NYI     # TODO. Unused. Remove after refactoring.
#NYI     my $self       = shift;
#NYI     my $sheetnames = shift;
#NYI     my $num_sheets = scalar @$sheetnames;
#NYI }
#NYI 
#NYI 
###############################################################################
#
# get_name().
#
# Retrieve the worksheet name.
#
method get_name {
    return $!name;
}


###############################################################################
#
# select()
#
# Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
# highlighted.
#
method select {

    $!hidden   = 0;    # Selected worksheet can't be hidden.
    $!selected = 1;
}


###############################################################################
#
# activate()
#
# Set this worksheet as the active worksheet, i.e. the worksheet that is
# displayed when the workbook is opened. Also set it as selected.
#
method activate {
    $!hidden   = 0;    # Active worksheet can't be hidden.
    $!selected = 1;
    $!activesheet = $!index;
}


###############################################################################
#
# hide()
#
# Hide this worksheet.
#
method hide {
    $!hidden = 1;

    # A hidden worksheet shouldn't be active or selected.
    $!selected    = 0;
    $!activesheet = 0;
    $!firstsheet  = 0;
}


###############################################################################
#
# set_first_sheet()
#
# Set this worksheet as the first visible sheet. This is necessary
# when there are a large number of worksheets and the activated
# worksheet is not visible on the screen.
#
method set_first_sheet {
    $!hidden = 0;    # Active worksheet can't be hidden.
    $!firstsheet = $!index;
}


###############################################################################
#
# protect( $password )
#
# Set the worksheet protection flags to prevent modification of worksheet
# objects.
#
#NYI sub protect {
#NYI 
#NYI     my $self     = shift;
#NYI     my $password = shift || '';
#NYI     my $options  = shift || {};
#NYI 
#NYI     if ( $password ne '' ) {
#NYI         $password = $self->_encode_password( $password );
#NYI     }
#NYI 
#NYI     # Default values for objects that can be protected.
#NYI     my %defaults = (
#NYI         sheet                 => 1,
#NYI         content               => 0,
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
#NYI 
#NYI     # Overwrite the defaults with user specified values.
#NYI     for my $key ( keys %{$options} ) {
#NYI 
#NYI         if ( exists $defaults{$key} ) {
#NYI             $defaults{$key} = $options->{$key};
#NYI         }
#NYI         else {
#NYI             warn "Unknown protection object: $key\n";
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the password after the user defined values.
#NYI     $defaults{password} = $password;
#NYI 
#NYI     $self->{_protect} = \%defaults;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _encode_password($password)
#NYI #
#NYI # Based on the algorithm provided by Daniel Rentz of OpenOffice.
#NYI #
#NYI sub _encode_password {
#NYI 
#NYI     use integer;
#NYI 
#NYI     my $self      = shift;
#NYI     my $plaintext = $_[0];
#NYI     my $password;
#NYI     my $count;
#NYI     my @chars;
#NYI     my $i = 0;
#NYI 
#NYI     $count = @chars = split //, $plaintext;
#NYI 
#NYI     foreach my $char ( @chars ) {
#NYI         my $low_15;
#NYI         my $high_15;
#NYI         $char    = ord( $char ) << ++$i;
#NYI         $low_15  = $char & 0x7fff;
#NYI         $high_15 = $char & 0x7fff << 15;
#NYI         $high_15 = $high_15 >> 15;
#NYI         $char    = $low_15 | $high_15;
#NYI     }
#NYI 
#NYI     $password = 0x0000;
#NYI     $password ^= $_ for @chars;
#NYI     $password ^= $count;
#NYI     $password ^= 0xCE4B;
#NYI 
#NYI     return sprintf "%X", $password;
#NYI }


# TODO: Check argument usage

###############################################################################
#
# set_column($firstcol, $lastcol, $width, $format, $hidden, $level)
#
# Set the width of a single column or a range of columns.
# See also: _write_col_info
#
method set_column(@data) {

    my $cell = @data[0];

    # Check for a cell reference in A1 notation and substitute row and column
    if $cell ~~ /^\D/ {
        @data = self.substitute_cellref( @data );

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift @data;    # $row1
        splice @data, 1, 1;    # $row2 # TODO
    }

    return if @data.elems < 3;       # Ensure at least $firstcol, $lastcol and $width
    return if not @data[0].defined;    # Columns must be defined.
    return if not @data[1].defined;

    # Assume second column is the same as first if 0. Avoids KB918419 bug.
    @data[1] = @data[0] if @data[1] == 0;

    # Ensure 2nd col is larger than first. Also for KB918419 bug.
    ( @data[0], @data[1] ) = ( @data[1], @data[0] ) if @data[0] > @data[1];


    # Check that cols are valid and store max and min values with default row.
    # NOTE: The check shouldn't modify the row dimensions and should only modify
    #       the column dimensions in certain cases.
    my $ignore_row = 1;
    my $ignore_col = 1;
#TODO: Fix next two lines
    $ignore_col = 0 if @data[3].defined;          # Column has a format.
    $ignore_col = 0 if @data[2] && @data[4];  # Column has a width but is hidden

    return -2
      if self.check_dimensions( 0, @data[0], $ignore_row, $ignore_col );
    return -2
      if self.check_dimensions( 0, @data[1], $ignore_row, $ignore_col );

    # Set the limits for the outline levels (0 <= x <= 7).
    @data[5] = 0 unless @data[5].defined;
    @data[5] = 0 if @data[5] < 0;
    @data[5] = 7 if @data[5] > 7;

    if @data[5] > $!outline_col_level {
        $!outline_col_level = @data[5];
    }

    # Store the column data based on the first column. Padded for sorting.
    %!colinfo{ sprintf "%05d", @data[0] } = [@data]; # TODO

    # Store the column change to allow optimisations.
    $!col_size_changed = 1;

    # Store the col sizes for use when calculating image vertices taking
    # hidden columns into account. Also store the column formats.
    my $width = @data[4] ?? 0 !! @data[2];    # Set width to zero if hidden.
    my $format = @data[3];

    my ( $firstcol, $lastcol ) = @data;

    for $firstcol .. $lastcol -> $col {
        %!col_sizes{$col} = $width;
        %!col_formats{$col} = $format if $format;
    }
}


#NYI ###############################################################################
#NYI #
#NYI # set_selection()
#NYI #
#NYI # Set which cell or cells are selected in a worksheet.
#NYI #
#NYI sub set_selection {
#NYI 
#NYI     my $self = shift;
#NYI     my $pane;
#NYI     my $active_cell;
#NYI     my $sqref;
#NYI 
#NYI     return unless @_;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column.
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI 
#NYI     # There should be either 2 or 4 arguments.
#NYI     if ( @_ == 2 ) {
#NYI 
#NYI         # Single cell selection.
#NYI         $active_cell = xl-rowcol-to-cell( $_[0], $_[1] );
#NYI         $sqref = $active_cell;
#NYI     }
#NYI     elsif ( @_ == 4 ) {
#NYI 
#NYI         # Range selection.
#NYI         $active_cell = xl-rowcol-to-cell( $_[0], $_[1] );
#NYI 
#NYI         my ( $row_first, $col_first, $row_last, $col_last ) = @_;
#NYI 
#NYI         # Swap last row/col for first row/col as necessary
#NYI         if ( $row_first > $row_last ) {
#NYI             ( $row_first, $row_last ) = ( $row_last, $row_first );
#NYI         }
#NYI 
#NYI         if ( $col_first > $col_last ) {
#NYI             ( $col_first, $col_last ) = ( $col_last, $col_first );
#NYI         }
#NYI 
#NYI         # If the first and last cell are the same write a single cell.
#NYI         if ( ( $row_first == $row_last ) && ( $col_first == $col_last ) ) {
#NYI             $sqref = $active_cell;
#NYI         }
#NYI         else {
#NYI             $sqref = xl-range( $row_first, $row_last, $col_first, $col_last );
#NYI         }
#NYI 
#NYI     }
#NYI     else {
#NYI 
#NYI         # User supplied wrong number or arguments.
#NYI         return;
#NYI     }
#NYI 
#NYI     # Selection isn't set for cell A1.
#NYI     return if $sqref eq 'A1';
#NYI 
#NYI     $self->{_selections} = [ [ $pane, $active_cell, $sqref ] ];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # freeze_panes( $row, $col, $top_row, $left_col )
#NYI #
#NYI # Set panes and mark them as frozen.
#NYI #
#NYI sub freeze_panes {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return unless @_;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column.
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     my $row      = shift;
#NYI     my $col      = shift || 0;
#NYI     my $top_row  = shift || $row;
#NYI     my $left_col = shift || $col;
#NYI     my $type     = shift || 0;
#NYI 
#NYI     $self->{_panes} = [ $row, $col, $top_row, $left_col, $type ];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # split_panes( $y, $x, $top_row, $left_col )
#NYI #
#NYI # Set panes and mark them as split.
#NYI #
#NYI # Implementers note. The API for this method doesn't map well from the XLS
#NYI # file format and isn't sufficient to describe all cases of split panes.
#NYI # It should probably be something like:
#NYI #
#NYI #     split_panes( $y, $x, $top_row, $left_col, $offset_row, $offset_col )
#NYI #
#NYI # I'll look at changing this if it becomes an issue.
#NYI #
#NYI sub split_panes {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Call freeze panes but add the type flag for split panes.
#NYI     $self->freeze_panes( @_[ 0 .. 3 ], 2 );
#NYI }
#NYI 
#NYI # Older method name for backwards compatibility.
#NYI *thaw_panes = *split_panes;
#NYI 
#NYI 
###############################################################################
#
# set_portrait()
#
# Set the page orientation as portrait.
#
method set_portrait {
    $!orientation        = 1;
    $!page_setup_changed = 1;
}


###############################################################################
#
# set_landscape()
#
# Set the page orientation as landscape.
#
method set_landscape {
    $!orientation        = 0;
    $!page_setup_changed = 1;
}


###############################################################################
#
# set_page_view()
#
# Set the page view mode for Mac Excel.
#
method set_page_view($view = 1) {
    $!page_view = $view;
}


###############################################################################
#
# set_tab_color()
#
# Set the colour of the worksheet tab.
#
method set_tab_color($colour) {
    $!tab_color = Excel::Writer::XLSX::Format::get_color( $colour );
}


###############################################################################
#
# set_paper()
#
# Set the paper type. Ex. 1 = US Letter, 9 = A4
#
method set_paper($paper-size) {
    if $paper-size {
        $!paper_size         = $paper-size;
        $!page_setup_changed = 1;
    }
}


#NYI ###############################################################################
#NYI #
#NYI # set_header()
#NYI #
#NYI # Set the page header caption and optional margin.
#NYI #
#NYI sub set_header {
#NYI 
#NYI     my $self    = shift;
#NYI     my $string  = $_[0] || '';
#NYI     my $margin  = $_[1] || 0.3;
#NYI     my $options = $_[2] || {};
#NYI 
#NYI 
#NYI     # Replace the Excel placeholder &[Picture] with the internal &G.
#NYI     $string =~ s/&\[Picture\]/&G/g;
#NYI 
#NYI     if ( length $string >= 255 ) {
#NYI         warn 'Header string must be less than 255 characters';
#NYI         return;
#NYI     }
#NYI 
#NYI     if ( defined $options->{align_with_margins} ) {
#NYI         $self->{_header_footer_aligns} = $options->{align_with_margins};
#NYI     }
#NYI 
#NYI     if ( defined $options->{scale_with_doc} ) {
#NYI         $self->{_header_footer_scales} = $options->{scale_with_doc};
#NYI     }
#NYI 
#NYI     # Reset the array in case the function is called more than once.
#NYI     $self->{_header_images} = [];
#NYI 
#NYI     if ( $options->{image_left} ) {
#NYI         push @{ $self->{_header_images} }, [ $options->{image_left}, 'LH' ];
#NYI     }
#NYI 
#NYI     if ( $options->{image_center} ) {
#NYI         push @{ $self->{_header_images} }, [ $options->{image_center}, 'CH' ];
#NYI     }
#NYI 
#NYI     if ( $options->{image_right} ) {
#NYI         push @{ $self->{_header_images} }, [ $options->{image_right}, 'RH' ];
#NYI     }
#NYI 
#NYI     my $placeholder_count = () = $string =~ /&G/g;
#NYI     my $image_count = @{ $self->{_header_images} };
#NYI 
#NYI     if ( $image_count != $placeholder_count ) {
#NYI         warn "Number of header images ($image_count) doesn't match placeholder "
#NYI           . "count ($placeholder_count) in string: $string\n";
#NYI         $self->{_header_images} = [];
#NYI         return;
#NYI     }
#NYI 
#NYI     if ( $image_count ) {
#NYI         $self->{_has_header_vml} = 1;
#NYI     }
#NYI 
#NYI     $self->{_header}                = $string;
#NYI     $self->{_margin_header}         = $margin;
#NYI     $self->{_header_footer_changed} = 1;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_footer()
#NYI #
#NYI # Set the page footer caption and optional margin.
#NYI #
#NYI sub set_footer {
#NYI 
#NYI     my $self    = shift;
#NYI     my $string  = $_[0] || '';
#NYI     my $margin  = $_[1] || 0.3;
#NYI     my $options = $_[2] || {};
#NYI 
#NYI 
#NYI     # Replace the Excel placeholder &[Picture] with the internal &G.
#NYI     $string =~ s/&\[Picture\]/&G/g;
#NYI 
#NYI     if ( length $string >= 255 ) {
#NYI         warn 'Footer string must be less than 255 characters';
#NYI         return;
#NYI     }
#NYI 
#NYI     if ( defined $options->{align_with_margins} ) {
#NYI         $self->{_header_footer_aligns} = $options->{align_with_margins};
#NYI     }
#NYI 
#NYI     if ( defined $options->{scale_with_doc} ) {
#NYI         $self->{_header_footer_scales} = $options->{scale_with_doc};
#NYI     }
#NYI 
#NYI     # Reset the array in case the function is called more than once.
#NYI     $self->{_footer_images} = [];
#NYI 
#NYI     if ( $options->{image_left} ) {
#NYI         push @{ $self->{_footer_images} }, [ $options->{image_left}, 'LF' ];
#NYI     }
#NYI 
#NYI     if ( $options->{image_center} ) {
#NYI         push @{ $self->{_footer_images} }, [ $options->{image_center}, 'CF' ];
#NYI     }
#NYI 
#NYI     if ( $options->{image_right} ) {
#NYI         push @{ $self->{_footer_images} }, [ $options->{image_right}, 'RF' ];
#NYI     }
#NYI 
#NYI     my $placeholder_count = () = $string =~ /&G/g;
#NYI     my $image_count = @{ $self->{_footer_images} };
#NYI 
#NYI     if ( $image_count != $placeholder_count ) {
#NYI         warn "Number of footer images ($image_count) doesn't match placeholder "
#NYI           . "count ($placeholder_count) in string: $string\n";
#NYI         $self->{_footer_images} = [];
#NYI         return;
#NYI     }
#NYI 
#NYI     if ( $image_count ) {
#NYI         $self->{_has_header_vml} = 1;
#NYI     }
#NYI 
#NYI     $self->{_footer}                = $string;
#NYI     $self->{_margin_footer}         = $margin;
#NYI     $self->{_header_footer_changed} = 1;
#NYI }


###############################################################################
#
# center_horizontally()
#
# Center the page horizontally.
#
method center_horizontally {
    $!hcenter               = 1;
    $!print_options_changed = 1;
}


###############################################################################
#
# center_vertically()
#
# Center the page horizontally.
#
method center_vertically {
    $!vcenter               = 1;
    $!print_options_changed = 1;
}


###############################################################################
#
# set_margins()
#
# Set all the page margins to the same value in inches.
#
method set_margins($margin) {
    self.set_margin_left( $margin );
    self.set_margin_right( $margin );
    self.set_margin_top( $margin );
    self.set_margin_bottom( $margin );
}


###############################################################################
#
# set_margins_LR()
#
# Set the left and right margins to the same value in inches.
#
method set_margins_LR($margin) {
    self.set_margin_left( $margin );
    self.set_margin_right( $margin );
}


###############################################################################
#
# set_margins_TB()
#
# Set the top and bottom margins to the same value in inches.
#
method set_margins_TB($margin) {
    self.set_margin_top( $margin );
    self.set_margin_bottom( $margin );
}


###############################################################################
#
# set_margin_left()
#
# Set the left margin in inches.
#
method set_margin_left($margin = 0.7) {
    $!margin_left = +$margin;
}


###############################################################################
#
# set_margin_right()
#
# Set the right margin in inches.
#
method set_margin_right($margin = 0.7) {
    $!margin_right = +$margin;
}


###############################################################################
#
# set_margin_top()
#
# Set the top margin in inches.
#
method set_margin_top($margin = 0.75) {
    $!margin_top = +$margin;
}


###############################################################################
#
# set_margin_bottom()
#
# Set the bottom margin in inches.
#
method set_margin_bottom($margin = 0.75) {
    $!margin_bottom = +$margin;
}


###############################################################################
#
# repeat_rows($first_row, $last_row)
#
# Set the rows to repeat at the top of each printed page.
#
method repeat_rows($row-min, $row-max) {
    $row-max //= $row-min; # row-max is optional

    # Convert to 1 based.
    $row-min++;
    $row-max++;

    my $area = '$' ~ $row-min ~ ':' ~ '$' ~ $row-max;

    # Build up the print titles "Sheet1!$1:$2"
    my $sheetname = quote-sheetname( $!name );
    $area = $sheetname ~ "!" ~ $area;

    $!repeat_rows = $area;
}


###############################################################################
#
# repeat_columns($first_col, $last_col)
#
# Set the columns to repeat at the left hand side of each printed page. This is
# stored as a <NamedRange> element.
#
method repeat_columns($col-min, $col-max) {
    # Check for a cell reference in A1 notation and substitute row and column
    if $col-min ~~ /^\D/ {
        (Nil, $col-min, Nil, $col-max) = self.substitute_cellref( $col-min, $col-max );
    }

    $col-max //= $col-min;    # Second col is optional

    # Convert to A notation.
    $col-min = xl-col-to-name( $col-min, 1 );
    $col-max = xl-col-to-name( $col-max, 1 );

    my $area = $col-min ~ ':' ~ $col-max;

    # Build up the print area range "=Sheet2!C1:C2"
    my $sheetname = quote-sheetname( $!name );
    $area = $sheetname ~ "!" ~ $area;

    $!repeat_cols = $area;
}


#NYI ###############################################################################
#NYI #
#NYI # print_area($first_row, $first_col, $last_row, $last_col)
#NYI #
#NYI # Set the print area in the current worksheet. This is stored as a <NamedRange>
#NYI # element.
#NYI #
#NYI sub print_area {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     return if @_ != 4;    # Require 4 parameters
#NYI 
#NYI     my ( $row1, $col1, $row2, $col2 ) = @_;
#NYI 
#NYI     # Ignore max print area since this is the same as no print area for Excel.
#NYI     if (    $row1 == 0
#NYI         and $col1 == 0
#NYI         and $row2 == $self->{_xls_rowmax} - 1
#NYI         and $col2 == $self->{_xls_colmax} - 1 )
#NYI     {
#NYI         return;
#NYI     }
#NYI 
#NYI     # Build up the print area range "=Sheet2!R1C1:R2C1"
#NYI     my $area = $self->_convert_name_area( $row1, $col1, $row2, $col2 );
#NYI 
#NYI     $self->{_print_area} = $area;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # autofilter($first_row, $first_col, $last_row, $last_col)
#NYI #
#NYI # Set the autofilter area in the worksheet.
#NYI #
#NYI sub autofilter {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     return if @_ != 4;    # Require 4 parameters
#NYI 
#NYI     my ( $row1, $col1, $row2, $col2 ) = @_;
#NYI 
#NYI     # Reverse max and min values if necessary.
#NYI     ( $row1, $row2 ) = ( $row2, $row1 ) if $row2 < $row1;
#NYI     ( $col1, $col2 ) = ( $col2, $col1 ) if $col2 < $col1;
#NYI 
#NYI     # Build up the print area range "Sheet1!$A$1:$C$13".
#NYI     my $area = $self->_convert_name_area( $row1, $col1, $row2, $col2 );
#NYI     my $ref = xl-range( $row1, $row2, $col1, $col2 );
#NYI 
#NYI     $self->{_autofilter}     = $area;
#NYI     $self->{_autofilter_ref} = $ref;
#NYI     $self->{_filter_range}   = [ $col1, $col2 ];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # filter_column($column, $criteria, ...)
#NYI #
#NYI # Set the column filter criteria.
#NYI #
#NYI sub filter_column {
#NYI 
#NYI     my $self       = shift;
#NYI     my $col        = $_[0];
#NYI     my $expression = $_[1];
#NYI 
#NYI     fail "Must call autofilter() before filter_column()"
#NYI       unless $self->{_autofilter};
#NYI     fail "Incorrect number of arguments to filter_column()"
#NYI       unless @_ == 2;
#NYI 
#NYI 
#NYI     # Check for a column reference in A1 notation and substitute.
#NYI     if ( $col =~ /^\D/ ) {
#NYI         my $col_letter = $col;
#NYI 
#NYI         # Convert col ref to a cell ref and then to a col number.
#NYI         ( undef, $col ) = $self->_substitute_cellref( $col . '1' );
#NYI 
#NYI         fail "Invalid column '$col_letter'" if $col >= $self->{_xls_colmax};
#NYI     }
#NYI 
#NYI     my ( $col_first, $col_last ) = @{ $self->{_filter_range} };
#NYI 
#NYI     # Reject column if it is outside filter range.
#NYI     if ( $col < $col_first or $col > $col_last ) {
#NYI         fail "Column '$col' outside autofilter() column range "
#NYI           . "($col_first .. $col_last)";
#NYI     }
#NYI 
#NYI 
#NYI     my @tokens = $self->_extract_filter_tokens( $expression );
#NYI 
#NYI     fail "Incorrect number of tokens in expression '$expression'"
#NYI       unless ( @tokens == 3 or @tokens == 7 );
#NYI 
#NYI 
#NYI     @tokens = $self->_parse_filter_expression( $expression, @tokens );
#NYI 
#NYI     # Excel handles single or double custom filters as default filters. We need
#NYI     # to check for them and handle them accordingly.
#NYI     if ( @tokens == 2 && $tokens[0] == 2 ) {
#NYI 
#NYI         # Single equality.
#NYI         $self->filter_column_list( $col, $tokens[1] );
#NYI     }
#NYI     elsif (@tokens == 5
#NYI         && $tokens[0] == 2
#NYI         && $tokens[2] == 1
#NYI         && $tokens[3] == 2 )
#NYI     {
#NYI 
#NYI         # Double equality with "or" operator.
#NYI         $self->filter_column_list( $col, $tokens[1], $tokens[4] );
#NYI     }
#NYI     else {
#NYI 
#NYI         # Non default custom filter.
#NYI         $self->{_filter_cols}->{$col} = [@tokens];
#NYI         $self->{_filter_type}->{$col} = 0;
#NYI 
#NYI     }
#NYI 
#NYI     $self->{_filter_on} = 1;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # filter_column_list($column, @matches )
#NYI #
#NYI # Set the column filter criteria in Excel 2007 list style.
#NYI #
#NYI sub filter_column_list {
#NYI 
#NYI     my $self   = shift;
#NYI     my $col    = shift;
#NYI     my @tokens = @_;
#NYI 
#NYI     fail "Must call autofilter() before filter_column_list()"
#NYI       unless $self->{_autofilter};
#NYI     fail "Incorrect number of arguments to filter_column_list()"
#NYI       unless @tokens;
#NYI 
#NYI     # Check for a column reference in A1 notation and substitute.
#NYI     if ( $col =~ /^\D/ ) {
#NYI         my $col_letter = $col;
#NYI 
#NYI         # Convert col ref to a cell ref and then to a col number.
#NYI         ( undef, $col ) = $self->_substitute_cellref( $col . '1' );
#NYI 
#NYI         fail "Invalid column '$col_letter'" if $col >= $self->{_xls_colmax};
#NYI     }
#NYI 
#NYI     my ( $col_first, $col_last ) = @{ $self->{_filter_range} };
#NYI 
#NYI     # Reject column if it is outside filter range.
#NYI     if ( $col < $col_first or $col > $col_last ) {
#NYI         fail "Column '$col' outside autofilter() column range "
#NYI           . "($col_first .. $col_last)";
#NYI     }
#NYI 
#NYI     $self->{_filter_cols}->{$col} = [@tokens];
#NYI     $self->{_filter_type}->{$col} = 1;           # Default style.
#NYI     $self->{_filter_on}           = 1;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _extract_filter_tokens($expression)
#NYI #
#NYI # Extract the tokens from the filter expression. The tokens are mainly non-
#NYI # whitespace groups. The only tricky part is to extract string tokens that
#NYI # contain whitespace and/or quoted double quotes (Excel's escaped quotes).
#NYI #
#NYI # Examples: 'x <  2000'
#NYI #           'x >  2000 and x <  5000'
#NYI #           'x = "foo"'
#NYI #           'x = "foo bar"'
#NYI #           'x = "foo "" bar"'
#NYI #
#NYI sub _extract_filter_tokens {
#NYI 
#NYI     my $self       = shift;
#NYI     my $expression = $_[0];
#NYI 
#NYI     return unless $expression;
#NYI 
#NYI     my @tokens = ( $expression =~ /"(?:[^"]|"")*"|\S+/g );    #"
#NYI 
#NYI     # Remove leading and trailing quotes and unescape other quotes
#NYI     for ( @tokens ) {
#NYI         s/^"//;                                               #"
#NYI         s/"$//;                                               #"
#NYI         s/""/"/g;                                             #"
#NYI     }
#NYI 
#NYI     return @tokens;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _parse_filter_expression(@token)
#NYI #
#NYI # Converts the tokens of a possibly conditional expression into 1 or 2
#NYI # sub expressions for further parsing.
#NYI #
#NYI # Examples:
#NYI #          ('x', '==', 2000) -> exp1
#NYI #          ('x', '>',  2000, 'and', 'x', '<', 5000) -> exp1 and exp2
#NYI #
#NYI sub _parse_filter_expression {
#NYI 
#NYI     my $self       = shift;
#NYI     my $expression = shift;
#NYI     my @tokens     = @_;
#NYI 
#NYI     # The number of tokens will be either 3 (for 1 expression)
#NYI     # or 7 (for 2  expressions).
#NYI     #
#NYI     if ( @tokens == 7 ) {
#NYI 
#NYI         my $conditional = $tokens[3];
#NYI 
#NYI         if ( $conditional =~ /^(and|&&)$/ ) {
#NYI             $conditional = 0;
#NYI         }
#NYI         elsif ( $conditional =~ /^(or|\|\|)$/ ) {
#NYI             $conditional = 1;
#NYI         }
#NYI         else {
#NYI             fail "Token '$conditional' is not a valid conditional "
#NYI               . "in filter expression '$expression'";
#NYI         }
#NYI 
#NYI         my @expression_1 =
#NYI           $self->_parse_filter_tokens( $expression, @tokens[ 0, 1, 2 ] );
#NYI         my @expression_2 =
#NYI           $self->_parse_filter_tokens( $expression, @tokens[ 4, 5, 6 ] );
#NYI 
#NYI         return ( @expression_1, $conditional, @expression_2 );
#NYI     }
#NYI     else {
#NYI         return $self->_parse_filter_tokens( $expression, @tokens );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _parse_filter_tokens(@token)
#NYI #
#NYI # Parse the 3 tokens of a filter expression and return the operator and token.
#NYI #
#NYI sub _parse_filter_tokens {
#NYI 
#NYI     my $self       = shift;
#NYI     my $expression = shift;
#NYI     my @tokens     = @_;
#NYI 
#NYI     my %operators = (
#NYI         '==' => 2,
#NYI         '='  => 2,
#NYI         '=~' => 2,
#NYI         'eq' => 2,
#NYI 
#NYI         '!=' => 5,
#NYI         '!~' => 5,
#NYI         'ne' => 5,
#NYI         '<>' => 5,
#NYI 
#NYI         '<'  => 1,
#NYI         '<=' => 3,
#NYI         '>'  => 4,
#NYI         '>=' => 6,
#NYI     );
#NYI 
#NYI     my $operator = $operators{ $tokens[1] };
#NYI     my $token    = $tokens[2];
#NYI 
#NYI 
#NYI     # Special handling of "Top" filter expressions.
#NYI     if ( $tokens[0] =~ /^top|bottom$/i ) {
#NYI 
#NYI         my $value = $tokens[1];
#NYI 
#NYI         if (   $value =~ /\D/
#NYI             or $value < 1
#NYI             or $value > 500 )
#NYI         {
#NYI             fail "The value '$value' in expression '$expression' "
#NYI               . "must be in the range 1 to 500";
#NYI         }
#NYI 
#NYI         $token = lc $token;
#NYI 
#NYI         if ( $token ne 'items' and $token ne '%' ) {
#NYI             fail "The type '$token' in expression '$expression' "
#NYI               . "must be either 'items' or '%'";
#NYI         }
#NYI 
#NYI         if ( $tokens[0] =~ /^top$/i ) {
#NYI             $operator = 30;
#NYI         }
#NYI         else {
#NYI             $operator = 32;
#NYI         }
#NYI 
#NYI         if ( $tokens[2] eq '%' ) {
#NYI             $operator++;
#NYI         }
#NYI 
#NYI         $token = $value;
#NYI     }
#NYI 
#NYI 
#NYI     if ( not $operator and $tokens[0] ) {
#NYI         fail "Token '$tokens[1]' is not a valid operator "
#NYI           . "in filter expression '$expression'";
#NYI     }
#NYI 
#NYI 
#NYI     # Special handling for Blanks/NonBlanks.
#NYI     if ( $token =~ /^blanks|nonblanks$/i ) {
#NYI 
#NYI         # Only allow Equals or NotEqual in this context.
#NYI         if ( $operator != 2 and $operator != 5 ) {
#NYI             fail "The operator '$tokens[1]' in expression '$expression' "
#NYI               . "is not valid in relation to Blanks/NonBlanks'";
#NYI         }
#NYI 
#NYI         $token = lc $token;
#NYI 
#NYI         # The operator should always be 2 (=) to flag a "simple" equality in
#NYI         # the binary record. Therefore we convert <> to =.
#NYI         if ( $token eq 'blanks' ) {
#NYI             if ( $operator == 5 ) {
#NYI                 $token = ' ';
#NYI             }
#NYI         }
#NYI         else {
#NYI             if ( $operator == 5 ) {
#NYI                 $operator = 2;
#NYI                 $token    = 'blanks';
#NYI             }
#NYI             else {
#NYI                 $operator = 5;
#NYI                 $token    = ' ';
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # if the string token contains an Excel match character then change the
#NYI     # operator type to indicate a non "simple" equality.
#NYI     if ( $operator == 2 and $token =~ /[*?]/ ) {
#NYI         $operator = 22;
#NYI     }
#NYI 
#NYI 
#NYI     return ( $operator, $token );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _convert_name_area($first_row, $first_col, $last_row, $last_col)
#NYI #
#NYI # Convert zero indexed rows and columns to the format required by worksheet
#NYI # named ranges, eg, "Sheet1!$A$1:$C$13".
#NYI #
#NYI sub _convert_name_area {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $row_num_1 = $_[0];
#NYI     my $col_num_1 = $_[1];
#NYI     my $row_num_2 = $_[2];
#NYI     my $col_num_2 = $_[3];
#NYI 
#NYI     my $range1       = '';
#NYI     my $range2       = '';
#NYI     my $row_col_only = 0;
#NYI     my $area;
#NYI 
#NYI     # Convert to A1 notation.
#NYI     my $col_char_1 = xl-col-to-name( $col_num_1, 1 );
#NYI     my $col_char_2 = xl-col-to-name( $col_num_2, 1 );
#NYI     my $row_char_1 = '$' . ( $row_num_1 + 1 );
#NYI     my $row_char_2 = '$' . ( $row_num_2 + 1 );
#NYI 
#NYI     # We need to handle some special cases that refer to rows or columns only.
#NYI     if ( $row_num_1 == 0 and $row_num_2 == $self->{_xls_rowmax} - 1 ) {
#NYI         $range1       = $col_char_1;
#NYI         $range2       = $col_char_2;
#NYI         $row_col_only = 1;
#NYI     }
#NYI     elsif ( $col_num_1 == 0 and $col_num_2 == $self->{_xls_colmax} - 1 ) {
#NYI         $range1       = $row_char_1;
#NYI         $range2       = $row_char_2;
#NYI         $row_col_only = 1;
#NYI     }
#NYI     else {
#NYI         $range1 = $col_char_1 . $row_char_1;
#NYI         $range2 = $col_char_2 . $row_char_2;
#NYI     }
#NYI 
#NYI     # A repeated range is only written once (if it isn't a special case).
#NYI     if ( $range1 eq $range2 && !$row_col_only ) {
#NYI         $area = $range1;
#NYI     }
#NYI     else {
#NYI         $area = $range1 . ':' . $range2;
#NYI     }
#NYI 
#NYI     # Build up the print area range "Sheet1!$A$1:$C$13".
#NYI     my $sheetname = quote-sheetname( $self->{_name} );
#NYI     $area = $sheetname . "!" . $area;
#NYI 
#NYI     return $area;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # hide_gridlines()
#NYI #
#NYI # Set the option to hide gridlines on the screen and the printed page.
#NYI #
#NYI # This was mainly useful for Excel 5 where printed gridlines were on by
#NYI # default.
#NYI #
#NYI sub hide_gridlines {
#NYI 
#NYI     my $self = shift;
#NYI     my $option =
#NYI       defined $_[0] ? $_[0] : 1;    # Default to hiding printed gridlines
#NYI 
#NYI     if ( $option == 0 ) {
#NYI         $self->{_print_gridlines}       = 1;    # 1 = display, 0 = hide
#NYI         $self->{_screen_gridlines}      = 1;
#NYI         $self->{_print_options_changed} = 1;
#NYI     }
#NYI     elsif ( $option == 1 ) {
#NYI         $self->{_print_gridlines}  = 0;
#NYI         $self->{_screen_gridlines} = 1;
#NYI     }
#NYI     else {
#NYI         $self->{_print_gridlines}  = 0;
#NYI         $self->{_screen_gridlines} = 0;
#NYI     }
#NYI }


###############################################################################
#
# print_row_col_headers()
#
# Set the option to print the row and column headers on the printed page.
# See also the _store_print_headers() method below.
#
method print_row_col_headers($headers = 1) {
    if $headers {
        $!print_headers         = 1;
        $!print_options_changed = 1;
    }
    else {
        $!print_headers = 0;
    }
}


###############################################################################
#
# fit_to_pages($width, $height)
#
# Store the vertical and horizontal number of pages that will define the
# maximum area printed.
#
method fit_to_pages($width = 1, $height = 1) {
    $!fit_page           = 1;
    $!fit_width          = $width;
    $!fit_height         = $height;
    $!page_setup_changed = 1;
}


###############################################################################
#
# set_h_pagebreaks(@breaks)
#
# Store the horizontal page breaks on a worksheet.
#
method set_h_pagebreaks(*@breaks) {
    @!hbreaks.append: @breaks;
}


###############################################################################
#
# set_v_pagebreaks(@breaks)
#
# Store the vertical page breaks on a worksheet.
#
method set_v_pagebreaks(*@breaks) {
    @!vbreaks.append: @breaks;
}


###############################################################################
#
# set_zoom( $scale )
#
# Set the worksheet zoom factor.
#
method set_zoom($scale = 100) {
    # Confine the scale to Excel's range
    if not 10 <= $scale <= 400 {
        warn "Zoom factor $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    $!zoom = $scale.int;
}


###############################################################################
#
# set_print_scale($scale)
#
# Set the scale factor for the printed page.
#
method set_print_scale($scale = 100) {
    # Confine the scale to Excel's range
    if not 10 <= $scale <= 400 {
        warn "Print scale $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    # Turn off "fit to page" option.
    $!fit_page = 0;

    $!print_scale        = $scale.int;
    $!page_setup_changed = 1;
}


###############################################################################
#
# print_black_and_white()
#
# Set the option to print the worksheet in black and white.
#
method print_black_and_white {
    $!black_white = 1;
}


###############################################################################
#
# keep_leading_zeros()
#
# Causes the write() method to treat integers with a leading zero as a string.
# This ensures that any leading zeros such, as in zip codes, are maintained.
#
method keep_leading_zeros($leading-zeros = 1) {
    $!leading_zeros = $leading-zeros;
}


###############################################################################
#
# show_comments()
#
# Make any comments in the worksheet visible.
#
method show_comments($visible = 1) {
    $!comments_visible = $visible;
}


###############################################################################
#
# set_comments_author()
#
# Set the default author of the cell comments.
#
method set_comments_author($author) {
    $!comments_author = $author if $author.defined;
}


#NYI ###############################################################################
#NYI #
#NYI # right_to_left()
#NYI #
#NYI # Display the worksheet right to left for some eastern versions of Excel.
#NYI #
#NYI sub right_to_left {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->{_right_to_left} = defined $_[0] ? $_[0] : 1;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # hide_zero()
#NYI #
#NYI # Hide cell zero values.
#NYI #
#NYI sub hide_zero {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->{_show_zeros} = defined $_[0] ? not $_[0] : 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # print_across()
#NYI #
#NYI # Set the order in which pages are printed.
#NYI #
#NYI sub print_across {
#NYI 
#NYI     my $self = shift;
#NYI     my $page_order = defined $_[0] ? $_[0] : 1;
#NYI 
#NYI     if ( $page_order ) {
#NYI         $self->{_page_order}         = 1;
#NYI         $self->{_page_setup_changed} = 1;
#NYI     }
#NYI     else {
#NYI         $self->{_page_order} = 0;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_start_page()
#NYI #
#NYI # Set the start page number.
#NYI #
#NYI sub set_start_page {
#NYI 
#NYI     my $self = shift;
#NYI     return unless defined $_[0];
#NYI 
#NYI     $self->{_page_start}   = $_[0];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_first_row_column()
#NYI #
#NYI # Set the topmost and leftmost visible row and column.
#NYI # TODO: Document this when tested fully for interaction with panes.
#NYI #
#NYI sub set_first_row_column {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $row = $_[0] || 0;
#NYI     my $col = $_[1] || 0;
#NYI 
#NYI     $row = $self->{_xls_rowmax} if $row > $self->{_xls_rowmax};
#NYI     $col = $self->{_xls_colmax} if $col > $self->{_xls_colmax};
#NYI 
#NYI     $self->{_first_row} = $row;
#NYI     $self->{_first_col} = $col;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_write_handler($re, $code_ref)
#NYI #
#NYI # Allow the user to add their own matches and handlers to the write() method.
#NYI #
#NYI sub add_write_handler {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return unless @_ == 2;
#NYI     return unless ref $_[1] eq 'CODE';
#NYI 
#NYI     push @{ $self->{_write_match} }, [@_];
#NYI }


###############################################################################
#
# write($row, $col, $token, $format)
#
# Parse $token and call appropriate write method. $row and $column are zero
# indexed. $format is optional.
#
# Returns: return value of called subroutine
#
# TODO:
method write(*@args) {
    # Check for a cell reference in A1 notation and substitute row and column
    if @args[0] ~~ /^\D/ {
        @args = self.substitute_cellref( @args );
    }

    my $token = @args[2];

    # Handle undefs as blanks
    $token = '' unless $token.defined;


    # First try user defined matches.
    for @!write_match -> @aref {
        my $re  = @aref[0];
        my $sub = @aref[1];

        if $token ~~ /<$re>/ {
            my $match = &$sub( self, @args );
            return $match if $match.defined;
        }
    }


#NYI     # Match an array ref.
#NYI     if ( ref $token eq "ARRAY" ) {
#NYI         return $self->write_row( @_ );
#NYI     }

#NYI     # Match integer with leading zero(s)
#NYI     elsif ( $self->{_leading_zeros} and $token =~ /^0\d+$/ ) {
#NYI         return $self->write_string( @_ );
#NYI     }

#NYI     # Match number
#NYI     elsif ( $token =~ /^([+-]?)(?=[0-9]|\.[0-9])[0-9]*(\.[0-9]*)?([Ee]([+-]?[0-9]+))?$/ ) {
#NYI         return $self->write_number( @_ );
#NYI     }

#NYI     # Match http, https or ftp URL
#NYI     elsif ( $token =~ m|^[fh]tt?ps?://| ) {
#NYI         return $self->write_url( @_ );
#NYI     }

#NYI     # Match mailto:
#NYI     elsif ( $token =~ m/^mailto:/ ) {
#NYI         return $self->write_url( @_ );
#NYI     }

#NYI     # Match internal or external sheet link
#NYI     elsif ( $token =~ m[^(?:in|ex)ternal:] ) {
#NYI         return $self->write_url( @_ );
#NYI     }

#NYI     # Match formula
#NYI     elsif ( $token =~ /^=/ ) {
#NYI         return $self->write_formula( @_ );
#NYI     }

#NYI     # Match array formula
#NYI     elsif ( $token =~ /^{=.*}$/ ) {
#NYI         return $self->write_formula( @_ );
#NYI     }

#NYI     # Match blank
#NYI     elsif ( $token eq '' ) {
#NYI         splice @_, 2, 1;    # remove the empty string from the parameter list
#NYI         return $self->write_blank( @_ );
#NYI     }

#NYI     # Default: match string
#NYI     else {
#NYI         return $self->write_string( @_ );
#NYI     }
}


#NYI ###############################################################################
#NYI #
#NYI # write_row($row, $col, $array_ref, $format)
#NYI #
#NYI # Write a row of data starting from ($row, $col). Call write_col() if any of
#NYI # the elements of the array ref are in turn array refs. This allows the writing
#NYI # of 1D or 2D arrays of data in one go.
#NYI #
#NYI # Returns: the first encountered error value or zero for no errors
#NYI #
#NYI sub write_row {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Catch non array refs passed by user.
#NYI     if ( ref $_[2] ne 'ARRAY' ) {
#NYI         fail "Not an array ref in call to write_row()$!";
#NYI     }
#NYI 
#NYI     my $row     = shift;
#NYI     my $col     = shift;
#NYI     my $tokens  = shift;
#NYI     my @options = @_;
#NYI     my $error   = 0;
#NYI     my $ret;
#NYI 
#NYI     for my $token ( @$tokens ) {
#NYI 
#NYI         # Check for nested arrays
#NYI         if ( ref $token eq "ARRAY" ) {
#NYI             $ret = $self->write_col( $row, $col, $token, @options );
#NYI         }
#NYI         else {
#NYI             $ret = $self->write( $row, $col, $token, @options );
#NYI         }
#NYI 
#NYI         # Return only the first error encountered, if any.
#NYI         $error ||= $ret;
#NYI         $col++;
#NYI     }
#NYI 
#NYI     return $error;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_col($row, $col, $array_ref, $format)
#NYI #
#NYI # Write a column of data starting from ($row, $col). Call write_row() if any of
#NYI # the elements of the array ref are in turn array refs. This allows the writing
#NYI # of 1D or 2D arrays of data in one go.
#NYI #
#NYI # Returns: the first encountered error value or zero for no errors
#NYI #
#NYI sub write_col {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Catch non array refs passed by user.
#NYI     if ( ref $_[2] ne 'ARRAY' ) {
#NYI         fail "Not an array ref in call to write_col()$!";
#NYI     }
#NYI 
#NYI     my $row     = shift;
#NYI     my $col     = shift;
#NYI     my $tokens  = shift;
#NYI     my @options = @_;
#NYI     my $error   = 0;
#NYI     my $ret;
#NYI 
#NYI     for my $token ( @$tokens ) {
#NYI 
#NYI         # write() will deal with any nested arrays
#NYI         $ret = $self->write( $row, $col, $token, @options );
#NYI 
#NYI         # Return only the first error encountered, if any.
#NYI         $error ||= $ret;
#NYI         $row++;
#NYI     }
#NYI 
#NYI     return $error;
#NYI }


###############################################################################
#
# write_comment($row, $col, $comment)
#
# Write a comment to the specified row and column (zero indexed).
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
method write_comment(*@options) {
    # Check for a cell reference in A1 notation and substitute row and column
    if ( @options[0] ~~ /^\D/ ) {
        (@options) = self.substitute_cellref( @options );
    }

    if @options.elems < 3 { return -1 }    # Check the number of args

    # Check for pairs of optional arguments, i.e. an odd number of args.
    fail "Uneven number of additional arguments" unless @options.elems %% 2;

    my $row = @options[0];
    my $col = @options[1];

    # Check that row and col are valid and store max and min values
    return -2 if self.check_dimensions( $row, $col );

    $!has_vml      = 1;
    $!has_comments = 1;

    # Process the properties of the cell comment.
    %!comments{$row}{$col} = [ self.comment_params( @options ) ];
}


###############################################################################
#
# write_number($row, $col, $num, $format)
#
# Write a double to the specified row and column (zero indexed).
# An integer can be written as a double. Excel will display an
# integer. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
method write_number(*@args) {

    # Check for a cell reference in A1 notation and substitute row and column
    if ( @args[0] ~~ /^\D/ ) {
        @args = self.substitute_cellref( @args );
    }

    if ( @args.elems < 3 ) { return -1 }    # Check the number of args


    my $row  =  @args[0];              # Zero indexed row
    my $col  =  @args[1];              # Zero indexed column
    my $num  = +@args[2];
    my $xf   =  @args[3];              # The cell format
    my $type =  'n';                   # The data type

    # Check that row and col are valid and store max and min values
    return -2 if self.check_dimensions( $row, $col );

    # Write previous row if in in-line string optimization mode.
    if $!optimization == 1 && $row > $!previous_row {
        self.write_single_row( $row );
    }

    %!table{$row}{$col} = [ $type, $num, $xf ];

    return 0;
}


#NYI ###############################################################################
#NYI #
#NYI # write_string ($row, $col, $string, $format)
#NYI #
#NYI # Write a string to the specified row and column (zero indexed).
#NYI # $format is optional.
#NYI # Returns  0 : normal termination
#NYI #         -1 : insufficient number of arguments
#NYI #         -2 : row or column out of range
#NYI #         -3 : long string truncated to 32767 chars
#NYI #
#NYI sub write_string {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     if ( @_ < 3 ) { return -1 }    # Check the number of args
#NYI 
#NYI     my $row  = $_[0];              # Zero indexed row
#NYI     my $col  = $_[1];              # Zero indexed column
#NYI     my $str  = $_[2];
#NYI     my $xf   = $_[3];              # The cell format
#NYI     my $type = 's';                # The data type
#NYI     my $index;
#NYI     my $str_error = 0;
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI     # Check that the string is < 32767 chars
#NYI     if ( length $str > $self->{_xls_strmax} ) {
#NYI         $str = substr( $str, 0, $self->{_xls_strmax} );
#NYI         $str_error = -3;
#NYI     }
#NYI 
#NYI     # Write a shared string or an in-line string based on optimisation level.
#NYI     if ( $self->{_optimization} == 0 ) {
#NYI         $index = $self->_get_shared_string_index( $str );
#NYI     }
#NYI     else {
#NYI         $index = $str;
#NYI     }
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row}->{$col} = [ $type, $index, $xf ];
#NYI 
#NYI     return $str_error;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_rich_string( $row, $column, $format, $string, ..., $cell_format )
#NYI #
#NYI # The write_rich_string() method is used to write strings with multiple formats.
#NYI # The method receives string fragments prefixed by format objects. The final
#NYI # format object is used as the cell format.
#NYI #
#NYI # Returns  0 : normal termination.
#NYI #         -1 : insufficient number of arguments.
#NYI #         -2 : row or column out of range.
#NYI #         -3 : long string truncated to 32767 chars.
#NYI #         -4 : 2 consecutive formats used.
#NYI #
#NYI sub write_rich_string {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     if ( @_ < 3 ) { return -1 }    # Check the number of args
#NYI 
#NYI     my $row    = shift;            # Zero indexed row.
#NYI     my $col    = shift;            # Zero indexed column.
#NYI     my $str    = '';
#NYI     my $xf     = undef;
#NYI     my $type   = 's';              # The data type.
#NYI     my $length = 0;                # String length.
#NYI     my $index;
#NYI     my $str_error = 0;
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI 
#NYI     # If the last arg is a format we use it as the cell format.
#NYI     if ( ref $_[-1] ) {
#NYI         $xf = pop @_;
#NYI     }
#NYI 
#NYI 
#NYI     # Create a temp XML::Writer object and use it to write the rich string
#NYI     # XML to a string.
#NYI     open my $str_fh, '>', \$str or die "Failed to open filehandle: $!";
#NYI     binmode $str_fh, ':utf8';
#NYI 
#NYI     my $writer = Excel::Writer::XLSX::Package::XMLwriter->new( $str_fh );
#NYI 
#NYI     $self->{_rstring} = $writer;
#NYI 
#NYI     # Create a temp format with the default font for unformatted fragments.
#NYI     my $default = Excel::Writer::XLSX::Format->new();
#NYI 
#NYI     # Convert the list of $format, $string tokens to pairs of ($format, $string)
#NYI     # except for the first $string fragment which doesn't require a default
#NYI     # formatting run. Use the default for strings without a leading format.
#NYI     my @fragments;
#NYI     my $last = 'format';
#NYI     my $pos  = 0;
#NYI 
#NYI     for my $token ( @_ ) {
#NYI         if ( !ref $token ) {
#NYI 
#NYI             # Token is a string.
#NYI             if ( $last ne 'format' ) {
#NYI 
#NYI                 # If previous token wasn't a format add one before the string.
#NYI                 push @fragments, ( $default, $token );
#NYI             }
#NYI             else {
#NYI 
#NYI                 # If previous token was a format just add the string.
#NYI                 push @fragments, $token;
#NYI             }
#NYI 
#NYI             $length += length $token;    # Keep track of actual string length.
#NYI             $last = 'string';
#NYI         }
#NYI         else {
#NYI 
#NYI             # Can't allow 2 formats in a row.
#NYI             if ( $last eq 'format' && $pos > 0 ) {
#NYI                 return -4;
#NYI             }
#NYI 
#NYI             # Token is a format object. Add it to the fragment list.
#NYI             push @fragments, $token;
#NYI             $last = 'format';
#NYI         }
#NYI 
#NYI         $pos++;
#NYI     }
#NYI 
#NYI 
#NYI     # If the first token is a string start the <r> element.
#NYI     if ( !ref $fragments[0] ) {
#NYI         $self->{_rstring}->xml_start_tag( 'r' );
#NYI     }
#NYI 
#NYI     # Write the XML elements for the $format $string fragments.
#NYI     for my $token ( @fragments ) {
#NYI         if ( ref $token ) {
#NYI 
#NYI             # Write the font run.
#NYI             $self->{_rstring}->xml_start_tag( 'r' );
#NYI             $self->_write_font( $token );
#NYI         }
#NYI         else {
#NYI 
#NYI             # Write the string fragment part, with whitespace handling.
#NYI             my @attributes = ();
#NYI 
#NYI             if ( $token =~ /^\s/ || $token =~ /\s$/ ) {
#NYI                 push @attributes, ( 'xml:space' => 'preserve' );
#NYI             }
#NYI 
#NYI             $self->{_rstring}->xml_data_element( 't', $token, @attributes );
#NYI             $self->{_rstring}->xml_end_tag( 'r' );
#NYI         }
#NYI     }
#NYI 
#NYI     # Check that the string is < 32767 chars.
#NYI     if ( $length > $self->{_xls_strmax} ) {
#NYI         return -3;
#NYI     }
#NYI 
#NYI 
#NYI     # Write a shared string or an in-line string based on optimisation level.
#NYI     if ( $self->{_optimization} == 0 ) {
#NYI         $index = $self->_get_shared_string_index( $str );
#NYI     }
#NYI     else {
#NYI         $index = $str;
#NYI     }
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row}->{$col} = [ $type, $index, $xf ];
#NYI 
#NYI     return 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_blank($row, $col, $format)
#NYI #
#NYI # Write a blank cell to the specified row and column (zero indexed).
#NYI # A blank cell is used to specify formatting without adding a string
#NYI # or a number.
#NYI #
#NYI # A blank cell without a format serves no purpose. Therefore, we don't write
#NYI # a BLANK record unless a format is specified. This is mainly an optimisation
#NYI # for the write_row() and write_col() methods.
#NYI #
#NYI # Returns  0 : normal termination (including no format)
#NYI #         -1 : insufficient number of arguments
#NYI #         -2 : row or column out of range
#NYI #
#NYI sub write_blank {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Check the number of args
#NYI     return -1 if @_ < 2;
#NYI 
#NYI     # Don't write a blank cell unless it has a format
#NYI     return 0 if not defined $_[2];
#NYI 
#NYI     my $row  = $_[0];    # Zero indexed row
#NYI     my $col  = $_[1];    # Zero indexed column
#NYI     my $xf   = $_[2];    # The cell format
#NYI     my $type = 'b';      # The data type
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row}->{$col} = [ $type, undef, $xf ];
#NYI 
#NYI     return 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_formula($row, $col, $formula, $format)
#NYI #
#NYI # Write a formula to the specified row and column (zero indexed).
#NYI #
#NYI # $format is optional.
#NYI #
#NYI # Returns  0 : normal termination
#NYI #         -1 : insufficient number of arguments
#NYI #         -2 : row or column out of range
#NYI #
#NYI sub write_formula {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     if ( @_ < 3 ) { return -1 }    # Check the number of args
#NYI 
#NYI     my $row     = $_[0];           # Zero indexed row
#NYI     my $col     = $_[1];           # Zero indexed column
#NYI     my $formula = $_[2];           # The formula text string
#NYI     my $xf      = $_[3];           # The format object.
#NYI     my $value   = $_[4];           # Optional formula value.
#NYI     my $type    = 'f';             # The data type
#NYI 
#NYI     # Hand off array formulas.
#NYI     if ( $formula =~ /^{=.*}$/ ) {
#NYI         return $self->write_array_formula( $row, $col, $row, $col, $formula,
#NYI             $xf, $value );
#NYI     }
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI     # Remove the = sign if it exists.
#NYI     $formula =~ s/^=//;
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row}->{$col} = [ $type, $formula, $xf, $value ];
#NYI 
#NYI     return 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_array_formula($row1, $col1, $row2, $col2, $formula, $format)
#NYI #
#NYI # Write an array formula to the specified row and column (zero indexed).
#NYI #
#NYI # $format is optional.
#NYI #
#NYI # Returns  0 : normal termination
#NYI #         -1 : insufficient number of arguments
#NYI #         -2 : row or column out of range
#NYI #
#NYI sub write_array_formula {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     if ( @_ < 5 ) { return -1 }    # Check the number of args
#NYI 
#NYI     my $row1    = $_[0];           # First row
#NYI     my $col1    = $_[1];           # First column
#NYI     my $row2    = $_[2];           # Last row
#NYI     my $col2    = $_[3];           # Last column
#NYI     my $formula = $_[4];           # The formula text string
#NYI     my $xf      = $_[5];           # The format object.
#NYI     my $value   = $_[6];           # Optional formula value.
#NYI     my $type    = 'a';             # The data type
#NYI 
#NYI     # Swap last row/col with first row/col as necessary
#NYI     ( $row1, $row2 ) = ( $row2, $row1 ) if $row1 > $row2;
#NYI     ( $col1, $col2 ) = ( $col1, $col2 ) if $col1 > $col2;
#NYI 
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row2, $col2 );
#NYI 
#NYI 
#NYI     # Define array range
#NYI     my $range;
#NYI 
#NYI     if ( $row1 == $row2 and $col1 == $col2 ) {
#NYI         $range = xl-rowcol-to-cell( $row1, $col1 );
#NYI 
#NYI     }
#NYI     else {
#NYI         $range =
#NYI             xl-rowcol-to-cell( $row1, $col1 ) . ':'
#NYI           . xl-rowcol-to-cell( $row2, $col2 );
#NYI     }
#NYI 
#NYI     # Remove array formula braces and the leading =.
#NYI     $formula =~ s/^{(.*)}$/$1/;
#NYI     $formula =~ s/^=//;
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     my $row = $row1;
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row1}->{$col1} =
#NYI       [ $type, $formula, $xf, $range, $value ];
#NYI 
#NYI 
#NYI     # Pad out the rest of the area with formatted zeroes.
#NYI     if ( !$self->{_optimization} ) {
#NYI         for my $row ( $row1 .. $row2 ) {
#NYI             for my $col ( $col1 .. $col2 ) {
#NYI                 next if $row == $row1 and $col == $col1;
#NYI                 $self->write_number( $row, $col, 0, $xf );
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     return 0;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # write_boolean($row, $col, $value, $format)
#NYI #
#NYI # Write a boolean value to the specified row and column (zero indexed).
#NYI #
#NYI # Returns  0 : normal termination (including no format)
#NYI #         -2 : row or column out of range
#NYI #
#NYI sub write_boolean {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     my $row  = $_[0];            # Zero indexed row
#NYI     my $col  = $_[1];            # Zero indexed column
#NYI     my $val  = $_[2] ? 1 : 0;    # Boolean value.
#NYI     my $xf   = $_[3];            # The cell format
#NYI     my $type = 'l';              # The data type
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row}->{$col} = [ $type, $val, $xf ];
#NYI 
#NYI     return 0;
#NYI }


###############################################################################
#
# outline_settings($visible, $symbols_below, $symbols_right, $auto_style)
#
# This method sets the properties for outlining and grouping. The defaults
# correspond to Excel's defaults.
#
method outline_settings($visible = 1, $symbols-below = 1, $symbols-right = 1, $auto-style = 0) {
    $!outline_on    = $visible;
    $!outline_below = $symbols-below;
    $!outline_right = $symbols-right;
    $!outline_style = $auto-style;

    $!outline_changed = 1;
}


###############################################################################
#
# Escape urls like Excel.
#
method escape_url($url) {

    # Don't escape URL if it looks already escaped.
    return $url if $url ~~ / '%' <[0..9 a..f A..F]> ** 2/;

    # Escape the URL escape symbol.
    $url ~~ s:g/\%/%25/;

    # Escape whitespace in URL.
    $url ~~ s:g/<[\s \x00]>/%20/;

    # Escape other special characters in URL.
    $url ~~ s:g/(<["<>[\]`^{}]>)/{sprintf '%%%x', $0.ord}/;

    return $url;
}


###############################################################################
#
# write_url($row, $col, $url, $string, $format)
#
# Write a hyperlink. This is comprised of two elements: the visible label and
# the invisible link. The visible label is the same as the link unless an
# alternative string is specified. The label is written using the
# write_string() method. Therefore the max characters string limit applies.
# $string and $format are optional and their order is interchangeable.
#
# The hyperlink can be to a http, ftp, mail, internal sheet, or external
# directory url.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 32767 chars
#         -4 : URL longer than 255 characters
#         -5 : Exceeds limit of 65_530 urls per worksheet
#
#NYI sub write_url {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     if ( @_ < 3 ) { return -1 }    # Check the number of args
#NYI 
#NYI 
#NYI     # Reverse the order of $string and $format if necessary. We work on a copy
#NYI     # in order to protect the callers args. We don't use "local @_" in case of
#NYI     # perl50005 threads.
#NYI     my @args = @_;
#NYI     ( $args[3], $args[4] ) = ( $args[4], $args[3] ) if ref $args[3];
#NYI 
#NYI 
#NYI     my $row       = $args[0];    # Zero indexed row
#NYI     my $col       = $args[1];    # Zero indexed column
#NYI     my $url       = $args[2];    # URL string
#NYI     my $str       = $args[3];    # Alternative label
#NYI     my $xf        = $args[4];    # Cell format
#NYI     my $tip       = $args[5];    # Tool tip
#NYI     my $type      = 'l';         # XML data type
#NYI     my $link_type = 1;
#NYI     my $external  = 0;
#NYI 
#NYI     # The displayed string defaults to the url string.
#NYI     $str = $url unless defined $str;
#NYI 
#NYI     # Remove the URI scheme from internal links.
#NYI     if ( $url =~ s/^internal:// ) {
#NYI         $str =~ s/^internal://;
#NYI         $link_type = 2;
#NYI     }
#NYI 
#NYI     # Remove the URI scheme from external links and change the directory
#NYI     # separator from Unix to Dos.
#NYI     if ( $url =~ s/^external:// ) {
#NYI         $str =~ s/^external://;
#NYI         $url =~ s[/][\\]g;
#NYI         $str =~ s[/][\\]g;
#NYI         $external = 1;
#NYI     }
#NYI 
#NYI     # Strip the mailto header.
#NYI     $str =~ s/^mailto://;
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI     # Check that the string is < 32767 chars
#NYI     my $str_error = 0;
#NYI     if ( length $str > $self->{_xls_strmax} ) {
#NYI         $str = substr( $str, 0, $self->{_xls_strmax} );
#NYI         $str_error = -3;
#NYI     }
#NYI 
#NYI     # Copy string for use in hyperlink elements.
#NYI     my $url_str = $str;
#NYI 
#NYI     # External links to URLs and to other Excel workbooks have slightly
#NYI     # different characteristics that we have to account for.
#NYI     if ( $link_type == 1 ) {
#NYI 
#NYI         # Split url into the link and optional anchor/location.
#NYI         ( $url, $url_str ) = split /#/, $url, 2;
#NYI 
#NYI         $url = _escape_url( $url );
#NYI 
#NYI         # Escape the anchor for hyperlink style urls only.
#NYI         if ( $url_str && !$external ) {
#NYI             $url_str = _escape_url( $url_str );
#NYI         }
#NYI 
#NYI         # Add the file:/// URI to the url for Windows style "C:/" link and
#NYI         # Network shares.
#NYI         if ( $url =~ m{^\w:} || $url =~ m{^\\\\} ) {
#NYI             $url = 'file:///' . $url;
#NYI         }
#NYI 
#NYI         # Convert a ./dir/file.xlsx link to dir/file.xlsx.
#NYI         $url =~ s{^.\\}{};
#NYI     }
#NYI 
#NYI     # Excel limits the escaped URL and location/anchor to 255 characters.
#NYI     my $tmp_url_str = $url_str || '';
#NYI 
#NYI     if ( length $url > 255 || length $tmp_url_str > 255 ) {
#NYI         warn "Ignoring URL '$url' where link or anchor > 255 characters "
#NYI           . "since it exceeds Excel's limit for URLS. See LIMITATIONS "
#NYI           . "section of the Excel::Writer::XLSX documentation.";
#NYI         return -4;
#NYI     }
#NYI 
#NYI     # Check the limit of URLS per worksheet.
#NYI     $self->{_hlink_count}++;
#NYI 
#NYI     if ( $self->{_hlink_count} > 65_530 ) {
#NYI         warn "Ignoring URL '$url' since it exceeds Excel's limit of 65,530 "
#NYI           . "URLS per worksheet. See LIMITATIONS section of the "
#NYI           . "Excel::Writer::XLSX documentation.";
#NYI         return -5;
#NYI     }
#NYI 
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     # Write the hyperlink string.
#NYI     $self->write_string( $row, $col, $str, $xf );
#NYI 
#NYI     # Store the hyperlink data in a separate structure.
#NYI     $self->{_hyperlinks}->{$row}->{$col} = {
#NYI         _link_type => $link_type,
#NYI         _url       => $url,
#NYI         _str       => $url_str,
#NYI         _tip       => $tip
#NYI     };
#NYI 
#NYI     return $str_error;
#NYI }


###############################################################################
#
# write_date_time ($row, $col, $string, $format)
#
# Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
# number representing an Excel date. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : Invalid date_time, written as string
#
#NYI sub write_date_time {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     if ( @_ < 3 ) { return -1 }    # Check the number of args
#NYI 
#NYI     my $row  = $_[0];              # Zero indexed row
#NYI     my $col  = $_[1];              # Zero indexed column
#NYI     my $str  = $_[2];
#NYI     my $xf   = $_[3];              # The cell format
#NYI     my $type = 'n';                # The data type
#NYI 
#NYI 
#NYI     # Check that row and col are valid and store max and min values
#NYI     return -2 if $self->_check_dimensions( $row, $col );
#NYI 
#NYI     my $str_error = 0;
#NYI     my $date_time = $self->convert_date_time( $str );
#NYI 
#NYI     # If the date isn't valid then write it as a string.
#NYI     if ( !defined $date_time ) {
#NYI         return $self->write_string( @_ );
#NYI     }
#NYI 
#NYI     # Write previous row if in in-line string optimization mode.
#NYI     if ( $self->{_optimization} == 1 && $row > $self->{_previous_row} ) {
#NYI         $self->_write_single_row( $row );
#NYI     }
#NYI 
#NYI     $self->{_table}->{$row}->{$col} = [ $type, $date_time, $xf ];
#NYI 
#NYI     return $str_error;
#NYI }


###############################################################################
#
# convert_date_time($date_time_string)
#
# The function takes a date and time in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format
# and converts it to a decimal number representing a valid Excel date.
#
# Dates and times in Excel are represented by real numbers. The integer part of
# the number stores the number of days since the epoch and the fractional part
# stores the percentage of the day in seconds. The epoch can be either 1900 or
# 1904.
#
# Parameter: Date and time string in one of the following formats:
#               yyyy-mm-ddThh:mm:ss.ss  # Standard
#               yyyy-mm-ddT             # Date only
#                         Thh:mm:ss.ss  # Time only
#
# Returns:
#            A decimal number representing a valid Excel date, or
#            undef if the date is invalid.
#
#NYI sub convert_date_time {
#NYI 
#NYI     my $self      = shift;
#NYI     my $date_time = $_[0];
#NYI 
#NYI     my $days    = 0;    # Number of days since epoch
#NYI     my $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
#NYI 
#NYI     my ( $year, $month, $day );
#NYI     my ( $hour, $min,   $sec );
#NYI 
#NYI 
#NYI     # Strip leading and trailing whitespace.
#NYI     $date_time =~ s/^\s+//;
#NYI     $date_time =~ s/\s+$//;
#NYI 
#NYI     # Check for invalid date char.
#NYI     return if $date_time =~ /[^0-9T:\-\.Z]/;
#NYI 
#NYI     # Check for "T" after date or before time.
#NYI     return unless $date_time =~ /\dT|T\d/;
#NYI 
#NYI     # Strip trailing Z in ISO8601 date.
#NYI     $date_time =~ s/Z$//;
#NYI 
#NYI 
#NYI     # Split into date and time.
#NYI     my ( $date, $time ) = split /T/, $date_time;
#NYI 
#NYI 
#NYI     # We allow the time portion of the input DateTime to be optional.
#NYI     if ( $time ne '' ) {
#NYI 
#NYI         # Match hh:mm:ss.sss+ where the seconds are optional
#NYI         if ( $time =~ /^(\d\d):(\d\d)(:(\d\d(\.\d+)?))?/ ) {
#NYI             $hour = $1;
#NYI             $min  = $2;
#NYI             $sec  = $4 || 0;
#NYI         }
#NYI         else {
#NYI             return undef;    # Not a valid time format.
#NYI         }
#NYI 
#NYI         # Some boundary checks
#NYI         return if $hour >= 24;
#NYI         return if $min >= 60;
#NYI         return if $sec >= 60;
#NYI 
#NYI         # Excel expresses seconds as a fraction of the number in 24 hours.
#NYI         $seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
#NYI     }
#NYI 
#NYI 
#NYI     # We allow the date portion of the input DateTime to be optional.
#NYI     return $seconds if $date eq '';
#NYI 
#NYI 
#NYI     # Match date as yyyy-mm-dd.
#NYI     if ( $date =~ /^(\d\d\d\d)-(\d\d)-(\d\d)$/ ) {
#NYI         $year  = $1;
#NYI         $month = $2;
#NYI         $day   = $3;
#NYI     }
#NYI     else {
#NYI         return undef;    # Not a valid date format.
#NYI     }
#NYI 
#NYI     # Set the epoch as 1900 or 1904. Defaults to 1900.
#NYI     my $date_1904 = $self->{_date_1904};
#NYI 
#NYI 
#NYI     # Special cases for Excel.
#NYI     if ( not $date_1904 ) {
#NYI         return $seconds      if $date eq '1899-12-31';    # Excel 1900 epoch
#NYI         return $seconds      if $date eq '1900-01-00';    # Excel 1900 epoch
#NYI         return 60 + $seconds if $date eq '1900-02-29';    # Excel false leapday
#NYI     }
#NYI 
#NYI 
#NYI     # We calculate the date by calculating the number of days since the epoch
#NYI     # and adjust for the number of leap days. We calculate the number of leap
#NYI     # days by normalising the year in relation to the epoch. Thus the year 2000
#NYI     # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
#NYI     #
#NYI     my $epoch  = $date_1904 ? 1904 : 1900;
#NYI     my $offset = $date_1904 ? 4    : 0;
#NYI     my $norm   = 300;
#NYI     my $range  = $year - $epoch;
#NYI 
#NYI 
#NYI     # Set month days and check for leap year.
#NYI     my @mdays = ( 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );
#NYI     my $leap = 0;
#NYI     $leap = 1 if $year % 4 == 0 and $year % 100 or $year % 400 == 0;
#NYI     $mdays[1] = 29 if $leap;
#NYI 
#NYI 
#NYI     # Some boundary checks
#NYI     return if $year < $epoch or $year > 9999;
#NYI     return if $month < 1     or $month > 12;
#NYI     return if $day < 1       or $day > $mdays[ $month - 1 ];
#NYI 
#NYI     # Accumulate the number of days since the epoch.
#NYI     $days = $day;    # Add days for current month
#NYI     $days += $mdays[$_] for 0 .. $month - 2;    # Add days for past months
#NYI     $days += $range * 365;                      # Add days for past years
#NYI     $days += int( ( $range ) / 4 );             # Add leapdays
#NYI     $days -= int( ( $range + $offset ) / 100 ); # Subtract 100 year leapdays
#NYI     $days += int( ( $range + $offset + $norm ) / 400 );  # Add 400 year leapdays
#NYI     $days -= $leap;                                      # Already counted above
#NYI 
#NYI 
#NYI     # Adjust for Excel erroneously treating 1900 as a leap year.
#NYI     $days++ if $date_1904 == 0 and $days > 59;
#NYI 
#NYI     return $days + $seconds;
#NYI }


###############################################################################
#
# set_row($row, $height, $XF, $hidden, $level, $collapsed)
#
# This method is used to set the height and XF format for a row.
#
#NYI sub set_row {
#NYI 
#NYI     my $self      = shift;
#NYI     my $row       = shift;         # Row Number.
#NYI     my $height    = shift;         # Row height.
#NYI     my $xf        = shift;         # Format object.
#NYI     my $hidden    = shift || 0;    # Hidden flag.
#NYI     my $level     = shift || 0;    # Outline level.
#NYI     my $collapsed = shift || 0;    # Collapsed row.
#NYI     my $min_col   = 0;
#NYI 
#NYI     return unless defined $row;    # Ensure at least $row is specified.
#NYI 
#NYI     # Get the default row height.
#NYI     my $default_height = $self->{_default_row_height};
#NYI 
#NYI     # Use min col in _check_dimensions(). Default to 0 if undefined.
#NYI     if ( defined $self->{_dim_colmin} ) {
#NYI         $min_col = $self->{_dim_colmin};
#NYI     }
#NYI 
#NYI     # Check that row is valid.
#NYI     return -2 if $self->_check_dimensions( $row, $min_col );
#NYI 
#NYI     $height = $default_height if !defined $height;
#NYI 
#NYI     # If the height is 0 the row is hidden and the height is the default.
#NYI     if ( $height == 0 ) {
#NYI         $hidden = 1;
#NYI         $height = $default_height;
#NYI     }
#NYI 
#NYI     # Set the limits for the outline levels (0 <= x <= 7).
#NYI     $level = 0 if $level < 0;
#NYI     $level = 7 if $level > 7;
#NYI 
#NYI     if ( $level > $self->{_outline_row_level} ) {
#NYI         $self->{_outline_row_level} = $level;
#NYI     }
#NYI 
#NYI     # Store the row properties.
#NYI     $self->{_set_rows}->{$row} = [ $height, $xf, $hidden, $level, $collapsed ];
#NYI 
#NYI     # Store the row change to allow optimisations.
#NYI     $self->{_row_size_changed} = 1;
#NYI 
#NYI     if ($hidden) {
#NYI         $height = 0;
#NYI     }
#NYI 
#NYI     # Store the row sizes for use when calculating image vertices.
#NYI     $self->{_row_sizes}->{$row} = $height;
#NYI }


###############################################################################
#
# set_default_row()
#
# Set the default row properties
#
#NYI sub set_default_row {
#NYI 
#NYI     my $self        = shift;
#NYI     my $height      = shift || $self->{_original_row_height};
#NYI     my $zero_height = shift || 0;
#NYI 
#NYI     if ( $height != $self->{_original_row_height} ) {
#NYI         $self->{_default_row_height} = $height;
#NYI 
#NYI         # Store the row change to allow optimisations.
#NYI         $self->{_row_size_changed} = 1;
#NYI     }
#NYI 
#NYI     if ( $zero_height ) {
#NYI         $self->{_default_row_zeroed} = 1;
#NYI     }
#NYI }


###############################################################################
#
# merge_range($first_row, $first_col, $last_row, $last_col, $string, $format)
#
# Merge a range of cells. The first cell should contain the data and the others
# should be blank. All cells should contain the same format.
#
#NYI sub merge_range {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI     fail "Incorrect number of arguments" if @_ < 6;
#NYI     fail "Fifth parameter must be a format object" unless ref $_[5];
#NYI 
#NYI     my $row_first  = shift;
#NYI     my $col_first  = shift;
#NYI     my $row_last   = shift;
#NYI     my $col_last   = shift;
#NYI     my $string     = shift;
#NYI     my $format     = shift;
#NYI     my @extra_args = @_;      # For write_url().
#NYI 
#NYI     # Excel doesn't allow a single cell to be merged
#NYI     if ( $row_first == $row_last and $col_first == $col_last ) {
#NYI         fail "Can't merge single cell";
#NYI     }
#NYI 
#NYI     # Swap last row/col with first row/col as necessary
#NYI     ( $row_first, $row_last ) = ( $row_last, $row_first )
#NYI       if $row_first > $row_last;
#NYI     ( $col_first, $col_last ) = ( $col_last, $col_first )
#NYI       if $col_first > $col_last;
#NYI 
#NYI     # Check that column number is valid and store the max value
#NYI     return if $self->_check_dimensions( $row_last, $col_last );
#NYI 
#NYI     # Store the merge range.
#NYI     push @{ $self->{_merge} }, [ $row_first, $col_first, $row_last, $col_last ];
#NYI 
#NYI     # Write the first cell
#NYI     $self->write( $row_first, $col_first, $string, $format, @extra_args );
#NYI 
#NYI     # Pad out the rest of the area with formatted blank cells.
#NYI     for my $row ( $row_first .. $row_last ) {
#NYI         for my $col ( $col_first .. $col_last ) {
#NYI             next if $row == $row_first and $col == $col_first;
#NYI             $self->write_blank( $row, $col, $format );
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# merge_range_type()
#
# Same as merge_range() above except the type of write() is specified.
#
#NYI sub merge_range_type {
#NYI 
#NYI     my $self = shift;
#NYI     my $type = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     my $row_first = shift;
#NYI     my $col_first = shift;
#NYI     my $row_last  = shift;
#NYI     my $col_last  = shift;
#NYI     my $format;
#NYI 
#NYI     # Get the format. It can be in different positions for the different types.
#NYI     if (   $type eq 'array_formula'
#NYI         || $type eq 'blank'
#NYI         || $type eq 'rich_string' )
#NYI     {
#NYI 
#NYI         # The format is the last element.
#NYI         $format = $_[-1];
#NYI     }
#NYI     else {
#NYI 
#NYI         # Or else it is after the token.
#NYI         $format = $_[1];
#NYI     }
#NYI 
#NYI     # Check that there is a format object.
#NYI     fail "Format object missing or in an incorrect position"
#NYI       unless ref $format;
#NYI 
#NYI     # Excel doesn't allow a single cell to be merged
#NYI     if ( $row_first == $row_last and $col_first == $col_last ) {
#NYI         fail "Can't merge single cell";
#NYI     }
#NYI 
#NYI     # Swap last row/col with first row/col as necessary
#NYI     ( $row_first, $row_last ) = ( $row_last, $row_first )
#NYI       if $row_first > $row_last;
#NYI     ( $col_first, $col_last ) = ( $col_last, $col_first )
#NYI       if $col_first > $col_last;
#NYI 
#NYI     # Check that column number is valid and store the max value
#NYI     return if $self->_check_dimensions( $row_last, $col_last );
#NYI 
#NYI     # Store the merge range.
#NYI     push @{ $self->{_merge} }, [ $row_first, $col_first, $row_last, $col_last ];
#NYI 
#NYI     # Write the first cell
#NYI     if ( $type eq 'string' ) {
#NYI         $self->write_string( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'number' ) {
#NYI         $self->write_number( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'blank' ) {
#NYI         $self->write_blank( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'date_time' ) {
#NYI         $self->write_date_time( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'rich_string' ) {
#NYI         $self->write_rich_string( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'url' ) {
#NYI         $self->write_url( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'formula' ) {
#NYI         $self->write_formula( $row_first, $col_first, @_ );
#NYI     }
#NYI     elsif ( $type eq 'array_formula' ) {
#NYI         $self->write_formula_array( $row_first, $col_first, @_ );
#NYI     }
#NYI     else {
#NYI         fail "Unknown type '$type'";
#NYI     }
#NYI 
#NYI     # Pad out the rest of the area with formatted blank cells.
#NYI     for my $row ( $row_first .. $row_last ) {
#NYI         for my $col ( $col_first .. $col_last ) {
#NYI             next if $row == $row_first and $col == $col_first;
#NYI             $self->write_blank( $row, $col, $format );
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# data_validation($row, $col, {...})
#
# This method handles the interface to Excel data validation.
# Somewhat ironically this requires a lot of validation code since the
# interface is flexible and covers a several types of data validation.
#
# We allow data validation to be called on one cell or a range of cells. The
# hashref contains the validation parameters and must be the last param:
#    data_validation($row, $col, {...})
#    data_validation($first_row, $first_col, $last_row, $last_col, {...})
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : incorrect parameter.
#
#NYI sub data_validation {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Check for a valid number of args.
#NYI     if ( @_ != 5 && @_ != 3 ) { return -1 }
#NYI 
#NYI     # The final hashref contains the validation parameters.
#NYI     my $param = pop;
#NYI 
#NYI     # Make the last row/col the same as the first if not defined.
#NYI     my ( $row1, $col1, $row2, $col2 ) = @_;
#NYI     if ( !defined $row2 ) {
#NYI         $row2 = $row1;
#NYI         $col2 = $col1;
#NYI     }
#NYI 
#NYI     # Check that row and col are valid without storing the values.
#NYI     return -2 if $self->_check_dimensions( $row1, $col1, 1, 1 );
#NYI     return -2 if $self->_check_dimensions( $row2, $col2, 1, 1 );
#NYI 
#NYI 
#NYI     # Check that the last parameter is a hash list.
#NYI     if ( ref $param ne 'HASH' ) {
#NYI         warn "Last parameter '$param' in data_validation() must be a hash ref";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # List of valid input parameters.
#NYI     my %valid_parameter = (
#NYI         validate      => 1,
#NYI         criteria      => 1,
#NYI         value         => 1,
#NYI         source        => 1,
#NYI         minimum       => 1,
#NYI         maximum       => 1,
#NYI         ignore_blank  => 1,
#NYI         dropdown      => 1,
#NYI         show_input    => 1,
#NYI         input_title   => 1,
#NYI         input_message => 1,
#NYI         show_error    => 1,
#NYI         error_title   => 1,
#NYI         error_message => 1,
#NYI         error_type    => 1,
#NYI         other_cells   => 1,
#NYI     );
#NYI 
#NYI     # Check for valid input parameters.
#NYI     for my $param_key ( keys %$param ) {
#NYI         if ( not exists $valid_parameter{$param_key} ) {
#NYI             warn "Unknown parameter '$param_key' in data_validation()";
#NYI             return -3;
#NYI         }
#NYI     }
#NYI 
#NYI     # Map alternative parameter names 'source' or 'minimum' to 'value'.
#NYI     $param->{value} = $param->{source}  if defined $param->{source};
#NYI     $param->{value} = $param->{minimum} if defined $param->{minimum};
#NYI 
#NYI     # 'validate' is a required parameter.
#NYI     if ( not exists $param->{validate} ) {
#NYI         warn "Parameter 'validate' is required in data_validation()";
#NYI         return -3;
#NYI     }
#NYI 
#NYI 
#NYI     # List of  valid validation types.
#NYI     my %valid_type = (
#NYI         'any'          => 'none',
#NYI         'any value'    => 'none',
#NYI         'whole number' => 'whole',
#NYI         'whole'        => 'whole',
#NYI         'integer'      => 'whole',
#NYI         'decimal'      => 'decimal',
#NYI         'list'         => 'list',
#NYI         'date'         => 'date',
#NYI         'time'         => 'time',
#NYI         'text length'  => 'textLength',
#NYI         'length'       => 'textLength',
#NYI         'custom'       => 'custom',
#NYI     );
#NYI 
#NYI 
#NYI     # Check for valid validation types.
#NYI     if ( not exists $valid_type{ lc( $param->{validate} ) } ) {
#NYI         warn "Unknown validation type '$param->{validate}' for parameter "
#NYI           . "'validate' in data_validation()";
#NYI         return -3;
#NYI     }
#NYI     else {
#NYI         $param->{validate} = $valid_type{ lc( $param->{validate} ) };
#NYI     }
#NYI 
#NYI     # No action is required for validation type 'any'
#NYI     # unless there are input messages.
#NYI     if (   $param->{validate} eq 'none'
#NYI         && !defined $param->{input_message}
#NYI         && !defined $param->{input_title} )
#NYI     {
#NYI         return 0;
#NYI     }
#NYI 
#NYI     # The any, list and custom validations don't have a criteria
#NYI     # so we use a default of 'between'.
#NYI     if (   $param->{validate} eq 'none'
#NYI         || $param->{validate} eq 'list'
#NYI         || $param->{validate} eq 'custom' )
#NYI     {
#NYI         $param->{criteria} = 'between';
#NYI         $param->{maximum}  = undef;
#NYI     }
#NYI 
#NYI     # 'criteria' is a required parameter.
#NYI     if ( not exists $param->{criteria} ) {
#NYI         warn "Parameter 'criteria' is required in data_validation()";
#NYI         return -3;
#NYI     }
#NYI 
#NYI 
#NYI     # List of valid criteria types.
#NYI     my %criteria_type = (
#NYI         'between'                  => 'between',
#NYI         'not between'              => 'notBetween',
#NYI         'equal to'                 => 'equal',
#NYI         '='                        => 'equal',
#NYI         '=='                       => 'equal',
#NYI         'not equal to'             => 'notEqual',
#NYI         '!='                       => 'notEqual',
#NYI         '<>'                       => 'notEqual',
#NYI         'greater than'             => 'greaterThan',
#NYI         '>'                        => 'greaterThan',
#NYI         'less than'                => 'lessThan',
#NYI         '<'                        => 'lessThan',
#NYI         'greater than or equal to' => 'greaterThanOrEqual',
#NYI         '>='                       => 'greaterThanOrEqual',
#NYI         'less than or equal to'    => 'lessThanOrEqual',
#NYI         '<='                       => 'lessThanOrEqual',
#NYI     );
#NYI 
#NYI     # Check for valid criteria types.
#NYI     if ( not exists $criteria_type{ lc( $param->{criteria} ) } ) {
#NYI         warn "Unknown criteria type '$param->{criteria}' for parameter "
#NYI           . "'criteria' in data_validation()";
#NYI         return -3;
#NYI     }
#NYI     else {
#NYI         $param->{criteria} = $criteria_type{ lc( $param->{criteria} ) };
#NYI     }
#NYI 
#NYI 
#NYI     # 'Between' and 'Not between' criteria require 2 values.
#NYI     if ( $param->{criteria} eq 'between' || $param->{criteria} eq 'notBetween' )
#NYI     {
#NYI         if ( not exists $param->{maximum} ) {
#NYI             warn "Parameter 'maximum' is required in data_validation() "
#NYI               . "when using 'between' or 'not between' criteria";
#NYI             return -3;
#NYI         }
#NYI     }
#NYI     else {
#NYI         $param->{maximum} = undef;
#NYI     }
#NYI 
#NYI 
#NYI     # List of valid error dialog types.
#NYI     my %error_type = (
#NYI         'stop'        => 0,
#NYI         'warning'     => 1,
#NYI         'information' => 2,
#NYI     );
#NYI 
#NYI     # Check for valid error dialog types.
#NYI     if ( not exists $param->{error_type} ) {
#NYI         $param->{error_type} = 0;
#NYI     }
#NYI     elsif ( not exists $error_type{ lc( $param->{error_type} ) } ) {
#NYI         warn "Unknown criteria type '$param->{error_type}' for parameter "
#NYI           . "'error_type' in data_validation()";
#NYI         return -3;
#NYI     }
#NYI     else {
#NYI         $param->{error_type} = $error_type{ lc( $param->{error_type} ) };
#NYI     }
#NYI 
#NYI 
#NYI     # Convert date/times value if required.
#NYI     if ( $param->{validate} eq 'date' || $param->{validate} eq 'time' ) {
#NYI         if ( $param->{value} =~ /T/ ) {
#NYI             my $date_time = $self->convert_date_time( $param->{value} );
#NYI 
#NYI             if ( !defined $date_time ) {
#NYI                 warn "Invalid date/time value '$param->{value}' "
#NYI                   . "in data_validation()";
#NYI                 return -3;
#NYI             }
#NYI             else {
#NYI                 $param->{value} = $date_time;
#NYI             }
#NYI         }
#NYI         if ( defined $param->{maximum} && $param->{maximum} =~ /T/ ) {
#NYI             my $date_time = $self->convert_date_time( $param->{maximum} );
#NYI 
#NYI             if ( !defined $date_time ) {
#NYI                 warn "Invalid date/time value '$param->{maximum}' "
#NYI                   . "in data_validation()";
#NYI                 return -3;
#NYI             }
#NYI             else {
#NYI                 $param->{maximum} = $date_time;
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     # Check that the input title doesn't exceed the maximum length.
#NYI     if ( $param->{input_title} and length $param->{input_title} > 32 ) {
#NYI         warn "Length of input title '$param->{input_title}'"
#NYI           . " exceeds Excel's limit of 32";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # Check that the error title don't exceed the maximum length.
#NYI     if ( $param->{error_title} and length $param->{error_title} > 32 ) {
#NYI         warn "Length of error title '$param->{error_title}'"
#NYI           . " exceeds Excel's limit of 32";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # Check that the input message don't exceed the maximum length.
#NYI     if ( $param->{input_message} and length $param->{input_message} > 255 ) {
#NYI         warn "Length of input message '$param->{input_message}'"
#NYI           . " exceeds Excel's limit of 255";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # Check that the error message don't exceed the maximum length.
#NYI     if ( $param->{error_message} and length $param->{error_message} > 255 ) {
#NYI         warn "Length of error message '$param->{error_message}'"
#NYI           . " exceeds Excel's limit of 255";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # Check that the input list don't exceed the maximum length.
#NYI     if ( $param->{validate} eq 'list' ) {
#NYI 
#NYI         if ( ref $param->{value} eq 'ARRAY' ) {
#NYI 
#NYI             my $formula = join ',', @{ $param->{value} };
#NYI             if ( length $formula > 255 ) {
#NYI                 warn "Length of list items '$formula' exceeds Excel's "
#NYI                   . "limit of 255, use a formula range instead";
#NYI                 return -3;
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     # Set some defaults if they haven't been defined by the user.
#NYI     $param->{ignore_blank} = 1 if !defined $param->{ignore_blank};
#NYI     $param->{dropdown}     = 1 if !defined $param->{dropdown};
#NYI     $param->{show_input}   = 1 if !defined $param->{show_input};
#NYI     $param->{show_error}   = 1 if !defined $param->{show_error};
#NYI 
#NYI 
#NYI     # These are the cells to which the validation is applied.
#NYI     $param->{cells} = [ [ $row1, $col1, $row2, $col2 ] ];
#NYI 
#NYI     # A (for now) undocumented parameter to pass additional cell ranges.
#NYI     if ( exists $param->{other_cells} ) {
#NYI 
#NYI         push @{ $param->{cells} }, @{ $param->{other_cells} };
#NYI     }
#NYI 
#NYI     # Store the validation information until we close the worksheet.
#NYI     push @{ $self->{_validations} }, $param;
#NYI }


###############################################################################
#
# conditional_formatting($row, $col, {...})
#
# This method handles the interface to Excel conditional formatting.
#
# We allow the format to be called on one cell or a range of cells. The
# hashref contains the formatting parameters and must be the last param:
#    conditional_formatting($row, $col, {...})
#    conditional_formatting($first_row, $first_col, $last_row, $last_col, {...})
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : incorrect parameter.
#
#NYI sub conditional_formatting {
#NYI 
#NYI     my $self       = shift;
#NYI     my $user_range = '';
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI 
#NYI         # Check for a user defined multiple range like B3:K6,B8:K11.
#NYI         if ( $_[0] =~ /,/ ) {
#NYI             $user_range = $_[0];
#NYI             $user_range =~ s/^=//;
#NYI             $user_range =~ s/\s*,\s*/ /g;
#NYI             $user_range =~ s/\$//g;
#NYI         }
#NYI 
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # The final hashref contains the validation parameters.
#NYI     my $options = pop;
#NYI 
#NYI     # Make the last row/col the same as the first if not defined.
#NYI     my ( $row1, $col1, $row2, $col2 ) = @_;
#NYI     if ( !defined $row2 ) {
#NYI         $row2 = $row1;
#NYI         $col2 = $col1;
#NYI     }
#NYI 
#NYI     # Check that row and col are valid without storing the values.
#NYI     return -2 if $self->_check_dimensions( $row1, $col1, 1, 1 );
#NYI     return -2 if $self->_check_dimensions( $row2, $col2, 1, 1 );
#NYI 
#NYI 
#NYI     # Check that the last parameter is a hash list.
#NYI     if ( ref $options ne 'HASH' ) {
#NYI         warn "Last parameter in conditional_formatting() "
#NYI           . "must be a hash ref";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # Copy the user params.
#NYI     my $param = {%$options};
#NYI 
#NYI     # List of valid input parameters.
#NYI     my %valid_parameter = (
#NYI         type          => 1,
#NYI         format        => 1,
#NYI         criteria      => 1,
#NYI         value         => 1,
#NYI         minimum       => 1,
#NYI         maximum       => 1,
#NYI         stop_if_true  => 1,
#NYI         min_type      => 1,
#NYI         mid_type      => 1,
#NYI         max_type      => 1,
#NYI         min_value     => 1,
#NYI         mid_value     => 1,
#NYI         max_value     => 1,
#NYI         min_color     => 1,
#NYI         mid_color     => 1,
#NYI         max_color     => 1,
#NYI         bar_color     => 1,
#NYI         icon_style    => 1,
#NYI         reverse_icons => 1,
#NYI         icons_only    => 1,
#NYI         icons         => 1,
#NYI     );
#NYI 
#NYI     # Check for valid input parameters.
#NYI     for my $param_key ( keys %$param ) {
#NYI         if ( not exists $valid_parameter{$param_key} ) {
#NYI             warn "Unknown parameter '$param_key' in conditional_formatting()";
#NYI             return -3;
#NYI         }
#NYI     }
#NYI 
#NYI     # 'type' is a required parameter.
#NYI     if ( not exists $param->{type} ) {
#NYI         warn "Parameter 'type' is required in conditional_formatting()";
#NYI         return -3;
#NYI     }
#NYI 
#NYI 
#NYI     # List of  valid validation types.
#NYI     my %valid_type = (
#NYI         'cell'          => 'cellIs',
#NYI         'date'          => 'date',
#NYI         'time'          => 'time',
#NYI         'average'       => 'aboveAverage',
#NYI         'duplicate'     => 'duplicateValues',
#NYI         'unique'        => 'uniqueValues',
#NYI         'top'           => 'top10',
#NYI         'bottom'        => 'top10',
#NYI         'text'          => 'text',
#NYI         'time_period'   => 'timePeriod',
#NYI         'blanks'        => 'containsBlanks',
#NYI         'no_blanks'     => 'notContainsBlanks',
#NYI         'errors'        => 'containsErrors',
#NYI         'no_errors'     => 'notContainsErrors',
#NYI         '2_color_scale' => '2_color_scale',
#NYI         '3_color_scale' => '3_color_scale',
#NYI         'data_bar'      => 'dataBar',
#NYI         'formula'       => 'expression',
#NYI         'icon_set'      => 'iconSet',
#NYI     );
#NYI 
#NYI     # Check for valid validation types.
#NYI     if ( not exists $valid_type{ lc( $param->{type} ) } ) {
#NYI         warn "Unknown validation type '$param->{type}' for parameter "
#NYI           . "'type' in conditional_formatting()";
#NYI         return -3;
#NYI     }
#NYI     else {
#NYI         $param->{direction} = 'bottom' if $param->{type} eq 'bottom';
#NYI         $param->{type} = $valid_type{ lc( $param->{type} ) };
#NYI     }
#NYI 
#NYI 
#NYI     # List of valid criteria types.
#NYI     my %criteria_type = (
#NYI         'between'                  => 'between',
#NYI         'not between'              => 'notBetween',
#NYI         'equal to'                 => 'equal',
#NYI         '='                        => 'equal',
#NYI         '=='                       => 'equal',
#NYI         'not equal to'             => 'notEqual',
#NYI         '!='                       => 'notEqual',
#NYI         '<>'                       => 'notEqual',
#NYI         'greater than'             => 'greaterThan',
#NYI         '>'                        => 'greaterThan',
#NYI         'less than'                => 'lessThan',
#NYI         '<'                        => 'lessThan',
#NYI         'greater than or equal to' => 'greaterThanOrEqual',
#NYI         '>='                       => 'greaterThanOrEqual',
#NYI         'less than or equal to'    => 'lessThanOrEqual',
#NYI         '<='                       => 'lessThanOrEqual',
#NYI         'containing'               => 'containsText',
#NYI         'not containing'           => 'notContains',
#NYI         'begins with'              => 'beginsWith',
#NYI         'ends with'                => 'endsWith',
#NYI         'yesterday'                => 'yesterday',
#NYI         'today'                    => 'today',
#NYI         'last 7 days'              => 'last7Days',
#NYI         'last week'                => 'lastWeek',
#NYI         'this week'                => 'thisWeek',
#NYI         'next week'                => 'nextWeek',
#NYI         'last month'               => 'lastMonth',
#NYI         'this month'               => 'thisMonth',
#NYI         'next month'               => 'nextMonth',
#NYI     );
#NYI 
#NYI     # Check for valid criteria types.
#NYI     if ( defined $param->{criteria}
#NYI         && exists $criteria_type{ lc( $param->{criteria} ) } )
#NYI     {
#NYI         $param->{criteria} = $criteria_type{ lc( $param->{criteria} ) };
#NYI     }
#NYI 
#NYI     # Convert date/times value if required.
#NYI     if ( $param->{type} eq 'date' || $param->{type} eq 'time' ) {
#NYI         $param->{type} = 'cellIs';
#NYI 
#NYI         if ( defined $param->{value} && $param->{value} =~ /T/ ) {
#NYI             my $date_time = $self->convert_date_time( $param->{value} );
#NYI 
#NYI             if ( !defined $date_time ) {
#NYI                 warn "Invalid date/time value '$param->{value}' "
#NYI                   . "in conditional_formatting()";
#NYI                 return -3;
#NYI             }
#NYI             else {
#NYI                 $param->{value} = $date_time;
#NYI             }
#NYI         }
#NYI 
#NYI         if ( defined $param->{minimum} && $param->{minimum} =~ /T/ ) {
#NYI             my $date_time = $self->convert_date_time( $param->{minimum} );
#NYI 
#NYI             if ( !defined $date_time ) {
#NYI                 warn "Invalid date/time value '$param->{minimum}' "
#NYI                   . "in conditional_formatting()";
#NYI                 return -3;
#NYI             }
#NYI             else {
#NYI                 $param->{minimum} = $date_time;
#NYI             }
#NYI         }
#NYI 
#NYI         if ( defined $param->{maximum} && $param->{maximum} =~ /T/ ) {
#NYI             my $date_time = $self->convert_date_time( $param->{maximum} );
#NYI 
#NYI             if ( !defined $date_time ) {
#NYI                 warn "Invalid date/time value '$param->{maximum}' "
#NYI                   . "in conditional_formatting()";
#NYI                 return -3;
#NYI             }
#NYI             else {
#NYI                 $param->{maximum} = $date_time;
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # List of valid icon styles.
#NYI     my %icon_set_styles = (
#NYI         "3_arrows"                => "3Arrows",            # 1
#NYI         "3_flags"                 => "3Flags",             # 2
#NYI         "3_traffic_lights_rimmed" => "3TrafficLights2",    # 3
#NYI         "3_symbols_circled"       => "3Symbols",           # 4
#NYI         "4_arrows"                => "4Arrows",            # 5
#NYI         "4_red_to_black"          => "4RedToBlack",        # 6
#NYI         "4_traffic_lights"        => "4TrafficLights",     # 7
#NYI         "5_arrows_gray"           => "5ArrowsGray",        # 8
#NYI         "5_quarters"              => "5Quarters",          # 9
#NYI         "3_arrows_gray"           => "3ArrowsGray",        # 10
#NYI         "3_traffic_lights"        => "3TrafficLights",     # 11
#NYI         "3_signs"                 => "3Signs",             # 12
#NYI         "3_symbols"               => "3Symbols2",          # 13
#NYI         "4_arrows_gray"           => "4ArrowsGray",        # 14
#NYI         "4_ratings"               => "4Rating",            # 15
#NYI         "5_arrows"                => "5Arrows",            # 16
#NYI         "5_ratings"               => "5Rating",            # 17
#NYI     );
#NYI 
#NYI 
#NYI     # Set properties for icon sets.
#NYI     if ( $param->{type} eq 'iconSet' ) {
#NYI 
#NYI         if ( !defined $param->{icon_style} ) {
#NYI             warn "The 'icon_style' parameter must be specified when "
#NYI               . "'type' == 'icon_set' in conditional_formatting()";
#NYI             return -3;
#NYI         }
#NYI 
#NYI         # Check for valid icon styles.
#NYI         if ( not exists $icon_set_styles{ $param->{icon_style} } ) {
#NYI             warn "Unknown icon style '$param->{icon_style}' for parameter "
#NYI               . "'icon_style' in conditional_formatting()";
#NYI             return -3;
#NYI         }
#NYI         else {
#NYI             $param->{icon_style} = $icon_set_styles{ $param->{icon_style} };
#NYI         }
#NYI 
#NYI         # Set the number of icons for the icon style.
#NYI         $param->{total_icons} = 3;
#NYI         if ( $param->{icon_style} =~ /^4/ ) {
#NYI             $param->{total_icons} = 4;
#NYI         }
#NYI         elsif ( $param->{icon_style} =~ /^5/ ) {
#NYI             $param->{total_icons} = 5;
#NYI         }
#NYI 
#NYI         $param->{icons} =
#NYI           $self->_set_icon_properties( $param->{total_icons}, $param->{icons} );
#NYI     }
#NYI 
#NYI 
#NYI     # Set the formatting range.
#NYI     my $range      = '';
#NYI     my $start_cell = '';    # Use for formulas.
#NYI 
#NYI     # Swap last row/col for first row/col as necessary
#NYI     if ( $row1 > $row2 ) {
#NYI         ( $row1, $row2 ) = ( $row2, $row1 );
#NYI     }
#NYI 
#NYI     if ( $col1 > $col2 ) {
#NYI         ( $col1, $col2 ) = ( $col2, $col1 );
#NYI     }
#NYI 
#NYI     # If the first and last cell are the same write a single cell.
#NYI     if ( ( $row1 == $row2 ) && ( $col1 == $col2 ) ) {
#NYI         $range = xl-rowcol-to-cell( $row1, $col1 );
#NYI         $start_cell = $range;
#NYI     }
#NYI     else {
#NYI         $range = xl-range( $row1, $row2, $col1, $col2 );
#NYI         $start_cell = xl-rowcol-to-cell( $row1, $col1 );
#NYI     }
#NYI 
#NYI     # Override with user defined multiple range if provided.
#NYI     if ( $user_range ) {
#NYI         $range = $user_range;
#NYI     }
#NYI 
#NYI     # Get the dxf format index.
#NYI     if ( defined $param->{format} && ref $param->{format} ) {
#NYI         $param->{format} = $param->{format}->get_dxf_index();
#NYI     }
#NYI 
#NYI     # Set the priority based on the order of adding.
#NYI     $param->{priority} = $self->{_dxf_priority}++;
#NYI 
#NYI     # Special handling of text criteria.
#NYI     if ( $param->{type} eq 'text' ) {
#NYI 
#NYI         if ( $param->{criteria} eq 'containsText' ) {
#NYI             $param->{type}    = 'containsText';
#NYI             $param->{formula} = sprintf 'NOT(ISERROR(SEARCH("%s",%s)))',
#NYI               $param->{value}, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'notContains' ) {
#NYI             $param->{type}    = 'notContainsText';
#NYI             $param->{formula} = sprintf 'ISERROR(SEARCH("%s",%s))',
#NYI               $param->{value}, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'beginsWith' ) {
#NYI             $param->{type}    = 'beginsWith';
#NYI             $param->{formula} = sprintf 'LEFT(%s,%d)="%s"',
#NYI               $start_cell, length( $param->{value} ), $param->{value};
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'endsWith' ) {
#NYI             $param->{type}    = 'endsWith';
#NYI             $param->{formula} = sprintf 'RIGHT(%s,%d)="%s"',
#NYI               $start_cell, length( $param->{value} ), $param->{value};
#NYI         }
#NYI         else {
#NYI             warn "Invalid text criteria '$param->{criteria}' "
#NYI               . "in conditional_formatting()";
#NYI         }
#NYI     }
#NYI 
#NYI     # Special handling of time time_period criteria.
#NYI     if ( $param->{type} eq 'timePeriod' ) {
#NYI 
#NYI         if ( $param->{criteria} eq 'yesterday' ) {
#NYI             $param->{formula} = sprintf 'FLOOR(%s,1)=TODAY()-1', $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'today' ) {
#NYI             $param->{formula} = sprintf 'FLOOR(%s,1)=TODAY()', $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'tomorrow' ) {
#NYI             $param->{formula} = sprintf 'FLOOR(%s,1)=TODAY()+1', $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'last7Days' ) {
#NYI             $param->{formula} =
#NYI               sprintf 'AND(TODAY()-FLOOR(%s,1)<=6,FLOOR(%s,1)<=TODAY())',
#NYI               $start_cell, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'lastWeek' ) {
#NYI             $param->{formula} =
#NYI               sprintf 'AND(TODAY()-ROUNDDOWN(%s,0)>=(WEEKDAY(TODAY())),'
#NYI               . 'TODAY()-ROUNDDOWN(%s,0)<(WEEKDAY(TODAY())+7))',
#NYI               $start_cell, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'thisWeek' ) {
#NYI             $param->{formula} =
#NYI               sprintf 'AND(TODAY()-ROUNDDOWN(%s,0)<=WEEKDAY(TODAY())-1,'
#NYI               . 'ROUNDDOWN(%s,0)-TODAY()<=7-WEEKDAY(TODAY()))',
#NYI               $start_cell, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'nextWeek' ) {
#NYI             $param->{formula} =
#NYI               sprintf 'AND(ROUNDDOWN(%s,0)-TODAY()>(7-WEEKDAY(TODAY())),'
#NYI               . 'ROUNDDOWN(%s,0)-TODAY()<(15-WEEKDAY(TODAY())))',
#NYI               $start_cell, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'lastMonth' ) {
#NYI             $param->{formula} =
#NYI               sprintf
#NYI               'AND(MONTH(%s)=MONTH(TODAY())-1,OR(YEAR(%s)=YEAR(TODAY()),'
#NYI               . 'AND(MONTH(%s)=1,YEAR(A1)=YEAR(TODAY())-1)))',
#NYI               $start_cell, $start_cell, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'thisMonth' ) {
#NYI             $param->{formula} =
#NYI               sprintf 'AND(MONTH(%s)=MONTH(TODAY()),YEAR(%s)=YEAR(TODAY()))',
#NYI               $start_cell, $start_cell;
#NYI         }
#NYI         elsif ( $param->{criteria} eq 'nextMonth' ) {
#NYI             $param->{formula} =
#NYI               sprintf
#NYI               'AND(MONTH(%s)=MONTH(TODAY())+1,OR(YEAR(%s)=YEAR(TODAY()),'
#NYI               . 'AND(MONTH(%s)=12,YEAR(%s)=YEAR(TODAY())+1)))',
#NYI               $start_cell, $start_cell, $start_cell, $start_cell;
#NYI         }
#NYI         else {
#NYI             warn "Invalid time_period criteria '$param->{criteria}' "
#NYI               . "in conditional_formatting()";
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # Special handling of blanks/error types.
#NYI     if ( $param->{type} eq 'containsBlanks' ) {
#NYI         $param->{formula} = sprintf 'LEN(TRIM(%s))=0', $start_cell;
#NYI     }
#NYI 
#NYI     if ( $param->{type} eq 'notContainsBlanks' ) {
#NYI         $param->{formula} = sprintf 'LEN(TRIM(%s))>0', $start_cell;
#NYI     }
#NYI 
#NYI     if ( $param->{type} eq 'containsErrors' ) {
#NYI         $param->{formula} = sprintf 'ISERROR(%s)', $start_cell;
#NYI     }
#NYI 
#NYI     if ( $param->{type} eq 'notContainsErrors' ) {
#NYI         $param->{formula} = sprintf 'NOT(ISERROR(%s))', $start_cell;
#NYI     }
#NYI 
#NYI 
#NYI     # Special handling for 2 color scale.
#NYI     if ( $param->{type} eq '2_color_scale' ) {
#NYI         $param->{type} = 'colorScale';
#NYI 
#NYI         # Color scales don't use any additional formatting.
#NYI         $param->{format} = undef;
#NYI 
#NYI         # Turn off 3 color parameters.
#NYI         $param->{mid_type}  = undef;
#NYI         $param->{mid_color} = undef;
#NYI 
#NYI         $param->{min_type}  ||= 'min';
#NYI         $param->{max_type}  ||= 'max';
#NYI         $param->{min_value} ||= 0;
#NYI         $param->{max_value} ||= 0;
#NYI         $param->{min_color} ||= '#FF7128';
#NYI         $param->{max_color} ||= '#FFEF9C';
#NYI 
#NYI         $param->{max_color} = $self->_get_palette_color( $param->{max_color} );
#NYI         $param->{min_color} = $self->_get_palette_color( $param->{min_color} );
#NYI     }
#NYI 
#NYI 
#NYI     # Special handling for 3 color scale.
#NYI     if ( $param->{type} eq '3_color_scale' ) {
#NYI         $param->{type} = 'colorScale';
#NYI 
#NYI         # Color scales don't use any additional formatting.
#NYI         $param->{format} = undef;
#NYI 
#NYI         $param->{min_type}  ||= 'min';
#NYI         $param->{mid_type}  ||= 'percentile';
#NYI         $param->{max_type}  ||= 'max';
#NYI         $param->{min_value} ||= 0;
#NYI         $param->{mid_value} = 50 unless defined $param->{mid_value};
#NYI         $param->{max_value} ||= 0;
#NYI         $param->{min_color} ||= '#F8696B';
#NYI         $param->{mid_color} ||= '#FFEB84';
#NYI         $param->{max_color} ||= '#63BE7B';
#NYI 
#NYI         $param->{max_color} = $self->_get_palette_color( $param->{max_color} );
#NYI         $param->{mid_color} = $self->_get_palette_color( $param->{mid_color} );
#NYI         $param->{min_color} = $self->_get_palette_color( $param->{min_color} );
#NYI     }
#NYI 
#NYI 
#NYI     # Special handling for data bar.
#NYI     if ( $param->{type} eq 'dataBar' ) {
#NYI 
#NYI         # Color scales don't use any additional formatting.
#NYI         $param->{format} = undef;
#NYI 
#NYI         $param->{min_type}  ||= 'min';
#NYI         $param->{max_type}  ||= 'max';
#NYI         $param->{min_value} ||= 0;
#NYI         $param->{max_value} ||= 0;
#NYI         $param->{bar_color} ||= '#638EC6';
#NYI 
#NYI         $param->{bar_color} = $self->_get_palette_color( $param->{bar_color} );
#NYI     }
#NYI 
#NYI 
#NYI     # Store the validation information until we close the worksheet.
#NYI     push @{ $self->{_cond_formats}->{$range} }, $param;
#NYI }


###############################################################################
#
# Set the sub-properites for icons.
#
#NYI sub _set_icon_properties {
#NYI 
#NYI     my $self        = shift;
#NYI     my $total_icons = shift;
#NYI     my $user_props  = shift;
#NYI     my $props       = [];
#NYI 
#NYI     # Set the default icon properties.
#NYI     for ( 0 .. $total_icons - 1 ) {
#NYI         push @$props,
#NYI           {
#NYI             criteria => 0,
#NYI             value    => 0,
#NYI             type     => 'percent'
#NYI           };
#NYI     }
#NYI 
#NYI     # Set the default icon values based on the number of icons.
#NYI     if ( $total_icons == 3 ) {
#NYI         $props->[0]->{value} = 67;
#NYI         $props->[1]->{value} = 33;
#NYI     }
#NYI 
#NYI     if ( $total_icons == 4 ) {
#NYI         $props->[0]->{value} = 75;
#NYI         $props->[1]->{value} = 50;
#NYI         $props->[2]->{value} = 25;
#NYI     }
#NYI 
#NYI     if ( $total_icons == 5 ) {
#NYI         $props->[0]->{value} = 80;
#NYI         $props->[1]->{value} = 60;
#NYI         $props->[2]->{value} = 40;
#NYI         $props->[3]->{value} = 20;
#NYI     }
#NYI 
#NYI     # Overwrite default properties with user defined properties.
#NYI     if ( defined $user_props ) {
#NYI 
#NYI         # Ensure we don't set user properties for lowest icon.
#NYI         my $max_data = @$user_props;
#NYI         if ( $max_data >= $total_icons ) {
#NYI             $max_data = $total_icons -1;
#NYI         }
#NYI 
#NYI         for my $i ( 0 .. $max_data - 1 ) {
#NYI 
#NYI             # Set the user defined 'value' property.
#NYI             if ( defined $user_props->[$i]->{value} ) {
#NYI                 $props->[$i]->{value} = $user_props->[$i]->{value};
#NYI                 $props->[$i]->{value} =~ s/^=//;
#NYI             }
#NYI 
#NYI             # Set the user defined 'type' property.
#NYI             if ( defined $user_props->[$i]->{type} ) {
#NYI 
#NYI                 my $type = $user_props->[$i]->{type};
#NYI 
#NYI                 if (   $type ne 'percent'
#NYI                     && $type ne 'percentile'
#NYI                     && $type ne 'number'
#NYI                     && $type ne 'formula' )
#NYI                 {
#NYI                     warn "Unknown icon property type '$props->{type}' for sub-"
#NYI                       . "property 'type' in conditional_formatting()";
#NYI                 }
#NYI                 else {
#NYI                     $props->[$i]->{type} = $type;
#NYI 
#NYI                     if ( $props->[$i]->{type} eq 'number' ) {
#NYI                         $props->[$i]->{type} = 'num';
#NYI                     }
#NYI                 }
#NYI             }
#NYI 
#NYI             # Set the user defined 'criteria' property.
#NYI             if ( defined $user_props->[$i]->{criteria}
#NYI                 && $user_props->[$i]->{criteria} eq '>' )
#NYI             {
#NYI                 $props->[$i]->{criteria} = 1;
#NYI             }
#NYI 
#NYI         }
#NYI 
#NYI     }
#NYI 
#NYI     return $props;
#NYI }


###############################################################################
#
# add_table()
#
# Add an Excel table to a worksheet.
#
#NYI sub add_table {
#NYI 
#NYI     my $self       = shift;
#NYI     my $user_range = '';
#NYI     my %table;
#NYI     my @col_formats;
#NYI 
#NYI     # We would need to order the write statements very carefully within this
#NYI     # function to support optimisation mode. Disable add_table() when it is
#NYI     # on for now.
#NYI     if ( $self->{_optimization} == 1 ) {
#NYI         warn "add_table() isn't supported when set_optimization() is on";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( @_ && $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Check for a valid number of args.
#NYI     if ( @_ < 4 ) {
#NYI         warn "Not enough parameters to add_table()";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     my ( $row1, $col1, $row2, $col2 ) = @_;
#NYI 
#NYI     # Check that row and col are valid without storing the values.
#NYI     return -2 if $self->_check_dimensions( $row1, $col1, 1, 1 );
#NYI     return -2 if $self->_check_dimensions( $row2, $col2, 1, 1 );
#NYI 
#NYI 
#NYI     # The final hashref contains the validation parameters.
#NYI     my $param = $_[4] || {};
#NYI 
#NYI     # Check that the last parameter is a hash list.
#NYI     if ( ref $param ne 'HASH' ) {
#NYI         warn "Last parameter '$param' in add_table() must be a hash ref";
#NYI         return -3;
#NYI     }
#NYI 
#NYI 
#NYI     # List of valid input parameters.
#NYI     my %valid_parameter = (
#NYI         autofilter     => 1,
#NYI         banded_columns => 1,
#NYI         banded_rows    => 1,
#NYI         columns        => 1,
#NYI         data           => 1,
#NYI         first_column   => 1,
#NYI         header_row     => 1,
#NYI         last_column    => 1,
#NYI         name           => 1,
#NYI         style          => 1,
#NYI         total_row      => 1,
#NYI     );
#NYI 
#NYI     # Check for valid input parameters.
#NYI     for my $param_key ( keys %$param ) {
#NYI         if ( not exists $valid_parameter{$param_key} ) {
#NYI             warn "Unknown parameter '$param_key' in add_table()";
#NYI             return -3;
#NYI         }
#NYI     }
#NYI 
#NYI     # Turn on Excel's defaults.
#NYI     $param->{banded_rows} = 1 if !defined $param->{banded_rows};
#NYI     $param->{header_row}  = 1 if !defined $param->{header_row};
#NYI     $param->{autofilter}  = 1 if !defined $param->{autofilter};
#NYI 
#NYI     # Set the table options.
#NYI     $table{_show_first_col}   = $param->{first_column}   ? 1 : 0;
#NYI     $table{_show_last_col}    = $param->{last_column}    ? 1 : 0;
#NYI     $table{_show_row_stripes} = $param->{banded_rows}    ? 1 : 0;
#NYI     $table{_show_col_stripes} = $param->{banded_columns} ? 1 : 0;
#NYI     $table{_header_row_count} = $param->{header_row}     ? 1 : 0;
#NYI     $table{_totals_row_shown} = $param->{total_row}      ? 1 : 0;
#NYI 
#NYI 
#NYI     # Set the table name.
#NYI     if ( defined $param->{name} ) {
#NYI         my $name = $param->{name};
#NYI 
#NYI         # Warn if the name contains invalid chars as defined by Excel help.
#NYI         if ( $name !~ m/^[\w\\][\w\\.]*$/ || $name =~ m/^\d/ ) {
#NYI             warn "Invalid character in name '$name' used in add_table()";
#NYI             return -3;
#NYI         }
#NYI 
#NYI         # Warn if the name looks like a cell name.
#NYI         if ( $name =~ m/^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$/ ) {
#NYI             warn "Invalid name '$name' looks like a cell name in add_table()";
#NYI             return -3;
#NYI         }
#NYI 
#NYI         # Warn if the name looks like a R1C1.
#NYI         if ( $name =~ m/^[rcRC]$/ || $name =~ m/^[rcRC]\d+[rcRC]\d+$/ ) {
#NYI             warn "Invalid name '$name' like a RC cell ref in add_table()";
#NYI             return -3;
#NYI         }
#NYI 
#NYI         $table{_name} = $param->{name};
#NYI     }
#NYI 
#NYI     # Set the table style.
#NYI     if ( defined $param->{style} ) {
#NYI         $table{_style} = $param->{style};
#NYI 
#NYI         # Remove whitespace from style name.
#NYI         $table{_style} =~ s/\s//g;
#NYI     }
#NYI     else {
#NYI         $table{_style} = "TableStyleMedium9";
#NYI     }
#NYI 
#NYI 
#NYI     # Swap last row/col for first row/col as necessary.
#NYI     if ( $row1 > $row2 ) {
#NYI         ( $row1, $row2 ) = ( $row2, $row1 );
#NYI     }
#NYI 
#NYI     if ( $col1 > $col2 ) {
#NYI         ( $col1, $col2 ) = ( $col2, $col1 );
#NYI     }
#NYI 
#NYI 
#NYI     # Set the data range rows (without the header and footer).
#NYI     my $first_data_row = $row1;
#NYI     my $last_data_row  = $row2;
#NYI     $first_data_row++ if $param->{header_row};
#NYI     $last_data_row--  if $param->{total_row};
#NYI 
#NYI 
#NYI     # Set the table and autofilter ranges.
#NYI     $table{_range}   = xl-range( $row1, $row2,          $col1, $col2 );
#NYI     $table{_a_range} = xl-range( $row1, $last_data_row, $col1, $col2 );
#NYI 
#NYI 
#NYI     # If the header row if off the default is to turn autofilter off.
#NYI     if ( !$param->{header_row} ) {
#NYI         $param->{autofilter} = 0;
#NYI     }
#NYI 
#NYI     # Set the autofilter range.
#NYI     if ( $param->{autofilter} ) {
#NYI         $table{_autofilter} = $table{_a_range};
#NYI     }
#NYI 
#NYI     # Add the table columns.
#NYI     my %seen_names;
#NYI     my $col_id = 1;
#NYI     for my $col_num ( $col1 .. $col2 ) {
#NYI 
#NYI         # Set up the default column data.
#NYI         my $col_data = {
#NYI             _id             => $col_id,
#NYI             _name           => 'Column' . $col_id,
#NYI             _total_string   => '',
#NYI             _total_function => '',
#NYI             _formula        => '',
#NYI             _format         => undef,
#NYI             _name_format    => undef,
#NYI         };
#NYI 
#NYI         # Overwrite the defaults with any use defined values.
#NYI         if ( $param->{columns} ) {
#NYI 
#NYI             # Check if there are user defined values for this column.
#NYI             if ( my $user_data = $param->{columns}->[ $col_id - 1 ] ) {
#NYI 
#NYI                 # Map user defined values to internal values.
#NYI                 $col_data->{_name} = $user_data->{header}
#NYI                   if $user_data->{header};
#NYI 
#NYI                 # Excel requires unique case insensitive header names.
#NYI                 my $name = $col_data->{_name};
#NYI                 my $key = lc $name;
#NYI                 if (exists $seen_names{$key}) {
#NYI                     warn "add_table() contains duplicate name: '$name'";
#NYI                     return -1;
#NYI                 }
#NYI                 else {
#NYI                     $seen_names{$key} = 1;
#NYI                 }
#NYI 
#NYI                 # Get the header format if defined.
#NYI                 $col_data->{_name_format} = $user_data->{header_format};
#NYI 
#NYI                 # Handle the column formula.
#NYI                 if ( $user_data->{formula} ) {
#NYI                     my $formula = $user_data->{formula};
#NYI 
#NYI                     # Remove the leading = from formula.
#NYI                     $formula =~ s/^=//;
#NYI 
#NYI                     # Covert Excel 2010 "@" ref to 2007 "#This Row".
#NYI                     $formula =~ s/@/[#This Row],/g;
#NYI 
#NYI                     $col_data->{_formula} = $formula;
#NYI 
#NYI                     for my $row ( $first_data_row .. $last_data_row ) {
#NYI                         $self->write_formula( $row, $col_num, $formula,
#NYI                             $user_data->{format} );
#NYI                     }
#NYI                 }
#NYI 
#NYI                 # Handle the function for the total row.
#NYI                 if ( $user_data->{total_function} ) {
#NYI                     my $function = $user_data->{total_function};
#NYI 
#NYI                     # Massage the function name.
#NYI                     $function = lc $function;
#NYI                     $function =~ s/_//g;
#NYI                     $function =~ s/\s//g;
#NYI 
#NYI                     $function = 'countNums' if $function eq 'countnums';
#NYI                     $function = 'stdDev'    if $function eq 'stddev';
#NYI 
#NYI                     $col_data->{_total_function} = $function;
#NYI 
#NYI                     my $formula = _table_function_to_formula(
#NYI                         $function,
#NYI                         $col_data->{_name}
#NYI 
#NYI                     );
#NYI 
#NYI                     my $value = $user_data->{total_value} || 0;
#NYI 
#NYI                     $self->write_formula( $row2, $col_num, $formula,
#NYI                         $user_data->{format}, $value );
#NYI 
#NYI                 }
#NYI                 elsif ( $user_data->{total_string} ) {
#NYI 
#NYI                     # Total label only (not a function).
#NYI                     my $total_string = $user_data->{total_string};
#NYI                     $col_data->{_total_string} = $total_string;
#NYI 
#NYI                     $self->write_string( $row2, $col_num, $total_string,
#NYI                         $user_data->{format} );
#NYI                 }
#NYI 
#NYI                 # Get the dxf format index.
#NYI                 if ( defined $user_data->{format} && ref $user_data->{format} )
#NYI                 {
#NYI                     $col_data->{_format} =
#NYI                       $user_data->{format}->get_dxf_index();
#NYI                 }
#NYI 
#NYI                 # Store the column format for writing the cell data.
#NYI                 # It doesn't matter if it is undefined.
#NYI                 $col_formats[ $col_id - 1 ] = $user_data->{format};
#NYI             }
#NYI         }
#NYI 
#NYI         # Store the column data.
#NYI         push @{ $table{_columns} }, $col_data;
#NYI 
#NYI         # Write the column headers to the worksheet.
#NYI         if ( $param->{header_row} ) {
#NYI             $self->write_string( $row1, $col_num, $col_data->{_name},
#NYI                 $col_data->{_name_format} );
#NYI         }
#NYI 
#NYI         $col_id++;
#NYI     }    # Table columns.
#NYI 
#NYI 
#NYI     # Write the cell data if supplied.
#NYI     if ( my $data = $param->{data} ) {
#NYI 
#NYI         my $i = 0;    # For indexing the row data.
#NYI         for my $row ( $first_data_row .. $last_data_row ) {
#NYI             my $j = 0;    # For indexing the col data.
#NYI 
#NYI             for my $col ( $col1 .. $col2 ) {
#NYI 
#NYI                 my $token = $data->[$i]->[$j];
#NYI 
#NYI                 if ( defined $token ) {
#NYI                     $self->write( $row, $col, $token, $col_formats[$j] );
#NYI                 }
#NYI 
#NYI                 $j++;
#NYI             }
#NYI             $i++;
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # Store the table data.
#NYI     push @{ $self->{_tables} }, \%table;
#NYI 
#NYI     return \%table;
#NYI }


###############################################################################
#
# add_sparkline()
#
# Add sparklines to the worksheet.
#
#NYI sub add_sparkline {
#NYI 
#NYI     my $self      = shift;
#NYI     my $param     = shift;
#NYI     my $sparkline = {};
#NYI 
#NYI     # Check that the last parameter is a hash list.
#NYI     if ( ref $param ne 'HASH' ) {
#NYI         warn "Parameter list in add_sparkline() must be a hash ref";
#NYI         return -1;
#NYI     }
#NYI 
#NYI     # List of valid input parameters.
#NYI     my %valid_parameter = (
#NYI         location        => 1,
#NYI         range           => 1,
#NYI         type            => 1,
#NYI         high_point      => 1,
#NYI         low_point       => 1,
#NYI         negative_points => 1,
#NYI         first_point     => 1,
#NYI         last_point      => 1,
#NYI         markers         => 1,
#NYI         style           => 1,
#NYI         series_color    => 1,
#NYI         negative_color  => 1,
#NYI         markers_color   => 1,
#NYI         first_color     => 1,
#NYI         last_color      => 1,
#NYI         high_color      => 1,
#NYI         low_color       => 1,
#NYI         max             => 1,
#NYI         min             => 1,
#NYI         axis            => 1,
#NYI         reverse         => 1,
#NYI         empty_cells     => 1,
#NYI         show_hidden     => 1,
#NYI         plot_hidden     => 1,
#NYI         date_axis       => 1,
#NYI         weight          => 1,
#NYI     );
#NYI 
#NYI     # Check for valid input parameters.
#NYI     for my $param_key ( keys %$param ) {
#NYI         if ( not exists $valid_parameter{$param_key} ) {
#NYI             warn "Unknown parameter '$param_key' in add_sparkline()";
#NYI             return -2;
#NYI         }
#NYI     }
#NYI 
#NYI     # 'location' is a required parameter.
#NYI     if ( not exists $param->{location} ) {
#NYI         warn "Parameter 'location' is required in add_sparkline()";
#NYI         return -3;
#NYI     }
#NYI 
#NYI     # 'range' is a required parameter.
#NYI     if ( not exists $param->{range} ) {
#NYI         warn "Parameter 'range' is required in add_sparkline()";
#NYI         return -3;
#NYI     }
#NYI 
#NYI 
#NYI     # Handle the sparkline type.
#NYI     my $type = $param->{type} || 'line';
#NYI 
#NYI     if ( $type ne 'line' && $type ne 'column' && $type ne 'win_loss' ) {
#NYI         warn "Parameter 'type' must be 'line', 'column' "
#NYI           . "or 'win_loss' in add_sparkline()";
#NYI         return -4;
#NYI     }
#NYI 
#NYI     $type = 'stacked' if $type eq 'win_loss';
#NYI     $sparkline->{_type} = $type;
#NYI 
#NYI 
#NYI     # We handle single location/range values or array refs of values.
#NYI     if ( ref $param->{location} ) {
#NYI         $sparkline->{_locations} = $param->{location};
#NYI         $sparkline->{_ranges}    = $param->{range};
#NYI     }
#NYI     else {
#NYI         $sparkline->{_locations} = [ $param->{location} ];
#NYI         $sparkline->{_ranges}    = [ $param->{range} ];
#NYI     }
#NYI 
#NYI     my $range_count    = @{ $sparkline->{_ranges} };
#NYI     my $location_count = @{ $sparkline->{_locations} };
#NYI 
#NYI     # The ranges and locations must match.
#NYI     if ( $range_count != $location_count ) {
#NYI         warn "Must have the same number of location and range "
#NYI           . "parameters in add_sparkline()";
#NYI         return -5;
#NYI     }
#NYI 
#NYI     # Store the count.
#NYI     $sparkline->{_count} = @{ $sparkline->{_locations} };
#NYI 
#NYI 
#NYI     # Get the worksheet name for the range conversion below.
#NYI     my $sheetname = quote-sheetname( $self->{_name} );
#NYI 
#NYI     # Cleanup the input ranges.
#NYI     for my $range ( @{ $sparkline->{_ranges} } ) {
#NYI 
#NYI         # Remove the absolute reference $ symbols.
#NYI         $range =~ s{\$}{}g;
#NYI 
#NYI         # Remove the = from xl-range-formula(.
#NYI         $range =~ s{^=}{};
#NYI 
#NYI         # Convert a simple range into a full Sheet1!A1:D1 range.
#NYI         if ( $range !~ /!/ ) {
#NYI             $range = $sheetname . "!" . $range;
#NYI         }
#NYI     }
#NYI 
#NYI     # Cleanup the input locations.
#NYI     for my $location ( @{ $sparkline->{_locations} } ) {
#NYI         $location =~ s{\$}{}g;
#NYI     }
#NYI 
#NYI     # Map options.
#NYI     $sparkline->{_high}     = $param->{high_point};
#NYI     $sparkline->{_low}      = $param->{low_point};
#NYI     $sparkline->{_negative} = $param->{negative_points};
#NYI     $sparkline->{_first}    = $param->{first_point};
#NYI     $sparkline->{_last}     = $param->{last_point};
#NYI     $sparkline->{_markers}  = $param->{markers};
#NYI     $sparkline->{_min}      = $param->{min};
#NYI     $sparkline->{_max}      = $param->{max};
#NYI     $sparkline->{_axis}     = $param->{axis};
#NYI     $sparkline->{_reverse}  = $param->{reverse};
#NYI     $sparkline->{_hidden}   = $param->{show_hidden};
#NYI     $sparkline->{_weight}   = $param->{weight};
#NYI 
#NYI     # Map empty cells options.
#NYI     my $empty = $param->{empty_cells} || '';
#NYI 
#NYI     if ( $empty eq 'zero' ) {
#NYI         $sparkline->{_empty} = 0;
#NYI     }
#NYI     elsif ( $empty eq 'connect' ) {
#NYI         $sparkline->{_empty} = 'span';
#NYI     }
#NYI     else {
#NYI         $sparkline->{_empty} = 'gap';
#NYI     }
#NYI 
#NYI 
#NYI     # Map the date axis range.
#NYI     my $date_range = $param->{date_axis};
#NYI 
#NYI     if ( $date_range && $date_range !~ /!/ ) {
#NYI         $date_range = $sheetname . "!" . $date_range;
#NYI     }
#NYI     $sparkline->{_date_axis} = $date_range;
#NYI 
#NYI 
#NYI     # Set the sparkline styles.
#NYI     my $style_id = $param->{style} || 0;
#NYI     my $style = $Excel::Writer::XLSX::Package::Theme::spark_styles[$style_id];
#NYI 
#NYI     $sparkline->{_series_color}   = $style->{series};
#NYI     $sparkline->{_negative_color} = $style->{negative};
#NYI     $sparkline->{_markers_color}  = $style->{markers};
#NYI     $sparkline->{_first_color}    = $style->{first};
#NYI     $sparkline->{_last_color}     = $style->{last};
#NYI     $sparkline->{_high_color}     = $style->{high};
#NYI     $sparkline->{_low_color}      = $style->{low};
#NYI 
#NYI     # Override the style colours with user defined colors.
#NYI     $self->_set_spark_color( $sparkline, $param, 'series_color' );
#NYI     $self->_set_spark_color( $sparkline, $param, 'negative_color' );
#NYI     $self->_set_spark_color( $sparkline, $param, 'markers_color' );
#NYI     $self->_set_spark_color( $sparkline, $param, 'first_color' );
#NYI     $self->_set_spark_color( $sparkline, $param, 'last_color' );
#NYI     $self->_set_spark_color( $sparkline, $param, 'high_color' );
#NYI     $self->_set_spark_color( $sparkline, $param, 'low_color' );
#NYI 
#NYI     push @{ $self->{_sparklines} }, $sparkline;
#NYI }


###############################################################################
#
# insert_button()
#
# Insert a button form object into the worksheet.
#
#NYI sub insert_button {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Check the number of args.
#NYI     if ( @_ < 3 ) { return -1 }
#NYI 
#NYI     my $button = $self->_button_params( @_ );
#NYI 
#NYI     push @{ $self->{_buttons_array} }, $button;
#NYI 
#NYI     $self->{_has_vml} = 1;
#NYI }


###############################################################################
#
# set_vba_name()
#
# Set the VBA name for the worksheet.
#
#NYI sub set_vba_name {
#NYI 
#NYI     my $self         = shift;
#NYI     my $vba_codemame = shift;
#NYI 
#NYI     if ( $vba_codemame ) {
#NYI         $self->{_vba_codename} = $vba_codemame;
#NYI     }
#NYI     else {
#NYI         $self->{_vba_codename} = $self->{_name};
#NYI     }
#NYI }


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _table_function_to_formula
#
# Convert a table total function to a worksheet formula.
#
#NYI sub _table_function_to_formula {
#NYI 
#NYI     my $function = shift;
#NYI     my $col_name = shift;
#NYI     my $formula  = '';
#NYI 
#NYI     my %subtotals = (
#NYI         average   => 101,
#NYI         countNums => 102,
#NYI         count     => 103,
#NYI         max       => 104,
#NYI         min       => 105,
#NYI         stdDev    => 107,
#NYI         sum       => 109,
#NYI         var       => 110,
#NYI     );
#NYI 
#NYI     if ( exists $subtotals{$function} ) {
#NYI         my $func_num = $subtotals{$function};
#NYI         $formula = qq{SUBTOTAL($func_num,[$col_name])};
#NYI     }
#NYI     else {
#NYI         warn "Unsupported function '$function' in add_table()";
#NYI     }
#NYI 
#NYI     return $formula;
#NYI }


###############################################################################
#
# _set_spark_color()
#
# Set the sparkline colour.
#
#NYI sub _set_spark_color {
#NYI 
#NYI     my $self        = shift;
#NYI     my $sparkline   = shift;
#NYI     my $param       = shift;
#NYI     my $user_color  = shift;
#NYI     my $spark_color = '_' . $user_color;
#NYI 
#NYI     return unless $param->{$user_color};
#NYI 
#NYI     $sparkline->{$spark_color} =
#NYI       { _rgb => $self->_get_palette_color( $param->{$user_color} ) };
#NYI }


###############################################################################
#
# _get_palette_color()
#
# Convert from an Excel internal colour index to a XML style #RRGGBB index
# based on the default or user defined values in the Workbook palette.
#
#NYI sub _get_palette_color {
#NYI 
#NYI     my $self    = shift;
#NYI     my $index   = shift;
#NYI     my $palette = $self->{_palette};
#NYI 
#NYI     # Handle colours in #XXXXXX RGB format.
#NYI     if ( $index =~ m/^#([0-9A-F]{6})$/i ) {
#NYI         return "FF" . uc( $1 );
#NYI     }
#NYI 
#NYI     # Adjust the colour index.
#NYI     $index -= 8;
#NYI 
#NYI     # Palette is passed in from the Workbook class.
#NYI     my @rgb = @{ $palette->[$index] };
#NYI 
#NYI     return sprintf "FF%02X%02X%02X", @rgb[0, 1, 2];
#NYI }


###############################################################################
#
# _substitute_cellref()
#
# Substitute an Excel cell reference in A1 notation for  zero based row and
# column values in an argument list.
#
# Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
#
method substitute_cellref($cell, *@args) {
    $cell .=  uc;

    # Convert a column range: 'A:A' or 'B:G'.
    # A range such as A:A is equivalent to A1:Rowmax, so add rows as required
    if $cell ~~ /
                 \$?
                 (<[A..Z]> ** 1..3)
                 ':'
                 \$?
                 (<[A..Z]> ** 1..3)
                / {
        my ( $row1, $col1 ) = self.cell_to_rowcol( $0 ~ '1' );
        my ( $row2, $col2 ) = self.cell_to_rowcol( $1 ~ $!xls_rowmax );
        return $row1, $col1, $row2, $col2, |@args;
    }

    # Convert a cell range: 'A1:B7'
    if $cell ~~ /
                 \$?
                 (<[A..Z]> ** 1..3 \$? \d+)
                 ':'
                 \$?
                 (<[A..Z]> ** 1..3 \$? \d+)
                / {
        my ( $row1, $col1 ) = self.cell_to_rowcol( $0 );
        my ( $row2, $col2 ) = self.cell_to_rowcol( $1 );
        return $row1, $col1, $row2, $col2, |@args;
    }

    # Convert a cell reference: 'A1' or 'AD2000'
    if $cell ~~ /
                 \$?
                 (<[A..Z]> ** 1..3 \$? \d+)
                / {
        my ( $row1, $col1 ) = self.cell_to_rowcol( $0 );
        return $row1, $col1, |@args;

    }

    fail( "Unknown cell reference $cell" );
}


###############################################################################
#
# _cell_to_rowcol($cell_ref)
#
# Convert an Excel cell reference in A1 notation to a zero based row and column
# reference; converts C1 to (0, 2).
#
# See also: http://www.perlmonks.org/index.pl?node_id=270352
#
# Returns: ($row, $col, $row_absolute, $col_absolute)
#
#
method cell_to_rowcol($cell) {
    $cell ~~ /
              (\$?)
              (<[A..Z]> ** 1..3 )
              (\$?)
              (\d+)
             /;

    my $col-abs = $0 eq "" ?? 0 !! 1;
    my $col     = $1;
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

    # TODO Check row and column range
    return $row, $col, $row-abs, $col-abs;
}


###############################################################################
#
# _xl_rowcol_to_cell($row, $col)
#
# Optimised version of xl_rowcol_to_cell from Utility.pm for the inner loop
# of _write_cell().
#

our @col_names = ( 'A' .. 'XFD' ); # CHECK

method xl-rowcol-to-cell($row, $col) {
    return @col_names[ $col ] ~ ( $row + 1 );
}


###############################################################################
#
# _sort_pagebreaks()
#
# This is an internal method that is used to filter elements of the array of
# pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
#   1. Removes duplicate entries from the list.
#   2. Sorts the list.
#   3. Removes 0 from the list if present.
#
#NYI sub _sort_pagebreaks {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return () unless @_;
#NYI 
#NYI     my %hash;
#NYI     my @array;
#NYI 
#NYI     @hash{@_} = undef;    # Hash slice to remove duplicates
#NYI     @array = sort { $a <=> $b } keys %hash;    # Numerical sort
#NYI     shift @array if $array[0] == 0;            # Remove zero
#NYI 
#NYI     # The Excel 2007 specification says that the maximum number of page breaks
#NYI     # is 1026. However, in practice it is actually 1023.
#NYI     my $max_num_breaks = 1023;
#NYI     splice( @array, $max_num_breaks ) if @array > $max_num_breaks;
#NYI 
#NYI     return @array;
#NYI }


###############################################################################
#
# _check_dimensions($row, $col, $ignore_row, $ignore_col)
#
# Check that $row and $col are valid and store max and min values for use in
# other methods/elements.
#
# The $ignore_row/$ignore_col flags is used to indicate that we wish to
# perform the dimension check without storing the value.
#
# The ignore flags are use by set_row() and data_validate.
#
#NYI sub _check_dimensions {
#NYI 
#NYI     my $self       = shift;
#NYI     my $row        = $_[0];
#NYI     my $col        = $_[1];
#NYI     my $ignore_row = $_[2];
#NYI     my $ignore_col = $_[3];
#NYI 
#NYI 
#NYI     return -2 if not defined $row;
#NYI     return -2 if $row >= $self->{_xls_rowmax};
#NYI 
#NYI     return -2 if not defined $col;
#NYI     return -2 if $col >= $self->{_xls_colmax};
#NYI 
#NYI     # In optimization mode we don't change dimensions for rows that are
#NYI     # already written.
#NYI     if ( !$ignore_row && !$ignore_col && $self->{_optimization} == 1 ) {
#NYI         return -2 if $row < $self->{_previous_row};
#NYI     }
#NYI 
#NYI     if ( !$ignore_row ) {
#NYI 
#NYI         if ( not defined $self->{_dim_rowmin} or $row < $self->{_dim_rowmin} ) {
#NYI             $self->{_dim_rowmin} = $row;
#NYI         }
#NYI 
#NYI         if ( not defined $self->{_dim_rowmax} or $row > $self->{_dim_rowmax} ) {
#NYI             $self->{_dim_rowmax} = $row;
#NYI         }
#NYI     }
#NYI 
#NYI     if ( !$ignore_col ) {
#NYI 
#NYI         if ( not defined $self->{_dim_colmin} or $col < $self->{_dim_colmin} ) {
#NYI             $self->{_dim_colmin} = $col;
#NYI         }
#NYI 
#NYI         if ( not defined $self->{_dim_colmax} or $col > $self->{_dim_colmax} ) {
#NYI             $self->{_dim_colmax} = $col;
#NYI         }
#NYI     }
#NYI 
#NYI     return 0;
#NYI }


###############################################################################
#
#  _position_object_pixels()
#
# Calculate the vertices that define the position of a graphical object within
# the worksheet in pixels.
#
#         +------------+------------+
#         |     A      |      B     |
#   +-----+------------+------------+
#   |     |(x1,y1)     |            |
#   |  1  |(A1)._______|______      |
#   |     |    |              |     |
#   |     |    |              |     |
#   +-----+----|    Object    |-----+
#   |     |    |              |     |
#   |  2  |    |______________.     |
#   |     |            |        (B2)|
#   |     |            |     (x2,y2)|
#   +---- +------------+------------+
#
# Example of an object that covers some of the area from cell A1 to cell B2.
#
# Based on the width and height of the object we need to calculate 8 vars:
#
#     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
#
# We also calculate the absolute x and y position of the top left vertex of
# the object. This is required for images.
#
#    $x_abs, $y_abs
#
# The width and height of the cells that the object occupies can be variable
# and have to be taken into account.
#
# The values of $col_start and $row_start are passed in from the calling
# function. The values of $col_end and $row_end are calculated by subtracting
# the width and height of the object from the width and height of the
# underlying cells.
#
#NYI sub _position_object_pixels {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $col_start;    # Col containing upper left corner of object.
#NYI     my $x1;           # Distance to left side of object.
#NYI 
#NYI     my $row_start;    # Row containing top left corner of object.
#NYI     my $y1;           # Distance to top of object.
#NYI 
#NYI     my $col_end;      # Col containing lower right corner of object.
#NYI     my $x2;           # Distance to right side of object.
#NYI 
#NYI     my $row_end;      # Row containing bottom right corner of object.
#NYI     my $y2;           # Distance to bottom of object.
#NYI 
#NYI     my $width;        # Width of object frame.
#NYI     my $height;       # Height of object frame.
#NYI 
#NYI     my $x_abs = 0;    # Absolute distance to left side of object.
#NYI     my $y_abs = 0;    # Absolute distance to top  side of object.
#NYI 
#NYI     ( $col_start, $row_start, $x1, $y1, $width, $height ) = @_;
#NYI 
#NYI     # Adjust start column for negative offsets.
#NYI     while ( $x1 < 0 && $col_start > 0) {
#NYI         $x1 += $self->_size_col( $col_start  - 1);
#NYI         $col_start--;
#NYI     }
#NYI 
#NYI     # Adjust start row for negative offsets.
#NYI     while ( $y1 < 0 && $row_start > 0) {
#NYI         $y1 += $self->_size_row( $row_start - 1);
#NYI         $row_start--;
#NYI     }
#NYI 
#NYI     # Ensure that the image isn't shifted off the page at top left.
#NYI     $x1 = 0 if $x1 < 0;
#NYI     $y1 = 0 if $y1 < 0;
#NYI 
#NYI     # Calculate the absolute x offset of the top-left vertex.
#NYI     if ( $self->{_col_size_changed} ) {
#NYI         for my $col_id ( 0 .. $col_start -1 ) {
#NYI             $x_abs += $self->_size_col( $col_id );
#NYI         }
#NYI     }
#NYI     else {
#NYI         # Optimisation for when the column widths haven't changed.
#NYI         $x_abs += $self->{_default_col_pixels} * $col_start;
#NYI     }
#NYI 
#NYI     $x_abs += $x1;
#NYI 
#NYI     # Calculate the absolute y offset of the top-left vertex.
#NYI     # Store the column change to allow optimisations.
#NYI     if ( $self->{_row_size_changed} ) {
#NYI         for my $row_id ( 0 .. $row_start -1 ) {
#NYI             $y_abs += $self->_size_row( $row_id );
#NYI         }
#NYI     }
#NYI     else {
#NYI         # Optimisation for when the row heights haven't changed.
#NYI         $y_abs += $self->{_default_row_pixels} * $row_start;
#NYI     }
#NYI 
#NYI     $y_abs += $y1;
#NYI 
#NYI 
#NYI     # Adjust start column for offsets that are greater than the col width.
#NYI     while ( $x1 >= $self->_size_col( $col_start ) ) {
#NYI         $x1 -= $self->_size_col( $col_start );
#NYI         $col_start++;
#NYI     }
#NYI 
#NYI     # Adjust start row for offsets that are greater than the row height.
#NYI     while ( $y1 >= $self->_size_row( $row_start ) ) {
#NYI         $y1 -= $self->_size_row( $row_start );
#NYI         $row_start++;
#NYI     }
#NYI 
#NYI     # Initialise end cell to the same as the start cell.
#NYI     $col_end = $col_start;
#NYI     $row_end = $row_start;
#NYI 
#NYI     $width  = $width + $x1;
#NYI     $height = $height + $y1;
#NYI 
#NYI 
#NYI     # Subtract the underlying cell widths to find the end cell of the object.
#NYI     while ( $width >= $self->_size_col( $col_end ) ) {
#NYI         $width -= $self->_size_col( $col_end );
#NYI         $col_end++;
#NYI     }
#NYI 
#NYI 
#NYI     # Subtract the underlying cell heights to find the end cell of the object.
#NYI     while ( $height >= $self->_size_row( $row_end ) ) {
#NYI         $height -= $self->_size_row( $row_end );
#NYI         $row_end++;
#NYI     }
#NYI 
#NYI     # The end vertices are whatever is left from the width and height.
#NYI     $x2 = $width;
#NYI     $y2 = $height;
#NYI 
#NYI     return (
#NYI         $col_start, $row_start, $x1, $y1,
#NYI         $col_end,   $row_end,   $x2, $y2,
#NYI         $x_abs,     $y_abs
#NYI 
#NYI     );
#NYI }


###############################################################################
#
#  _position_object_emus()
#
# Calculate the vertices that define the position of a graphical object within
# the worksheet in EMUs.
#
# The vertices are expressed as English Metric Units (EMUs). There are 12,700
# EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
#
#NYI sub _position_object_emus {
#NYI 
#NYI     my $self       = shift;
#NYI 
#NYI     my (
#NYI         $col_start, $row_start, $x1, $y1,
#NYI         $col_end,   $row_end,   $x2, $y2,
#NYI         $x_abs,     $y_abs
#NYI 
#NYI     ) = $self->_position_object_pixels( @_ );
#NYI 
#NYI     # Convert the pixel values to EMUs. See above.
#NYI     $x1    = int( 0.5 + 9_525 * $x1 );
#NYI     $y1    = int( 0.5 + 9_525 * $y1 );
#NYI     $x2    = int( 0.5 + 9_525 * $x2 );
#NYI     $y2    = int( 0.5 + 9_525 * $y2 );
#NYI     $x_abs = int( 0.5 + 9_525 * $x_abs );
#NYI     $y_abs = int( 0.5 + 9_525 * $y_abs );
#NYI 
#NYI     return (
#NYI         $col_start, $row_start, $x1, $y1,
#NYI         $col_end,   $row_end,   $x2, $y2,
#NYI         $x_abs,     $y_abs
#NYI 
#NYI     );
#NYI }


###############################################################################
#
#  _position_shape_emus()
#
# Calculate the vertices that define the position of a shape object within
# the worksheet in EMUs.  Save the vertices with the object.
#
# The vertices are expressed as English Metric Units (EMUs). There are 12,700
# EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
#
#NYI sub _position_shape_emus {
#NYI 
#NYI     my $self  = shift;
#NYI     my $shape = shift;
#NYI 
#NYI     my (
#NYI         $col_start, $row_start, $x1, $y1,    $col_end,
#NYI         $row_end,   $x2,        $y2, $x_abs, $y_abs
#NYI       )
#NYI       = $self->_position_object_pixels(
#NYI         $shape->{_column_start},
#NYI         $shape->{_row_start},
#NYI         $shape->{_x_offset},
#NYI         $shape->{_y_offset},
#NYI         $shape->{_width} * $shape->{_scale_x},
#NYI         $shape->{_height} * $shape->{_scale_y},
#NYI         $shape->{_drawing}
#NYI       );
#NYI 
#NYI     # Now that x2/y2 have been calculated with a potentially negative
#NYI     # width/height we use the absolute value and convert to EMUs.
#NYI     $shape->{_width_emu}  = int( abs( $shape->{_width} * 9_525 ) );
#NYI     $shape->{_height_emu} = int( abs( $shape->{_height} * 9_525 ) );
#NYI 
#NYI     $shape->{_column_start} = int( $col_start );
#NYI     $shape->{_row_start}    = int( $row_start );
#NYI     $shape->{_column_end}   = int( $col_end );
#NYI     $shape->{_row_end}      = int( $row_end );
#NYI 
#NYI     # Convert the pixel values to EMUs. See above.
#NYI     $shape->{_x1}    = int( $x1 * 9_525 );
#NYI     $shape->{_y1}    = int( $y1 * 9_525 );
#NYI     $shape->{_x2}    = int( $x2 * 9_525 );
#NYI     $shape->{_y2}    = int( $y2 * 9_525 );
#NYI     $shape->{_x_abs} = int( $x_abs * 9_525 );
#NYI     $shape->{_y_abs} = int( $y_abs * 9_525 );
#NYI }

###############################################################################
#
# _size_col($col)
#
# Convert the width of a cell from user's units to pixels. Excel rounds the
# column width to the nearest pixel. If the width hasn't been set by the user
# we use the default value. If the column is hidden it has a value of zero.
#
#NYI sub _size_col {
#NYI 
#NYI     my $self = shift;
#NYI     my $col  = shift;
#NYI 
#NYI     my $max_digit_width = 7;    # For Calabri 11.
#NYI     my $padding         = 5;
#NYI     my $pixels;
#NYI 
#NYI     # Look up the cell value to see if it has been changed.
#NYI     if ( exists $self->{_col_sizes}->{$col}
#NYI         and defined $self->{_col_sizes}->{$col} )
#NYI     {
#NYI         my $width = $self->{_col_sizes}->{$col};
#NYI 
#NYI         # Convert to pixels.
#NYI         if ( $width == 0 ) {
#NYI             $pixels = 0;
#NYI         }
#NYI         elsif ( $width < 1 ) {
#NYI             $pixels = int( $width * ( $max_digit_width + $padding ) + 0.5 );
#NYI         }
#NYI         else {
#NYI             $pixels = int( $width * $max_digit_width + 0.5 ) + $padding;
#NYI         }
#NYI     }
#NYI     else {
#NYI         $pixels = $self->{_default_col_pixels};
#NYI     }
#NYI 
#NYI     return $pixels;
#NYI }


###############################################################################
#
# _size_row($row)
#
# Convert the height of a cell from user's units to pixels. If the height
# hasn't been set by the user we use the default value. If the row is hidden
# it has a value of zero.
#
#NYI sub _size_row {
#NYI 
#NYI     my $self = shift;
#NYI     my $row  = shift;
#NYI     my $pixels;
#NYI 
#NYI     # Look up the cell value to see if it has been changed
#NYI     if ( exists $self->{_row_sizes}->{$row} ) {
#NYI         my $height = $self->{_row_sizes}->{$row};
#NYI 
#NYI         if ( $height == 0 ) {
#NYI             $pixels = 0;
#NYI         }
#NYI         else {
#NYI             $pixels = int( 4 / 3 * $height );
#NYI         }
#NYI     }
#NYI     else {
#NYI         $pixels = int( 4 / 3 * $self->{_default_row_height} );
#NYI     }
#NYI 
#NYI     return $pixels;
#NYI }


###############################################################################
#
# _get_shared_string_index()
#
# Add a string to the shared string table, if it isn't already there, and
# return the string index.
#
#NYI sub _get_shared_string_index {
#NYI 
#NYI     my $self = shift;
#NYI     my $str  = shift;
#NYI 
#NYI     # Add the string to the shared string table.
#NYI     if ( not exists ${ $self->{_str_table} }->{$str} ) {
#NYI         ${ $self->{_str_table} }->{$str} = ${ $self->{_str_unique} }++;
#NYI     }
#NYI 
#NYI     ${ $self->{_str_total} }++;
#NYI     my $index = ${ $self->{_str_table} }->{$str};
#NYI 
#NYI     return $index;
#NYI }


###############################################################################
#
# insert_chart( $row, $col, $chart, $x, $y, $x_scale, $y_scale )
#
# Insert a chart into a worksheet. The $chart argument should be a Chart
# object or else it is assumed to be a filename of an external binary file.
# The latter is for backwards compatibility.
#
#NYI sub insert_chart {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column.
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     my $row      = $_[0];
#NYI     my $col      = $_[1];
#NYI     my $chart    = $_[2];
#NYI     my $x_offset = $_[3] || 0;
#NYI     my $y_offset = $_[4] || 0;
#NYI     my $x_scale  = $_[5] || 1;
#NYI     my $y_scale  = $_[6] || 1;
#NYI 
#NYI     fail "Insufficient arguments in insert_chart()" unless @_ >= 3;
#NYI 
#NYI     if ( ref $chart ) {
#NYI 
#NYI         # Check for a Chart object.
#NYI         fail "Not a Chart object in insert_chart()"
#NYI           unless $chart->isa( 'Excel::Writer::XLSX::Chart' );
#NYI 
#NYI         # Check that the chart is an embedded style chart.
#NYI         fail "Not a embedded style Chart object in insert_chart()"
#NYI           unless $chart->{_embedded};
#NYI 
#NYI     }
#NYI 
#NYI     # Ensure a chart isn't inserted more than once.
#NYI     if (   $chart->{_already_inserted}
#NYI         || $chart->{_combined} && $chart->{_combined}->{_already_inserted} )
#NYI     {
#NYI         warn "Chart cannot be inserted in a worksheet more than once";
#NYI         return;
#NYI     }
#NYI     else {
#NYI         $chart->{_already_inserted} = 1;
#NYI 
#NYI         if ( $chart->{_combined} ) {
#NYI             $chart->{_combined}->{_already_inserted} = 1;
#NYI         }
#NYI     }
#NYI 
#NYI     # Use the values set with $chart->set_size(), if any.
#NYI     $x_scale  = $chart->{_x_scale}  if $chart->{_x_scale} != 1;
#NYI     $y_scale  = $chart->{_y_scale}  if $chart->{_y_scale} != 1;
#NYI     $x_offset = $chart->{_x_offset} if $chart->{_x_offset};
#NYI     $y_offset = $chart->{_y_offset} if $chart->{_y_offset};
#NYI 
#NYI     push @{ $self->{_charts} },
#NYI       [ $row, $col, $chart, $x_offset, $y_offset, $x_scale, $y_scale ];
#NYI }


###############################################################################
#
# _prepare_chart()
#
# Set up chart/drawings.
#
#NYI sub _prepare_chart {
#NYI 
#NYI     my $self         = shift;
#NYI     my $index        = shift;
#NYI     my $chart_id     = shift;
#NYI     my $drawing_id   = shift;
#NYI     my $drawing_type = 1;
#NYI 
#NYI     my ( $row, $col, $chart, $x_offset, $y_offset, $x_scale, $y_scale ) =
#NYI       @{ $self->{_charts}->[$index] };
#NYI 
#NYI     $chart->{_id} = $chart_id - 1;
#NYI 
#NYI     # Use user specified dimensions, if any.
#NYI     my $width  = $chart->{_width}  if $chart->{_width};
#NYI     my $height = $chart->{_height} if $chart->{_height};
#NYI 
#NYI     $width  = int( 0.5 + ( $width  * $x_scale ) );
#NYI     $height = int( 0.5 + ( $height * $y_scale ) );
#NYI 
#NYI     my @dimensions =
#NYI       $self->_position_object_emus( $col, $row, $x_offset, $y_offset, $width,
#NYI         $height);
#NYI 
#NYI     # Set the chart name for the embedded object if it has been specified.
#NYI     my $name = $chart->{_chart_name};
#NYI 
#NYI     # Create a Drawing object to use with worksheet unless one already exists.
#NYI     if ( !$self->{_drawing} ) {
#NYI 
#NYI         my $drawing = Excel::Writer::XLSX::Drawing->new();
#NYI         $drawing->_add_drawing_object( $drawing_type, @dimensions, 0, 0,
#NYI             $name );
#NYI         $drawing->{_embedded} = 1;
#NYI 
#NYI         $self->{_drawing} = $drawing;
#NYI 
#NYI         push @{ $self->{_external_drawing_links} },
#NYI           [ '/drawing', '../drawings/drawing' . $drawing_id . '.xml' ];
#NYI     }
#NYI     else {
#NYI         my $drawing = $self->{_drawing};
#NYI         $drawing->_add_drawing_object( $drawing_type, @dimensions, 0, 0,
#NYI             $name );
#NYI 
#NYI     }
#NYI 
#NYI     push @{ $self->{_drawing_links} },
#NYI       [ '/chart', '../charts/chart' . $chart_id . '.xml' ];
#NYI }


###############################################################################
#
# _get_range_data
#
# Returns a range of data from the worksheet _table to be used in chart
# cached data. Strings are returned as SST ids and decoded in the workbook.
# Return undefs for data that doesn't exist since Excel can chart series
# with data missing.
#
#NYI sub _get_range_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return () if $self->{_optimization};
#NYI 
#NYI     my @data;
#NYI     my ( $row_start, $col_start, $row_end, $col_end ) = @_;
#NYI 
#NYI     # TODO. Check for worksheet limits.
#NYI 
#NYI     # Iterate through the table data.
#NYI     for my $row_num ( $row_start .. $row_end ) {
#NYI 
#NYI         # Store undef if row doesn't exist.
#NYI         if ( !exists $self->{_table}->{$row_num} ) {
#NYI             push @data, undef;
#NYI             next;
#NYI         }
#NYI 
#NYI         for my $col_num ( $col_start .. $col_end ) {
#NYI 
#NYI             if ( my $cell = $self->{_table}->{$row_num}->{$col_num} ) {
#NYI 
#NYI                 my $type  = $cell->[0];
#NYI                 my $token = $cell->[1];
#NYI 
#NYI 
#NYI                 if ( $type eq 'n' ) {
#NYI 
#NYI                     # Store a number.
#NYI                     push @data, $token;
#NYI                 }
#NYI                 elsif ( $type eq 's' ) {
#NYI 
#NYI                     # Store a string.
#NYI                     if ( $self->{_optimization} == 0 ) {
#NYI                         push @data, { 'sst_id' => $token };
#NYI                     }
#NYI                     else {
#NYI                         push @data, $token;
#NYI                     }
#NYI                 }
#NYI                 elsif ( $type eq 'f' ) {
#NYI 
#NYI                     # Store a formula.
#NYI                     push @data, $cell->[3] || 0;
#NYI                 }
#NYI                 elsif ( $type eq 'a' ) {
#NYI 
#NYI                     # Store an array formula.
#NYI                     push @data, $cell->[4] || 0;
#NYI                 }
#NYI                 elsif ( $type eq 'b' ) {
#NYI 
#NYI                     # Store a empty cell.
#NYI                     push @data, '';
#NYI                 }
#NYI             }
#NYI             else {
#NYI 
#NYI                 # Store undef if col doesn't exist.
#NYI                 push @data, undef;
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     return @data;
#NYI }


###############################################################################
#
# insert_image( $row, $col, $filename, $x, $y, $x_scale, $y_scale )
#
# Insert an image into the worksheet.
#
#NYI sub insert_image {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column.
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     my $row      = $_[0];
#NYI     my $col      = $_[1];
#NYI     my $image    = $_[2];
#NYI     my $x_offset = $_[3] || 0;
#NYI     my $y_offset = $_[4] || 0;
#NYI     my $x_scale  = $_[5] || 1;
#NYI     my $y_scale  = $_[6] || 1;
#NYI 
#NYI     fail "Insufficient arguments in insert_image()" unless @_ >= 3;
#NYI     fail "Couldn't locate $image: $!" unless -e $image;
#NYI 
#NYI     push @{ $self->{_images} },
#NYI       [ $row, $col, $image, $x_offset, $y_offset, $x_scale, $y_scale ];
#NYI }


###############################################################################
#
# _prepare_image()
#
# Set up image/drawings.
#
#NYI sub _prepare_image {
#NYI 
#NYI     my $self         = shift;
#NYI     my $index        = shift;
#NYI     my $image_id     = shift;
#NYI     my $drawing_id   = shift;
#NYI     my $width        = shift;
#NYI     my $height       = shift;
#NYI     my $name         = shift;
#NYI     my $image_type   = shift;
#NYI     my $x_dpi        = shift;
#NYI     my $y_dpi        = shift;
#NYI     my $drawing_type = 2;
#NYI     my $drawing;
#NYI 
#NYI     my ( $row, $col, $image, $x_offset, $y_offset, $x_scale, $y_scale ) =
#NYI       @{ $self->{_images}->[$index] };
#NYI 
#NYI     $width  *= $x_scale;
#NYI     $height *= $y_scale;
#NYI 
#NYI     $width  *= 96 / $x_dpi;
#NYI     $height *= 96 / $y_dpi;
#NYI 
#NYI     my @dimensions =
#NYI       $self->_position_object_emus( $col, $row, $x_offset, $y_offset, $width,
#NYI         $height);
#NYI 
#NYI     # Convert from pixels to emus.
#NYI     $width  = int( 0.5 + ( $width * 9_525 ) );
#NYI     $height = int( 0.5 + ( $height * 9_525 ) );
#NYI 
#NYI     # Create a Drawing object to use with worksheet unless one already exists.
#NYI     if ( !$self->{_drawing} ) {
#NYI 
#NYI         $drawing = Excel::Writer::XLSX::Drawing->new();
#NYI         $drawing->{_embedded} = 1;
#NYI 
#NYI         $self->{_drawing} = $drawing;
#NYI 
#NYI         push @{ $self->{_external_drawing_links} },
#NYI           [ '/drawing', '../drawings/drawing' . $drawing_id . '.xml' ];
#NYI     }
#NYI     else {
#NYI         $drawing = $self->{_drawing};
#NYI     }
#NYI 
#NYI     $drawing->_add_drawing_object( $drawing_type, @dimensions, $width, $height,
#NYI         $name );
#NYI 
#NYI 
#NYI     push @{ $self->{_drawing_links} },
#NYI       [ '/image', '../media/image' . $image_id . '.' . $image_type ];
#NYI }


###############################################################################
#
# _prepare_header_image()
#
# Set up an image without a drawing object for header/footer images.
#
#NYI sub _prepare_header_image {
#NYI 
#NYI     my $self       = shift;
#NYI     my $image_id   = shift;
#NYI     my $width      = shift;
#NYI     my $height     = shift;
#NYI     my $name       = shift;
#NYI     my $image_type = shift;
#NYI     my $position   = shift;
#NYI     my $x_dpi      = shift;
#NYI     my $y_dpi      = shift;
#NYI 
#NYI     # Strip the extension from the filename.
#NYI     $name =~ s/\.[^\.]+$//;
#NYI 
#NYI     push @{ $self->{_header_images_array} },
#NYI       [ $width, $height, $name, $position, $x_dpi, $y_dpi ];
#NYI 
#NYI     push @{ $self->{_vml_drawing_links} },
#NYI       [ '/image', '../media/image' . $image_id . '.' . $image_type ];
#NYI }


###############################################################################
#
# insert_shape( $row, $col, $shape, $x, $y, $x_scale, $y_scale )
#
# Insert a shape into the worksheet.
#
#NYI sub insert_shape {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Check for a cell reference in A1 notation and substitute row and column.
#NYI     if ( $_[0] =~ /^\D/ ) {
#NYI         @_ = $self->_substitute_cellref( @_ );
#NYI     }
#NYI 
#NYI     # Check the number of arguments.
#NYI     fail "Insufficient arguments in insert_shape()" unless @_ >= 3;
#NYI 
#NYI     my $shape = $_[2];
#NYI 
#NYI     # Verify we are being asked to insert a "shape" object.
#NYI     fail "Not a Shape object in insert_shape()"
#NYI       unless $shape->isa( 'Excel::Writer::XLSX::Shape' );
#NYI 
#NYI     # Set the shape properties.
#NYI     $shape->{_row_start}    = $_[0];
#NYI     $shape->{_column_start} = $_[1];
#NYI     $shape->{_x_offset}     = $_[3] || 0;
#NYI     $shape->{_y_offset}     = $_[4] || 0;
#NYI 
#NYI     # Override shape scale if supplied as an argument.  Otherwise, use the
#NYI     # existing shape scale factors.
#NYI     $shape->{_scale_x} = $_[5] if defined $_[5];
#NYI     $shape->{_scale_y} = $_[6] if defined $_[6];
#NYI 
#NYI     # Assign a shape ID.
#NYI     my $needs_id = 1;
#NYI     while ( $needs_id ) {
#NYI         my $id = $shape->{_id} || 0;
#NYI         my $used = exists $self->{_shape_hash}->{$id} ? 1 : 0;
#NYI 
#NYI         # Test if shape ID is already used. Otherwise assign a new one.
#NYI         if ( !$used && $id != 0 ) {
#NYI             $needs_id = 0;
#NYI         }
#NYI         else {
#NYI             $shape->{_id} = ++$self->{_last_shape_id};
#NYI         }
#NYI     }
#NYI 
#NYI     $shape->{_element} = $#{ $self->{_shapes} } + 1;
#NYI 
#NYI     # Allow lookup of entry into shape array by shape ID.
#NYI     $self->{_shape_hash}->{ $shape->{_id} } = $shape->{_element};
#NYI 
#NYI     # Create link to Worksheet color palette.
#NYI     $shape->{_palette} = $self->{_palette};
#NYI 
#NYI     if ( $shape->{_stencil} ) {
#NYI 
#NYI         # Insert a copy of the shape, not a reference so that the shape is
#NYI         # used as a stencil. Previously stamped copies don't get modified
#NYI         # if the stencil is modified.
#NYI         my $insert = { %{$shape} };
#NYI 
#NYI        # For connectors change x/y coords based on location of connected shapes.
#NYI         $self->_auto_locate_connectors( $insert );
#NYI 
#NYI         # Bless the copy into this class, so AUTOLOADED _get, _set methods
#NYI         #still work on the child.
#NYI         bless $insert, ref $shape;
#NYI 
#NYI         push @{ $self->{_shapes} }, $insert;
#NYI         return $insert;
#NYI     }
#NYI     else {
#NYI 
#NYI        # For connectors change x/y coords based on location of connected shapes.
#NYI         $self->_auto_locate_connectors( $shape );
#NYI 
#NYI         # Insert a link to the shape on the list of shapes. Connection to
#NYI         # the parent shape is maintained
#NYI         push @{ $self->{_shapes} }, $shape;
#NYI         return $shape;
#NYI     }
#NYI }


###############################################################################
#
# _prepare_shape()
#
# Set up drawing shapes
#
#NYI sub _prepare_shape {
#NYI 
#NYI     my $self       = shift;
#NYI     my $index      = shift;
#NYI     my $drawing_id = shift;
#NYI     my $shape      = $self->{_shapes}->[$index];
#NYI     my $drawing;
#NYI     my $drawing_type = 3;
#NYI 
#NYI     # Create a Drawing object to use with worksheet unless one already exists.
#NYI     if ( !$self->{_drawing} ) {
#NYI 
#NYI         $drawing              = Excel::Writer::XLSX::Drawing->new();
#NYI         $drawing->{_embedded} = 1;
#NYI         $self->{_drawing}     = $drawing;
#NYI 
#NYI         push @{ $self->{_external_drawing_links} },
#NYI           [ '/drawing', '../drawings/drawing' . $drawing_id . '.xml' ];
#NYI 
#NYI         $self->{_has_shapes} = 1;
#NYI     }
#NYI     else {
#NYI         $drawing = $self->{_drawing};
#NYI     }
#NYI 
#NYI     # Validate the he shape against various rules.
#NYI     $self->_validate_shape( $shape, $index );
#NYI 
#NYI     $self->_position_shape_emus( $shape );
#NYI 
#NYI     my @dimensions = (
#NYI         $shape->{_column_start}, $shape->{_row_start},
#NYI         $shape->{_x1},           $shape->{_y1},
#NYI         $shape->{_column_end},   $shape->{_row_end},
#NYI         $shape->{_x2},           $shape->{_y2},
#NYI         $shape->{_x_abs},        $shape->{_y_abs},
#NYI         $shape->{_width_emu},    $shape->{_height_emu},
#NYI     );
#NYI 
#NYI     $drawing->_add_drawing_object( $drawing_type, @dimensions, $shape->{_name},
#NYI         $shape );
#NYI }


###############################################################################
#
# _auto_locate_connectors()
#
# Re-size connector shapes if they are connected to other shapes.
#
#NYI sub _auto_locate_connectors {
#NYI 
#NYI     my $self  = shift;
#NYI     my $shape = shift;
#NYI 
#NYI     # Valid connector shapes.
#NYI     my $connector_shapes = {
#NYI         straightConnector => 1,
#NYI         Connector         => 1,
#NYI         bentConnector     => 1,
#NYI         curvedConnector   => 1,
#NYI         line              => 1,
#NYI     };
#NYI 
#NYI     my $shape_base = $shape->{_type};
#NYI 
#NYI     # Remove the number of segments from end of type.
#NYI     chop $shape_base;
#NYI 
#NYI     $shape->{_connect} = $connector_shapes->{$shape_base} ? 1 : 0;
#NYI 
#NYI     return unless $shape->{_connect};
#NYI 
#NYI     # Both ends have to be connected to size it.
#NYI     return unless ( $shape->{_start} and $shape->{_end} );
#NYI 
#NYI     # Both ends need to provide info about where to connect.
#NYI     return unless ( $shape->{_start_side} and $shape->{_end_side} );
#NYI 
#NYI     my $sid = $shape->{_start};
#NYI     my $eid = $shape->{_end};
#NYI 
#NYI     my $slink_id = $self->{_shape_hash}->{$sid};
#NYI     my ( $sls, $els );
#NYI     if ( defined $slink_id ) {
#NYI         $sls = $self->{_shapes}->[$slink_id];    # Start linked shape.
#NYI     }
#NYI     else {
#NYI         warn "missing start connection for '$shape->{_name}', id=$sid\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     my $elink_id = $self->{_shape_hash}->{$eid};
#NYI     if ( defined $elink_id ) {
#NYI         $els = $self->{_shapes}->[$elink_id];    # Start linked shape.
#NYI     }
#NYI     else {
#NYI         warn "missing end connection for '$shape->{_name}', id=$eid\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     # Assume shape connections are to the middle of an object, and
#NYI     # not a corner (for now).
#NYI     my $connect_type = $shape->{_start_side} . $shape->{_end_side};
#NYI     my $smidx        = $sls->{_x_offset} + $sls->{_width} / 2;
#NYI     my $emidx        = $els->{_x_offset} + $els->{_width} / 2;
#NYI     my $smidy        = $sls->{_y_offset} + $sls->{_height} / 2;
#NYI     my $emidy        = $els->{_y_offset} + $els->{_height} / 2;
#NYI     my $netx         = abs( $smidx - $emidx );
#NYI     my $nety         = abs( $smidy - $emidy );
#NYI 
#NYI     if ( $connect_type eq 'bt' ) {
#NYI         my $sy = $sls->{_y_offset} + $sls->{_height};
#NYI         my $ey = $els->{_y_offset};
#NYI 
#NYI         $shape->{_width} = abs( int( $emidx - $smidx ) );
#NYI         $shape->{_x_offset} = int( min( $smidx, $emidx ) );
#NYI         $shape->{_height} =
#NYI           abs(
#NYI             int( $els->{_y_offset} - ( $sls->{_y_offset} + $sls->{_height} ) )
#NYI           );
#NYI         $shape->{_y_offset} = int(
#NYI             min( ( $sls->{_y_offset} + $sls->{_height} ), $els->{_y_offset} ) );
#NYI         $shape->{_flip_h} = ( $smidx < $emidx ) ? 1 : 0;
#NYI         $shape->{_rotation} = 90;
#NYI 
#NYI         if ( $sy > $ey ) {
#NYI             $shape->{_flip_v} = 1;
#NYI 
#NYI             # Create 3 adjustments for an end shape vertically above a
#NYI             # start shape. Adjustments count from the upper left object.
#NYI             if ( $#{ $shape->{_adjustments} } < 0 ) {
#NYI                 $shape->{_adjustments} = [ -10, 50, 110 ];
#NYI             }
#NYI 
#NYI             $shape->{_type} = 'bentConnector5';
#NYI         }
#NYI     }
#NYI     elsif ( $connect_type eq 'rl' ) {
#NYI         $shape->{_width} =
#NYI           abs(
#NYI             int( $els->{_x_offset} - ( $sls->{_x_offset} + $sls->{_width} ) ) );
#NYI         $shape->{_height} = abs( int( $emidy - $smidy ) );
#NYI         $shape->{_x_offset} =
#NYI           min( $sls->{_x_offset} + $sls->{_width}, $els->{_x_offset} );
#NYI         $shape->{_y_offset} = min( $smidy, $emidy );
#NYI 
#NYI         $shape->{_flip_h} = 1 if ( $smidx < $emidx ) and ( $smidy > $emidy );
#NYI         $shape->{_flip_h} = 1 if ( $smidx > $emidx ) and ( $smidy < $emidy );
#NYI         if ( $smidx > $emidx ) {
#NYI 
#NYI             # Create 3 adjustments if end shape is left of start
#NYI             if ( $#{ $shape->{_adjustments} } < 0 ) {
#NYI                 $shape->{_adjustments} = [ -10, 50, 110 ];
#NYI             }
#NYI 
#NYI             $shape->{_type} = 'bentConnector5';
#NYI         }
#NYI     }
#NYI     else {
#NYI         warn "Connection $connect_type not implemented yet\n";
#NYI     }
#NYI }


###############################################################################
#
# _validate_shape()
#
# Check shape attributes to ensure they are valid.
#
#NYI sub _validate_shape {
#NYI 
#NYI     my $self  = shift;
#NYI     my $shape = shift;
#NYI     my $index = shift;
#NYI 
#NYI     if ( !grep ( /^$shape->{_align}$/, qw[l ctr r just] ) ) {
#NYI         fail "Shape $index ($shape->{_type}) alignment ($shape->{align}), "
#NYI           . "not in ('l', 'ctr', 'r', 'just')\n";
#NYI     }
#NYI 
#NYI     if ( !grep ( /^$shape->{_valign}$/, qw[t ctr b] ) ) {
#NYI         fail "Shape $index ($shape->{_type}) vertical alignment "
#NYI           . "($shape->{valign}), not ('t', 'ctr', 'b')\n";
#NYI     }
#NYI }


###############################################################################
#
# _prepare_vml_objects()
#
# Turn the HoH that stores the comments into an array for easier handling
# and set the external links for comments and buttons.
#
#NYI sub _prepare_vml_objects {
#NYI 
#NYI     my $self           = shift;
#NYI     my $vml_data_id    = shift;
#NYI     my $vml_shape_id   = shift;
#NYI     my $vml_drawing_id = shift;
#NYI     my $comment_id     = shift;
#NYI     my @comments;
#NYI 
#NYI 
#NYI     # We sort the comments by row and column but that isn't strictly required.
#NYI     my @rows = sort { $a <=> $b } keys %{ $self->{_comments} };
#NYI 
#NYI     for my $row ( @rows ) {
#NYI         my @cols = sort { $a <=> $b } keys %{ $self->{_comments}->{$row} };
#NYI 
#NYI         for my $col ( @cols ) {
#NYI 
#NYI             # Set comment visibility if required and not already user defined.
#NYI             if ( $self->{_comments_visible} ) {
#NYI                 if ( !defined $self->{_comments}->{$row}->{$col}->[4] ) {
#NYI                     $self->{_comments}->{$row}->{$col}->[4] = 1;
#NYI                 }
#NYI             }
#NYI 
#NYI             # Set comment author if not already user defined.
#NYI             if ( !defined $self->{_comments}->{$row}->{$col}->[3] ) {
#NYI                 $self->{_comments}->{$row}->{$col}->[3] =
#NYI                   $self->{_comments_author};
#NYI             }
#NYI 
#NYI             push @comments, $self->{_comments}->{$row}->{$col};
#NYI         }
#NYI     }
#NYI 
#NYI     push @{ $self->{_external_vml_links} },
#NYI       [ '/vmlDrawing', '../drawings/vmlDrawing' . $vml_drawing_id . '.vml' ];
#NYI 
#NYI     if ( $self->{_has_comments} ) {
#NYI 
#NYI         $self->{_comments_array} = \@comments;
#NYI 
#NYI         push @{ $self->{_external_comment_links} },
#NYI           [ '/comments', '../comments' . $comment_id . '.xml' ];
#NYI     }
#NYI 
#NYI     my $count         = scalar @comments;
#NYI     my $start_data_id = $vml_data_id;
#NYI 
#NYI     # The VML o:idmap data id contains a comma separated range when there is
#NYI     # more than one 1024 block of comments, like this: data="1,2".
#NYI     for my $i ( 1 .. int( $count / 1024 ) ) {
#NYI         $vml_data_id = "$vml_data_id," . ( $start_data_id + $i );
#NYI     }
#NYI 
#NYI     $self->{_vml_data_id}  = $vml_data_id;
#NYI     $self->{_vml_shape_id} = $vml_shape_id;
#NYI 
#NYI     return $count;
#NYI }


###############################################################################
#
# _prepare_header_vml_objects()
#
# Set up external linkage for VML header/footer images.
#
#NYI sub _prepare_header_vml_objects {
#NYI 
#NYI     my $self           = shift;
#NYI     my $vml_header_id  = shift;
#NYI     my $vml_drawing_id = shift;
#NYI 
#NYI     $self->{_vml_header_id} = $vml_header_id;
#NYI 
#NYI     push @{ $self->{_external_vml_links} },
#NYI       [ '/vmlDrawing', '../drawings/vmlDrawing' . $vml_drawing_id . '.vml' ];
#NYI }


###############################################################################
#
# _prepare_tables()
#
# Set the table ids for the worksheet tables.
#
#NYI sub _prepare_tables {
#NYI 
#NYI     my $self     = shift;
#NYI     my $table_id = shift;
#NYI     my $seen     = shift;
#NYI 
#NYI 
#NYI     for my $table ( @{ $self->{_tables} } ) {
#NYI 
#NYI         $table-> {_id} = $table_id;
#NYI 
#NYI         # Set the table name unless defined by the user.
#NYI         if ( !defined $table->{_name} ) {
#NYI 
#NYI             # Set a default name.
#NYI             $table->{_name} = 'Table' . $table_id;
#NYI         }
#NYI 
#NYI         # Check for duplicate table names.
#NYI         my $name = lc $table->{_name};
#NYI 
#NYI         if ( exists $seen->{$name} ) {
#NYI             die "error: invalid duplicate table name '$table->{_name}' found";
#NYI         }
#NYI         else {
#NYI             $seen->{$name} = 1;
#NYI         }
#NYI 
#NYI         # Store the link used for the rels file.
#NYI         my $link = [ '/table', '../tables/table' . $table_id . '.xml' ];
#NYI 
#NYI         push @{ $self->{_external_table_links} }, $link;
#NYI         $table_id++;
#NYI     }
#NYI }


###############################################################################
#
# _comment_params()
#
# This method handles the additional optional parameters to write_comment() as
# well as calculating the comment object position and vertices.
#
#NYI sub _comment_params {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $row    = shift;
#NYI     my $col    = shift;
#NYI     my $string = shift;
#NYI 
#NYI     my $default_width  = 128;
#NYI     my $default_height = 74;
#NYI 
#NYI     my %params = (
#NYI         author     => undef,
#NYI         color      => 81,
#NYI         start_cell => undef,
#NYI         start_col  => undef,
#NYI         start_row  => undef,
#NYI         visible    => undef,
#NYI         width      => $default_width,
#NYI         height     => $default_height,
#NYI         x_offset   => undef,
#NYI         x_scale    => 1,
#NYI         y_offset   => undef,
#NYI         y_scale    => 1,
#NYI     );
#NYI 
#NYI 
#NYI     # Overwrite the defaults with any user supplied values. Incorrect or
#NYI     # misspelled parameters are silently ignored.
#NYI     %params = ( %params, @_ );
#NYI 
#NYI 
#NYI     # Ensure that a width and height have been set.
#NYI     $params{width}  = $default_width  if not $params{width};
#NYI     $params{height} = $default_height if not $params{height};
#NYI 
#NYI 
#NYI     # Limit the string to the max number of chars.
#NYI     my $max_len = 32767;
#NYI 
#NYI     if ( length( $string ) > $max_len ) {
#NYI         $string = substr( $string, 0, $max_len );
#NYI     }
#NYI 
#NYI 
#NYI     # Set the comment background colour.
#NYI     my $color    = $params{color};
#NYI     my $color_id = &Excel::Writer::XLSX::Format::_get_color( $color );
#NYI 
#NYI     if ( $color_id =~ m/^#[0-9A-F]{6}$/i ) {
#NYI         $params{color} = $color_id;
#NYI     }
#NYI     elsif ( $color_id == 0 ) {
#NYI         $params{color} = '#ffffe1';
#NYI     }
#NYI     else {
#NYI         my $palette = $self->{_palette};
#NYI 
#NYI         # Get the RGB color from the palette.
#NYI         my @rgb = @{ $palette->[ $color_id - 8 ] };
#NYI         my $rgb_color = sprintf "%02x%02x%02x", @rgb[0, 1, 2];
#NYI 
#NYI         # Minor modification to allow comparison testing. Change RGB colors
#NYI         # from long format, ffcc00 to short format fc0 used by VML.
#NYI         $rgb_color =~ s/^([0-9a-f])\1([0-9a-f])\2([0-9a-f])\3$/$1$2$3/;
#NYI 
#NYI         $params{color} = sprintf "#%s [%d]", $rgb_color, $color_id;
#NYI     }
#NYI 
#NYI 
#NYI     # Convert a cell reference to a row and column.
#NYI     if ( defined $params{start_cell} ) {
#NYI         my ( $row, $col ) = $self->_substitute_cellref( $params{start_cell} );
#NYI         $params{start_row} = $row;
#NYI         $params{start_col} = $col;
#NYI     }
#NYI 
#NYI 
#NYI     # Set the default start cell and offsets for the comment. These are
#NYI     # generally fixed in relation to the parent cell. However there are
#NYI     # some edge cases for cells at the, er, edges.
#NYI     #
#NYI     my $row_max = $self->{_xls_rowmax};
#NYI     my $col_max = $self->{_xls_colmax};
#NYI 
#NYI     if ( not defined $params{start_row} ) {
#NYI 
#NYI         if    ( $row == 0 )            { $params{start_row} = 0 }
#NYI         elsif ( $row == $row_max - 3 ) { $params{start_row} = $row_max - 7 }
#NYI         elsif ( $row == $row_max - 2 ) { $params{start_row} = $row_max - 6 }
#NYI         elsif ( $row == $row_max - 1 ) { $params{start_row} = $row_max - 5 }
#NYI         else                           { $params{start_row} = $row - 1 }
#NYI     }
#NYI 
#NYI     if ( not defined $params{y_offset} ) {
#NYI 
#NYI         if    ( $row == 0 )            { $params{y_offset} = 2 }
#NYI         elsif ( $row == $row_max - 3 ) { $params{y_offset} = 16 }
#NYI         elsif ( $row == $row_max - 2 ) { $params{y_offset} = 16 }
#NYI         elsif ( $row == $row_max - 1 ) { $params{y_offset} = 14 }
#NYI         else                           { $params{y_offset} = 10 }
#NYI     }
#NYI 
#NYI     if ( not defined $params{start_col} ) {
#NYI 
#NYI         if    ( $col == $col_max - 3 ) { $params{start_col} = $col_max - 6 }
#NYI         elsif ( $col == $col_max - 2 ) { $params{start_col} = $col_max - 5 }
#NYI         elsif ( $col == $col_max - 1 ) { $params{start_col} = $col_max - 4 }
#NYI         else                           { $params{start_col} = $col + 1 }
#NYI     }
#NYI 
#NYI     if ( not defined $params{x_offset} ) {
#NYI 
#NYI         if    ( $col == $col_max - 3 ) { $params{x_offset} = 49 }
#NYI         elsif ( $col == $col_max - 2 ) { $params{x_offset} = 49 }
#NYI         elsif ( $col == $col_max - 1 ) { $params{x_offset} = 49 }
#NYI         else                           { $params{x_offset} = 15 }
#NYI     }
#NYI 
#NYI 
#NYI     # Scale the size of the comment box if required.
#NYI     if ( $params{x_scale} ) {
#NYI         $params{width} = $params{width} * $params{x_scale};
#NYI     }
#NYI 
#NYI     if ( $params{y_scale} ) {
#NYI         $params{height} = $params{height} * $params{y_scale};
#NYI     }
#NYI 
#NYI     # Round the dimensions to the nearest pixel.
#NYI     $params{width}  = int( 0.5 + $params{width} );
#NYI     $params{height} = int( 0.5 + $params{height} );
#NYI 
#NYI     # Calculate the positions of comment object.
#NYI     my @vertices = $self->_position_object_pixels(
#NYI         $params{start_col}, $params{start_row}, $params{x_offset},
#NYI         $params{y_offset},  $params{width},     $params{height}
#NYI     );
#NYI 
#NYI     # Add the width and height for VML.
#NYI     push @vertices, ( $params{width}, $params{height} );
#NYI 
#NYI     return (
#NYI         $row,
#NYI         $col,
#NYI         $string,
#NYI 
#NYI         $params{author},
#NYI         $params{visible},
#NYI         $params{color},
#NYI 
#NYI         [@vertices]
#NYI     );
#NYI }


###############################################################################
#
# _button_params()
#
# This method handles the parameters passed to insert_button() as well as
# calculating the comment object position and vertices.
#
#NYI sub _button_params {
#NYI 
#NYI     my $self   = shift;
#NYI     my $row    = shift;
#NYI     my $col    = shift;
#NYI     my $params = shift;
#NYI     my $button = { _row => $row, _col => $col };
#NYI 
#NYI     my $button_number = 1 + @{ $self->{_buttons_array} };
#NYI 
#NYI     # Set the button caption.
#NYI     my $caption = $params->{caption};
#NYI 
#NYI     # Set a default caption if none was specified by user.
#NYI     if ( !defined $caption ) {
#NYI         $caption = 'Button ' . $button_number;
#NYI     }
#NYI 
#NYI     $button->{_font}->{_caption} = $caption;
#NYI 
#NYI 
#NYI     # Set the macro name.
#NYI     if ( $params->{macro} ) {
#NYI         $button->{_macro} = '[0]!' . $params->{macro};
#NYI     }
#NYI     else {
#NYI         $button->{_macro} = '[0]!Button' . $button_number . '_Click';
#NYI     }
#NYI 
#NYI 
#NYI     # Ensure that a width and height have been set.
#NYI     my $default_width  = $self->{_default_col_pixels};
#NYI     my $default_height = $self->{_default_row_pixels};
#NYI     $params->{width}  = $default_width  if !$params->{width};
#NYI     $params->{height} = $default_height if !$params->{height};
#NYI 
#NYI     # Set the x/y offsets.
#NYI     $params->{x_offset}  = 0  if !$params->{x_offset};
#NYI     $params->{y_offset}  = 0  if !$params->{y_offset};
#NYI 
#NYI     # Scale the size of the comment box if required.
#NYI     if ( $params->{x_scale} ) {
#NYI         $params->{width} = $params->{width} * $params->{x_scale};
#NYI     }
#NYI 
#NYI     if ( $params->{y_scale} ) {
#NYI         $params->{height} = $params->{height} * $params->{y_scale};
#NYI     }
#NYI 
#NYI     # Round the dimensions to the nearest pixel.
#NYI     $params->{width}  = int( 0.5 + $params->{width} );
#NYI     $params->{height} = int( 0.5 + $params->{height} );
#NYI 
#NYI     $params->{start_row} = $row;
#NYI     $params->{start_col} = $col;
#NYI 
#NYI     # Calculate the positions of comment object.
#NYI     my @vertices = $self->_position_object_pixels(
#NYI         $params->{start_col}, $params->{start_row}, $params->{x_offset},
#NYI         $params->{y_offset},  $params->{width},     $params->{height}
#NYI     );
#NYI 
#NYI     # Add the width and height for VML.
#NYI     push @vertices, ( $params->{width}, $params->{height} );
#NYI 
#NYI     $button->{_vertices} = \@vertices;
#NYI 
#NYI     return $button;
#NYI }


###############################################################################
#
# Deprecated methods for backwards compatibility.
#
###############################################################################


#NYI # This method was mainly only required for Excel 5.
#NYI sub write_url_range { }
#NYI 
#NYI # Deprecated UTF-16 method required for the Excel 5 format.
#NYI sub write_utf16be_string {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Convert A1 notation if present.
#NYI     @_ = $self->_substitute_cellref( @_ ) if $_[0] =~ /^\D/;
#NYI 
#NYI     # Check the number of args.
#NYI     return -1 if @_ < 3;
#NYI 
#NYI     # Convert UTF16 string to UTF8.
#NYI     require Encode;
#NYI     my $utf8_string = Encode::decode( 'UTF-16BE', $_[2] );
#NYI 
#NYI     return $self->write_string( $_[0], $_[1], $utf8_string, $_[3] );
#NYI }
#NYI 
#NYI # Deprecated UTF-16 method required for the Excel 5 format.
#NYI sub write_utf16le_string {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Convert A1 notation if present.
#NYI     @_ = $self->_substitute_cellref( @_ ) if $_[0] =~ /^\D/;
#NYI 
#NYI     # Check the number of args.
#NYI     return -1 if @_ < 3;
#NYI 
#NYI     # Convert UTF16 string to UTF8.
#NYI     require Encode;
#NYI     my $utf8_string = Encode::decode( 'UTF-16LE', $_[2] );
#NYI 
#NYI     return $self->write_string( $_[0], $_[1], $utf8_string, $_[3] );
#NYI }
#NYI 
#NYI # No longer required. Was used to avoid slow formula parsing.
#NYI sub store_formula {
#NYI 
#NYI     my $self   = shift;
#NYI     my $string = shift;
#NYI 
#NYI     my @tokens = split /(\$?[A-I]?[A-Z]\$?\d+)/, $string;
#NYI 
#NYI     return \@tokens;
#NYI }
#NYI 
#NYI # No longer required. Was used to avoid slow formula parsing.
#NYI sub repeat_formula {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Convert A1 notation if present.
#NYI     @_ = $self->_substitute_cellref( @_ ) if $_[0] =~ /^\D/;
#NYI 
#NYI     if ( @_ < 2 ) { return -1 }    # Check the number of args
#NYI 
#NYI     my $row         = shift;       # Zero indexed row
#NYI     my $col         = shift;       # Zero indexed column
#NYI     my $formula_ref = shift;       # Array ref with formula tokens
#NYI     my $format      = shift;       # XF format
#NYI     my @pairs       = @_;          # Pattern/replacement pairs
#NYI 
#NYI 
#NYI     # Enforce an even number of arguments in the pattern/replacement list.
#NYI     fail "Odd number of elements in pattern/replacement list" if @pairs % 2;
#NYI 
#NYI     # Check that $formula is an array ref.
#NYI     fail "Not a valid formula" if ref $formula_ref ne 'ARRAY';
#NYI 
#NYI     my @tokens = @$formula_ref;
#NYI 
#NYI     # Allow the user to specify the result of the formula by appending a
#NYI     # result => $value pair to the end of the arguments.
#NYI     my $value = undef;
#NYI     if ( @pairs && $pairs[-2] eq 'result' ) {
#NYI         $value = pop @pairs;
#NYI         pop @pairs;
#NYI     }
#NYI 
#NYI     # Make the substitutions.
#NYI     while ( @pairs ) {
#NYI         my $pattern = shift @pairs;
#NYI         my $replace = shift @pairs;
#NYI 
#NYI         foreach my $token ( @tokens ) {
#NYI             last if $token =~ s/$pattern/$replace/;
#NYI         }
#NYI     }
#NYI 
#NYI     my $formula = join '', @tokens;
#NYI 
#NYI     return $self->write_formula( $row, $col, $formula, $format, $value );
#NYI }


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# write_worksheet()
#
# Write the <worksheet> element. This is the root element of Worksheet.
#
method write_worksheet {
    my $schema                 = 'http://schemas.openxmlformats.org/';
    my $xmlns                  = $schema ~ 'spreadsheetml/2006/main';
    my $xmlns_r                = $schema ~ 'officeDocument/2006/relationships';
    my $xmlns_mc               = $schema ~ 'markup-compatibility/2006';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns_r,
    );

    if $!excel_version == 2010 {
        @attributes.push: 'xmlns:mc' => $xmlns_mc;

        @attributes.push:
               'xmlns:x14ac' => 'http://schemas.microsoft.com/'
             ~ 'office/spreadsheetml/2009/9/ac';

        @attributes.push: 'mc:Ignorable' => 'x14ac';
    }

    self.xml_start_tag( 'worksheet', @attributes );
}


###############################################################################
#
# _write_sheet_pr()
#
# Write the <sheetPr> element for Sheet level properties.
#
#NYI sub _write_sheet_pr {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     if (   !$self->{_fit_page}
#NYI         && !$self->{_filter_on}
#NYI         && !$self->{_tab_color}
#NYI         && !$self->{_outline_changed}
#NYI         && !$self->{_vba_codename} )
#NYI     {
#NYI         return;
#NYI     }
#NYI 
#NYI 
#NYI     my $codename = $self->{_vba_codename};
#NYI     push @attributes, ( 'codeName'   => $codename ) if $codename;
#NYI     push @attributes, ( 'filterMode' => 1 )         if $self->{_filter_on};
#NYI 
#NYI     if (   $self->{_fit_page}
#NYI         || $self->{_tab_color}
#NYI         || $self->{_outline_changed} )
#NYI     {
#NYI         $self->xml_start_tag( 'sheetPr', @attributes );
#NYI         $self->_write_tab_color();
#NYI         $self->_write_outline_pr();
#NYI         $self->_write_page_set_up_pr();
#NYI         $self->xml_end_tag( 'sheetPr' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'sheetPr', @attributes );
#NYI     }
#NYI }


##############################################################################
#
# _write_page_set_up_pr()
#
# Write the <pageSetUpPr> element.
#
#NYI sub _write_page_set_up_pr {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return unless $self->{_fit_page};
#NYI 
#NYI     my @attributes = ( 'fitToPage' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'pageSetUpPr', @attributes );
#NYI }


###############################################################################
#
# _write_dimension()
#
# Write the <dimension> element. This specifies the range of cells in the
# worksheet. As a special case, empty spreadsheets use 'A1' as a range.
#
#NYI sub _write_dimension {
#NYI 
#NYI     my $self = shift;
#NYI     my $ref;
#NYI 
#NYI     if ( !defined $self->{_dim_rowmin} && !defined $self->{_dim_colmin} ) {
#NYI 
#NYI         # If the min dims are undefined then no dimensions have been set
#NYI         # and we use the default 'A1'.
#NYI         $ref = 'A1';
#NYI     }
#NYI     elsif ( !defined $self->{_dim_rowmin} && defined $self->{_dim_colmin} ) {
#NYI 
#NYI         # If the row dims aren't set but the column dims are then they
#NYI         # have been changed via set_column().
#NYI 
#NYI         if ( $self->{_dim_colmin} == $self->{_dim_colmax} ) {
#NYI 
#NYI             # The dimensions are a single cell and not a range.
#NYI             $ref = xl-rowcol-to-cell( 0, $self->{_dim_colmin} );
#NYI         }
#NYI         else {
#NYI 
#NYI             # The dimensions are a cell range.
#NYI             my $cell_1 = xl-rowcol-to-cell( 0, $self->{_dim_colmin} );
#NYI             my $cell_2 = xl-rowcol-to-cell( 0, $self->{_dim_colmax} );
#NYI 
#NYI             $ref = $cell_1 . ':' . $cell_2;
#NYI         }
#NYI 
#NYI     }
#NYI     elsif ($self->{_dim_rowmin} == $self->{_dim_rowmax}
#NYI         && $self->{_dim_colmin} == $self->{_dim_colmax} )
#NYI     {
#NYI 
#NYI         # The dimensions are a single cell and not a range.
#NYI         $ref = xl-rowcol-to-cell( $self->{_dim_rowmin}, $self->{_dim_colmin} );
#NYI     }
#NYI     else {
#NYI 
#NYI         # The dimensions are a cell range.
#NYI         my $cell_1 =
#NYI           xl-rowcol-to-cell( $self->{_dim_rowmin}, $self->{_dim_colmin} );
#NYI         my $cell_2 =
#NYI           xl-rowcol-to-cell( $self->{_dim_rowmax}, $self->{_dim_colmax} );
#NYI 
#NYI         $ref = $cell_1 . ':' . $cell_2;
#NYI     }
#NYI 
#NYI 
#NYI     my @attributes = ( 'ref' => $ref );
#NYI 
#NYI     $self->xml_empty_tag( 'dimension', @attributes );
#NYI }


###############################################################################
#
# _write_sheet_views()
#
# Write the <sheetViews> element.
#
#NYI sub _write_sheet_views {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ();
#NYI 
#NYI     $self->xml_start_tag( 'sheetViews', @attributes );
#NYI     $self->_write_sheet_view();
#NYI     $self->xml_end_tag( 'sheetViews' );
#NYI }


###############################################################################
#
# _write_sheet_view()
#
# Write the <sheetView> element.
#
# Sample structure:
#     <sheetView
#         showGridLines="0"
#         showRowColHeaders="0"
#         showZeros="0"
#         rightToLeft="1"
#         tabSelected="1"
#         showRuler="0"
#         showOutlineSymbols="0"
#         view="pageLayout"
#         zoomScale="121"
#         zoomScaleNormal="121"
#         workbookViewId="0"
#      />
#
#NYI sub _write_sheet_view {
#NYI 
#NYI     my $self             = shift;
#NYI     my $gridlines        = $self->{_screen_gridlines};
#NYI     my $show_zeros       = $self->{_show_zeros};
#NYI     my $right_to_left    = $self->{_right_to_left};
#NYI     my $tab_selected     = $self->{_selected};
#NYI     my $view             = $self->{_page_view};
#NYI     my $zoom             = $self->{_zoom};
#NYI     my $workbook_view_id = 0;
#NYI     my @attributes       = ();
#NYI 
#NYI     # Hide screen gridlines if required
#NYI     if ( !$gridlines ) {
#NYI         push @attributes, ( 'showGridLines' => 0 );
#NYI     }
#NYI 
#NYI     # Hide zeroes in cells.
#NYI     if ( !$show_zeros ) {
#NYI         push @attributes, ( 'showZeros' => 0 );
#NYI     }
#NYI 
#NYI     # Display worksheet right to left for Hebrew, Arabic and others.
#NYI     if ( $right_to_left ) {
#NYI         push @attributes, ( 'rightToLeft' => 1 );
#NYI     }
#NYI 
#NYI     # Show that the sheet tab is selected.
#NYI     if ( $tab_selected ) {
#NYI         push @attributes, ( 'tabSelected' => 1 );
#NYI     }
#NYI 
#NYI 
#NYI     # Turn outlines off. Also required in the outlinePr element.
#NYI     if ( !$self->{_outline_on} ) {
#NYI         push @attributes, ( "showOutlineSymbols" => 0 );
#NYI     }
#NYI 
#NYI     # Set the page view/layout mode if required.
#NYI     # TODO. Add pageBreakPreview mode when requested.
#NYI     if ( $view ) {
#NYI         push @attributes, ( 'view' => 'pageLayout' );
#NYI     }
#NYI 
#NYI     # Set the zoom level.
#NYI     if ( $zoom != 100 ) {
#NYI         push @attributes, ( 'zoomScale' => $zoom ) unless $view;
#NYI         push @attributes, ( 'zoomScaleNormal' => $zoom )
#NYI           if $self->{_zoom_scale_normal};
#NYI     }
#NYI 
#NYI     push @attributes, ( 'workbookViewId' => $workbook_view_id );
#NYI 
#NYI     if ( @{ $self->{_panes} } || @{ $self->{_selections} } ) {
#NYI         $self->xml_start_tag( 'sheetView', @attributes );
#NYI         $self->_write_panes();
#NYI         $self->_write_selections();
#NYI         $self->xml_end_tag( 'sheetView' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'sheetView', @attributes );
#NYI     }
#NYI }


###############################################################################
#
# _write_selections()
#
# Write the <selection> elements.
#
#NYI sub _write_selections {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     for my $selection ( @{ $self->{_selections} } ) {
#NYI         $self->_write_selection( @$selection );
#NYI     }
#NYI }


###############################################################################
#
# _write_selection()
#
# Write the <selection> element.
#
#NYI sub _write_selection {
#NYI 
#NYI     my $self        = shift;
#NYI     my $pane        = shift;
#NYI     my $active_cell = shift;
#NYI     my $sqref       = shift;
#NYI     my @attributes  = ();
#NYI 
#NYI     push @attributes, ( 'pane'       => $pane )        if $pane;
#NYI     push @attributes, ( 'activeCell' => $active_cell ) if $active_cell;
#NYI     push @attributes, ( 'sqref'      => $sqref )       if $sqref;
#NYI 
#NYI     $self->xml_empty_tag( 'selection', @attributes );
#NYI }


###############################################################################
#
# _write_sheet_format_pr()
#
# Write the <sheetFormatPr> element.
#
#NYI sub _write_sheet_format_pr {
#NYI 
#NYI     my $self               = shift;
#NYI     my $base_col_width     = 10;
#NYI     my $default_row_height = $self->{_default_row_height};
#NYI     my $row_level          = $self->{_outline_row_level};
#NYI     my $col_level          = $self->{_outline_col_level};
#NYI     my $zero_height        = $self->{_default_row_zeroed};
#NYI 
#NYI     my @attributes = ( 'defaultRowHeight' => $default_row_height );
#NYI 
#NYI     if ( $self->{_default_row_height} != $self->{_original_row_height} ) {
#NYI         push @attributes, ( 'customHeight' => 1 );
#NYI     }
#NYI 
#NYI     if ( $self->{_default_row_zeroed} ) {
#NYI         push @attributes, ( 'zeroHeight' => 1 );
#NYI     }
#NYI 
#NYI     push @attributes, ( 'outlineLevelRow' => $row_level ) if $row_level;
#NYI     push @attributes, ( 'outlineLevelCol' => $col_level ) if $col_level;
#NYI 
#NYI     if ( $self->{_excel_version} == 2010 ) {
#NYI         push @attributes, ( 'x14ac:dyDescent' => '0.25' );
#NYI     }
#NYI 
#NYI     $self->xml_empty_tag( 'sheetFormatPr', @attributes );
#NYI }


##############################################################################
#
# _write_cols()
#
# Write the <cols> element and <col> sub elements.
#
#NYI sub _write_cols {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Exit unless some column have been formatted.
#NYI     return unless %{ $self->{_colinfo} };
#NYI 
#NYI     $self->xml_start_tag( 'cols' );
#NYI 
#NYI     for my $col ( sort keys %{ $self->{_colinfo} } ) {
#NYI         $self->_write_col_info( @{ $self->{_colinfo}->{$col} } );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'cols' );
#NYI }


##############################################################################
#
# _write_col_info()
#
# Write the <col> element.
#
#NYI sub _write_col_info {
#NYI 
#NYI     my $self         = shift;
#NYI     my $min          = $_[0] || 0;    # First formatted column.
#NYI     my $max          = $_[1] || 0;    # Last formatted column.
#NYI     my $width        = $_[2];         # Col width in user units.
#NYI     my $format       = $_[3];         # Format index.
#NYI     my $hidden       = $_[4] || 0;    # Hidden flag.
#NYI     my $level        = $_[5] || 0;    # Outline level.
#NYI     my $collapsed    = $_[6] || 0;    # Outline level.
#NYI     my $custom_width = 1;
#NYI     my $xf_index     = 0;
#NYI 
#NYI     # Get the format index.
#NYI     if ( ref( $format ) ) {
#NYI         $xf_index = $format->get_xf_index();
#NYI     }
#NYI 
#NYI     # Set the Excel default col width.
#NYI     if ( !defined $width ) {
#NYI         if ( !$hidden ) {
#NYI             $width        = 8.43;
#NYI             $custom_width = 0;
#NYI         }
#NYI         else {
#NYI             $width = 0;
#NYI         }
#NYI     }
#NYI     else {
#NYI 
#NYI         # Width is defined but same as default.
#NYI         if ( $width == 8.43 ) {
#NYI             $custom_width = 0;
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # Convert column width from user units to character width.
#NYI     my $max_digit_width = 7;    # For Calabri 11.
#NYI     my $padding         = 5;
#NYI 
#NYI     if ( $width > 0 ) {
#NYI         if ( $width < 1 ) {
#NYI             $width =
#NYI               int( ( int( $width * ($max_digit_width + $padding) + 0.5 ) ) /
#NYI                   $max_digit_width *
#NYI                   256 ) / 256;
#NYI         }
#NYI         else {
#NYI             $width =
#NYI               int( ( int( $width * $max_digit_width + 0.5 ) + $padding ) /
#NYI                   $max_digit_width *
#NYI                   256 ) / 256;
#NYI         }
#NYI     }
#NYI 
#NYI     my @attributes = (
#NYI         'min'   => $min + 1,
#NYI         'max'   => $max + 1,
#NYI         'width' => $width,
#NYI     );
#NYI 
#NYI     push @attributes, ( 'style'        => $xf_index ) if $xf_index;
#NYI     push @attributes, ( 'hidden'       => 1 )         if $hidden;
#NYI     push @attributes, ( 'customWidth'  => 1 )         if $custom_width;
#NYI     push @attributes, ( 'outlineLevel' => $level )    if $level;
#NYI     push @attributes, ( 'collapsed'    => 1 )         if $collapsed;
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'col', @attributes );
#NYI }


###############################################################################
#
# _write_sheet_data()
#
# Write the <sheetData> element.
#
#NYI sub _write_sheet_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     if ( not defined $self->{_dim_rowmin} ) {
#NYI 
#NYI         # If the dimensions aren't defined then there is no data to write.
#NYI         $self->xml_empty_tag( 'sheetData' );
#NYI     }
#NYI     else {
#NYI         $self->xml_start_tag( 'sheetData' );
#NYI         $self->_write_rows();
#NYI         $self->xml_end_tag( 'sheetData' );
#NYI 
#NYI     }
#NYI 
#NYI }


###############################################################################
#
# _write_optimized_sheet_data()
#
# Write the <sheetData> element when the memory optimisation is on. In which
# case we read the data stored in the temp file and rewrite it to the XML
# sheet file.
#
#NYI sub _write_optimized_sheet_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     if ( not defined $self->{_dim_rowmin} ) {
#NYI 
#NYI         # If the dimensions aren't defined then there is no data to write.
#NYI         $self->xml_empty_tag( 'sheetData' );
#NYI     }
#NYI     else {
#NYI 
#NYI         $self->xml_start_tag( 'sheetData' );
#NYI 
#NYI         my $xlsx_fh = $self->xml_get_fh();
#NYI         my $cell_fh = $self->{_cell_data_fh};
#NYI 
#NYI         my $buffer;
#NYI 
#NYI         # Rewind the temp file.
#NYI         seek $cell_fh, 0, 0;
#NYI 
#NYI         while ( read( $cell_fh, $buffer, 4_096 ) ) {
#NYI             local $\ = undef;    # Protect print from -l on commandline.
#NYI             print $xlsx_fh $buffer;
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'sheetData' );
#NYI     }
#NYI }


###############################################################################
#
# _write_rows()
#
# Write out the worksheet data as a series of rows and cells.
#
#NYI sub _write_rows {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_calculate_spans();
#NYI 
#NYI     for my $row_num ( $self->{_dim_rowmin} .. $self->{_dim_rowmax} ) {
#NYI 
#NYI         # Skip row if it doesn't contain row formatting, cell data or a comment.
#NYI         if (   !$self->{_set_rows}->{$row_num}
#NYI             && !$self->{_table}->{$row_num}
#NYI             && !$self->{_comments}->{$row_num} )
#NYI         {
#NYI             next;
#NYI         }
#NYI 
#NYI         my $span_index = int( $row_num / 16 );
#NYI         my $span       = $self->{_row_spans}->[$span_index];
#NYI 
#NYI         # Write the cells if the row contains data.
#NYI         if ( my $row_ref = $self->{_table}->{$row_num} ) {
#NYI 
#NYI             if ( !$self->{_set_rows}->{$row_num} ) {
#NYI                 $self->_write_row( $row_num, $span );
#NYI             }
#NYI             else {
#NYI                 $self->_write_row( $row_num, $span,
#NYI                     @{ $self->{_set_rows}->{$row_num} } );
#NYI             }
#NYI 
#NYI 
#NYI             for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
#NYI                 if ( my $col_ref = $self->{_table}->{$row_num}->{$col_num} ) {
#NYI                     $self->_write_cell( $row_num, $col_num, $col_ref );
#NYI                 }
#NYI             }
#NYI 
#NYI             $self->xml_end_tag( 'row' );
#NYI         }
#NYI         elsif ( $self->{_comments}->{$row_num} ) {
#NYI 
#NYI             $self->_write_empty_row( $row_num, $span,
#NYI                 @{ $self->{_set_rows}->{$row_num} } );
#NYI         }
#NYI         else {
#NYI 
#NYI             # Row attributes only.
#NYI             $self->_write_empty_row( $row_num, $span,
#NYI                 @{ $self->{_set_rows}->{$row_num} } );
#NYI         }
#NYI     }
#NYI }


###############################################################################
#
# _write_single_row()
#
# Write out the worksheet data as a single row with cells. This method is
# used when memory optimisation is on. A single row is written and the data
# table is reset. That way only one row of data is kept in memory at any one
# time. We don't write span data in the optimised case since it is optional.
#
#NYI sub _write_single_row {
#NYI 
#NYI     my $self        = shift;
#NYI     my $current_row = shift || 0;
#NYI     my $row_num     = $self->{_previous_row};
#NYI 
#NYI     # Set the new previous row as the current row.
#NYI     $self->{_previous_row} = $current_row;
#NYI 
#NYI     # Skip row if it doesn't contain row formatting, cell data or a comment.
#NYI     if (   !$self->{_set_rows}->{$row_num}
#NYI         && !$self->{_table}->{$row_num}
#NYI         && !$self->{_comments}->{$row_num} )
#NYI     {
#NYI         return;
#NYI     }
#NYI 
#NYI     # Write the cells if the row contains data.
#NYI     if ( my $row_ref = $self->{_table}->{$row_num} ) {
#NYI 
#NYI         if ( !$self->{_set_rows}->{$row_num} ) {
#NYI             $self->_write_row( $row_num );
#NYI         }
#NYI         else {
#NYI             $self->_write_row( $row_num, undef,
#NYI                 @{ $self->{_set_rows}->{$row_num} } );
#NYI         }
#NYI 
#NYI         for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
#NYI             if ( my $col_ref = $self->{_table}->{$row_num}->{$col_num} ) {
#NYI                 $self->_write_cell( $row_num, $col_num, $col_ref );
#NYI             }
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'row' );
#NYI     }
#NYI     else {
#NYI 
#NYI         # Row attributes or comments only.
#NYI         $self->_write_empty_row( $row_num, undef,
#NYI             @{ $self->{_set_rows}->{$row_num} } );
#NYI     }
#NYI 
#NYI     # Reset table.
#NYI     $self->{_table} = {};
#NYI 
#NYI }


###############################################################################
#
# _calculate_spans()
#
# Calculate the "spans" attribute of the <row> tag. This is an XLSX
# optimisation and isn't strictly required. However, it makes comparing
# files easier.
#
# The span is the same for each block of 16 rows.
#
#NYI sub _calculate_spans {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @spans;
#NYI     my $span_min;
#NYI     my $span_max;
#NYI 
#NYI     for my $row_num ( $self->{_dim_rowmin} .. $self->{_dim_rowmax} ) {
#NYI 
#NYI         # Calculate spans for cell data.
#NYI         if ( my $row_ref = $self->{_table}->{$row_num} ) {
#NYI 
#NYI             for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
#NYI                 if ( my $col_ref = $self->{_table}->{$row_num}->{$col_num} ) {
#NYI 
#NYI                     if ( !defined $span_min ) {
#NYI                         $span_min = $col_num;
#NYI                         $span_max = $col_num;
#NYI                     }
#NYI                     else {
#NYI                         $span_min = $col_num if $col_num < $span_min;
#NYI                         $span_max = $col_num if $col_num > $span_max;
#NYI                     }
#NYI                 }
#NYI             }
#NYI         }
#NYI 
#NYI         # Calculate spans for comments.
#NYI         if ( defined $self->{_comments}->{$row_num} ) {
#NYI 
#NYI             for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
#NYI                 if ( defined $self->{_comments}->{$row_num}->{$col_num} ) {
#NYI 
#NYI                     if ( !defined $span_min ) {
#NYI                         $span_min = $col_num;
#NYI                         $span_max = $col_num;
#NYI                     }
#NYI                     else {
#NYI                         $span_min = $col_num if $col_num < $span_min;
#NYI                         $span_max = $col_num if $col_num > $span_max;
#NYI                     }
#NYI                 }
#NYI             }
#NYI         }
#NYI 
#NYI         if ( ( ( $row_num + 1 ) % 16 == 0 )
#NYI             || $row_num == $self->{_dim_rowmax} )
#NYI         {
#NYI             my $span_index = int( $row_num / 16 );
#NYI 
#NYI             if ( defined $span_min ) {
#NYI                 $span_min++;
#NYI                 $span_max++;
#NYI                 $spans[$span_index] = "$span_min:$span_max";
#NYI                 $span_min = undef;
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     $self->{_row_spans} = \@spans;
#NYI }


###############################################################################
#
# _write_row()
#
# Write the <row> element.
#
#NYI sub _write_row {
#NYI 
#NYI     my $self      = shift;
#NYI     my $r         = shift;
#NYI     my $spans     = shift;
#NYI     my $height    = shift;
#NYI     my $format    = shift;
#NYI     my $hidden    = shift || 0;
#NYI     my $level     = shift || 0;
#NYI     my $collapsed = shift || 0;
#NYI     my $empty_row = shift || 0;
#NYI     my $xf_index  = 0;
#NYI 
#NYI     $height = $self->{_default_row_height} if !defined $height;
#NYI 
#NYI     my @attributes = ( 'r' => $r + 1 );
#NYI 
#NYI     # Get the format index.
#NYI     if ( ref( $format ) ) {
#NYI         $xf_index = $format->get_xf_index();
#NYI     }
#NYI 
#NYI     push @attributes, ( 'spans'        => $spans )    if defined $spans;
#NYI     push @attributes, ( 's'            => $xf_index ) if $xf_index;
#NYI     push @attributes, ( 'customFormat' => 1 )         if $format;
#NYI 
#NYI     if ( $height != $self->{_original_row_height} ) {
#NYI         push @attributes, ( 'ht' => $height );
#NYI     }
#NYI 
#NYI     push @attributes, ( 'hidden'       => 1 )         if $hidden;
#NYI 
#NYI     if ( $height != $self->{_original_row_height} ) {
#NYI         push @attributes, ( 'customHeight' => 1 );
#NYI     }
#NYI 
#NYI     push @attributes, ( 'outlineLevel' => $level )    if $level;
#NYI     push @attributes, ( 'collapsed'    => 1 )         if $collapsed;
#NYI 
#NYI     if ( $self->{_excel_version} == 2010 ) {
#NYI         push @attributes, ( 'x14ac:dyDescent' => '0.25' );
#NYI     }
#NYI 
#NYI     if ( $empty_row ) {
#NYI         $self->xml_empty_tag_unencoded( 'row', @attributes );
#NYI     }
#NYI     else {
#NYI         $self->xml_start_tag_unencoded( 'row', @attributes );
#NYI     }
#NYI }


###############################################################################
#
# _write_empty_row()
#
# Write and empty <row> element, i.e., attributes only, no cell data.
#
#NYI sub _write_empty_row {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Set the $empty_row parameter.
#NYI     $_[7] = 1;
#NYI 
#NYI     $self->_write_row( @_ );
#NYI }


###############################################################################
#
# _write_cell()
#
# Write the <cell> element. This is the innermost loop so efficiency is
# important where possible. The basic methodology is that the data of every
# cell type is passed in as follows:
#
#      [ $row, $col, $aref]
#
# The aref, called $cell below, contains the following structure in all types:
#
#     [ $type, $token, $xf, @args ]
#
# Where $type:  represents the cell type, such as string, number, formula, etc.
#       $token: is the actual data for the string, number, formula, etc.
#       $xf:    is the XF format object.
#       @args:  additional args relevant to the specific data type.
#
#NYI sub _write_cell {
#NYI 
#NYI     my $self     = shift;
#NYI     my $row      = shift;
#NYI     my $col      = shift;
#NYI     my $cell     = shift;
#NYI     my $type     = $cell->[0];
#NYI     my $token    = $cell->[1];
#NYI     my $xf       = $cell->[2];
#NYI     my $xf_index = 0;
#NYI 
#NYI     my %error_codes = (
#NYI         '#DIV/0!' => 1,
#NYI         '#N/A'    => 1,
#NYI         '#NAME?'  => 1,
#NYI         '#NULL!'  => 1,
#NYI         '#NUM!'   => 1,
#NYI         '#REF!'   => 1,
#NYI         '#VALUE!' => 1,
#NYI     );
#NYI 
#NYI     my %boolean = ( 'TRUE' => 1, 'FALSE' => 0 );
#NYI 
#NYI     # Get the format index.
#NYI     if ( ref( $xf ) ) {
#NYI         $xf_index = $xf->get_xf_index();
#NYI     }
#NYI 
#NYI     my $range = _xl-rowcol-to-cell( $row, $col );
#NYI     my @attributes = ( 'r' => $range );
#NYI 
#NYI     # Add the cell format index.
#NYI     if ( $xf_index ) {
#NYI         push @attributes, ( 's' => $xf_index );
#NYI     }
#NYI     elsif ( $self->{_set_rows}->{$row} && $self->{_set_rows}->{$row}->[1] ) {
#NYI         my $row_xf = $self->{_set_rows}->{$row}->[1];
#NYI         push @attributes, ( 's' => $row_xf->get_xf_index() );
#NYI     }
#NYI     elsif ( $self->{_col_formats}->{$col} ) {
#NYI         my $col_xf = $self->{_col_formats}->{$col};
#NYI         push @attributes, ( 's' => $col_xf->get_xf_index() );
#NYI     }
#NYI 
#NYI 
#NYI     # Write the various cell types.
#NYI     if ( $type eq 'n' ) {
#NYI 
#NYI         # Write a number.
#NYI         $self->xml_number_element( $token, @attributes );
#NYI     }
#NYI     elsif ( $type eq 's' ) {
#NYI 
#NYI         # Write a string.
#NYI         if ( $self->{_optimization} == 0 ) {
#NYI             $self->xml_string_element( $token, @attributes );
#NYI         }
#NYI         else {
#NYI 
#NYI             my $string = $token;
#NYI 
#NYI             # Escape control characters. See SharedString.pm for details.
#NYI             $string =~ s/(_x[0-9a-fA-F]{4}_)/_x005F$1/g;
#NYI             $string =~ s/([\x00-\x08\x0B-\x1F])/sprintf "_x%04X_", ord($1)/eg;
#NYI 
#NYI             # Write any rich strings without further tags.
#NYI             if ( $string =~ m{^<r>} && $string =~ m{</r>$} ) {
#NYI 
#NYI                 $self->xml_rich_inline_string( $string, @attributes );
#NYI             }
#NYI             else {
#NYI 
#NYI                 # Add attribute to preserve leading or trailing whitespace.
#NYI                 my $preserve = 0;
#NYI                 if ( $string =~ /^\s/ || $string =~ /\s$/ ) {
#NYI                     $preserve = 1;
#NYI                 }
#NYI 
#NYI                 $self->xml_inline_string( $string, $preserve, @attributes );
#NYI             }
#NYI         }
#NYI     }
#NYI     elsif ( $type eq 'f' ) {
#NYI 
#NYI         # Write a formula.
#NYI         my $value = $cell->[3] || 0;
#NYI 
#NYI         # Check if the formula value is a string.
#NYI         if (   $value
#NYI             && $value !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/ )
#NYI         {
#NYI             if ( exists $boolean{$value} ) {
#NYI                 push @attributes, ( 't' => 'b' );
#NYI                 $value = $boolean{$value};
#NYI             }
#NYI             elsif ( exists $error_codes{$value} ) {
#NYI                 push @attributes, ( 't' => 'e' );
#NYI             }
#NYI             else {
#NYI                 push @attributes, ( 't' => 'str' );
#NYI                 $value = Excel::Writer::XLSX::Package::XMLwriter::_escape_data(
#NYI                     $value );
#NYI             }
#NYI         }
#NYI 
#NYI         $self->xml_formula_element( $token, $value, @attributes );
#NYI 
#NYI     }
#NYI     elsif ( $type eq 'a' ) {
#NYI 
#NYI         # Write an array formula.
#NYI         $self->xml_start_tag( 'c', @attributes );
#NYI         $self->_write_cell_array_formula( $token, $cell->[3] );
#NYI         $self->_write_cell_value( $cell->[4] );
#NYI         $self->xml_end_tag( 'c' );
#NYI     }
#NYI     elsif ( $type eq 'l' ) {
#NYI 
#NYI         # Write a boolean value.
#NYI         push @attributes, ( 't' => 'b' );
#NYI 
#NYI         $self->xml_start_tag( 'c', @attributes );
#NYI         $self->_write_cell_value( $cell->[1] );
#NYI         $self->xml_end_tag( 'c' );
#NYI     }
#NYI     elsif ( $type eq 'b' ) {
#NYI 
#NYI         # Write a empty cell.
#NYI         $self->xml_empty_tag( 'c', @attributes );
#NYI     }
#NYI }


###############################################################################
#
# _write_cell_value()
#
# Write the cell value <v> element.
#
#NYI sub _write_cell_value {
#NYI 
#NYI     my $self = shift;
#NYI     my $value = defined $_[0] ? $_[0] : '';
#NYI 
#NYI     $self->xml_data_element( 'v', $value );
#NYI }


###############################################################################
#
# _write_cell_formula()
#
# Write the cell formula <f> element.
#
#NYI sub _write_cell_formula {
#NYI 
#NYI     my $self = shift;
#NYI     my $formula = defined $_[0] ? $_[0] : '';
#NYI 
#NYI     $self->xml_data_element( 'f', $formula );
#NYI }


###############################################################################
#
# _write_cell_array_formula()
#
# Write the cell array formula <f> element.
#
#NYI sub _write_cell_array_formula {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI     my $range   = shift;
#NYI 
#NYI     my @attributes = ( 't' => 'array', 'ref' => $range );
#NYI 
#NYI     $self->xml_data_element( 'f', $formula, @attributes );
#NYI }


##############################################################################
#
# _write_sheet_calc_pr()
#
# Write the <sheetCalcPr> element for the worksheet calculation properties.
#
#NYI sub _write_sheet_calc_pr {
#NYI 
#NYI     my $self              = shift;
#NYI     my $full_calc_on_load = 1;
#NYI 
#NYI     my @attributes = ( 'fullCalcOnLoad' => $full_calc_on_load );
#NYI 
#NYI     $self->xml_empty_tag( 'sheetCalcPr', @attributes );
#NYI }


###############################################################################
#
# _write_phonetic_pr()
#
# Write the <phoneticPr> element.
#
#NYI sub _write_phonetic_pr {
#NYI 
#NYI     my $self    = shift;
#NYI     my $font_id = 0;
#NYI     my $type    = 'noConversion';
#NYI 
#NYI     my @attributes = (
#NYI         'fontId' => $font_id,
#NYI         'type'   => $type,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'phoneticPr', @attributes );
#NYI }


###############################################################################
#
# _write_page_margins()
#
# Write the <pageMargins> element.
#
#NYI sub _write_page_margins {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'left'   => $self->{_margin_left},
#NYI         'right'  => $self->{_margin_right},
#NYI         'top'    => $self->{_margin_top},
#NYI         'bottom' => $self->{_margin_bottom},
#NYI         'header' => $self->{_margin_header},
#NYI         'footer' => $self->{_margin_footer},
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'pageMargins', @attributes );
#NYI }


###############################################################################
#
# _write_page_setup()
#
# Write the <pageSetup> element.
#
# The following is an example taken from Excel.
#
# <pageSetup
#     paperSize="9"
#     scale="110"
#     fitToWidth="2"
#     fitToHeight="2"
#     pageOrder="overThenDown"
#     orientation="portrait"
#     blackAndWhite="1"
#     draft="1"
#     horizontalDpi="200"
#     verticalDpi="200"
#     r:id="rId1"
# />
#
#NYI sub _write_page_setup {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     return unless $self->{_page_setup_changed};
#NYI 
#NYI     # Set paper size.
#NYI     if ( $self->{_paper_size} ) {
#NYI         push @attributes, ( 'paperSize' => $self->{_paper_size} );
#NYI     }
#NYI 
#NYI     # Set the print_scale
#NYI     if ( $self->{_print_scale} != 100 ) {
#NYI         push @attributes, ( 'scale' => $self->{_print_scale} );
#NYI     }
#NYI 
#NYI     # Set the "Fit to page" properties.
#NYI     if ( $self->{_fit_page} && $self->{_fit_width} != 1 ) {
#NYI         push @attributes, ( 'fitToWidth' => $self->{_fit_width} );
#NYI     }
#NYI 
#NYI     if ( $self->{_fit_page} && $self->{_fit_height} != 1 ) {
#NYI         push @attributes, ( 'fitToHeight' => $self->{_fit_height} );
#NYI     }
#NYI 
#NYI     # Set the page print direction.
#NYI     if ( $self->{_page_order} ) {
#NYI         push @attributes, ( 'pageOrder' => "overThenDown" );
#NYI     }
#NYI 
#NYI     # Set start page.
#NYI     if ( $self->{_page_start} > 1 ) {
#NYI         push @attributes, ( 'firstPageNumber' => $self->{_page_start} );
#NYI     }
#NYI 
#NYI     # Set page orientation.
#NYI     if ( $self->{_orientation} == 0 ) {
#NYI         push @attributes, ( 'orientation' => 'landscape' );
#NYI     }
#NYI     else {
#NYI         push @attributes, ( 'orientation' => 'portrait' );
#NYI     }
#NYI 
#NYI     # Set print in black and white option.
#NYI     if ( $self->{_black_white} ) {
#NYI         push @attributes, ( 'blackAndWhite' => 1 );
#NYI     }
#NYI 
#NYI     # Set start page.
#NYI     if ( $self->{_page_start} != 0 ) {
#NYI         push @attributes, ( 'useFirstPageNumber' => 1 );
#NYI     }
#NYI 
#NYI     # Set the DPI. Mainly only for testing.
#NYI     if ( $self->{_horizontal_dpi} ) {
#NYI         push @attributes, ( 'horizontalDpi' => $self->{_horizontal_dpi} );
#NYI     }
#NYI 
#NYI     if ( $self->{_vertical_dpi} ) {
#NYI         push @attributes, ( 'verticalDpi' => $self->{_vertical_dpi} );
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'pageSetup', @attributes );
#NYI }


##############################################################################
#
# _write_merge_cells()
#
# Write the <mergeCells> element.
#
#NYI sub _write_merge_cells {
#NYI 
#NYI     my $self         = shift;
#NYI     my $merged_cells = $self->{_merge};
#NYI     my $count        = @$merged_cells;
#NYI 
#NYI     return unless $count;
#NYI 
#NYI     my @attributes = ( 'count' => $count );
#NYI 
#NYI     $self->xml_start_tag( 'mergeCells', @attributes );
#NYI 
#NYI     for my $merged_range ( @$merged_cells ) {
#NYI 
#NYI         # Write the mergeCell element.
#NYI         $self->_write_merge_cell( $merged_range );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'mergeCells' );
#NYI }


##############################################################################
#
# _write_merge_cell()
#
# Write the <mergeCell> element.
#
#NYI sub _write_merge_cell {
#NYI 
#NYI     my $self         = shift;
#NYI     my $merged_range = shift;
#NYI     my ( $row_min, $col_min, $row_max, $col_max ) = @$merged_range;
#NYI 
#NYI 
#NYI     # Convert the merge dimensions to a cell range.
#NYI     my $cell_1 = xl-rowcol-to-cell( $row_min, $col_min );
#NYI     my $cell_2 = xl-rowcol-to-cell( $row_max, $col_max );
#NYI     my $ref    = $cell_1 . ':' . $cell_2;
#NYI 
#NYI     my @attributes = ( 'ref' => $ref );
#NYI 
#NYI     $self->xml_empty_tag( 'mergeCell', @attributes );
#NYI }


##############################################################################
#
# _write_print_options()
#
# Write the <printOptions> element.
#
#NYI sub _write_print_options {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     return unless $self->{_print_options_changed};
#NYI 
#NYI     # Set horizontal centering.
#NYI     if ( $self->{_hcenter} ) {
#NYI         push @attributes, ( 'horizontalCentered' => 1 );
#NYI     }
#NYI 
#NYI     # Set vertical centering.
#NYI     if ( $self->{_vcenter} ) {
#NYI         push @attributes, ( 'verticalCentered' => 1 );
#NYI     }
#NYI 
#NYI     # Enable row and column headers.
#NYI     if ( $self->{_print_headers} ) {
#NYI         push @attributes, ( 'headings' => 1 );
#NYI     }
#NYI 
#NYI     # Set printed gridlines.
#NYI     if ( $self->{_print_gridlines} ) {
#NYI         push @attributes, ( 'gridLines' => 1 );
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'printOptions', @attributes );
#NYI }


##############################################################################
#
# _write_header_footer()
#
# Write the <headerFooter> element.
#
#NYI sub _write_header_footer {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     if ( !$self->{_header_footer_scales} ) {
#NYI         push @attributes, ( 'scaleWithDoc' => 0 );
#NYI     }
#NYI 
#NYI     if ( !$self->{_header_footer_aligns} ) {
#NYI         push @attributes, ( 'alignWithMargins' => 0 );
#NYI     }
#NYI 
#NYI     if ( $self->{_header_footer_changed} ) {
#NYI         $self->xml_start_tag( 'headerFooter', @attributes );
#NYI         $self->_write_odd_header() if $self->{_header};
#NYI         $self->_write_odd_footer() if $self->{_footer};
#NYI         $self->xml_end_tag( 'headerFooter' );
#NYI     }
#NYI     elsif ( $self->{_excel2003_style} ) {
#NYI         $self->xml_empty_tag( 'headerFooter', @attributes );
#NYI     }
#NYI }


##############################################################################
#
# _write_odd_header()
#
# Write the <oddHeader> element.
#
method write_odd_header {
    my $data = $!header;
    self.xml_data_element( 'oddHeader', $data );
}


##############################################################################
#
# _write_odd_footer()
#
# Write the <oddFooter> element.
#
method write_odd_footer {
    self.xml_data_element( 'oddFooter', $!footer );
}


##############################################################################
#
# _write_row_breaks()
#
# Write the <rowBreaks> element.
#
#NYI sub _write_row_breaks {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @page_breaks = $self->_sort_pagebreaks( @{ $self->{_hbreaks} } );
#NYI     my $count       = scalar @page_breaks;
#NYI 
#NYI     return unless @page_breaks;
#NYI 
#NYI     my @attributes = (
#NYI         'count'            => $count,
#NYI         'manualBreakCount' => $count,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'rowBreaks', @attributes );
#NYI 
#NYI     for my $row_num ( @page_breaks ) {
#NYI         $self->_write_brk( $row_num, 16383 );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'rowBreaks' );
#NYI }


##############################################################################
#
# _write_col_breaks()
#
# Write the <colBreaks> element.
#
#NYI sub _write_col_breaks {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @page_breaks = $self->_sort_pagebreaks( @{ $self->{_vbreaks} } );
#NYI     my $count       = scalar @page_breaks;
#NYI 
#NYI     return unless @page_breaks;
#NYI 
#NYI     my @attributes = (
#NYI         'count'            => $count,
#NYI         'manualBreakCount' => $count,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'colBreaks', @attributes );
#NYI 
#NYI     for my $col_num ( @page_breaks ) {
#NYI         $self->_write_brk( $col_num, 1048575 );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'colBreaks' );
#NYI }


##############################################################################
#
# _write_brk()
#
# Write the <brk> element.
#
method write_brk($id, $max) {
    my $man  = 1;

    my @attributes = (
        'id'  => $id,
        'max' => $max,
        'man' => $man,
    );

    self.xml_empty_tag( 'brk', @attributes );
}


##############################################################################
#
# _write_auto_filter()
#
# Write the <autoFilter> element.
#
#NYI sub _write_auto_filter {
#NYI 
#NYI     my $self = shift;
#NYI     my $ref  = $self->{_autofilter_ref};
#NYI 
#NYI     return unless $ref;
#NYI 
#NYI     my @attributes = ( 'ref' => $ref );
#NYI 
#NYI     if ( $self->{_filter_on} ) {
#NYI 
#NYI         # Autofilter defined active filters.
#NYI         $self->xml_start_tag( 'autoFilter', @attributes );
#NYI 
#NYI         $self->_write_autofilters();
#NYI 
#NYI         $self->xml_end_tag( 'autoFilter' );
#NYI 
#NYI     }
#NYI     else {
#NYI 
#NYI         # Autofilter defined without active filters.
#NYI         $self->xml_empty_tag( 'autoFilter', @attributes );
#NYI     }
#NYI 
#NYI }


###############################################################################
#
# _write_autofilters()
#
# Function to iterate through the columns that form part of an autofilter
# range and write the appropriate filters.
#
#NYI sub _write_autofilters {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my ( $col1, $col2 ) = @{ $self->{_filter_range} };
#NYI 
#NYI     for my $col ( $col1 .. $col2 ) {
#NYI 
#NYI         # Skip if column doesn't have an active filter.
#NYI         next unless $self->{_filter_cols}->{$col};
#NYI 
#NYI         # Retrieve the filter tokens and write the autofilter records.
#NYI         my @tokens = @{ $self->{_filter_cols}->{$col} };
#NYI         my $type   = $self->{_filter_type}->{$col};
#NYI 
#NYI         # Filters are relative to first column in the autofilter.
#NYI         $self->_write_filter_column( $col - $col1, $type, \@tokens );
#NYI     }
#NYI }


##############################################################################
#
# _write_filter_column()
#
# Write the <filterColumn> element.
#
#NYI sub _write_filter_column {
#NYI 
#NYI     my $self    = shift;
#NYI     my $col_id  = shift;
#NYI     my $type    = shift;
#NYI     my $filters = shift;
#NYI 
#NYI     my @attributes = ( 'colId' => $col_id );
#NYI 
#NYI     $self->xml_start_tag( 'filterColumn', @attributes );
#NYI 
#NYI 
#NYI     if ( $type == 1 ) {
#NYI 
#NYI         # Type == 1 is the new XLSX style filter.
#NYI         $self->_write_filters( @$filters );
#NYI 
#NYI     }
#NYI     else {
#NYI 
#NYI         # Type == 0 is the classic "custom" filter.
#NYI         $self->_write_custom_filters( @$filters );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'filterColumn' );
#NYI }


##############################################################################
#
# _write_filters()
#
# Write the <filters> element.
#
#NYI sub _write_filters {
#NYI 
#NYI     my $self    = shift;
#NYI     my @filters = @_;
#NYI 
#NYI     if ( @filters == 1 && $filters[0] eq 'blanks' ) {
#NYI 
#NYI         # Special case for blank cells only.
#NYI         $self->xml_empty_tag( 'filters', 'blank' => 1 );
#NYI     }
#NYI     else {
#NYI 
#NYI         # General case.
#NYI         $self->xml_start_tag( 'filters' );
#NYI 
#NYI         for my $filter ( @filters ) {
#NYI             $self->_write_filter( $filter );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'filters' );
#NYI     }
#NYI }


##############################################################################
#
# _write_filter()
#
# Write the <filter> element.
#
#NYI sub _write_filter {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'filter', @attributes );
#NYI }


##############################################################################
#
# _write_custom_filters()
#
# Write the <customFilters> element.
#
#NYI sub _write_custom_filters {
#NYI 
#NYI     my $self   = shift;
#NYI     my @tokens = @_;
#NYI 
#NYI     if ( @tokens == 2 ) {
#NYI 
#NYI         # One filter expression only.
#NYI         $self->xml_start_tag( 'customFilters' );
#NYI         $self->_write_custom_filter( @tokens );
#NYI         $self->xml_end_tag( 'customFilters' );
#NYI 
#NYI     }
#NYI     else {
#NYI 
#NYI         # Two filter expressions.
#NYI 
#NYI         my @attributes;
#NYI 
#NYI         # Check if the "join" operand is "and" or "or".
#NYI         if ( $tokens[2] == 0 ) {
#NYI             @attributes = ( 'and' => 1 );
#NYI         }
#NYI         else {
#NYI             @attributes = ( 'and' => 0 );
#NYI         }
#NYI 
#NYI         # Write the two custom filters.
#NYI         $self->xml_start_tag( 'customFilters', @attributes );
#NYI         $self->_write_custom_filter( $tokens[0], $tokens[1] );
#NYI         $self->_write_custom_filter( $tokens[3], $tokens[4] );
#NYI         $self->xml_end_tag( 'customFilters' );
#NYI     }
#NYI }


##############################################################################
#
# _write_custom_filter()
#
# Write the <customFilter> element.
#
#NYI sub _write_custom_filter {
#NYI 
#NYI     my $self       = shift;
#NYI     my $operator   = shift;
#NYI     my $val        = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     my %operators = (
#NYI         1  => 'lessThan',
#NYI         2  => 'equal',
#NYI         3  => 'lessThanOrEqual',
#NYI         4  => 'greaterThan',
#NYI         5  => 'notEqual',
#NYI         6  => 'greaterThanOrEqual',
#NYI         22 => 'equal',
#NYI     );
#NYI 
#NYI 
#NYI     # Convert the operator from a number to a descriptive string.
#NYI     if ( defined $operators{$operator} ) {
#NYI         $operator = $operators{$operator};
#NYI     }
#NYI     else {
#NYI         fail "Unknown operator = $operator\n";
#NYI     }
#NYI 
#NYI     # The 'equal' operator is the default attribute and isn't stored.
#NYI     push @attributes, ( 'operator' => $operator ) unless $operator eq 'equal';
#NYI     push @attributes, ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'customFilter', @attributes );
#NYI }


##############################################################################
#
# _write_hyperlinks()
#
# Process any stored hyperlinks in row/col order and write the <hyperlinks>
# element. The attributes are different for internal and external links.
#
#NYI sub _write_hyperlinks {
#NYI 
#NYI     my $self = shift;
#NYI     my @hlink_refs;
#NYI 
#NYI     # Sort the hyperlinks into row order.
#NYI     my @row_nums = sort { $a <=> $b } keys %{ $self->{_hyperlinks} };
#NYI 
#NYI     # Exit if there are no hyperlinks to process.
#NYI     return if !@row_nums;
#NYI 
#NYI     # Iterate over the rows.
#NYI     for my $row_num ( @row_nums ) {
#NYI 
#NYI         # Sort the hyperlinks into column order.
#NYI         my @col_nums = sort { $a <=> $b }
#NYI           keys %{ $self->{_hyperlinks}->{$row_num} };
#NYI 
#NYI         # Iterate over the columns.
#NYI         for my $col_num ( @col_nums ) {
#NYI 
#NYI             # Get the link data for this cell.
#NYI             my $link      = $self->{_hyperlinks}->{$row_num}->{$col_num};
#NYI             my $link_type = $link->{_link_type};
#NYI 
#NYI 
#NYI             # If the cell isn't a string then we have to add the url as
#NYI             # the string to display.
#NYI             my $display;
#NYI             if (   $self->{_table}
#NYI                 && $self->{_table}->{$row_num}
#NYI                 && $self->{_table}->{$row_num}->{$col_num} )
#NYI             {
#NYI                 my $cell = $self->{_table}->{$row_num}->{$col_num};
#NYI                 $display = $link->{_url} if $cell->[0] ne 's';
#NYI             }
#NYI 
#NYI 
#NYI             if ( $link_type == 1 ) {
#NYI 
#NYI                 # External link with rel file relationship.
#NYI                 push @hlink_refs,
#NYI                   [
#NYI                     $link_type,    $row_num,
#NYI                     $col_num,      ++$self->{_rel_count},
#NYI                     $link->{_str}, $display,
#NYI                     $link->{_tip}
#NYI                   ];
#NYI 
#NYI                 # Links for use by the packager.
#NYI                 push @{ $self->{_external_hyper_links} },
#NYI                   [ '/hyperlink', $link->{_url}, 'External' ];
#NYI             }
#NYI             else {
#NYI 
#NYI                 # Internal link with rel file relationship.
#NYI                 push @hlink_refs,
#NYI                   [
#NYI                     $link_type,    $row_num,      $col_num,
#NYI                     $link->{_url}, $link->{_str}, $link->{_tip}
#NYI                   ];
#NYI             }
#NYI         }
#NYI     }
#NYI 
#NYI     # Write the hyperlink elements.
#NYI     $self->xml_start_tag( 'hyperlinks' );
#NYI 
#NYI     for my $aref ( @hlink_refs ) {
#NYI         my ( $type, @args ) = @$aref;
#NYI 
#NYI         if ( $type == 1 ) {
#NYI             $self->_write_hyperlink_external( @args );
#NYI         }
#NYI         elsif ( $type == 2 ) {
#NYI             $self->_write_hyperlink_internal( @args );
#NYI         }
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'hyperlinks' );
#NYI }


##############################################################################
#
# _write_hyperlink_external()
#
# Write the <hyperlink> element for external links.
#
#NYI sub _write_hyperlink_external {
#NYI 
#NYI     my $self     = shift;
#NYI     my $row      = shift;
#NYI     my $col      = shift;
#NYI     my $id       = shift;
#NYI     my $location = shift;
#NYI     my $display  = shift;
#NYI     my $tooltip  = shift;
#NYI 
#NYI     my $ref = xl-rowcol-to-cell( $row, $col );
#NYI     my $r_id = 'rId' . $id;
#NYI 
#NYI     my @attributes = (
#NYI         'ref'  => $ref,
#NYI         'r:id' => $r_id,
#NYI     );
#NYI 
#NYI     push @attributes, ( 'location' => $location ) if defined $location;
#NYI     push @attributes, ( 'display' => $display )   if defined $display;
#NYI     push @attributes, ( 'tooltip'  => $tooltip )  if defined $tooltip;
#NYI 
#NYI     $self->xml_empty_tag( 'hyperlink', @attributes );
#NYI }


##############################################################################
#
# _write_hyperlink_internal()
#
# Write the <hyperlink> element for internal links.
#
#NYI sub _write_hyperlink_internal {
#NYI 
#NYI     my $self     = shift;
#NYI     my $row      = shift;
#NYI     my $col      = shift;
#NYI     my $location = shift;
#NYI     my $display  = shift;
#NYI     my $tooltip  = shift;
#NYI 
#NYI     my $ref = xl-rowcol-to-cell( $row, $col );
#NYI 
#NYI     my @attributes = ( 'ref' => $ref, 'location' => $location );
#NYI 
#NYI     push @attributes, ( 'tooltip' => $tooltip ) if defined $tooltip;
#NYI     push @attributes, ( 'display' => $display );
#NYI 
#NYI     $self->xml_empty_tag( 'hyperlink', @attributes );
#NYI }


##############################################################################
#
# _write_panes()
#
# Write the frozen or split <pane> elements.
#
#NYI sub _write_panes {
#NYI 
#NYI     my $self  = shift;
#NYI     my @panes = @{ $self->{_panes} };
#NYI 
#NYI     return unless @panes;
#NYI 
#NYI     if ( $panes[4] == 2 ) {
#NYI         $self->_write_split_panes( @panes );
#NYI     }
#NYI     else {
#NYI         $self->_write_freeze_panes( @panes );
#NYI     }
#NYI }


##############################################################################
#
# _write_freeze_panes()
#
# Write the <pane> element for freeze panes.
#
#NYI sub _write_freeze_panes {
#NYI 
#NYI     my $self = shift;
#NYI     my @attributes;
#NYI 
#NYI     my ( $row, $col, $top_row, $left_col, $type ) = @_;
#NYI 
#NYI     my $y_split       = $row;
#NYI     my $x_split       = $col;
#NYI     my $top_left_cell = xl-rowcol-to-cell( $top_row, $left_col );
#NYI     my $active_pane;
#NYI     my $state;
#NYI     my $active_cell;
#NYI     my $sqref;
#NYI 
#NYI     # Move user cell selection to the panes.
#NYI     if ( @{ $self->{_selections} } ) {
#NYI         ( undef, $active_cell, $sqref ) = @{ $self->{_selections}->[0] };
#NYI         $self->{_selections} = [];
#NYI     }
#NYI 
#NYI     # Set the active pane.
#NYI     if ( $row && $col ) {
#NYI         $active_pane = 'bottomRight';
#NYI 
#NYI         my $row_cell = xl-rowcol-to-cell( $row, 0 );
#NYI         my $col_cell = xl-rowcol-to-cell( 0,    $col );
#NYI 
#NYI         push @{ $self->{_selections} },
#NYI           (
#NYI             [ 'topRight',    $col_cell,    $col_cell ],
#NYI             [ 'bottomLeft',  $row_cell,    $row_cell ],
#NYI             [ 'bottomRight', $active_cell, $sqref ]
#NYI           );
#NYI     }
#NYI     elsif ( $col ) {
#NYI         $active_pane = 'topRight';
#NYI         push @{ $self->{_selections} }, [ 'topRight', $active_cell, $sqref ];
#NYI     }
#NYI     else {
#NYI         $active_pane = 'bottomLeft';
#NYI         push @{ $self->{_selections} }, [ 'bottomLeft', $active_cell, $sqref ];
#NYI     }
#NYI 
#NYI     # Set the pane type.
#NYI     if ( $type == 0 ) {
#NYI         $state = 'frozen';
#NYI     }
#NYI     elsif ( $type == 1 ) {
#NYI         $state = 'frozenSplit';
#NYI     }
#NYI     else {
#NYI         $state = 'split';
#NYI     }
#NYI 
#NYI 
#NYI     push @attributes, ( 'xSplit' => $x_split ) if $x_split;
#NYI     push @attributes, ( 'ySplit' => $y_split ) if $y_split;
#NYI 
#NYI     push @attributes, ( 'topLeftCell' => $top_left_cell );
#NYI     push @attributes, ( 'activePane'  => $active_pane );
#NYI     push @attributes, ( 'state'       => $state );
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'pane', @attributes );
#NYI }


##############################################################################
#
# _write_split_panes()
#
# Write the <pane> element for split panes.
#
# See also, implementers note for split_panes().
#
#NYI sub _write_split_panes {
#NYI 
#NYI     my $self = shift;
#NYI     my @attributes;
#NYI     my $y_split;
#NYI     my $x_split;
#NYI     my $has_selection = 0;
#NYI     my $active_pane;
#NYI     my $active_cell;
#NYI     my $sqref;
#NYI 
#NYI     my ( $row, $col, $top_row, $left_col, $type ) = @_;
#NYI     $y_split = $row;
#NYI     $x_split = $col;
#NYI 
#NYI     # Move user cell selection to the panes.
#NYI     if ( @{ $self->{_selections} } ) {
#NYI         ( undef, $active_cell, $sqref ) = @{ $self->{_selections}->[0] };
#NYI         $self->{_selections} = [];
#NYI         $has_selection = 1;
#NYI     }
#NYI 
#NYI     # Convert the row and col to 1/20 twip units with padding.
#NYI     $y_split = int( 20 * $y_split + 300 ) if $y_split;
#NYI     $x_split = $self->_calculate_x_split_width( $x_split ) if $x_split;
#NYI 
#NYI     # For non-explicit topLeft definitions, estimate the cell offset based
#NYI     # on the pixels dimensions. This is only a workaround and doesn't take
#NYI     # adjusted cell dimensions into account.
#NYI     if ( $top_row == $row && $left_col == $col ) {
#NYI         $top_row  = int( 0.5 + ( $y_split - 300 ) / 20 / 15 );
#NYI         $left_col = int( 0.5 + ( $x_split - 390 ) / 20 / 3 * 4 / 64 );
#NYI     }
#NYI 
#NYI     my $top_left_cell = xl-rowcol-to-cell( $top_row, $left_col );
#NYI 
#NYI     # If there is no selection set the active cell to the top left cell.
#NYI     if ( !$has_selection ) {
#NYI         $active_cell = $top_left_cell;
#NYI         $sqref       = $top_left_cell;
#NYI     }
#NYI 
#NYI     # Set the Cell selections.
#NYI     if ( $row && $col ) {
#NYI         $active_pane = 'bottomRight';
#NYI 
#NYI         my $row_cell = xl-rowcol-to-cell( $top_row, 0 );
#NYI         my $col_cell = xl-rowcol-to-cell( 0,        $left_col );
#NYI 
#NYI         push @{ $self->{_selections} },
#NYI           (
#NYI             [ 'topRight',    $col_cell,    $col_cell ],
#NYI             [ 'bottomLeft',  $row_cell,    $row_cell ],
#NYI             [ 'bottomRight', $active_cell, $sqref ]
#NYI           );
#NYI     }
#NYI     elsif ( $col ) {
#NYI         $active_pane = 'topRight';
#NYI         push @{ $self->{_selections} }, [ 'topRight', $active_cell, $sqref ];
#NYI     }
#NYI     else {
#NYI         $active_pane = 'bottomLeft';
#NYI         push @{ $self->{_selections} }, [ 'bottomLeft', $active_cell, $sqref ];
#NYI     }
#NYI 
#NYI     push @attributes, ( 'xSplit' => $x_split ) if $x_split;
#NYI     push @attributes, ( 'ySplit' => $y_split ) if $y_split;
#NYI     push @attributes, ( 'topLeftCell' => $top_left_cell );
#NYI     push @attributes, ( 'activePane' => $active_pane ) if $has_selection;
#NYI 
#NYI     $self->xml_empty_tag( 'pane', @attributes );
#NYI }


##############################################################################
#
# _calculate_x_split_width()
#
# Convert column width from user units to pane split width.
#
#NYI sub _calculate_x_split_width {
#NYI 
#NYI     my $self  = shift;
#NYI     my $width = shift;
#NYI 
#NYI     my $max_digit_width = 7;    # For Calabri 11.
#NYI     my $padding         = 5;
#NYI     my $pixels;
#NYI 
#NYI     # Convert to pixels.
#NYI     if ( $width < 1 ) {
#NYI         $pixels = int( $width * ( $max_digit_width + $padding ) + 0.5 );
#NYI     }
#NYI     else {
#NYI           $pixels = int( $width * $max_digit_width + 0.5 ) + $padding;
#NYI     }
#NYI 
#NYI     # Convert to points.
#NYI     my $points = $pixels * 3 / 4;
#NYI 
#NYI     # Convert to twips (twentieths of a point).
#NYI     my $twips = $points * 20;
#NYI 
#NYI     # Add offset/padding.
#NYI     $width = $twips + 390;
#NYI 
#NYI     return $width;
#NYI }


##############################################################################
#
# _write_tab_color()
#
# Write the <tabColor> element.
#
#NYI sub _write_tab_color {
#NYI 
#NYI     my $self        = shift;
#NYI     my $color_index = $self->{_tab_color};
#NYI 
#NYI     return unless $color_index;
#NYI 
#NYI     my $rgb = $self->_get_palette_color( $color_index );
#NYI 
#NYI     my @attributes = ( 'rgb' => $rgb );
#NYI 
#NYI     $self->xml_empty_tag( 'tabColor', @attributes );
#NYI }


##############################################################################
#
# _write_outline_pr()
#
# Write the <outlinePr> element.
#
#NYI sub _write_outline_pr {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     return unless $self->{_outline_changed};
#NYI 
#NYI     push @attributes, ( "applyStyles"        => 1 ) if $self->{_outline_style};
#NYI     push @attributes, ( "summaryBelow"       => 0 ) if !$self->{_outline_below};
#NYI     push @attributes, ( "summaryRight"       => 0 ) if !$self->{_outline_right};
#NYI     push @attributes, ( "showOutlineSymbols" => 0 ) if !$self->{_outline_on};
#NYI 
#NYI     $self->xml_empty_tag( 'outlinePr', @attributes );
#NYI }


##############################################################################
#
# _write_sheet_protection()
#
# Write the <sheetProtection> element.
#
#NYI sub _write_sheet_protection {
#NYI 
#NYI     my $self = shift;
#NYI     my @attributes;
#NYI 
#NYI     return unless $self->{_protect};
#NYI 
#NYI     my %arg = %{ $self->{_protect} };
#NYI 
#NYI     push @attributes, ( "password"    => $arg{password} ) if $arg{password};
#NYI     push @attributes, ( "sheet"       => 1 )              if $arg{sheet};
#NYI     push @attributes, ( "content"     => 1 )              if $arg{content};
#NYI     push @attributes, ( "objects"     => 1 )              if !$arg{objects};
#NYI     push @attributes, ( "scenarios"   => 1 )              if !$arg{scenarios};
#NYI     push @attributes, ( "formatCells" => 0 )              if $arg{format_cells};
#NYI     push @attributes, ( "formatColumns"    => 0 ) if $arg{format_columns};
#NYI     push @attributes, ( "formatRows"       => 0 ) if $arg{format_rows};
#NYI     push @attributes, ( "insertColumns"    => 0 ) if $arg{insert_columns};
#NYI     push @attributes, ( "insertRows"       => 0 ) if $arg{insert_rows};
#NYI     push @attributes, ( "insertHyperlinks" => 0 ) if $arg{insert_hyperlinks};
#NYI     push @attributes, ( "deleteColumns"    => 0 ) if $arg{delete_columns};
#NYI     push @attributes, ( "deleteRows"       => 0 ) if $arg{delete_rows};
#NYI 
#NYI     push @attributes, ( "selectLockedCells" => 1 )
#NYI       if !$arg{select_locked_cells};
#NYI 
#NYI     push @attributes, ( "sort"        => 0 ) if $arg{sort};
#NYI     push @attributes, ( "autoFilter"  => 0 ) if $arg{autofilter};
#NYI     push @attributes, ( "pivotTables" => 0 ) if $arg{pivot_tables};
#NYI 
#NYI     push @attributes, ( "selectUnlockedCells" => 1 )
#NYI       if !$arg{select_unlocked_cells};
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'sheetProtection', @attributes );
#NYI }


##############################################################################
#
# _write_drawings()
#
# Write the <drawing> elements.
#
#NYI sub _write_drawings {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return unless $self->{_drawing};
#NYI 
#NYI     $self->_write_drawing( ++$self->{_rel_count} );
#NYI }


##############################################################################
#
# _write_drawing()
#
# Write the <drawing> element.
#
#NYI sub _write_drawing {
#NYI 
#NYI     my $self = shift;
#NYI     my $id   = shift;
#NYI     my $r_id = 'rId' . $id;
#NYI 
#NYI     my @attributes = ( 'r:id' => $r_id );
#NYI 
#NYI     $self->xml_empty_tag( 'drawing', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_legacy_drawing()
#NYI #
#NYI # Write the <legacyDrawing> element.
#NYI #
#NYI sub _write_legacy_drawing {
#NYI 
#NYI     my $self = shift;
#NYI     my $id;
#NYI 
#NYI     return unless $self->{_has_vml};
#NYI 
#NYI     # Increment the relationship id for any drawings or comments.
#NYI     $id = ++$self->{_rel_count};
#NYI 
#NYI     my @attributes = ( 'r:id' => 'rId' . $id );
#NYI 
#NYI     $self->xml_empty_tag( 'legacyDrawing', @attributes );
#NYI }



##############################################################################
#
# _write_legacy_drawing_hf()
#
# Write the <legacyDrawingHF> element.
#
#NYI sub _write_legacy_drawing_hf {
#NYI 
#NYI     my $self = shift;
#NYI     my $id;
#NYI 
#NYI     return unless $self->{_has_header_vml};
#NYI 
#NYI     # Increment the relationship id for any drawings or comments.
#NYI     $id = ++$self->{_rel_count};
#NYI 
#NYI     my @attributes = ( 'r:id' => 'rId' . $id );
#NYI 
#NYI     $self->xml_empty_tag( 'legacyDrawingHF', @attributes );
#NYI }


#
# Note, the following font methods are, more or less, duplicated from the
# Excel::Writer::XLSX::Package::Styles class. I will look at implementing
# this is a cleaner encapsulated mode at a later stage.
#


##############################################################################
#
# _write_font()
#
# Write the <font> element.
#
#NYI sub _write_font {
#NYI 
#NYI     my $self   = shift;
#NYI     my $format = shift;
#NYI 
#NYI     $self->{_rstring}->xml_start_tag( 'rPr' );
#NYI 
#NYI     $self->{_rstring}->xml_empty_tag( 'b' )       if $format->{_bold};
#NYI     $self->{_rstring}->xml_empty_tag( 'i' )       if $format->{_italic};
#NYI     $self->{_rstring}->xml_empty_tag( 'strike' )  if $format->{_font_strikeout};
#NYI     $self->{_rstring}->xml_empty_tag( 'outline' ) if $format->{_font_outline};
#NYI     $self->{_rstring}->xml_empty_tag( 'shadow' )  if $format->{_font_shadow};
#NYI 
#NYI     # Handle the underline variants.
#NYI     $self->_write_underline( $format->{_underline} ) if $format->{_underline};
#NYI 
#NYI     $self->_write_vert_align( 'superscript' ) if $format->{_font_script} == 1;
#NYI     $self->_write_vert_align( 'subscript' )   if $format->{_font_script} == 2;
#NYI 
#NYI     $self->{_rstring}->xml_empty_tag( 'sz', 'val', $format->{_size} );
#NYI 
#NYI     if ( my $theme = $format->{_theme} ) {
#NYI         $self->_write_rstring_color( 'theme' => $theme );
#NYI     }
#NYI     elsif ( my $color = $format->{_color} ) {
#NYI         $color = $self->_get_palette_color( $color );
#NYI 
#NYI         $self->_write_rstring_color( 'rgb' => $color );
#NYI     }
#NYI     else {
#NYI         $self->_write_rstring_color( 'theme' => 1 );
#NYI     }
#NYI 
#NYI     $self->{_rstring}->xml_empty_tag( 'rFont', 'val', $format->{_font} );
#NYI     $self->{_rstring}
#NYI       ->xml_empty_tag( 'family', 'val', $format->{_font_family} );
#NYI 
#NYI     if ( $format->{_font} eq 'Calibri' && !$format->{_hyperlink} ) {
#NYI         $self->{_rstring}
#NYI           ->xml_empty_tag( 'scheme', 'val', $format->{_font_scheme} );
#NYI     }
#NYI 
#NYI     $self->{_rstring}->xml_end_tag( 'rPr' );
#NYI }


###############################################################################
#
# _write_underline()
#
# Write the underline font element.
#
#NYI sub _write_underline {
#NYI 
#NYI     my $self      = shift;
#NYI     my $underline = shift;
#NYI     my @attributes;
#NYI 
#NYI     # Handle the underline variants.
#NYI     if ( $underline == 2 ) {
#NYI         @attributes = ( val => 'double' );
#NYI     }
#NYI     elsif ( $underline == 33 ) {
#NYI         @attributes = ( val => 'singleAccounting' );
#NYI     }
#NYI     elsif ( $underline == 34 ) {
#NYI         @attributes = ( val => 'doubleAccounting' );
#NYI     }
#NYI     else {
#NYI         @attributes = ();    # Default to single underline.
#NYI     }
#NYI 
#NYI     $self->{_rstring}->xml_empty_tag( 'u', @attributes );
#NYI 
#NYI }


##############################################################################
#
# _write_vert_align()
#
# Write the <vertAlign> font sub-element.
#
#NYI sub _write_vert_align {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->{_rstring}->xml_empty_tag( 'vertAlign', @attributes );
#NYI }


##############################################################################
#
# _write_rstring_color()
#
# Write the <color> element.
#
#NYI sub _write_rstring_color {
#NYI 
#NYI     my $self  = shift;
#NYI     my $name  = shift;
#NYI     my $value = shift;
#NYI 
#NYI     my @attributes = ( $name => $value );
#NYI 
#NYI     $self->{_rstring}->xml_empty_tag( 'color', @attributes );
#NYI }


#
# End font duplication code.
#


##############################################################################
#
# _write_data_validations()
#
# Write the <dataValidations> element.
#
#NYI sub _write_data_validations {
#NYI 
#NYI     my $self        = shift;
#NYI     my @validations = @{ $self->{_validations} };
#NYI     my $count       = @validations;
#NYI 
#NYI     return unless $count;
#NYI 
#NYI     my @attributes = ( 'count' => $count );
#NYI 
#NYI     $self->xml_start_tag( 'dataValidations', @attributes );
#NYI 
#NYI     for my $validation ( @validations ) {
#NYI 
#NYI         # Write the dataValidation element.
#NYI         $self->_write_data_validation( $validation );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'dataValidations' );
#NYI }


##############################################################################
#
# _write_data_validation()
#
# Write the <dataValidation> element.
#
#NYI sub _write_data_validation {
#NYI 
#NYI     my $self       = shift;
#NYI     my $param      = shift;
#NYI     my $sqref      = '';
#NYI     my @attributes = ();
#NYI 
#NYI 
#NYI     # Set the cell range(s) for the data validation.
#NYI     for my $cells ( @{ $param->{cells} } ) {
#NYI 
#NYI         # Add a space between multiple cell ranges.
#NYI         $sqref .= ' ' if $sqref ne '';
#NYI 
#NYI         my ( $row_first, $col_first, $row_last, $col_last ) = @$cells;
#NYI 
#NYI         # Swap last row/col for first row/col as necessary
#NYI         if ( $row_first > $row_last ) {
#NYI             ( $row_first, $row_last ) = ( $row_last, $row_first );
#NYI         }
#NYI 
#NYI         if ( $col_first > $col_last ) {
#NYI             ( $col_first, $col_last ) = ( $col_last, $col_first );
#NYI         }
#NYI 
#NYI         # If the first and last cell are the same write a single cell.
#NYI         if ( ( $row_first == $row_last ) && ( $col_first == $col_last ) ) {
#NYI             $sqref .= xl-rowcol-to-cell( $row_first, $col_first );
#NYI         }
#NYI         else {
#NYI             $sqref .= xl-range( $row_first, $row_last, $col_first, $col_last );
#NYI         }
#NYI     }


#NYI     if ( $param->{validate} ne 'none' ) {
#NYI 
#NYI         push @attributes, ( 'type' => $param->{validate} );
#NYI 
#NYI         if ( $param->{criteria} ne 'between' ) {
#NYI             push @attributes, ( 'operator' => $param->{criteria} );
#NYI         }
#NYI 
#NYI     }
#NYI 
#NYI     if ( $param->{error_type} ) {
#NYI         push @attributes, ( 'errorStyle' => 'warning' )
#NYI           if $param->{error_type} == 1;
#NYI         push @attributes, ( 'errorStyle' => 'information' )
#NYI           if $param->{error_type} == 2;
#NYI     }
#NYI 
#NYI     push @attributes, ( 'allowBlank'       => 1 ) if $param->{ignore_blank};
#NYI     push @attributes, ( 'showDropDown'     => 1 ) if !$param->{dropdown};
#NYI     push @attributes, ( 'showInputMessage' => 1 ) if $param->{show_input};
#NYI     push @attributes, ( 'showErrorMessage' => 1 ) if $param->{show_error};
#NYI 
#NYI     push @attributes, ( 'errorTitle' => $param->{error_title} )
#NYI       if $param->{error_title};
#NYI 
#NYI     push @attributes, ( 'error' => $param->{error_message} )
#NYI       if $param->{error_message};
#NYI 
#NYI     push @attributes, ( 'promptTitle' => $param->{input_title} )
#NYI       if $param->{input_title};
#NYI 
#NYI     push @attributes, ( 'prompt' => $param->{input_message} )
#NYI       if $param->{input_message};
#NYI 
#NYI     push @attributes, ( 'sqref' => $sqref );
#NYI 
#NYI     if ( $param->{validate} eq 'none' ) {
#NYI         $self->xml_empty_tag( 'dataValidation', @attributes );
#NYI     }
#NYI     else {
#NYI         $self->xml_start_tag( 'dataValidation', @attributes );
#NYI 
#NYI         # Write the formula1 element.
#NYI         $self->_write_formula_1( $param->{value} );
#NYI 
#NYI         # Write the formula2 element.
#NYI         $self->_write_formula_2( $param->{maximum} )
#NYI           if defined $param->{maximum};
#NYI 
#NYI         $self->xml_end_tag( 'dataValidation' );
#NYI     }
#NYI }


##############################################################################
#
# _write_formula_1()
#
# Write the <formula1> element.
#
#NYI sub _write_formula_1 {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI 
#NYI     # Convert a list array ref into a comma separated string.
#NYI     if ( ref $formula eq 'ARRAY' ) {
#NYI         $formula = join ',', @$formula;
#NYI         $formula = qq("$formula");
#NYI     }
#NYI 
#NYI     $formula =~ s/^=//;    # Remove formula symbol.
#NYI 
#NYI     $self->xml_data_element( 'formula1', $formula );
#NYI }


##############################################################################
#
# _write_formula_2()
#
# Write the <formula2> element.
#
#NYI sub _write_formula_2 {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI 
#NYI     $formula =~ s/^=//;    # Remove formula symbol.
#NYI 
#NYI     $self->xml_data_element( 'formula2', $formula );
#NYI }


##############################################################################
#
# _write_conditional_formats()
#
# Write the Worksheet conditional formats.
#
#NYI sub _write_conditional_formats {
#NYI 
#NYI     my $self   = shift;
#NYI     my @ranges = sort keys %{ $self->{_cond_formats} };
#NYI 
#NYI     return unless scalar @ranges;
#NYI 
#NYI     for my $range ( @ranges ) {
#NYI         $self->_write_conditional_formatting( $range,
#NYI             $self->{_cond_formats}->{$range} );
#NYI     }
#NYI }


##############################################################################
#
# _write_conditional_formatting()
#
# Write the <conditionalFormatting> element.
#
#NYI sub _write_conditional_formatting {
#NYI 
#NYI     my $self   = shift;
#NYI     my $range  = shift;
#NYI     my $params = shift;
#NYI 
#NYI     my @attributes = ( 'sqref' => $range );
#NYI 
#NYI     $self->xml_start_tag( 'conditionalFormatting', @attributes );
#NYI 
#NYI     for my $param ( @$params ) {
#NYI 
#NYI         # Write the cfRule element.
#NYI         $self->_write_cf_rule( $param );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'conditionalFormatting' );
#NYI }

##############################################################################
#
# _write_cf_rule()
#
# Write the <cfRule> element.
#
#NYI sub _write_cf_rule {
#NYI 
#NYI     my $self  = shift;
#NYI     my $param = shift;
#NYI 
#NYI     my @attributes = ( 'type' => $param->{type} );
#NYI 
#NYI     push @attributes, ( 'dxfId' => $param->{format} )
#NYI       if defined $param->{format};
#NYI 
#NYI     push @attributes, ( 'priority' => $param->{priority} );
#NYI 
#NYI     push @attributes, ( 'stopIfTrue' => 1 )
#NYI       if $param->{stop_if_true};
#NYI 
#NYI     if ( $param->{type} eq 'cellIs' ) {
#NYI         push @attributes, ( 'operator' => $param->{criteria} );
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI 
#NYI         if ( defined $param->{minimum} && defined $param->{maximum} ) {
#NYI             $self->_write_formula( $param->{minimum} );
#NYI             $self->_write_formula( $param->{maximum} );
#NYI         }
#NYI         else {
#NYI             $self->_write_formula( $param->{value} );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ( $param->{type} eq 'aboveAverage' ) {
#NYI         if ( $param->{criteria} =~ /below/ ) {
#NYI             push @attributes, ( 'aboveAverage' => 0 );
#NYI         }
#NYI 
#NYI         if ( $param->{criteria} =~ /equal/ ) {
#NYI             push @attributes, ( 'equalAverage' => 1 );
#NYI         }
#NYI 
#NYI         if ( $param->{criteria} =~ /([123]) std dev/ ) {
#NYI             push @attributes, ( 'stdDev' => $1 );
#NYI         }
#NYI 
#NYI         $self->xml_empty_tag( 'cfRule', @attributes );
#NYI     }
#NYI     elsif ( $param->{type} eq 'top10' ) {
#NYI         if ( defined $param->{criteria} && $param->{criteria} eq '%' ) {
#NYI             push @attributes, ( 'percent' => 1 );
#NYI         }
#NYI 
#NYI         if ( $param->{direction} ) {
#NYI             push @attributes, ( 'bottom' => 1 );
#NYI         }
#NYI 
#NYI         my $rank = $param->{value} || 10;
#NYI         push @attributes, ( 'rank' => $rank );
#NYI 
#NYI         $self->xml_empty_tag( 'cfRule', @attributes );
#NYI     }
#NYI     elsif ( $param->{type} eq 'duplicateValues' ) {
#NYI         $self->xml_empty_tag( 'cfRule', @attributes );
#NYI     }
#NYI     elsif ( $param->{type} eq 'uniqueValues' ) {
#NYI         $self->xml_empty_tag( 'cfRule', @attributes );
#NYI     }
#NYI     elsif ($param->{type} eq 'containsText'
#NYI         || $param->{type} eq 'notContainsText'
#NYI         || $param->{type} eq 'beginsWith'
#NYI         || $param->{type} eq 'endsWith' )
#NYI     {
#NYI         push @attributes, ( 'operator' => $param->{criteria} );
#NYI         push @attributes, ( 'text'     => $param->{value} );
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_formula( $param->{formula} );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ( $param->{type} eq 'timePeriod' ) {
#NYI         push @attributes, ( 'timePeriod' => $param->{criteria} );
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_formula( $param->{formula} );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ($param->{type} eq 'containsBlanks'
#NYI         || $param->{type} eq 'notContainsBlanks'
#NYI         || $param->{type} eq 'containsErrors'
#NYI         || $param->{type} eq 'notContainsErrors' )
#NYI     {
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_formula( $param->{formula} );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ( $param->{type} eq 'colorScale' ) {
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_color_scale( $param );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ( $param->{type} eq 'dataBar' ) {
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_data_bar( $param );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ( $param->{type} eq 'expression' ) {
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_formula( $param->{criteria} );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI     elsif ( $param->{type} eq 'iconSet' ) {
#NYI 
#NYI         $self->xml_start_tag( 'cfRule', @attributes );
#NYI         $self->_write_icon_set( $param );
#NYI         $self->xml_end_tag( 'cfRule' );
#NYI     }
#NYI }


##############################################################################
#
# _write_icon_set()
#
# Write the <iconSet> element.
#
#NYI sub _write_icon_set {
#NYI 
#NYI     my $self        = shift;
#NYI     my $param       = shift;
#NYI     my $icon_style  = $param->{icon_style};
#NYI     my $total_icons = $param->{total_icons};
#NYI     my $icons       = $param->{icons};
#NYI     my $i;
#NYI 
#NYI     my @attributes = ();
#NYI 
#NYI     # Don't set attribute for default style.
#NYI     if ( $icon_style ne '3TrafficLights' ) {
#NYI         @attributes = ( 'iconSet' => $icon_style );
#NYI     }
#NYI 
#NYI     if ( exists $param->{'icons_only'} && $param->{'icons_only'} ) {
#NYI         push @attributes, ( 'showValue' => 0 );
#NYI     }
#NYI 
#NYI     if ( exists $param->{'reverse_icons'} && $param->{'reverse_icons'} ) {
#NYI         push @attributes, ( 'reverse' => 1 );
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'iconSet', @attributes );
#NYI 
#NYI     # Write the properites for different icon styles.
#NYI     for my $icon ( reverse @{ $param->{icons} } ) {
#NYI         $self->_write_cfvo(
#NYI             $icon->{'type'},
#NYI             $icon->{'value'},
#NYI             $icon->{'criteria'}
#NYI         );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'iconSet' );
#NYI }

##############################################################################
#
# _write_formula()
#
# Write the <formula> element.
#
#NYI sub _write_formula {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     # Remove equality from formula.
#NYI     $data =~ s/^=//;
#NYI 
#NYI     $self->xml_data_element( 'formula', $data );
#NYI }


##############################################################################
#
# _write_color_scale()
#
# Write the <colorScale> element.
#
#NYI sub _write_color_scale {
#NYI 
#NYI     my $self  = shift;
#NYI     my $param = shift;
#NYI 
#NYI     $self->xml_start_tag( 'colorScale' );
#NYI 
#NYI     $self->_write_cfvo( $param->{min_type}, $param->{min_value} );
#NYI 
#NYI     if ( defined $param->{mid_type} ) {
#NYI         $self->_write_cfvo( $param->{mid_type}, $param->{mid_value} );
#NYI     }
#NYI 
#NYI     $self->_write_cfvo( $param->{max_type}, $param->{max_value} );
#NYI 
#NYI     $self->_write_color( 'rgb' => $param->{min_color} );
#NYI 
#NYI     if ( defined $param->{mid_color} ) {
#NYI         $self->_write_color( 'rgb' => $param->{mid_color} );
#NYI     }
#NYI 
#NYI     $self->_write_color( 'rgb' => $param->{max_color} );
#NYI 
#NYI     $self->xml_end_tag( 'colorScale' );
#NYI }


##############################################################################
#
# _write_data_bar()
#
# Write the <dataBar> element.
#
#NYI sub _write_data_bar {
#NYI 
#NYI     my $self  = shift;
#NYI     my $param = shift;
#NYI 
#NYI     $self->xml_start_tag( 'dataBar' );
#NYI 
#NYI     $self->_write_cfvo( $param->{min_type}, $param->{min_value} );
#NYI     $self->_write_cfvo( $param->{max_type}, $param->{max_value} );
#NYI 
#NYI     $self->_write_color( 'rgb' => $param->{bar_color} );
#NYI 
#NYI     $self->xml_end_tag( 'dataBar' );
#NYI }


##############################################################################
#
# _write_cfvo()
#
# Write the <cfvo> element.
#
#NYI sub _write_cfvo {
#NYI 
#NYI     my $self     = shift;
#NYI     my $type     = shift;
#NYI     my $value    = shift;
#NYI     my $criteria = shift;
#NYI 
#NYI     my @attributes = (
#NYI         'type' => $type,
#NYI         'val'  => $value
#NYI     );
#NYI 
#NYI     if ( $criteria ) {
#NYI         push @attributes, ( 'gte', 0 );
#NYI     }
#NYI 
#NYI     $self->xml_empty_tag( 'cfvo', @attributes );
#NYI }


##############################################################################
#
# _write_color()
#
# Write the <color> element.
#
#NYI sub _write_color {
#NYI 
#NYI     my $self  = shift;
#NYI     my $name  = shift;
#NYI     my $value = shift;
#NYI 
#NYI     my @attributes = ( $name => $value );
#NYI 
#NYI     $self->xml_empty_tag( 'color', @attributes );
#NYI }


##############################################################################
#
# _write_table_parts()
#
# Write the <tableParts> element.
#
#NYI sub _write_table_parts {
#NYI 
#NYI     my $self   = shift;
#NYI     my @tables = @{ $self->{_tables} };
#NYI     my $count  = scalar @tables;
#NYI 
#NYI     # Return if worksheet doesn't contain any tables.
#NYI     return unless $count;
#NYI 
#NYI     my @attributes = ( 'count' => $count, );
#NYI 
#NYI     $self->xml_start_tag( 'tableParts', @attributes );
#NYI 
#NYI     for my $table ( @tables ) {
#NYI 
#NYI         # Write the tablePart element.
#NYI         $self->_write_table_part( ++$self->{_rel_count} );
#NYI 
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'tableParts' );
#NYI }


##############################################################################
#
# _write_table_part()
#
# Write the <tablePart> element.
#
#NYI sub _write_table_part {
#NYI 
#NYI     my $self = shift;
#NYI     my $id   = shift;
#NYI     my $r_id = 'rId' . $id;
#NYI 
#NYI     my @attributes = ( 'r:id' => $r_id, );
#NYI 
#NYI     $self->xml_empty_tag( 'tablePart', @attributes );
#NYI }


##############################################################################
#
# _write_ext_sparklines()
#
# Write the <extLst> element and sparkline subelements.
#
#NYI sub _write_ext_sparklines {
#NYI 
#NYI     my $self       = shift;
#NYI     my @sparklines = @{ $self->{_sparklines} };
#NYI     my $count      = scalar @sparklines;
#NYI 
#NYI     # Return if worksheet doesn't contain any sparklines.
#NYI     return unless $count;
#NYI 
#NYI 
#NYI     # Write the extLst element.
#NYI     $self->xml_start_tag( 'extLst' );
#NYI 
#NYI     # Write the ext element.
#NYI     $self->_write_ext();
#NYI 
#NYI     # Write the x14:sparklineGroups element.
#NYI     $self->_write_sparkline_groups();
#NYI 
#NYI     # Write the sparkline elements.
#NYI     for my $sparkline ( reverse @sparklines ) {
#NYI 
#NYI         # Write the x14:sparklineGroup element.
#NYI         $self->_write_sparkline_group( $sparkline );
#NYI 
#NYI         # Write the x14:colorSeries element.
#NYI         $self->_write_color_series( $sparkline->{_series_color} );
#NYI 
#NYI         # Write the x14:colorNegative element.
#NYI         $self->_write_color_negative( $sparkline->{_negative_color} );
#NYI 
#NYI         # Write the x14:colorAxis element.
#NYI         $self->_write_color_axis();
#NYI 
#NYI         # Write the x14:colorMarkers element.
#NYI         $self->_write_color_markers( $sparkline->{_markers_color} );
#NYI 
#NYI         # Write the x14:colorFirst element.
#NYI         $self->_write_color_first( $sparkline->{_first_color} );
#NYI 
#NYI         # Write the x14:colorLast element.
#NYI         $self->_write_color_last( $sparkline->{_last_color} );
#NYI 
#NYI         # Write the x14:colorHigh element.
#NYI         $self->_write_color_high( $sparkline->{_high_color} );
#NYI 
#NYI         # Write the x14:colorLow element.
#NYI         $self->_write_color_low( $sparkline->{_low_color} );
#NYI 
#NYI         if ( $sparkline->{_date_axis} ) {
#NYI             $self->xml_data_element( 'xm:f', $sparkline->{_date_axis} );
#NYI         }
#NYI 
#NYI         $self->_write_sparklines( $sparkline );
#NYI 
#NYI         $self->xml_end_tag( 'x14:sparklineGroup' );
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_end_tag( 'x14:sparklineGroups' );
#NYI     $self->xml_end_tag( 'ext' );
#NYI     $self->xml_end_tag( 'extLst' );
#NYI }


##############################################################################
#
# _write_sparklines()
#
# Write the <x14:sparklines> element and <x14:sparkline> subelements.
#
#NYI sub _write_sparklines {
#NYI 
#NYI     my $self      = shift;
#NYI     my $sparkline = shift;
#NYI 
#NYI     # Write the sparkline elements.
#NYI     $self->xml_start_tag( 'x14:sparklines' );
#NYI 
#NYI     for my $i ( 0 .. $sparkline->{_count} - 1 ) {
#NYI         my $range    = $sparkline->{_ranges}->[$i];
#NYI         my $location = $sparkline->{_locations}->[$i];
#NYI 
#NYI         $self->xml_start_tag( 'x14:sparkline' );
#NYI         $self->xml_data_element( 'xm:f',     $range );
#NYI         $self->xml_data_element( 'xm:sqref', $location );
#NYI         $self->xml_end_tag( 'x14:sparkline' );
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_end_tag( 'x14:sparklines' );
#NYI }


##############################################################################
#
# _write_ext()
#
# Write the <ext> element.
#
#NYI sub _write_ext {
#NYI 
#NYI     my $self       = shift;
#NYI     my $schema     = 'http://schemas.microsoft.com/office/';
#NYI     my $xmlns_x_14 = $schema . 'spreadsheetml/2009/9/main';
#NYI     my $uri        = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns:x14' => $xmlns_x_14,
#NYI         'uri'       => $uri,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'ext', @attributes );
#NYI }


##############################################################################
#
# _write_sparkline_groups()
#
# Write the <x14:sparklineGroups> element.
#
#NYI sub _write_sparkline_groups {
#NYI 
#NYI     my $self     = shift;
#NYI     my $xmlns_xm = 'http://schemas.microsoft.com/office/excel/2006/main';
#NYI 
#NYI     my @attributes = ( 'xmlns:xm' => $xmlns_xm );
#NYI 
#NYI     $self->xml_start_tag( 'x14:sparklineGroups', @attributes );
#NYI 
#NYI }


##############################################################################
#
# _write_sparkline_group()
#
# Write the <x14:sparklineGroup> element.
#
# Example for order.
#
# <x14:sparklineGroup
#     manualMax="0"
#     manualMin="0"
#     lineWeight="2.25"
#     type="column"
#     dateAxis="1"
#     displayEmptyCellsAs="span"
#     markers="1"
#     high="1"
#     low="1"
#     first="1"
#     last="1"
#     negative="1"
#     displayXAxis="1"
#     displayHidden="1"
#     minAxisType="custom"
#     maxAxisType="custom"
#     rightToLeft="1">
#
#NYI sub _write_sparkline_group {
#NYI 
#NYI     my $self     = shift;
#NYI     my $opts     = shift;
#NYI     my $empty    = $opts->{_empty};
#NYI     my $user_max = 0;
#NYI     my $user_min = 0;
#NYI     my @a;
#NYI 
#NYI     if ( defined $opts->{_max} ) {
#NYI 
#NYI         if ( $opts->{_max} eq 'group' ) {
#NYI             $opts->{_cust_max} = 'group';
#NYI         }
#NYI         else {
#NYI             push @a, ( 'manualMax' => $opts->{_max} );
#NYI             $opts->{_cust_max} = 'custom';
#NYI         }
#NYI     }
#NYI 
#NYI     if ( defined $opts->{_min} ) {
#NYI 
#NYI         if ( $opts->{_min} eq 'group' ) {
#NYI             $opts->{_cust_min} = 'group';
#NYI         }
#NYI         else {
#NYI             push @a, ( 'manualMin' => $opts->{_min} );
#NYI             $opts->{_cust_min} = 'custom';
#NYI         }
#NYI     }
#NYI 
#NYI 
#NYI     # Ignore the default type attribute (line).
#NYI     if ( $opts->{_type} ne 'line' ) {
#NYI         push @a, ( 'type' => $opts->{_type} );
#NYI     }
#NYI 
#NYI     push @a, ( 'lineWeight' => $opts->{_weight} ) if $opts->{_weight};
#NYI     push @a, ( 'dateAxis' => 1 ) if $opts->{_date_axis};
#NYI     push @a, ( 'displayEmptyCellsAs' => $empty ) if $empty;
#NYI 
#NYI     push @a, ( 'markers'       => 1 )                  if $opts->{_markers};
#NYI     push @a, ( 'high'          => 1 )                  if $opts->{_high};
#NYI     push @a, ( 'low'           => 1 )                  if $opts->{_low};
#NYI     push @a, ( 'first'         => 1 )                  if $opts->{_first};
#NYI     push @a, ( 'last'          => 1 )                  if $opts->{_last};
#NYI     push @a, ( 'negative'      => 1 )                  if $opts->{_negative};
#NYI     push @a, ( 'displayXAxis'  => 1 )                  if $opts->{_axis};
#NYI     push @a, ( 'displayHidden' => 1 )                  if $opts->{_hidden};
#NYI     push @a, ( 'minAxisType'   => $opts->{_cust_min} ) if $opts->{_cust_min};
#NYI     push @a, ( 'maxAxisType'   => $opts->{_cust_max} ) if $opts->{_cust_max};
#NYI     push @a, ( 'rightToLeft'   => 1 )                  if $opts->{_reverse};
#NYI 
#NYI     $self->xml_start_tag( 'x14:sparklineGroup', @a );
#NYI }


##############################################################################
#
# _write_spark_color()
#
# Helper function for the sparkline color functions below.
#
#NYI sub _write_spark_color {
#NYI 
#NYI     my $self    = shift;
#NYI     my $element = shift;
#NYI     my $color   = shift;
#NYI     my @attr;
#NYI 
#NYI     push @attr, ( 'rgb'   => $color->{_rgb} )   if defined $color->{_rgb};
#NYI     push @attr, ( 'theme' => $color->{_theme} ) if defined $color->{_theme};
#NYI     push @attr, ( 'tint'  => $color->{_tint} )  if defined $color->{_tint};
#NYI 
#NYI     $self->xml_empty_tag( $element, @attr );
#NYI }


##############################################################################
#
# _write_color_series()
#
# Write the <x14:colorSeries> element.
#
#NYI sub _write_color_series {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorSeries', @_ );
#NYI }


##############################################################################
#
# _write_color_negative()
#
# Write the <x14:colorNegative> element.
#
#NYI sub _write_color_negative {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorNegative', @_ );
#NYI }


##############################################################################
#
# _write_color_axis()
#
# Write the <x14:colorAxis> element.
#
#NYI sub _write_color_axis {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorAxis', { _rgb => 'FF000000' } );
#NYI }


##############################################################################
#
# _write_color_markers()
#
# Write the <x14:colorMarkers> element.
#
#NYI sub _write_color_markers {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorMarkers', @_ );
#NYI }


##############################################################################
#
# _write_color_first()
#
# Write the <x14:colorFirst> element.
#
#NYI sub _write_color_first {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorFirst', @_ );
#NYI }


##############################################################################
#
# _write_color_last()
#
# Write the <x14:colorLast> element.
#
#NYI sub _write_color_last {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorLast', @_ );
#NYI }


##############################################################################
#
# _write_color_high()
#
# Write the <x14:colorHigh> element.
#
#NYI sub _write_color_high {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorHigh', @_ );
#NYI }


##############################################################################
#
# _write_color_low()
#
# Write the <x14:colorLow> element.
#
#NYI sub _write_color_low {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->_write_spark_color( 'x14:colorLow', @_ );
#NYI }

=begin pod

=head1 NAME

Worksheet - A class for writing Excel Worksheets.

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
