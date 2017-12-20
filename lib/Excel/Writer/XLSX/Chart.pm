#NYI package Excel::Writer::XLSX::Chart;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Chart - A class for writing Excel Charts.
#NYI #
#NYI #
#NYI # Used in conjunction with Excel::Writer::XLSX.
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
#NYI use Excel::Writer::XLSX::Format;
#NYI use Excel::Writer::XLSX::Package::XMLwriter;
#NYI use Excel::Writer::XLSX::Utility qw(xl_cell_to_rowcol
#NYI   xl_rowcol_to_cell
#NYI   xl_col_to_name xl_range
#NYI   xl_range_formula
#NYI   quote_sheetname );
#NYI 
#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # factory()
#NYI #
#NYI # Factory method for returning chart objects based on their class type.
#NYI #
#NYI sub factory {
#NYI 
#NYI     my $current_class  = shift;
#NYI     my $chart_subclass = shift;
#NYI 
#NYI     $chart_subclass = ucfirst lc $chart_subclass;
#NYI 
#NYI     my $module = "Excel::Writer::XLSX::Chart::" . $chart_subclass;
#NYI 
#NYI     eval "require $module";
#NYI 
#NYI     # TODO. Need to re-raise this error from Workbook::add_chart().
#NYI     die "Chart type '$chart_subclass' not supported in add_chart()\n" if $@;
#NYI 
#NYI     my $fh = undef;
#NYI     return $module->new( $fh, @_ );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # new()
#NYI #
#NYI # Default constructor for sub-classes.
#NYI #
#NYI sub new {
#NYI 
#NYI     my $class = shift;
#NYI     my $fh    = shift;
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );
#NYI 
#NYI     $self->{_subtype}           = shift;
#NYI     $self->{_sheet_type}        = 0x0200;
#NYI     $self->{_orientation}       = 0x0;
#NYI     $self->{_series}            = [];
#NYI     $self->{_embedded}          = 0;
#NYI     $self->{_id}                = -1;
#NYI     $self->{_series_index}      = 0;
#NYI     $self->{_style_id}          = 2;
#NYI     $self->{_axis_ids}          = [];
#NYI     $self->{_axis2_ids}         = [];
#NYI     $self->{_cat_has_num_fmt}   = 0;
#NYI     $self->{_requires_category} = 0;
#NYI     $self->{_legend_position}   = 'right';
#NYI     $self->{_cat_axis_position} = 'b';
#NYI     $self->{_val_axis_position} = 'l';
#NYI     $self->{_formula_ids}       = {};
#NYI     $self->{_formula_data}      = [];
#NYI     $self->{_horiz_cat_axis}    = 0;
#NYI     $self->{_horiz_val_axis}    = 1;
#NYI     $self->{_protection}        = 0;
#NYI     $self->{_chartarea}         = {};
#NYI     $self->{_plotarea}          = {};
#NYI     $self->{_x_axis}            = {};
#NYI     $self->{_y_axis}            = {};
#NYI     $self->{_y2_axis}           = {};
#NYI     $self->{_x2_axis}           = {};
#NYI     $self->{_chart_name}        = '';
#NYI     $self->{_show_blanks}       = 'gap';
#NYI     $self->{_show_hidden_data}  = 0;
#NYI     $self->{_show_crosses}      = 1;
#NYI     $self->{_width}             = 480;
#NYI     $self->{_height}            = 288;
#NYI     $self->{_x_scale}           = 1;
#NYI     $self->{_y_scale}           = 1;
#NYI     $self->{_x_offset}          = 0;
#NYI     $self->{_y_offset}          = 0;
#NYI     $self->{_table}             = undef;
#NYI     $self->{_smooth_allowed}    = 0;
#NYI     $self->{_cross_between}     = 'between';
#NYI     $self->{_date_category}     = 0;
#NYI     $self->{_already_inserted}  = 0;
#NYI     $self->{_combined}          = undef;
#NYI     $self->{_is_secondary}      = 0;
#NYI 
#NYI     $self->{_label_positions}          = {};
#NYI     $self->{_label_position_default}   = '';
#NYI 
#NYI     bless $self, $class;
#NYI     $self->_set_default_properties();
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
#NYI     $self->xml_declaration();
#NYI 
#NYI     # Write the c:chartSpace element.
#NYI     $self->_write_chart_space();
#NYI 
#NYI     # Write the c:lang element.
#NYI     $self->_write_lang();
#NYI 
#NYI     # Write the c:style element.
#NYI     $self->_write_style();
#NYI 
#NYI     # Write the c:protection element.
#NYI     $self->_write_protection();
#NYI 
#NYI     # Write the c:chart element.
#NYI     $self->_write_chart();
#NYI 
#NYI     # Write the c:spPr element for the chartarea formatting.
#NYI     $self->_write_sp_pr( $self->{_chartarea} );
#NYI 
#NYI     # Write the c:printSettings element.
#NYI     $self->_write_print_settings() if $self->{_embedded};
#NYI 
#NYI     # Close the worksheet tag.
#NYI     $self->xml_end_tag( 'c:chartSpace' );
#NYI 
#NYI     # Close the XML writer filehandle.
#NYI     $self->xml_get_fh()->close();
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Public methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # add_series()
#NYI #
#NYI # Add a series and it's properties to a chart.
#NYI #
#NYI sub add_series {
#NYI 
#NYI     my $self = shift;
#NYI     my %arg  = @_;
#NYI 
#NYI     # Check that the required input has been specified.
#NYI     if ( !exists $arg{values} ) {
#NYI         croak "Must specify 'values' in add_series()";
#NYI     }
#NYI 
#NYI     if ( $self->{_requires_category} && !exists $arg{categories} ) {
#NYI         croak "Must specify 'categories' in add_series() for this chart type";
#NYI     }
#NYI 
#NYI     if ( @{ $self->{_series} } == 255 ) {
#NYI         carp "The maxiumn number of series that can be added to an "
#NYI           . "Excel Chart is 255";
#NYI         return
#NYI     }
#NYI 
#NYI     # Convert aref params into a formula string.
#NYI     my $values     = $self->_aref_to_formula( $arg{values} );
#NYI     my $categories = $self->_aref_to_formula( $arg{categories} );
#NYI 
#NYI     # Switch name and name_formula parameters if required.
#NYI     my ( $name, $name_formula ) =
#NYI       $self->_process_names( $arg{name}, $arg{name_formula} );
#NYI 
#NYI     # Get an id for the data equivalent to the range formula.
#NYI     my $cat_id  = $self->_get_data_id( $categories,   $arg{categories_data} );
#NYI     my $val_id  = $self->_get_data_id( $values,       $arg{values_data} );
#NYI     my $name_id = $self->_get_data_id( $name_formula, $arg{name_data} );
#NYI 
#NYI     # Set the line properties for the series.
#NYI     my $line = $self->_get_line_properties( $arg{line} );
#NYI 
#NYI     # Allow 'border' as a synonym for 'line' in bar/column style charts.
#NYI     if ( $arg{border} ) {
#NYI         $line = $self->_get_line_properties( $arg{border} );
#NYI     }
#NYI 
#NYI     # Set the fill properties for the series.
#NYI     my $fill = $self->_get_fill_properties( $arg{fill} );
#NYI 
#NYI     # Set the pattern properties for the series.
#NYI     my $pattern = $self->_get_pattern_properties( $arg{pattern} );
#NYI 
#NYI     # Set the gradient fill properties for the series.
#NYI     my $gradient = $self->_get_gradient_properties( $arg{gradient} );
#NYI 
#NYI     # Pattern fill overrides solid fill.
#NYI     if ( $pattern ) {
#NYI         $fill = undef;
#NYI     }
#NYI 
#NYI     # Gradient fill overrides solid and pattern fills.
#NYI     if ( $gradient ) {
#NYI         $pattern = undef;
#NYI         $fill    = undef;
#NYI     }
#NYI 
#NYI     # Set the marker properties for the series.
#NYI     my $marker = $self->_get_marker_properties( $arg{marker} );
#NYI 
#NYI     # Set the trendline properties for the series.
#NYI     my $trendline = $self->_get_trendline_properties( $arg{trendline} );
#NYI 
#NYI     # Set the line smooth property for the series.
#NYI     my $smooth = $arg{smooth};
#NYI 
#NYI     # Set the error bars properties for the series.
#NYI     my $y_error_bars = $self->_get_error_bars_properties( $arg{y_error_bars} );
#NYI     my $x_error_bars = $self->_get_error_bars_properties( $arg{x_error_bars} );
#NYI 
#NYI     # Set the point properties for the series.
#NYI     my $points = $self->_get_points_properties($arg{points});
#NYI 
#NYI     # Set the labels properties for the series.
#NYI     my $labels = $self->_get_labels_properties( $arg{data_labels} );
#NYI 
#NYI     # Set the "invert if negative" fill property.
#NYI     my $invert_if_neg = $arg{invert_if_negative};
#NYI 
#NYI     # Set the secondary axis properties.
#NYI     my $x2_axis = $arg{x2_axis};
#NYI     my $y2_axis = $arg{y2_axis};
#NYI 
#NYI     # Store secondary status for combined charts.
#NYI     if ($x2_axis || $y2_axis) {
#NYI         $self->{_is_secondary} = 1;
#NYI     }
#NYI 
#NYI     # Set the gap for Bar/Column charts.
#NYI     if ( defined $arg{gap} ) {
#NYI         if ($y2_axis) {
#NYI             $self->{_series_gap_2} = $arg{gap};
#NYI         }
#NYI         else {
#NYI             $self->{_series_gap_1} = $arg{gap};
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the overlap for Bar/Column charts.
#NYI     if ( defined $arg{overlap} ) {
#NYI         if ($y2_axis) {
#NYI             $self->{_series_overlap_2} = $arg{overlap};
#NYI         }
#NYI         else {
#NYI             $self->{_series_overlap_1} = $arg{overlap};
#NYI         }
#NYI     }
#NYI 
#NYI     # Add the user supplied data to the internal structures.
#NYI     %arg = (
#NYI         _values        => $values,
#NYI         _categories    => $categories,
#NYI         _name          => $name,
#NYI         _name_formula  => $name_formula,
#NYI         _name_id       => $name_id,
#NYI         _val_data_id   => $val_id,
#NYI         _cat_data_id   => $cat_id,
#NYI         _line          => $line,
#NYI         _fill          => $fill,
#NYI         _pattern       => $pattern,
#NYI         _gradient      => $gradient,
#NYI         _marker        => $marker,
#NYI         _trendline     => $trendline,
#NYI         _smooth        => $smooth,
#NYI         _labels        => $labels,
#NYI         _invert_if_neg => $invert_if_neg,
#NYI         _x2_axis       => $x2_axis,
#NYI         _y2_axis       => $y2_axis,
#NYI         _points        => $points,
#NYI         _error_bars =>
#NYI           { _x_error_bars => $x_error_bars, _y_error_bars => $y_error_bars },
#NYI     );
#NYI 
#NYI 
#NYI     push @{ $self->{_series} }, \%arg;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_x_axis()
#NYI #
#NYI # Set the properties of the X-axis.
#NYI #
#NYI sub set_x_axis {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $axis = $self->_convert_axis_args( $self->{_x_axis}, @_ );
#NYI 
#NYI     $self->{_x_axis} = $axis;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_y_axis()
#NYI #
#NYI # Set the properties of the Y-axis.
#NYI #
#NYI sub set_y_axis {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $axis = $self->_convert_axis_args( $self->{_y_axis}, @_ );
#NYI 
#NYI     $self->{_y_axis} = $axis;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_x2_axis()
#NYI #
#NYI # Set the properties of the secondary X-axis.
#NYI #
#NYI sub set_x2_axis {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $axis = $self->_convert_axis_args( $self->{_x2_axis}, @_ );
#NYI 
#NYI     $self->{_x2_axis} = $axis;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_y2_axis()
#NYI #
#NYI # Set the properties of the secondary Y-axis.
#NYI #
#NYI sub set_y2_axis {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $axis = $self->_convert_axis_args( $self->{_y2_axis}, @_ );
#NYI 
#NYI     $self->{_y2_axis} = $axis;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_title()
#NYI #
#NYI # Set the properties of the chart title.
#NYI #
#NYI sub set_title {
#NYI 
#NYI     my $self = shift;
#NYI     my %arg  = @_;
#NYI 
#NYI     my ( $name, $name_formula ) =
#NYI       $self->_process_names( $arg{name}, $arg{name_formula} );
#NYI 
#NYI     my $data_id = $self->_get_data_id( $name_formula, $arg{data} );
#NYI 
#NYI     $self->{_title_name}    = $name;
#NYI     $self->{_title_formula} = $name_formula;
#NYI     $self->{_title_data_id} = $data_id;
#NYI 
#NYI     # Set the font properties if present.
#NYI     $self->{_title_font} = $self->_convert_font_args( $arg{name_font} );
#NYI 
#NYI     # Set the title layout.
#NYI     $self->{_title_layout} = $self->_get_layout_properties( $arg{layout}, 1 );
#NYI 
#NYI     # Set the title overlay option.
#NYI     $self->{_title_overlay} = $arg{overlay};
#NYI 
#NYI     # Set the no automatic title option.
#NYI     $self->{_title_none} = $arg{none};
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_legend()
#NYI #
#NYI # Set the properties of the chart legend.
#NYI #
#NYI sub set_legend {
#NYI 
#NYI     my $self = shift;
#NYI     my %arg  = @_;
#NYI 
#NYI     $self->{_legend_position}      = $arg{position} || 'right';
#NYI     $self->{_legend_delete_series} = $arg{delete_series};
#NYI     $self->{_legend_font}          = $self->_convert_font_args( $arg{font} );
#NYI 
#NYI     # Set the legend layout.
#NYI     $self->{_legend_layout} = $self->_get_layout_properties( $arg{layout} );
#NYI 
#NYI     # Turn off the legend.
#NYI     if ( $arg{none} ) {
#NYI         $self->{_legend_position} = 'none';
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_plotarea()
#NYI #
#NYI # Set the properties of the chart plotarea.
#NYI #
#NYI sub set_plotarea {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Convert the user defined properties to internal properties.
#NYI     $self->{_plotarea} = $self->_get_area_properties( @_ );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_chartarea()
#NYI #
#NYI # Set the properties of the chart chartarea.
#NYI #
#NYI sub set_chartarea {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Convert the user defined properties to internal properties.
#NYI     $self->{_chartarea} = $self->_get_area_properties( @_ );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_style()
#NYI #
#NYI # Set on of the 48 built-in Excel chart styles. The default style is 2.
#NYI #
#NYI sub set_style {
#NYI 
#NYI     my $self = shift;
#NYI     my $style_id = defined $_[0] ? $_[0] : 2;
#NYI 
#NYI     if ( $style_id < 0 || $style_id > 48 ) {
#NYI         $style_id = 2;
#NYI     }
#NYI 
#NYI     $self->{_style_id} = $style_id;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # show_blanks_as()
#NYI #
#NYI # Set the option for displaying blank data in a chart. The default is 'gap'.
#NYI #
#NYI sub show_blanks_as {
#NYI 
#NYI     my $self   = shift;
#NYI     my $option = shift;
#NYI 
#NYI     return unless $option;
#NYI 
#NYI     my %valid = (
#NYI         gap  => 1,
#NYI         zero => 1,
#NYI         span => 1,
#NYI 
#NYI     );
#NYI 
#NYI     if ( !exists $valid{$option} ) {
#NYI         warn "Unknown show_blanks_as() option '$option'\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     $self->{_show_blanks} = $option;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # show_hidden_data()
#NYI #
#NYI # Display data in hidden rows or columns.
#NYI #
#NYI sub show_hidden_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->{_show_hidden_data} = 1;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_size()
#NYI #
#NYI # Set dimensions or scale for the chart.
#NYI #
#NYI sub set_size {
#NYI 
#NYI     my $self = shift;
#NYI     my %args = @_;
#NYI 
#NYI     $self->{_width}    = $args{width}    if $args{width};
#NYI     $self->{_height}   = $args{height}   if $args{height};
#NYI     $self->{_x_scale}  = $args{x_scale}  if $args{x_scale};
#NYI     $self->{_y_scale}  = $args{y_scale}  if $args{y_scale};
#NYI     $self->{_x_offset} = $args{x_offset} if $args{x_offset};
#NYI     $self->{_y_offset} = $args{y_offset} if $args{y_offset};
#NYI 
#NYI }
#NYI 
#NYI # Backward compatibility with poorly chosen method name.
#NYI *size = *set_size;
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_table()
#NYI #
#NYI # Set properties for an axis data table.
#NYI #
#NYI sub set_table {
#NYI 
#NYI     my $self = shift;
#NYI     my %args = @_;
#NYI 
#NYI     my %table = (
#NYI         _horizontal => 1,
#NYI         _vertical   => 1,
#NYI         _outline    => 1,
#NYI         _show_keys  => 0,
#NYI     );
#NYI 
#NYI     $table{_horizontal} = $args{horizontal} if defined $args{horizontal};
#NYI     $table{_vertical}   = $args{vertical}   if defined $args{vertical};
#NYI     $table{_outline}    = $args{outline}    if defined $args{outline};
#NYI     $table{_show_keys}  = $args{show_keys}  if defined $args{show_keys};
#NYI     $table{_font}       = $self->_convert_font_args( $args{font} );
#NYI 
#NYI     $self->{_table} = \%table;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_up_down_bars()
#NYI #
#NYI # Set properties for the chart up-down bars.
#NYI #
#NYI sub set_up_down_bars {
#NYI 
#NYI     my $self = shift;
#NYI     my %args = @_;
#NYI 
#NYI     # Map border to line.
#NYI     if ( defined $args{up}->{border} ) {
#NYI         $args{up}->{line} = $args{up}->{border};
#NYI     }
#NYI     if ( defined $args{down}->{border} ) {
#NYI         $args{down}->{line} = $args{down}->{border};
#NYI     }
#NYI 
#NYI     # Set the up and down bar properties.
#NYI     my $up_line   = $self->_get_line_properties( $args{up}->{line} );
#NYI     my $down_line = $self->_get_line_properties( $args{down}->{line} );
#NYI     my $up_fill   = $self->_get_fill_properties( $args{up}->{fill} );
#NYI     my $down_fill = $self->_get_fill_properties( $args{down}->{fill} );
#NYI 
#NYI     $self->{_up_down_bars} = {
#NYI         _up => {
#NYI             _line => $up_line,
#NYI             _fill => $up_fill,
#NYI         },
#NYI         _down => {
#NYI             _line => $down_line,
#NYI             _fill => $down_fill,
#NYI         },
#NYI     };
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_drop_lines()
#NYI #
#NYI # Set properties for the chart drop lines.
#NYI #
#NYI sub set_drop_lines {
#NYI 
#NYI     my $self = shift;
#NYI     my %args = @_;
#NYI 
#NYI     # Set the drop line properties.
#NYI     my $line = $self->_get_line_properties( $args{line} );
#NYI 
#NYI     $self->{_drop_lines} = { _line => $line };
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_high_low_lines()
#NYI #
#NYI # Set properties for the chart high-low lines.
#NYI #
#NYI sub set_high_low_lines {
#NYI 
#NYI     my $self = shift;
#NYI     my %args = @_;
#NYI 
#NYI     # Set the drop line properties.
#NYI     my $line = $self->_get_line_properties( $args{line} );
#NYI 
#NYI     $self->{_hi_low_lines} = { _line => $line };
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # combine()
#NYI #
#NYI # Add another chart to create a combined chart.
#NYI #
#NYI sub combine {
#NYI 
#NYI     my $self  = shift;
#NYI     my $chart = shift;
#NYI 
#NYI     $self->{_combined} = $chart;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Internal methods. The following section of methods are used for the internal
#NYI # structuring of the Chart object and file format.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _convert_axis_args()
#NYI #
#NYI # Convert user defined axis values into private hash values.
#NYI #
#NYI sub _convert_axis_args {
#NYI 
#NYI     my $self = shift;
#NYI     my $axis = shift;
#NYI     my %arg  = ( %{ $axis->{_defaults} }, @_ );
#NYI 
#NYI     my ( $name, $name_formula ) =
#NYI       $self->_process_names( $arg{name}, $arg{name_formula} );
#NYI 
#NYI     my $data_id = $self->_get_data_id( $name_formula, $arg{data} );
#NYI 
#NYI     $axis = {
#NYI         _defaults          => $axis->{_defaults},
#NYI         _name              => $name,
#NYI         _formula           => $name_formula,
#NYI         _data_id           => $data_id,
#NYI         _reverse           => $arg{reverse},
#NYI         _min               => $arg{min},
#NYI         _max               => $arg{max},
#NYI         _minor_unit        => $arg{minor_unit},
#NYI         _major_unit        => $arg{major_unit},
#NYI         _minor_unit_type   => $arg{minor_unit_type},
#NYI         _major_unit_type   => $arg{major_unit_type},
#NYI         _log_base          => $arg{log_base},
#NYI         _crossing          => $arg{crossing},
#NYI         _position_axis     => $arg{position_axis},
#NYI         _position          => $arg{position},
#NYI         _label_position    => $arg{label_position},
#NYI         _num_format        => $arg{num_format},
#NYI         _num_format_linked => $arg{num_format_linked},
#NYI         _interval_unit     => $arg{interval_unit},
#NYI         _interval_tick     => $arg{interval_tick},
#NYI         _visible           => defined $arg{visible} ? $arg{visible} : 1,
#NYI         _text_axis         => 0,
#NYI     };
#NYI 
#NYI     # Map major_gridlines properties.
#NYI     if ( $arg{major_gridlines} && $arg{major_gridlines}->{visible} ) {
#NYI         $axis->{_major_gridlines} =
#NYI           $self->_get_gridline_properties( $arg{major_gridlines} );
#NYI     }
#NYI 
#NYI     # Map minor_gridlines properties.
#NYI     if ( $arg{minor_gridlines} && $arg{minor_gridlines}->{visible} ) {
#NYI         $axis->{_minor_gridlines} =
#NYI           $self->_get_gridline_properties( $arg{minor_gridlines} );
#NYI     }
#NYI 
#NYI     # Convert the display units.
#NYI     $axis->{_display_units} = $self->_get_display_units( $arg{display_units} );
#NYI     if ( defined $arg{display_units_visible} ) {
#NYI         $axis->{_display_units_visible} = $arg{display_units_visible};
#NYI     }
#NYI     else {
#NYI         $axis->{_display_units_visible} = 1;
#NYI     }
#NYI 
#NYI     # Only use the first letter of bottom, top, left or right.
#NYI     if ( defined $axis->{_position} ) {
#NYI         $axis->{_position} = substr lc $axis->{_position}, 0, 1;
#NYI     }
#NYI 
#NYI     # Set the position for a category axis on or between the tick marks.
#NYI     if ( defined $axis->{_position_axis} ) {
#NYI         if ( $axis->{_position_axis} eq 'on_tick' ) {
#NYI             $axis->{_position_axis} = 'midCat';
#NYI         }
#NYI         elsif ( $axis->{_position_axis} eq 'between' ) {
#NYI 
#NYI             # Doesn't need to be modified.
#NYI         }
#NYI         else {
#NYI             # Otherwise use the default value.
#NYI             $axis->{_position_axis} = undef;
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the category axis as a date axis.
#NYI     if ( $arg{date_axis} ) {
#NYI         $self->{_date_category} = 1;
#NYI     }
#NYI 
#NYI     # Set the category axis as a text axis.
#NYI     if ( $arg{text_axis} ) {
#NYI         $self->{_date_category} = 0;
#NYI         $axis->{_text_axis} = 1;
#NYI     }
#NYI 
#NYI 
#NYI     # Set the font properties if present.
#NYI     $axis->{_num_font}  = $self->_convert_font_args( $arg{num_font} );
#NYI     $axis->{_name_font} = $self->_convert_font_args( $arg{name_font} );
#NYI 
#NYI     # Set the axis name layout.
#NYI     $axis->{_layout} = $self->_get_layout_properties( $arg{name_layout}, 1 );
#NYI 
#NYI     # Set the line properties for the axis.
#NYI     $axis->{_line} = $self->_get_line_properties( $arg{line} );
#NYI 
#NYI     # Set the fill properties for the axis.
#NYI     $axis->{_fill} = $self->_get_fill_properties( $arg{fill} );
#NYI 
#NYI     # Set the tick marker types.
#NYI     $axis->{_minor_tick_mark} = $self->_get_tick_type($arg{minor_tick_mark});
#NYI     $axis->{_major_tick_mark} = $self->_get_tick_type($arg{major_tick_mark});
#NYI 
#NYI 
#NYI     return $axis;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _convert_fonts_args()
#NYI #
#NYI # Convert user defined font values into private hash values.
#NYI #
#NYI sub _convert_font_args {
#NYI 
#NYI     my $self = shift;
#NYI     my $args = shift;
#NYI 
#NYI     return unless $args;
#NYI 
#NYI     my $font = {
#NYI         _name         => $args->{name},
#NYI         _color        => $args->{color},
#NYI         _size         => $args->{size},
#NYI         _bold         => $args->{bold},
#NYI         _italic       => $args->{italic},
#NYI         _underline    => $args->{underline},
#NYI         _pitch_family => $args->{pitch_family},
#NYI         _charset      => $args->{charset},
#NYI         _baseline     => $args->{baseline} || 0,
#NYI         _rotation     => $args->{rotation},
#NYI     };
#NYI 
#NYI     # Convert font size units.
#NYI     $font->{_size} *= 100 if $font->{_size};
#NYI 
#NYI     # Convert rotation into 60,000ths of a degree.
#NYI     if ( $font->{_rotation} ) {
#NYI         $font->{_rotation} = 60_000 * int( $font->{_rotation} );
#NYI     }
#NYI 
#NYI     return $font;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _aref_to_formula()
#NYI #
#NYI # Convert and aref of row col values to a range formula.
#NYI #
#NYI sub _aref_to_formula {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     # If it isn't an array ref it is probably a formula already.
#NYI     return $data if !ref $data;
#NYI 
#NYI     my $formula = xl_range_formula( @$data );
#NYI 
#NYI     return $formula;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _process_names()
#NYI #
#NYI # Switch name and name_formula parameters if required.
#NYI #
#NYI sub _process_names {
#NYI 
#NYI     my $self         = shift;
#NYI     my $name         = shift;
#NYI     my $name_formula = shift;
#NYI 
#NYI     if ( defined $name ) {
#NYI 
#NYI         if ( ref $name eq 'ARRAY' ) {
#NYI             my $cell = xl_rowcol_to_cell( $name->[1], $name->[2], 1, 1 );
#NYI             $name_formula = quote_sheetname( $name->[0] ) . '!' . $cell;
#NYI             $name         = '';
#NYI         }
#NYI         elsif ( $name =~ m/^=[^!]+!\$/ ) {
#NYI 
#NYI             # Name looks like a formula, use it to set name_formula.
#NYI             $name_formula = $name;
#NYI             $name         = '';
#NYI         }
#NYI     }
#NYI 
#NYI     return ( $name, $name_formula );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_data_type()
#NYI #
#NYI # Find the overall type of the data associated with a series.
#NYI #
#NYI # TODO. Need to handle date type.
#NYI #
#NYI sub _get_data_type {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     # Check for no data in the series.
#NYI     return 'none' if !defined $data;
#NYI     return 'none' if @$data == 0;
#NYI 
#NYI     if (ref $data->[0] eq 'ARRAY') {
#NYI         return 'multi_str'
#NYI     }
#NYI 
#NYI     # If the token isn't a number assume it is a string.
#NYI     for my $token ( @$data ) {
#NYI         next if !defined $token;
#NYI         return 'str'
#NYI           if $token !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/;
#NYI     }
#NYI 
#NYI     # The series data was all numeric.
#NYI     return 'num';
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_data_id()
#NYI #
#NYI # Assign an id to a each unique series formula or title/axis formula. Repeated
#NYI # formulas such as for categories get the same id. If the series or title
#NYI # has user specified data associated with it then that is also stored. This
#NYI # data is used to populate cached Excel data when creating a chart.
#NYI # If there is no user defined data then it will be populated by the parent
#NYI # workbook in Workbook::_add_chart_data()
#NYI #
#NYI sub _get_data_id {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI     my $data    = shift;
#NYI     my $id;
#NYI 
#NYI     # Ignore series without a range formula.
#NYI     return unless $formula;
#NYI 
#NYI     # Strip the leading '=' from the formula.
#NYI     $formula =~ s/^=//;
#NYI 
#NYI     # Store the data id in a hash keyed by the formula and store the data
#NYI     # in a separate array with the same id.
#NYI     if ( !exists $self->{_formula_ids}->{$formula} ) {
#NYI 
#NYI         # Haven't seen this formula before.
#NYI         $id = @{ $self->{_formula_data} };
#NYI 
#NYI         push @{ $self->{_formula_data} }, $data;
#NYI         $self->{_formula_ids}->{$formula} = $id;
#NYI     }
#NYI     else {
#NYI 
#NYI         # Formula already seen. Return existing id.
#NYI         $id = $self->{_formula_ids}->{$formula};
#NYI 
#NYI         # Store user defined data if it isn't already there.
#NYI         if ( !defined $self->{_formula_data}->[$id] ) {
#NYI             $self->{_formula_data}->[$id] = $data;
#NYI         }
#NYI     }
#NYI 
#NYI     return $id;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_color()
#NYI #
#NYI # Convert the user specified colour index or string to a rgb colour.
#NYI #
#NYI sub _get_color {
#NYI 
#NYI     my $self  = shift;
#NYI     my $color = shift;
#NYI 
#NYI     # Convert a HTML style #RRGGBB color.
#NYI     if ( defined $color and $color =~ /^#[0-9a-fA-F]{6}$/ ) {
#NYI         $color =~ s/^#//;
#NYI         return uc $color;
#NYI     }
#NYI 
#NYI     my $index = &Excel::Writer::XLSX::Format::_get_color( $color );
#NYI 
#NYI     # Set undefined colors to black.
#NYI     if ( !$index ) {
#NYI         $index = 0x08;
#NYI         warn "Unknown color '$color' used in chart formatting. "
#NYI           . "Converting to black.\n";
#NYI     }
#NYI 
#NYI     return $self->_get_palette_color( $index );
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_palette_color()
#NYI #
#NYI # Convert from an Excel internal colour index to a XML style #RRGGBB index
#NYI # based on the default or user defined values in the Workbook palette.
#NYI # Note: This version doesn't add an alpha channel.
#NYI #
#NYI sub _get_palette_color {
#NYI 
#NYI     my $self    = shift;
#NYI     my $index   = shift;
#NYI     my $palette = $self->{_palette};
#NYI 
#NYI     # Adjust the colour index.
#NYI     $index -= 8;
#NYI 
#NYI     # Palette is passed in from the Workbook class.
#NYI     my @rgb = @{ $palette->[$index] };
#NYI 
#NYI     return sprintf "%02X%02X%02X", @rgb[0, 1, 2];
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_swe_line_pattern()
#NYI #
#NYI # Get the Spreadsheet::WriteExcel line pattern for backward compatibility.
#NYI #
#NYI sub _get_swe_line_pattern {
#NYI 
#NYI     my $self    = shift;
#NYI     my $value   = lc shift;
#NYI     my $default = 'solid';
#NYI     my $pattern;
#NYI 
#NYI     my %patterns = (
#NYI         0              => 'solid',
#NYI         1              => 'dash',
#NYI         2              => 'dot',
#NYI         3              => 'dash_dot',
#NYI         4              => 'long_dash_dot_dot',
#NYI         5              => 'none',
#NYI         6              => 'solid',
#NYI         7              => 'solid',
#NYI         8              => 'solid',
#NYI         'solid'        => 'solid',
#NYI         'dash'         => 'dash',
#NYI         'dot'          => 'dot',
#NYI         'dash-dot'     => 'dash_dot',
#NYI         'dash-dot-dot' => 'long_dash_dot_dot',
#NYI         'none'         => 'none',
#NYI         'dark-gray'    => 'solid',
#NYI         'medium-gray'  => 'solid',
#NYI         'light-gray'   => 'solid',
#NYI     );
#NYI 
#NYI     if ( exists $patterns{$value} ) {
#NYI         $pattern = $patterns{$value};
#NYI     }
#NYI     else {
#NYI         $pattern = $default;
#NYI     }
#NYI 
#NYI     return $pattern;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_swe_line_weight()
#NYI #
#NYI # Get the Spreadsheet::WriteExcel line weight for backward compatibility.
#NYI #
#NYI sub _get_swe_line_weight {
#NYI 
#NYI     my $self    = shift;
#NYI     my $value   = lc shift;
#NYI     my $default = 1;
#NYI     my $weight;
#NYI 
#NYI     my %weights = (
#NYI         1          => 0.25,
#NYI         2          => 1,
#NYI         3          => 2,
#NYI         4          => 3,
#NYI         'hairline' => 0.25,
#NYI         'narrow'   => 1,
#NYI         'medium'   => 2,
#NYI         'wide'     => 3,
#NYI     );
#NYI 
#NYI     if ( exists $weights{$value} ) {
#NYI         $weight = $weights{$value};
#NYI     }
#NYI     else {
#NYI         $weight = $default;
#NYI     }
#NYI 
#NYI     return $weight;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_line_properties()
#NYI #
#NYI # Convert user defined line properties to the structure required internally.
#NYI #
#NYI sub _get_line_properties {
#NYI 
#NYI     my $self = shift;
#NYI     my $line = shift;
#NYI 
#NYI     return { _defined => 0 } unless $line;
#NYI 
#NYI     # Copy the user supplied properties.
#NYI     $line = { %$line };
#NYI 
#NYI     my %dash_types = (
#NYI         solid               => 'solid',
#NYI         round_dot           => 'sysDot',
#NYI         square_dot          => 'sysDash',
#NYI         dash                => 'dash',
#NYI         dash_dot            => 'dashDot',
#NYI         long_dash           => 'lgDash',
#NYI         long_dash_dot       => 'lgDashDot',
#NYI         long_dash_dot_dot   => 'lgDashDotDot',
#NYI         dot                 => 'dot',
#NYI         system_dash_dot     => 'sysDashDot',
#NYI         system_dash_dot_dot => 'sysDashDotDot',
#NYI     );
#NYI 
#NYI     # Check the dash type.
#NYI     my $dash_type = $line->{dash_type};
#NYI 
#NYI     if ( defined $dash_type ) {
#NYI         if ( exists $dash_types{$dash_type} ) {
#NYI             $line->{dash_type} = $dash_types{$dash_type};
#NYI         }
#NYI         else {
#NYI             warn "Unknown dash type '$dash_type'\n";
#NYI             return;
#NYI         }
#NYI     }
#NYI 
#NYI     $line->{_defined} = 1;
#NYI 
#NYI     return $line;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_fill_properties()
#NYI #
#NYI # Convert user defined fill properties to the structure required internally.
#NYI #
#NYI sub _get_fill_properties {
#NYI 
#NYI     my $self = shift;
#NYI     my $fill = shift;
#NYI 
#NYI     return { _defined => 0 } unless $fill;
#NYI 
#NYI     $fill->{_defined} = 1;
#NYI 
#NYI     return $fill;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_pattern_properties()
#NYI #
#NYI # Convert user defined pattern properties to the structure required internally.
#NYI #
#NYI sub _get_pattern_properties {
#NYI 
#NYI     my $self    = shift;
#NYI     my $args    = shift;
#NYI     my $pattern = {};
#NYI 
#NYI     return unless $args;
#NYI 
#NYI     # Check the pattern type is present.
#NYI     if ( !$args->{pattern} ) {
#NYI         carp "Pattern must include 'pattern'";
#NYI         return;
#NYI     }
#NYI 
#NYI     # Check the foreground color is present.
#NYI     if ( !$args->{fg_color} ) {
#NYI         carp "Pattern must include 'fg_color'";
#NYI         return;
#NYI     }
#NYI 
#NYI     my %types = (
#NYI         'percent_5'                 => 'pct5',
#NYI         'percent_10'                => 'pct10',
#NYI         'percent_20'                => 'pct20',
#NYI         'percent_25'                => 'pct25',
#NYI         'percent_30'                => 'pct30',
#NYI         'percent_40'                => 'pct40',
#NYI 
#NYI         'percent_50'                => 'pct50',
#NYI         'percent_60'                => 'pct60',
#NYI         'percent_70'                => 'pct70',
#NYI         'percent_75'                => 'pct75',
#NYI         'percent_80'                => 'pct80',
#NYI         'percent_90'                => 'pct90',
#NYI 
#NYI         'light_downward_diagonal'   => 'ltDnDiag',
#NYI         'light_upward_diagonal'     => 'ltUpDiag',
#NYI         'dark_downward_diagonal'    => 'dkDnDiag',
#NYI         'dark_upward_diagonal'      => 'dkUpDiag',
#NYI         'wide_downward_diagonal'    => 'wdDnDiag',
#NYI         'wide_upward_diagonal'      => 'wdUpDiag',
#NYI 
#NYI         'light_vertical'            => 'ltVert',
#NYI         'light_horizontal'          => 'ltHorz',
#NYI         'narrow_vertical'           => 'narVert',
#NYI         'narrow_horizontal'         => 'narHorz',
#NYI         'dark_vertical'             => 'dkVert',
#NYI         'dark_horizontal'           => 'dkHorz',
#NYI 
#NYI         'dashed_downward_diagonal'  => 'dashDnDiag',
#NYI         'dashed_upward_diagonal'    => 'dashUpDiag',
#NYI         'dashed_horizontal'         => 'dashHorz',
#NYI         'dashed_vertical'           => 'dashVert',
#NYI         'small_confetti'            => 'smConfetti',
#NYI         'large_confetti'            => 'lgConfetti',
#NYI 
#NYI         'zigzag'                    => 'zigZag',
#NYI         'wave'                      => 'wave',
#NYI         'diagonal_brick'            => 'diagBrick',
#NYI         'horizontal_brick'          => 'horzBrick',
#NYI         'weave'                     => 'weave',
#NYI         'plaid'                     => 'plaid',
#NYI 
#NYI         'divot'                     => 'divot',
#NYI         'dotted_grid'               => 'dotGrid',
#NYI         'dotted_diamond'            => 'dotDmnd',
#NYI         'shingle'                   => 'shingle',
#NYI         'trellis'                   => 'trellis',
#NYI         'sphere'                    => 'sphere',
#NYI 
#NYI         'small_grid'                => 'smGrid',
#NYI         'large_grid'                => 'lgGrid',
#NYI         'small_check'               => 'smCheck',
#NYI         'large_check'               => 'lgCheck',
#NYI         'outlined_diamond'          => 'openDmnd',
#NYI         'solid_diamond'             => 'solidDmnd',
#NYI     );
#NYI 
#NYI     # Check for valid types.
#NYI     my $pattern_type = $args->{pattern};
#NYI 
#NYI     if ( exists $types{$pattern_type} ) {
#NYI         $pattern->{pattern} = $types{$pattern_type};
#NYI     }
#NYI     else {
#NYI         carp "Unknown pattern type '$pattern_type'";
#NYI         return;
#NYI     }
#NYI 
#NYI     # Specify a default background color.
#NYI     if ( !$args->{bg_color} ) {
#NYI         $pattern->{bg_color} = '#FFFFFF';
#NYI     }
#NYI     else {
#NYI         $pattern->{bg_color} = $args->{bg_color};
#NYI     }
#NYI 
#NYI     $pattern->{fg_color} = $args->{fg_color};
#NYI 
#NYI     return $pattern;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_gradient_properties()
#NYI #
#NYI # Convert user defined gradient to the structure required internally.
#NYI #
#NYI sub _get_gradient_properties {
#NYI 
#NYI     my $self     = shift;
#NYI     my $args     = shift;
#NYI     my $gradient = {};
#NYI 
#NYI     my %types    = (
#NYI         linear      => 'linear',
#NYI         radial      => 'circle',
#NYI         rectangular => 'rect',
#NYI         path        => 'shape'
#NYI     );
#NYI 
#NYI     return unless $args;
#NYI 
#NYI     # Check the colors array exists and is valid.
#NYI     if ( !$args->{colors} || ref $args->{colors} ne 'ARRAY' ) {
#NYI         carp "Gradient must include colors array";
#NYI         return;
#NYI     }
#NYI 
#NYI     # Check the colors array has the required number of entries.
#NYI     if ( @{ $args->{colors} } < 2 ) {
#NYI         carp "Gradient colors array must at least 2 values";
#NYI         return;
#NYI     }
#NYI 
#NYI     $gradient->{_colors} = $args->{colors};
#NYI 
#NYI     if ( $args->{positions} ) {
#NYI 
#NYI         # Check the positions array has the right number of entries.
#NYI         if ( @{ $args->{positions} } != @{ $args->{colors} } ) {
#NYI             carp "Gradient positions not equal to number of colors";
#NYI             return;
#NYI         }
#NYI 
#NYI         # Check the positions are in the correct range.
#NYI         for my $pos ( @{ $args->{positions} } ) {
#NYI             if ( $pos < 0 || $pos > 100 ) {
#NYI                 carp "Gradient position '", $pos,
#NYI                   "' must be in range 0 <= pos <= 100";
#NYI                 return;
#NYI             }
#NYI         }
#NYI 
#NYI         $gradient->{_positions} = $args->{positions};
#NYI     }
#NYI     else {
#NYI         # Use the default gradient positions.
#NYI         if ( @{ $args->{colors} } == 2 ) {
#NYI             $gradient->{_positions} = [ 0, 100 ];
#NYI         }
#NYI         elsif ( @{ $args->{colors} } == 3 ) {
#NYI             $gradient->{_positions} = [ 0, 50, 100 ];
#NYI         }
#NYI         elsif ( @{ $args->{colors} } == 4 ) {
#NYI             $gradient->{_positions} = [ 0, 33, 66, 100 ];
#NYI         }
#NYI         else {
#NYI             carp "Must specify gradient positions";
#NYI             return;
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the gradient angle.
#NYI     if ( defined $args->{angle} ) {
#NYI         my $angle = $args->{angle};
#NYI 
#NYI         if ( $angle < 0 || $angle > 359.9 ) {
#NYI             carp "Gradient angle '", $angle,
#NYI               "' must be in range 0 <= pos < 360";
#NYI             return;
#NYI         }
#NYI         $gradient->{_angle} = $angle;
#NYI     }
#NYI     else {
#NYI         $gradient->{_angle} = 90;
#NYI     }
#NYI 
#NYI     # Set the gradient type.
#NYI     if ( defined $args->{type} ) {
#NYI         my $type = $args->{type};
#NYI 
#NYI         if ( !exists $types{$type} ) {
#NYI             carp "Unknown gradient type '", $type, "'";
#NYI             return;
#NYI         }
#NYI         $gradient->{_type} = $types{$type};
#NYI     }
#NYI     else {
#NYI         $gradient->{_type} = 'linear';
#NYI     }
#NYI 
#NYI     return $gradient;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_marker_properties()
#NYI #
#NYI # Convert user defined marker properties to the structure required internally.
#NYI #
#NYI sub _get_marker_properties {
#NYI 
#NYI     my $self   = shift;
#NYI     my $marker = shift;
#NYI 
#NYI     return if !$marker && ref $marker ne 'HASH';
#NYI 
#NYI     # Copy the user supplied properties.
#NYI     $marker = { %$marker };
#NYI 
#NYI     my %types = (
#NYI         automatic  => 'automatic',
#NYI         none       => 'none',
#NYI         square     => 'square',
#NYI         diamond    => 'diamond',
#NYI         triangle   => 'triangle',
#NYI         x          => 'x',
#NYI         star       => 'star',
#NYI         dot        => 'dot',
#NYI         short_dash => 'dot',
#NYI         dash       => 'dash',
#NYI         long_dash  => 'dash',
#NYI         circle     => 'circle',
#NYI         plus       => 'plus',
#NYI         picture    => 'picture',
#NYI     );
#NYI 
#NYI     # Check for valid types.
#NYI     my $marker_type = $marker->{type};
#NYI 
#NYI     if ( defined $marker_type ) {
#NYI         if ( $marker_type eq 'automatic' ) {
#NYI             $marker->{automatic} = 1;
#NYI         }
#NYI 
#NYI         if ( exists $types{$marker_type} ) {
#NYI             $marker->{type} = $types{$marker_type};
#NYI         }
#NYI         else {
#NYI             warn "Unknown marker type '$marker_type'\n";
#NYI             return;
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the line properties for the marker..
#NYI     my $line = $self->_get_line_properties( $marker->{line} );
#NYI 
#NYI     # Allow 'border' as a synonym for 'line'.
#NYI     if ( $marker->{border} ) {
#NYI         $line = $self->_get_line_properties( $marker->{border} );
#NYI     }
#NYI 
#NYI     # Set the fill properties for the marker.
#NYI     my $fill = $self->_get_fill_properties( $marker->{fill} );
#NYI 
#NYI     # Set the pattern properties for the series.
#NYI     my $pattern = $self->_get_pattern_properties( $marker->{pattern} );
#NYI 
#NYI     # Set the gradient fill properties for the series.
#NYI     my $gradient = $self->_get_gradient_properties( $marker->{gradient} );
#NYI 
#NYI     # Pattern fill overrides solid fill.
#NYI     if ( $pattern ) {
#NYI         $fill = undef;
#NYI     }
#NYI 
#NYI     # Gradient fill overrides solid and pattern fills.
#NYI     if ( $gradient ) {
#NYI         $pattern = undef;
#NYI         $fill    = undef;
#NYI     }
#NYI 
#NYI     $marker->{_line}     = $line;
#NYI     $marker->{_fill}     = $fill;
#NYI     $marker->{_pattern}  = $pattern;
#NYI     $marker->{_gradient} = $gradient;
#NYI 
#NYI     return $marker;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_trendline_properties()
#NYI #
#NYI # Convert user defined trendline properties to the structure required
#NYI # internally.
#NYI #
#NYI sub _get_trendline_properties {
#NYI 
#NYI     my $self      = shift;
#NYI     my $trendline = shift;
#NYI 
#NYI     return if !$trendline && ref $trendline ne 'HASH';
#NYI 
#NYI     # Copy the user supplied properties.
#NYI     $trendline = { %$trendline };
#NYI 
#NYI     my %types = (
#NYI         exponential    => 'exp',
#NYI         linear         => 'linear',
#NYI         log            => 'log',
#NYI         moving_average => 'movingAvg',
#NYI         polynomial     => 'poly',
#NYI         power          => 'power',
#NYI     );
#NYI 
#NYI     # Check the trendline type.
#NYI     my $trend_type = $trendline->{type};
#NYI 
#NYI     if ( exists $types{$trend_type} ) {
#NYI         $trendline->{type} = $types{$trend_type};
#NYI     }
#NYI     else {
#NYI         warn "Unknown trendline type '$trend_type'\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     # Set the line properties for the trendline..
#NYI     my $line = $self->_get_line_properties( $trendline->{line} );
#NYI 
#NYI     # Allow 'border' as a synonym for 'line'.
#NYI     if ( $trendline->{border} ) {
#NYI         $line = $self->_get_line_properties( $trendline->{border} );
#NYI     }
#NYI 
#NYI     # Set the fill properties for the trendline.
#NYI     my $fill = $self->_get_fill_properties( $trendline->{fill} );
#NYI 
#NYI     # Set the pattern properties for the series.
#NYI     my $pattern = $self->_get_pattern_properties( $trendline->{pattern} );
#NYI 
#NYI     # Set the gradient fill properties for the series.
#NYI     my $gradient = $self->_get_gradient_properties( $trendline->{gradient} );
#NYI 
#NYI     # Pattern fill overrides solid fill.
#NYI     if ( $pattern ) {
#NYI         $fill = undef;
#NYI     }
#NYI 
#NYI     # Gradient fill overrides solid and pattern fills.
#NYI     if ( $gradient ) {
#NYI         $pattern = undef;
#NYI         $fill    = undef;
#NYI     }
#NYI 
#NYI     $trendline->{_line}     = $line;
#NYI     $trendline->{_fill}     = $fill;
#NYI     $trendline->{_pattern}  = $pattern;
#NYI     $trendline->{_gradient} = $gradient;
#NYI 
#NYI     return $trendline;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_error_bars_properties()
#NYI #
#NYI # Convert user defined error bars properties to structure required internally.
#NYI #
#NYI sub _get_error_bars_properties {
#NYI 
#NYI     my $self = shift;
#NYI     my $args = shift;
#NYI 
#NYI     return if !$args && ref $args ne 'HASH';
#NYI 
#NYI     # Copy the user supplied properties.
#NYI     $args = { %$args };
#NYI 
#NYI 
#NYI     # Default values.
#NYI     my $error_bars = {
#NYI         _type         => 'fixedVal',
#NYI         _value        => 1,
#NYI         _endcap       => 1,
#NYI         _direction    => 'both',
#NYI         _plus_values  => [1],
#NYI         _minus_values => [1],
#NYI         _plus_data    => [],
#NYI         _minus_data   => [],
#NYI     };
#NYI 
#NYI     my %types = (
#NYI         fixed              => 'fixedVal',
#NYI         percentage         => 'percentage',
#NYI         standard_deviation => 'stdDev',
#NYI         standard_error     => 'stdErr',
#NYI         custom             => 'cust',
#NYI     );
#NYI 
#NYI     # Check the error bars type.
#NYI     my $error_type = $args->{type};
#NYI 
#NYI     if ( exists $types{$error_type} ) {
#NYI         $error_bars->{_type} = $types{$error_type};
#NYI     }
#NYI     else {
#NYI         warn "Unknown error bars type '$error_type'\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     # Set the value for error types that require it.
#NYI     if ( defined $args->{value} ) {
#NYI         $error_bars->{_value} = $args->{value};
#NYI     }
#NYI 
#NYI     # Set the end-cap style.
#NYI     if ( defined $args->{end_style} ) {
#NYI         $error_bars->{_endcap} = $args->{end_style};
#NYI     }
#NYI 
#NYI     # Set the error bar direction.
#NYI     if ( defined $args->{direction} ) {
#NYI         if ( $args->{direction} eq 'minus' ) {
#NYI             $error_bars->{_direction} = 'minus';
#NYI         }
#NYI         elsif ( $args->{direction} eq 'plus' ) {
#NYI             $error_bars->{_direction} = 'plus';
#NYI         }
#NYI         else {
#NYI             # Default to 'both'.
#NYI         }
#NYI     }
#NYI 
#NYI     # Set any custom values.
#NYI     if ( defined $args->{plus_values} ) {
#NYI         $error_bars->{_plus_values} = $args->{plus_values};
#NYI     }
#NYI     if ( defined $args->{minus_values} ) {
#NYI         $error_bars->{_minus_values} = $args->{minus_values};
#NYI     }
#NYI     if ( defined $args->{plus_data} ) {
#NYI         $error_bars->{_plus_data} = $args->{plus_data};
#NYI     }
#NYI     if ( defined $args->{minus_data} ) {
#NYI         $error_bars->{_minus_data} = $args->{minus_data};
#NYI     }
#NYI 
#NYI     # Set the line properties for the error bars.
#NYI     $error_bars->{_line} = $self->_get_line_properties( $args->{line} );
#NYI 
#NYI     return $error_bars;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_gridline_properties()
#NYI #
#NYI # Convert user defined gridline properties to the structure required internally.
#NYI #
#NYI sub _get_gridline_properties {
#NYI 
#NYI     my $self = shift;
#NYI     my $args = shift;
#NYI     my $gridline;
#NYI 
#NYI     # Set the visible property for the gridline.
#NYI     $gridline->{_visible} = $args->{visible};
#NYI 
#NYI     # Set the line properties for the gridline..
#NYI     $gridline->{_line} = $self->_get_line_properties( $args->{line} );
#NYI 
#NYI     return $gridline;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_labels_properties()
#NYI #
#NYI # Convert user defined labels properties to the structure required internally.
#NYI #
#NYI sub _get_labels_properties {
#NYI 
#NYI     my $self   = shift;
#NYI     my $labels = shift;
#NYI 
#NYI     return if !$labels && ref $labels ne 'HASH';
#NYI 
#NYI     # Copy the user supplied properties.
#NYI     $labels = { %$labels };
#NYI 
#NYI     # Map user defined label positions to Excel positions.
#NYI     if ( my $position = $labels->{position} ) {
#NYI 
#NYI         if ( exists $self->{_label_positions}->{$position} ) {
#NYI             if ($position eq $self->{_label_position_default}) {
#NYI                 $labels->{position} = undef;
#NYI             }
#NYI             else {
#NYI                 $labels->{position} = $self->{_label_positions}->{$position};
#NYI             }
#NYI         }
#NYI         else {
#NYI             carp "Unsupported label position '$position' for this chart type";
#NYI             return undef
#NYI         }
#NYI     }
#NYI 
#NYI     # Map the user defined label separator to the Excel separator.
#NYI     if ( my $separator = $labels->{separator} ) {
#NYI 
#NYI         my %separators = (
#NYI             ','  => ', ',
#NYI             ';'  => '; ',
#NYI             '.'  => '. ',
#NYI             "\n" => "\n",
#NYI             ' '  => ' '
#NYI         );
#NYI 
#NYI         if ( exists $separators{$separator} ) {
#NYI             $labels->{separator} = $separators{$separator};
#NYI         }
#NYI         else {
#NYI             carp "Unsupported label separator";
#NYI             return undef
#NYI         }
#NYI     }
#NYI 
#NYI     if ($labels->{font}) {
#NYI         $labels->{font} = $self->_convert_font_args( $labels->{font} );
#NYI     }
#NYI 
#NYI     return $labels;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_area_properties()
#NYI #
#NYI # Convert user defined area properties to the structure required internally.
#NYI #
#NYI sub _get_area_properties {
#NYI 
#NYI     my $self = shift;
#NYI     my %arg  = @_;
#NYI     my $area = {};
#NYI 
#NYI 
#NYI     # Map deprecated Spreadsheet::WriteExcel fill colour.
#NYI     if ( $arg{color} ) {
#NYI         $arg{fill}->{color} = $arg{color};
#NYI     }
#NYI 
#NYI     # Map deprecated Spreadsheet::WriteExcel line_weight.
#NYI     if ( $arg{line_weight} ) {
#NYI         my $width = $self->_get_swe_line_weight( $arg{line_weight} );
#NYI         $arg{border}->{width} = $width;
#NYI     }
#NYI 
#NYI     # Map deprecated Spreadsheet::WriteExcel line_pattern.
#NYI     if ( $arg{line_pattern} ) {
#NYI         my $pattern = $self->_get_swe_line_pattern( $arg{line_pattern} );
#NYI 
#NYI         if ( $pattern eq 'none' ) {
#NYI             $arg{border}->{none} = 1;
#NYI         }
#NYI         else {
#NYI             $arg{border}->{dash_type} = $pattern;
#NYI         }
#NYI     }
#NYI 
#NYI     # Map deprecated Spreadsheet::WriteExcel line colour.
#NYI     if ( $arg{line_color} ) {
#NYI         $arg{border}->{color} = $arg{line_color};
#NYI     }
#NYI 
#NYI 
#NYI     # Handle Excel::Writer::XLSX style properties.
#NYI 
#NYI     # Set the line properties for the chartarea.
#NYI     my $line = $self->_get_line_properties( $arg{line} );
#NYI 
#NYI     # Allow 'border' as a synonym for 'line'.
#NYI     if ( $arg{border} ) {
#NYI         $line = $self->_get_line_properties( $arg{border} );
#NYI     }
#NYI 
#NYI     # Set the fill properties for the chartarea.
#NYI     my $fill = $self->_get_fill_properties( $arg{fill} );
#NYI 
#NYI     # Set the pattern properties for the series.
#NYI     my $pattern = $self->_get_pattern_properties( $arg{pattern} );
#NYI 
#NYI     # Set the gradient fill properties for the series.
#NYI     my $gradient = $self->_get_gradient_properties( $arg{gradient} );
#NYI 
#NYI     # Pattern fill overrides solid fill.
#NYI     if ( $pattern ) {
#NYI         $fill = undef;
#NYI     }
#NYI 
#NYI     # Gradient fill overrides solid and pattern fills.
#NYI     if ( $gradient ) {
#NYI         $pattern = undef;
#NYI         $fill    = undef;
#NYI     }
#NYI 
#NYI     # Set the plotarea layout.
#NYI     my $layout = $self->_get_layout_properties( $arg{layout} );
#NYI 
#NYI     $area->{_line}     = $line;
#NYI     $area->{_fill}     = $fill;
#NYI     $area->{_pattern}  = $pattern;
#NYI     $area->{_gradient} = $gradient;
#NYI     $area->{_layout}   = $layout;
#NYI 
#NYI     return $area;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_layout_properties()
#NYI #
#NYI # Convert user defined layout properties to the format required internally.
#NYI #
#NYI sub _get_layout_properties {
#NYI 
#NYI     my $self    = shift;
#NYI     my $args    = shift;
#NYI     my $is_text = shift;
#NYI     my $layout  = {};
#NYI     my @properties;
#NYI     my %allowable;
#NYI 
#NYI     return if !$args;
#NYI 
#NYI     if ( $is_text ) {
#NYI         @properties = ( 'x', 'y' );
#NYI     }
#NYI     else {
#NYI         @properties = ( 'x', 'y', 'width', 'height' );
#NYI     }
#NYI 
#NYI     # Check for valid properties.
#NYI     @allowable{@properties} = undef;
#NYI 
#NYI     for my $key ( keys %$args ) {
#NYI 
#NYI         if ( !exists $allowable{$key} ) {
#NYI             warn "Property '$key' not allowed in layout options\n";
#NYI             return;
#NYI         }
#NYI     }
#NYI 
#NYI     # Set the layout properties.
#NYI     for my $property ( @properties ) {
#NYI 
#NYI         if ( !exists $args->{$property} ) {
#NYI             warn "Property '$property' must be specified in layout options\n";
#NYI             return;
#NYI         }
#NYI 
#NYI         my $value = $args->{$property};
#NYI 
#NYI         if ( $value !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/ ) {
#NYI             warn "Property '$property' value '$value' must be numeric"
#NYI               . " in layout options\n";
#NYI             return;
#NYI         }
#NYI 
#NYI         if ( $value < 0 || $value > 1 ) {
#NYI             warn "Property '$property' value '$value' must be in range "
#NYI               . "0 < x <= 1 in layout options\n";
#NYI             return;
#NYI         }
#NYI 
#NYI         # Convert to the format used by Excel for easier testing
#NYI         $layout->{$property} = sprintf "%.17g", $value;
#NYI     }
#NYI 
#NYI     return $layout;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_points_properties()
#NYI #
#NYI # Convert user defined points properties to structure required internally.
#NYI #
#NYI sub _get_points_properties {
#NYI 
#NYI     my $self        = shift;
#NYI     my $user_points = shift;
#NYI     my @points;
#NYI 
#NYI     return unless $user_points;
#NYI 
#NYI     for my $user_point ( @$user_points ) {
#NYI 
#NYI         my $point;
#NYI 
#NYI         if ( defined $user_point ) {
#NYI 
#NYI             # Set the line properties for the point.
#NYI             my $line = $self->_get_line_properties( $user_point->{line} );
#NYI 
#NYI             # Allow 'border' as a synonym for 'line'.
#NYI             if ( $user_point->{border} ) {
#NYI                 $line = $self->_get_line_properties( $user_point->{border} );
#NYI             }
#NYI 
#NYI             # Set the fill properties for the chartarea.
#NYI             my $fill = $self->_get_fill_properties( $user_point->{fill} );
#NYI 
#NYI 
#NYI             # Set the pattern properties for the series.
#NYI             my $pattern =
#NYI               $self->_get_pattern_properties( $user_point->{pattern} );
#NYI 
#NYI             # Set the gradient fill properties for the series.
#NYI             my $gradient =
#NYI               $self->_get_gradient_properties( $user_point->{gradient} );
#NYI 
#NYI             # Pattern fill overrides solid fill.
#NYI             if ( $pattern ) {
#NYI                 $fill = undef;
#NYI             }
#NYI 
#NYI             # Gradient fill overrides solid and pattern fills.
#NYI             if ( $gradient ) {
#NYI                 $pattern = undef;
#NYI                 $fill    = undef;
#NYI             }
#NYI                         # Gradient fill overrides solid fill.
#NYI             if ( $gradient ) {
#NYI                 $fill = undef;
#NYI             }
#NYI 
#NYI             $point->{_line}     = $line;
#NYI             $point->{_fill}     = $fill;
#NYI             $point->{_pattern}  = $pattern;
#NYI             $point->{_gradient} = $gradient;
#NYI         }
#NYI 
#NYI         push @points, $point;
#NYI     }
#NYI 
#NYI     return \@points;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_display_units()
#NYI #
#NYI # Convert user defined display units to internal units.
#NYI #
#NYI sub _get_display_units {
#NYI 
#NYI     my $self          = shift;
#NYI     my $display_units = shift;
#NYI 
#NYI     return if !$display_units;
#NYI 
#NYI     my %types = (
#NYI         'hundreds'          => 'hundreds',
#NYI         'thousands'         => 'thousands',
#NYI         'ten_thousands'     => 'tenThousands',
#NYI         'hundred_thousands' => 'hundredThousands',
#NYI         'millions'          => 'millions',
#NYI         'ten_millions'      => 'tenMillions',
#NYI         'hundred_millions'  => 'hundredMillions',
#NYI         'billions'          => 'billions',
#NYI         'trillions'         => 'trillions',
#NYI     );
#NYI 
#NYI     if ( exists $types{$display_units} ) {
#NYI         $display_units = $types{$display_units};
#NYI     }
#NYI     else {
#NYI         warn "Unknown display_units type '$display_units'\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     return $display_units;
#NYI }
#NYI 
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_tick_type()
#NYI #
#NYI # Convert user tick types to internal units.
#NYI #
#NYI sub _get_tick_type {
#NYI 
#NYI     my $self      = shift;
#NYI     my $tick_type = shift;
#NYI 
#NYI     return if !$tick_type;
#NYI 
#NYI     my %types = (
#NYI         'outside' => 'out',
#NYI         'inside'  => 'in',
#NYI         'none'    => 'none',
#NYI         'cross'   => 'cross',
#NYI     );
#NYI 
#NYI     if ( exists $types{$tick_type} ) {
#NYI         $tick_type = $types{$tick_type};
#NYI     }
#NYI     else {
#NYI         warn "Unknown tick_type type '$tick_type'\n";
#NYI         return;
#NYI     }
#NYI 
#NYI     return $tick_type;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_primary_axes_series()
#NYI #
#NYI # Returns series which use the primary axes.
#NYI #
#NYI sub _get_primary_axes_series {
#NYI 
#NYI     my $self = shift;
#NYI     my @primary_axes_series;
#NYI 
#NYI     for my $series ( @{ $self->{_series} } ) {
#NYI         push @primary_axes_series, $series unless $series->{_y2_axis};
#NYI     }
#NYI 
#NYI     return @primary_axes_series;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _get_secondary_axes_series()
#NYI #
#NYI # Returns series which use the secondary axes.
#NYI #
#NYI sub _get_secondary_axes_series {
#NYI 
#NYI     my $self = shift;
#NYI     my @secondary_axes_series;
#NYI 
#NYI     for my $series ( @{ $self->{_series} } ) {
#NYI         push @secondary_axes_series, $series if $series->{_y2_axis};
#NYI     }
#NYI 
#NYI     return @secondary_axes_series;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _add_axis_ids()
#NYI #
#NYI # Add unique ids for primary or secondary axes
#NYI #
#NYI sub _add_axis_ids {
#NYI 
#NYI     my $self       = shift;
#NYI     my %args       = @_;
#NYI     my $chart_id   = 5001 + $self->{_id};
#NYI     my $axis_count = 1 + @{ $self->{_axis2_ids} } + @{ $self->{_axis_ids} };
#NYI 
#NYI     my $id1 = sprintf '%04d%04d', $chart_id, $axis_count;
#NYI     my $id2 = sprintf '%04d%04d', $chart_id, $axis_count + 1;
#NYI 
#NYI     push @{ $self->{_axis_ids} },  $id1, $id2 if $args{primary_axes};
#NYI     push @{ $self->{_axis2_ids} }, $id1, $id2 if !$args{primary_axes};
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _get_font_style_attributes.
#NYI #
#NYI # Get the font style attributes from a font hashref.
#NYI #
#NYI sub _get_font_style_attributes {
#NYI 
#NYI     my $self = shift;
#NYI     my $font = shift;
#NYI 
#NYI     return () unless $font;
#NYI 
#NYI     my @attributes;
#NYI     push @attributes, ( 'sz' => $font->{_size} )   if $font->{_size};
#NYI     push @attributes, ( 'b'  => $font->{_bold} )   if defined $font->{_bold};
#NYI     push @attributes, ( 'i'  => $font->{_italic} ) if defined $font->{_italic};
#NYI     push @attributes, ( 'u' => 'sng' ) if defined $font->{_underline};
#NYI 
#NYI     # Turn off baseline when testing fonts that don't have it.
#NYI     if ($font->{_baseline} != -1) {
#NYI         push @attributes, ( 'baseline' => $font->{_baseline} );
#NYI     }
#NYI 
#NYI     return @attributes;
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _get_font_latin_attributes.
#NYI #
#NYI # Get the font latin attributes from a font hashref.
#NYI #
#NYI sub _get_font_latin_attributes {
#NYI 
#NYI     my $self = shift;
#NYI     my $font = shift;
#NYI 
#NYI     return () unless $font;
#NYI 
#NYI     my @attributes;
#NYI     push @attributes, ( 'typeface' => $font->{_name} ) if $font->{_name};
#NYI 
#NYI     push @attributes, ( 'pitchFamily' => $font->{_pitch_family} )
#NYI       if defined $font->{_pitch_family};
#NYI 
#NYI     push @attributes, ( 'charset' => $font->{_charset} )
#NYI       if defined $font->{_charset};
#NYI 
#NYI     return @attributes;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Config data.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _set_default_properties()
#NYI #
#NYI # Setup the default properties for a chart.
#NYI #
#NYI sub _set_default_properties {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Set the default axis properties.
#NYI     $self->{_x_axis}->{_defaults} = {
#NYI         num_format      => 'General',
#NYI         major_gridlines => { visible => 0 }
#NYI     };
#NYI 
#NYI     $self->{_y_axis}->{_defaults} = {
#NYI         num_format      => 'General',
#NYI         major_gridlines => { visible => 1 }
#NYI     };
#NYI 
#NYI     $self->{_x2_axis}->{_defaults} = {
#NYI         num_format     => 'General',
#NYI         label_position => 'none',
#NYI         crossing       => 'max',
#NYI         visible        => 0
#NYI     };
#NYI 
#NYI     $self->{_y2_axis}->{_defaults} = {
#NYI         num_format      => 'General',
#NYI         major_gridlines => { visible => 0 },
#NYI         position        => 'right',
#NYI         visible         => 1
#NYI     };
#NYI 
#NYI     $self->set_x_axis();
#NYI     $self->set_y_axis();
#NYI 
#NYI     $self->set_x2_axis();
#NYI     $self->set_y2_axis();
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # _set_embedded_config_data()
#NYI #
#NYI # Setup the default configuration data for an embedded chart.
#NYI #
#NYI sub _set_embedded_config_data {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->{_embedded} = 1;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # XML writing methods.
#NYI #
#NYI ###############################################################################
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_chart_space()
#NYI #
#NYI # Write the <c:chartSpace> element.
#NYI #
#NYI sub _write_chart_space {
#NYI 
#NYI     my $self    = shift;
#NYI     my $schema  = 'http://schemas.openxmlformats.org/';
#NYI     my $xmlns_c = $schema . 'drawingml/2006/chart';
#NYI     my $xmlns_a = $schema . 'drawingml/2006/main';
#NYI     my $xmlns_r = $schema . 'officeDocument/2006/relationships';
#NYI 
#NYI     my @attributes = (
#NYI         'xmlns:c' => $xmlns_c,
#NYI         'xmlns:a' => $xmlns_a,
#NYI         'xmlns:r' => $xmlns_r,
#NYI     );
#NYI 
#NYI     $self->xml_start_tag( 'c:chartSpace', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_lang()
#NYI #
#NYI # Write the <c:lang> element.
#NYI #
#NYI sub _write_lang {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 'en-US';
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:lang', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_style()
#NYI #
#NYI # Write the <c:style> element.
#NYI #
#NYI sub _write_style {
#NYI 
#NYI     my $self     = shift;
#NYI     my $style_id = $self->{_style_id};
#NYI 
#NYI     # Don't write an element for the default style, 2.
#NYI     return if $style_id == 2;
#NYI 
#NYI     my @attributes = ( 'val' => $style_id );
#NYI 
#NYI     $self->xml_empty_tag( 'c:style', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_chart()
#NYI #
#NYI # Write the <c:chart> element.
#NYI #
#NYI sub _write_chart {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:chart' );
#NYI 
#NYI     # Write the chart title elements.
#NYI 
#NYI     if ( $self->{_title_none} ) {
#NYI 
#NYI         # Turn off the title.
#NYI         $self->_write_auto_title_deleted();
#NYI     }
#NYI     else {
#NYI         my $title;
#NYI         if ( $title = $self->{_title_formula} ) {
#NYI             $self->_write_title_formula(
#NYI 
#NYI                 $title,
#NYI                 $self->{_title_data_id},
#NYI                 undef,
#NYI                 $self->{_title_font},
#NYI                 $self->{_title_layout},
#NYI                 $self->{_title_overlay}
#NYI             );
#NYI         }
#NYI         elsif ( $title = $self->{_title_name} ) {
#NYI             $self->_write_title_rich(
#NYI 
#NYI                 $title,
#NYI                 undef,
#NYI                 $self->{_title_font},
#NYI                 $self->{_title_layout},
#NYI                 $self->{_title_overlay}
#NYI             );
#NYI         }
#NYI     }
#NYI 
#NYI     # Write the c:plotArea element.
#NYI     $self->_write_plot_area();
#NYI 
#NYI     # Write the c:legend element.
#NYI     $self->_write_legend();
#NYI 
#NYI     # Write the c:plotVisOnly element.
#NYI     $self->_write_plot_vis_only();
#NYI 
#NYI     # Write the c:dispBlanksAs element.
#NYI     $self->_write_disp_blanks_as();
#NYI 
#NYI     $self->xml_end_tag( 'c:chart' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_disp_blanks_as()
#NYI #
#NYI # Write the <c:dispBlanksAs> element.
#NYI #
#NYI sub _write_disp_blanks_as {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = $self->{_show_blanks};
#NYI 
#NYI     # Ignore the default value.
#NYI     return if $val eq 'gap';
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:dispBlanksAs', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_plot_area()
#NYI #
#NYI # Write the <c:plotArea> element.
#NYI #
#NYI sub _write_plot_area {
#NYI 
#NYI     my $self = shift;
#NYI     my $second_chart = $self->{_combined};
#NYI 
#NYI     $self->xml_start_tag( 'c:plotArea' );
#NYI 
#NYI     # Write the c:layout element.
#NYI     $self->_write_layout( $self->{_plotarea}->{_layout}, 'plot' );
#NYI 
#NYI     # Write the subclass chart type elements for primary and secondary axes.
#NYI     $self->_write_chart_type( primary_axes => 1 );
#NYI     $self->_write_chart_type( primary_axes => 0 );
#NYI 
#NYI 
#NYI     # Configure a combined chart if present.
#NYI     if ( $second_chart ) {
#NYI 
#NYI         # Secondary axis has unique id otherwise use same as primary.
#NYI         if ( $second_chart->{_is_secondary} ) {
#NYI             $second_chart->{_id} = 1000 + $self->{_id};
#NYI         }
#NYI         else {
#NYI             $second_chart->{_id} = $self->{_id};
#NYI         }
#NYI 
#NYI         # Shart the same filehandle for writing.
#NYI         $second_chart->{_fh} = $self->{_fh};
#NYI 
#NYI         # Share series index with primary chart.
#NYI         $second_chart->{_series_index} = $self->{_series_index};
#NYI 
#NYI         # Write the subclass chart type elements for combined chart.
#NYI         $second_chart->_write_chart_type( primary_axes => 1 );
#NYI         $second_chart->_write_chart_type( primary_axes => 0 );
#NYI     }
#NYI 
#NYI     # Write the category and value elements for the primary axes.
#NYI     my @args = (
#NYI         x_axis   => $self->{_x_axis},
#NYI         y_axis   => $self->{_y_axis},
#NYI         axis_ids => $self->{_axis_ids}
#NYI     );
#NYI 
#NYI     if ( $self->{_date_category} ) {
#NYI         $self->_write_date_axis( @args );
#NYI     }
#NYI     else {
#NYI         $self->_write_cat_axis( @args );
#NYI     }
#NYI 
#NYI     $self->_write_val_axis( @args );
#NYI 
#NYI     # Write the category and value elements for the secondary axes.
#NYI     @args = (
#NYI         x_axis   => $self->{_x2_axis},
#NYI         y_axis   => $self->{_y2_axis},
#NYI         axis_ids => $self->{_axis2_ids}
#NYI     );
#NYI 
#NYI     $self->_write_val_axis( @args );
#NYI 
#NYI     # Write the secondary axis for the secondary chart.
#NYI     if ( $second_chart && $second_chart->{_is_secondary} ) {
#NYI 
#NYI         @args = (
#NYI              x_axis   => $second_chart->{_x2_axis},
#NYI              y_axis   => $second_chart->{_y2_axis},
#NYI              axis_ids => $second_chart->{_axis2_ids}
#NYI             );
#NYI 
#NYI         $second_chart->_write_val_axis( @args );
#NYI     }
#NYI 
#NYI 
#NYI     if ( $self->{_date_category} ) {
#NYI         $self->_write_date_axis( @args );
#NYI     }
#NYI     else {
#NYI         $self->_write_cat_axis( @args );
#NYI     }
#NYI 
#NYI     # Write the c:dTable element.
#NYI     $self->_write_d_table();
#NYI 
#NYI     # Write the c:spPr element for the plotarea formatting.
#NYI     $self->_write_sp_pr( $self->{_plotarea} );
#NYI 
#NYI     $self->xml_end_tag( 'c:plotArea' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_layout()
#NYI #
#NYI # Write the <c:layout> element.
#NYI #
#NYI sub _write_layout {
#NYI 
#NYI     my $self   = shift;
#NYI     my $layout = shift;
#NYI     my $type   = shift;
#NYI 
#NYI     if ( !$layout ) {
#NYI         # Automatic layout.
#NYI         $self->xml_empty_tag( 'c:layout' );
#NYI     }
#NYI     else {
#NYI         # User defined manual layout.
#NYI         $self->xml_start_tag( 'c:layout' );
#NYI         $self->_write_manual_layout( $layout, $type );
#NYI         $self->xml_end_tag( 'c:layout' );
#NYI     }
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_manual_layout()
#NYI #
#NYI # Write the <c:manualLayout> element.
#NYI #
#NYI sub _write_manual_layout {
#NYI 
#NYI     my $self   = shift;
#NYI     my $layout = shift;
#NYI     my $type   = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:manualLayout' );
#NYI 
#NYI     # Plotarea has a layoutTarget element.
#NYI     if ( $type eq 'plot' ) {
#NYI         $self->xml_empty_tag( 'c:layoutTarget', ( 'val' => 'inner' ) );
#NYI     }
#NYI 
#NYI     # Set the x, y positions.
#NYI     $self->xml_empty_tag( 'c:xMode', ( 'val' => 'edge' ) );
#NYI     $self->xml_empty_tag( 'c:yMode', ( 'val' => 'edge' ) );
#NYI     $self->xml_empty_tag( 'c:x', ( 'val' => $layout->{x} ) );
#NYI     $self->xml_empty_tag( 'c:y', ( 'val' => $layout->{y} ) );
#NYI 
#NYI     # For plotarea and legend set the width and height.
#NYI     if ( $type ne 'text' ) {
#NYI         $self->xml_empty_tag( 'c:w', ( 'val' => $layout->{width} ) );
#NYI         $self->xml_empty_tag( 'c:h', ( 'val' => $layout->{height} ) );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:manualLayout' );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_chart_type()
#NYI #
#NYI # Write the chart type element. This method should be overridden by the
#NYI # subclasses.
#NYI #
#NYI sub _write_chart_type {
#NYI 
#NYI     my $self = shift;
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_grouping()
#NYI #
#NYI # Write the <c:grouping> element.
#NYI #
#NYI sub _write_grouping {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:grouping', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_series()
#NYI #
#NYI # Write the series elements.
#NYI #
#NYI sub _write_series {
#NYI 
#NYI     my $self   = shift;
#NYI     my $series = shift;
#NYI 
#NYI     $self->_write_ser( $series );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_ser()
#NYI #
#NYI # Write the <c:ser> element.
#NYI #
#NYI sub _write_ser {
#NYI 
#NYI     my $self   = shift;
#NYI     my $series = shift;
#NYI     my $index  = $self->{_series_index}++;
#NYI 
#NYI     $self->xml_start_tag( 'c:ser' );
#NYI 
#NYI     # Write the c:idx element.
#NYI     $self->_write_idx( $index );
#NYI 
#NYI     # Write the c:order element.
#NYI     $self->_write_order( $index );
#NYI 
#NYI     # Write the series name.
#NYI     $self->_write_series_name( $series );
#NYI 
#NYI     # Write the c:spPr element.
#NYI     $self->_write_sp_pr( $series );
#NYI 
#NYI     # Write the c:marker element.
#NYI     $self->_write_marker( $series->{_marker} );
#NYI 
#NYI     # Write the c:invertIfNegative element.
#NYI     $self->_write_c_invert_if_negative( $series->{_invert_if_neg} );
#NYI 
#NYI     # Write the c:dPt element.
#NYI     $self->_write_d_pt( $series->{_points} );
#NYI 
#NYI     # Write the c:dLbls element.
#NYI     $self->_write_d_lbls( $series->{_labels} );
#NYI 
#NYI     # Write the c:trendline element.
#NYI     $self->_write_trendline( $series->{_trendline} );
#NYI 
#NYI     # Write the c:errBars element.
#NYI     $self->_write_error_bars( $series->{_error_bars} );
#NYI 
#NYI     # Write the c:cat element.
#NYI     $self->_write_cat( $series );
#NYI 
#NYI     # Write the c:val element.
#NYI     $self->_write_val( $series );
#NYI 
#NYI     # Write the c:smooth element.
#NYI     if ( $self->{_smooth_allowed} ) {
#NYI         $self->_write_c_smooth( $series->{_smooth} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:ser' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_idx()
#NYI #
#NYI # Write the <c:idx> element.
#NYI #
#NYI sub _write_idx {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:idx', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_order()
#NYI #
#NYI # Write the <c:order> element.
#NYI #
#NYI sub _write_order {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:order', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_series_name()
#NYI #
#NYI # Write the series name.
#NYI #
#NYI sub _write_series_name {
#NYI 
#NYI     my $self   = shift;
#NYI     my $series = shift;
#NYI 
#NYI     my $name;
#NYI     if ( $name = $series->{_name_formula} ) {
#NYI         $self->_write_tx_formula( $name, $series->{_name_id} );
#NYI     }
#NYI     elsif ( $name = $series->{_name} ) {
#NYI         $self->_write_tx_value( $name );
#NYI     }
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_cat()
#NYI #
#NYI # Write the <c:cat> element.
#NYI #
#NYI sub _write_cat {
#NYI 
#NYI     my $self    = shift;
#NYI     my $series  = shift;
#NYI     my $formula = $series->{_categories};
#NYI     my $data_id = $series->{_cat_data_id};
#NYI     my $data;
#NYI 
#NYI     if ( defined $data_id ) {
#NYI         $data = $self->{_formula_data}->[$data_id];
#NYI     }
#NYI 
#NYI     # Ignore <c:cat> elements for charts without category values.
#NYI     return unless $formula;
#NYI 
#NYI     $self->xml_start_tag( 'c:cat' );
#NYI 
#NYI     # Check the type of cached data.
#NYI     my $type = $self->_get_data_type( $data );
#NYI 
#NYI     if ( $type eq 'str' ) {
#NYI 
#NYI         $self->{_cat_has_num_fmt} = 0;
#NYI 
#NYI         # Write the c:numRef element.
#NYI         $self->_write_str_ref( $formula, $data, $type );
#NYI     }
#NYI     elsif ( $type eq 'multi_str') {
#NYI 
#NYI         $self->{_cat_has_num_fmt} = 0;
#NYI 
#NYI         # Write the c:multiLvLStrRef element.
#NYI         $self->_write_multi_lvl_str_ref( $formula, $data );
#NYI     }
#NYI     else {
#NYI 
#NYI         $self->{_cat_has_num_fmt} = 1;
#NYI 
#NYI         # Write the c:numRef element.
#NYI         $self->_write_num_ref( $formula, $data, $type );
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_end_tag( 'c:cat' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_val()
#NYI #
#NYI # Write the <c:val> element.
#NYI #
#NYI sub _write_val {
#NYI 
#NYI     my $self    = shift;
#NYI     my $series  = shift;
#NYI     my $formula = $series->{_values};
#NYI     my $data_id = $series->{_val_data_id};
#NYI     my $data    = $self->{_formula_data}->[$data_id];
#NYI 
#NYI     $self->xml_start_tag( 'c:val' );
#NYI 
#NYI     # Unlike Cat axes data should only be numeric.
#NYI 
#NYI     # Write the c:numRef element.
#NYI     $self->_write_num_ref( $formula, $data, 'num' );
#NYI 
#NYI     $self->xml_end_tag( 'c:val' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_num_ref()
#NYI #
#NYI # Write the <c:numRef> element.
#NYI #
#NYI sub _write_num_ref {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI     my $data    = shift;
#NYI     my $type    = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:numRef' );
#NYI 
#NYI     # Write the c:f element.
#NYI     $self->_write_series_formula( $formula );
#NYI 
#NYI     if ( $type eq 'num' ) {
#NYI 
#NYI         # Write the c:numCache element.
#NYI         $self->_write_num_cache( $data );
#NYI     }
#NYI     elsif ( $type eq 'str' ) {
#NYI 
#NYI         # Write the c:strCache element.
#NYI         $self->_write_str_cache( $data );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:numRef' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_str_ref()
#NYI #
#NYI # Write the <c:strRef> element.
#NYI #
#NYI sub _write_str_ref {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI     my $data    = shift;
#NYI     my $type    = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:strRef' );
#NYI 
#NYI     # Write the c:f element.
#NYI     $self->_write_series_formula( $formula );
#NYI 
#NYI     if ( $type eq 'num' ) {
#NYI 
#NYI         # Write the c:numCache element.
#NYI         $self->_write_num_cache( $data );
#NYI     }
#NYI     elsif ( $type eq 'str' ) {
#NYI 
#NYI         # Write the c:strCache element.
#NYI         $self->_write_str_cache( $data );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:strRef' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_multi_lvl_str_ref()
#NYI #
#NYI # Write the <c:multiLvLStrRef> element.
#NYI #
#NYI sub _write_multi_lvl_str_ref {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI     my $data    = shift;
#NYI     my $count   = @$data;
#NYI 
#NYI     return if !$count;
#NYI 
#NYI     $self->xml_start_tag( 'c:multiLvlStrRef' );
#NYI 
#NYI     # Write the c:f element.
#NYI     $self->_write_series_formula( $formula );
#NYI 
#NYI     $self->xml_start_tag( 'c:multiLvlStrCache' );
#NYI 
#NYI     # Write the c:ptCount element.
#NYI     $count = @{ $data->[-1] };
#NYI     $self->_write_pt_count( $count );
#NYI 
#NYI     # Write the data arrays in reverse order.
#NYI     for my $aref ( reverse @$data ) {
#NYI         $self->xml_start_tag( 'c:lvl' );
#NYI 
#NYI         for my $i ( 0 .. @$aref - 1 ) {
#NYI             # Write the c:pt element.
#NYI             $self->_write_pt( $i, $aref->[$i] );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'c:lvl' );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:multiLvlStrCache' );
#NYI 
#NYI     $self->xml_end_tag( 'c:multiLvlStrRef' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_series_formula()
#NYI #
#NYI # Write the <c:f> element.
#NYI #
#NYI sub _write_series_formula {
#NYI 
#NYI     my $self    = shift;
#NYI     my $formula = shift;
#NYI 
#NYI     # Strip the leading '=' from the formula.
#NYI     $formula =~ s/^=//;
#NYI 
#NYI     $self->xml_data_element( 'c:f', $formula );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_axis_ids()
#NYI #
#NYI # Write the <c:axId> elements for the primary or secondary axes.
#NYI #
#NYI sub _write_axis_ids {
#NYI 
#NYI     my $self = shift;
#NYI     my %args = @_;
#NYI 
#NYI     # Generate the axis ids.
#NYI     $self->_add_axis_ids( %args );
#NYI 
#NYI     if ( $args{primary_axes} ) {
#NYI 
#NYI         # Write the axis ids for the primary axes.
#NYI         $self->_write_axis_id( $self->{_axis_ids}->[0] );
#NYI         $self->_write_axis_id( $self->{_axis_ids}->[1] );
#NYI     }
#NYI     else {
#NYI         # Write the axis ids for the secondary axes.
#NYI         $self->_write_axis_id( $self->{_axis2_ids}->[0] );
#NYI         $self->_write_axis_id( $self->{_axis2_ids}->[1] );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_axis_id()
#NYI #
#NYI # Write the <c:axId> element.
#NYI #
#NYI sub _write_axis_id {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:axId', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_cat_axis()
#NYI #
#NYI # Write the <c:catAx> element. Usually the X axis.
#NYI #
#NYI sub _write_cat_axis {
#NYI 
#NYI     my $self     = shift;
#NYI     my %args     = @_;
#NYI     my $x_axis   = $args{x_axis};
#NYI     my $y_axis   = $args{y_axis};
#NYI     my $axis_ids = $args{axis_ids};
#NYI 
#NYI     # if there are no axis_ids then we don't need to write this element
#NYI     return unless $axis_ids;
#NYI     return unless scalar @$axis_ids;
#NYI 
#NYI     my $position = $self->{_cat_axis_position};
#NYI     my $horiz    = $self->{_horiz_cat_axis};
#NYI 
#NYI     # Overwrite the default axis position with a user supplied value.
#NYI     $position = $x_axis->{_position} || $position;
#NYI 
#NYI     $self->xml_start_tag( 'c:catAx' );
#NYI 
#NYI     $self->_write_axis_id( $axis_ids->[0] );
#NYI 
#NYI     # Write the c:scaling element.
#NYI     $self->_write_scaling( $x_axis->{_reverse} );
#NYI 
#NYI     $self->_write_delete( 1 ) unless $x_axis->{_visible};
#NYI 
#NYI     # Write the c:axPos element.
#NYI     $self->_write_axis_pos( $position, $y_axis->{_reverse} );
#NYI 
#NYI     # Write the c:majorGridlines element.
#NYI     $self->_write_major_gridlines( $x_axis->{_major_gridlines} );
#NYI 
#NYI     # Write the c:minorGridlines element.
#NYI     $self->_write_minor_gridlines( $x_axis->{_minor_gridlines} );
#NYI 
#NYI     # Write the axis title elements.
#NYI     my $title;
#NYI     if ( $title = $x_axis->{_formula} ) {
#NYI 
#NYI         $self->_write_title_formula( $title, $x_axis->{_data_id}, $horiz,
#NYI             $x_axis->{_name_font}, $x_axis->{_layout} );
#NYI     }
#NYI     elsif ( $title = $x_axis->{_name} ) {
#NYI         $self->_write_title_rich( $title, $horiz, $x_axis->{_name_font},
#NYI             $x_axis->{_layout} );
#NYI     }
#NYI 
#NYI     # Write the c:numFmt element.
#NYI     $self->_write_cat_number_format( $x_axis );
#NYI 
#NYI     # Write the c:majorTickMark element.
#NYI     $self->_write_major_tick_mark( $x_axis->{_major_tick_mark} );
#NYI 
#NYI     # Write the c:minorTickMark element.
#NYI     $self->_write_minor_tick_mark( $x_axis->{_minor_tick_mark} );
#NYI 
#NYI     # Write the c:tickLblPos element.
#NYI     $self->_write_tick_label_pos( $x_axis->{_label_position} );
#NYI 
#NYI     # Write the c:spPr element for the axis line.
#NYI     $self->_write_sp_pr( $x_axis );
#NYI 
#NYI     # Write the axis font elements.
#NYI     $self->_write_axis_font( $x_axis->{_num_font} );
#NYI 
#NYI     # Write the c:crossAx element.
#NYI     $self->_write_cross_axis( $axis_ids->[1] );
#NYI 
#NYI     if ( $self->{_show_crosses} || $x_axis->{_visible} ) {
#NYI 
#NYI         # Note, the category crossing comes from the value axis.
#NYI         if ( !defined $y_axis->{_crossing} || $y_axis->{_crossing} eq 'max' ) {
#NYI 
#NYI             # Write the c:crosses element.
#NYI             $self->_write_crosses( $y_axis->{_crossing} );
#NYI         }
#NYI         else {
#NYI 
#NYI             # Write the c:crossesAt element.
#NYI             $self->_write_c_crosses_at( $y_axis->{_crossing} );
#NYI         }
#NYI     }
#NYI 
#NYI     # Write the c:auto element.
#NYI     if (!$x_axis->{_text_axis}) {
#NYI         $self->_write_auto( 1 );
#NYI     }
#NYI 
#NYI     # Write the c:labelAlign element.
#NYI     $self->_write_label_align( 'ctr' );
#NYI 
#NYI     # Write the c:labelOffset element.
#NYI     $self->_write_label_offset( 100 );
#NYI 
#NYI     # Write the c:tickLblSkip element.
#NYI     $self->_write_tick_lbl_skip( $x_axis->{_interval_unit} );
#NYI 
#NYI     # Write the c:tickMarkSkip element.
#NYI     $self->_write_tick_mark_skip( $x_axis->{_interval_tick} );
#NYI 
#NYI     $self->xml_end_tag( 'c:catAx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_val_axis()
#NYI #
#NYI # Write the <c:valAx> element. Usually the Y axis.
#NYI #
#NYI # TODO. Maybe should have a _write_cat_val_axis() method as well for scatter.
#NYI #
#NYI sub _write_val_axis {
#NYI 
#NYI     my $self     = shift;
#NYI     my %args     = @_;
#NYI     my $x_axis   = $args{x_axis};
#NYI     my $y_axis   = $args{y_axis};
#NYI     my $axis_ids = $args{axis_ids};
#NYI     my $position = $args{position} || $self->{_val_axis_position};
#NYI     my $horiz    = $self->{_horiz_val_axis};
#NYI 
#NYI     return unless $axis_ids && scalar @$axis_ids;
#NYI 
#NYI     # Overwrite the default axis position with a user supplied value.
#NYI     $position = $y_axis->{_position} || $position;
#NYI 
#NYI     $self->xml_start_tag( 'c:valAx' );
#NYI 
#NYI     $self->_write_axis_id( $axis_ids->[1] );
#NYI 
#NYI     # Write the c:scaling element.
#NYI     $self->_write_scaling(
#NYI         $y_axis->{_reverse}, $y_axis->{_min},
#NYI         $y_axis->{_max},     $y_axis->{_log_base}
#NYI     );
#NYI 
#NYI     $self->_write_delete( 1 ) unless $y_axis->{_visible};
#NYI 
#NYI     # Write the c:axPos element.
#NYI     $self->_write_axis_pos( $position, $x_axis->{_reverse} );
#NYI 
#NYI     # Write the c:majorGridlines element.
#NYI     $self->_write_major_gridlines( $y_axis->{_major_gridlines} );
#NYI 
#NYI     # Write the c:minorGridlines element.
#NYI     $self->_write_minor_gridlines( $y_axis->{_minor_gridlines} );
#NYI 
#NYI     # Write the axis title elements.
#NYI     my $title;
#NYI     if ( $title = $y_axis->{_formula} ) {
#NYI         $self->_write_title_formula( $title, $y_axis->{_data_id}, $horiz,
#NYI             $y_axis->{_name_font}, $y_axis->{_layout} );
#NYI     }
#NYI     elsif ( $title = $y_axis->{_name} ) {
#NYI         $self->_write_title_rich( $title, $horiz, $y_axis->{_name_font},
#NYI             $y_axis->{_layout} );
#NYI     }
#NYI 
#NYI     # Write the c:numberFormat element.
#NYI     $self->_write_number_format( $y_axis );
#NYI 
#NYI     # Write the c:majorTickMark element.
#NYI     $self->_write_major_tick_mark( $y_axis->{_major_tick_mark} );
#NYI 
#NYI     # Write the c:minorTickMark element.
#NYI     $self->_write_minor_tick_mark( $y_axis->{_minor_tick_mark} );
#NYI 
#NYI     # Write the c:tickLblPos element.
#NYI     $self->_write_tick_label_pos( $y_axis->{_label_position} );
#NYI 
#NYI     # Write the c:spPr element for the axis line.
#NYI     $self->_write_sp_pr( $y_axis );
#NYI 
#NYI     # Write the axis font elements.
#NYI     $self->_write_axis_font( $y_axis->{_num_font} );
#NYI 
#NYI     # Write the c:crossAx element.
#NYI     $self->_write_cross_axis( $axis_ids->[0] );
#NYI 
#NYI     # Note, the category crossing comes from the value axis.
#NYI     if ( !defined $x_axis->{_crossing} || $x_axis->{_crossing} eq 'max' ) {
#NYI 
#NYI         # Write the c:crosses element.
#NYI         $self->_write_crosses( $x_axis->{_crossing} );
#NYI     }
#NYI     else {
#NYI 
#NYI         # Write the c:crossesAt element.
#NYI         $self->_write_c_crosses_at( $x_axis->{_crossing} );
#NYI     }
#NYI 
#NYI     # Write the c:crossBetween element.
#NYI     $self->_write_cross_between( $x_axis->{_position_axis} );
#NYI 
#NYI     # Write the c:majorUnit element.
#NYI     $self->_write_c_major_unit( $y_axis->{_major_unit} );
#NYI 
#NYI     # Write the c:minorUnit element.
#NYI     $self->_write_c_minor_unit( $y_axis->{_minor_unit} );
#NYI 
#NYI     # Write the c:dispUnits element.
#NYI     $self->_write_disp_units( $y_axis->{_display_units},
#NYI         $y_axis->{_display_units_visible} );
#NYI 
#NYI     $self->xml_end_tag( 'c:valAx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_cat_val_axis()
#NYI #
#NYI # Write the <c:valAx> element. This is for the second valAx in scatter plots.
#NYI # Usually the X axis.
#NYI #
#NYI sub _write_cat_val_axis {
#NYI 
#NYI     my $self     = shift;
#NYI     my %args     = @_;
#NYI     my $x_axis   = $args{x_axis};
#NYI     my $y_axis   = $args{y_axis};
#NYI     my $axis_ids = $args{axis_ids};
#NYI     my $position = $args{position} || $self->{_val_axis_position};
#NYI     my $horiz    = $self->{_horiz_val_axis};
#NYI 
#NYI     return unless $axis_ids && scalar @$axis_ids;
#NYI 
#NYI     # Overwrite the default axis position with a user supplied value.
#NYI     $position = $x_axis->{_position} || $position;
#NYI 
#NYI     $self->xml_start_tag( 'c:valAx' );
#NYI 
#NYI     $self->_write_axis_id( $axis_ids->[0] );
#NYI 
#NYI     # Write the c:scaling element.
#NYI     $self->_write_scaling(
#NYI         $x_axis->{_reverse}, $x_axis->{_min},
#NYI         $x_axis->{_max},     $x_axis->{_log_base}
#NYI     );
#NYI 
#NYI     $self->_write_delete( 1 ) unless $x_axis->{_visible};
#NYI 
#NYI     # Write the c:axPos element.
#NYI     $self->_write_axis_pos( $position, $y_axis->{_reverse} );
#NYI 
#NYI     # Write the c:majorGridlines element.
#NYI     $self->_write_major_gridlines( $x_axis->{_major_gridlines} );
#NYI 
#NYI     # Write the c:minorGridlines element.
#NYI     $self->_write_minor_gridlines( $x_axis->{_minor_gridlines} );
#NYI 
#NYI     # Write the axis title elements.
#NYI     my $title;
#NYI     if ( $title = $x_axis->{_formula} ) {
#NYI         $self->_write_title_formula( $title, $x_axis->{_data_id}, $horiz,
#NYI             $x_axis->{_name_font}, $x_axis->{_layout} );
#NYI     }
#NYI     elsif ( $title = $x_axis->{_name} ) {
#NYI         $self->_write_title_rich( $title, $horiz, $x_axis->{_name_font},
#NYI             $x_axis->{_layout} );
#NYI     }
#NYI 
#NYI     # Write the c:numberFormat element.
#NYI     $self->_write_number_format( $x_axis );
#NYI 
#NYI     # Write the c:majorTickMark element.
#NYI     $self->_write_major_tick_mark( $x_axis->{_major_tick_mark} );
#NYI 
#NYI     # Write the c:minorTickMark element.
#NYI     $self->_write_minor_tick_mark( $x_axis->{_minor_tick_mark} );
#NYI 
#NYI     # Write the c:tickLblPos element.
#NYI     $self->_write_tick_label_pos( $x_axis->{_label_position} );
#NYI 
#NYI     # Write the c:spPr element for the axis line.
#NYI     $self->_write_sp_pr( $x_axis );
#NYI 
#NYI     # Write the axis font elements.
#NYI     $self->_write_axis_font( $x_axis->{_num_font} );
#NYI 
#NYI     # Write the c:crossAx element.
#NYI     $self->_write_cross_axis( $axis_ids->[1] );
#NYI 
#NYI     # Note, the category crossing comes from the value axis.
#NYI     if ( !defined $y_axis->{_crossing} || $y_axis->{_crossing} eq 'max' ) {
#NYI 
#NYI         # Write the c:crosses element.
#NYI         $self->_write_crosses( $y_axis->{_crossing} );
#NYI     }
#NYI     else {
#NYI 
#NYI         # Write the c:crossesAt element.
#NYI         $self->_write_c_crosses_at( $y_axis->{_crossing} );
#NYI     }
#NYI 
#NYI     # Write the c:crossBetween element.
#NYI     $self->_write_cross_between( $y_axis->{_position_axis} );
#NYI 
#NYI     # Write the c:majorUnit element.
#NYI     $self->_write_c_major_unit( $x_axis->{_major_unit} );
#NYI 
#NYI     # Write the c:minorUnit element.
#NYI     $self->_write_c_minor_unit( $x_axis->{_minor_unit} );
#NYI 
#NYI     # Write the c:dispUnits element.
#NYI     $self->_write_disp_units( $x_axis->{_display_units},
#NYI         $x_axis->{_display_units_visible} );
#NYI 
#NYI     $self->xml_end_tag( 'c:valAx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_date_axis()
#NYI #
#NYI # Write the <c:dateAx> element. Usually the X axis.
#NYI #
#NYI sub _write_date_axis {
#NYI 
#NYI     my $self     = shift;
#NYI     my %args     = @_;
#NYI     my $x_axis   = $args{x_axis};
#NYI     my $y_axis   = $args{y_axis};
#NYI     my $axis_ids = $args{axis_ids};
#NYI 
#NYI     return unless $axis_ids && scalar @$axis_ids;
#NYI 
#NYI     my $position = $self->{_cat_axis_position};
#NYI 
#NYI     # Overwrite the default axis position with a user supplied value.
#NYI     $position = $x_axis->{_position} || $position;
#NYI 
#NYI     $self->xml_start_tag( 'c:dateAx' );
#NYI 
#NYI     $self->_write_axis_id( $axis_ids->[0] );
#NYI 
#NYI     # Write the c:scaling element.
#NYI     $self->_write_scaling(
#NYI         $x_axis->{_reverse}, $x_axis->{_min},
#NYI         $x_axis->{_max},     $x_axis->{_log_base}
#NYI     );
#NYI 
#NYI     $self->_write_delete( 1 ) unless $x_axis->{_visible};
#NYI 
#NYI     # Write the c:axPos element.
#NYI     $self->_write_axis_pos( $position, $y_axis->{_reverse} );
#NYI 
#NYI     # Write the c:majorGridlines element.
#NYI     $self->_write_major_gridlines( $x_axis->{_major_gridlines} );
#NYI 
#NYI     # Write the c:minorGridlines element.
#NYI     $self->_write_minor_gridlines( $x_axis->{_minor_gridlines} );
#NYI 
#NYI     # Write the axis title elements.
#NYI     my $title;
#NYI     if ( $title = $x_axis->{_formula} ) {
#NYI         $self->_write_title_formula( $title, $x_axis->{_data_id}, undef,
#NYI             $x_axis->{_name_font}, $x_axis->{_layout} );
#NYI     }
#NYI     elsif ( $title = $x_axis->{_name} ) {
#NYI         $self->_write_title_rich( $title, undef, $x_axis->{_name_font},
#NYI             $x_axis->{_layout} );
#NYI     }
#NYI 
#NYI     # Write the c:numFmt element.
#NYI     $self->_write_number_format( $x_axis );
#NYI 
#NYI     # Write the c:majorTickMark element.
#NYI     $self->_write_major_tick_mark( $x_axis->{_major_tick_mark} );
#NYI 
#NYI     # Write the c:minorTickMark element.
#NYI     $self->_write_minor_tick_mark( $x_axis->{_minor_tick_mark} );
#NYI 
#NYI     # Write the c:tickLblPos element.
#NYI     $self->_write_tick_label_pos( $x_axis->{_label_position} );
#NYI 
#NYI     # Write the c:spPr element for the axis line.
#NYI     $self->_write_sp_pr( $x_axis );
#NYI 
#NYI     # Write the axis font elements.
#NYI     $self->_write_axis_font( $x_axis->{_num_font} );
#NYI 
#NYI     # Write the c:crossAx element.
#NYI     $self->_write_cross_axis( $axis_ids->[1] );
#NYI 
#NYI     if ( $self->{_show_crosses} || $x_axis->{_visible} ) {
#NYI 
#NYI         # Note, the category crossing comes from the value axis.
#NYI         if ( !defined $y_axis->{_crossing} || $y_axis->{_crossing} eq 'max' ) {
#NYI 
#NYI             # Write the c:crosses element.
#NYI             $self->_write_crosses( $y_axis->{_crossing} );
#NYI         }
#NYI         else {
#NYI 
#NYI             # Write the c:crossesAt element.
#NYI             $self->_write_c_crosses_at( $y_axis->{_crossing} );
#NYI         }
#NYI     }
#NYI 
#NYI     # Write the c:auto element.
#NYI     $self->_write_auto( 1 );
#NYI 
#NYI     # Write the c:labelOffset element.
#NYI     $self->_write_label_offset( 100 );
#NYI 
#NYI     # Write the c:tickLblSkip element.
#NYI     $self->_write_tick_lbl_skip( $x_axis->{_interval_unit} );
#NYI 
#NYI     # Write the c:tickMarkSkip element.
#NYI     $self->_write_tick_mark_skip( $x_axis->{_interval_tick} );
#NYI 
#NYI     # Write the c:majorUnit element.
#NYI     $self->_write_c_major_unit( $x_axis->{_major_unit} );
#NYI 
#NYI     # Write the c:majorTimeUnit element.
#NYI     if ( defined $x_axis->{_major_unit} ) {
#NYI         $self->_write_c_major_time_unit( $x_axis->{_major_unit_type} );
#NYI     }
#NYI 
#NYI     # Write the c:minorUnit element.
#NYI     $self->_write_c_minor_unit( $x_axis->{_minor_unit} );
#NYI 
#NYI     # Write the c:minorTimeUnit element.
#NYI     if ( defined $x_axis->{_minor_unit} ) {
#NYI         $self->_write_c_minor_time_unit( $x_axis->{_minor_unit_type} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:dateAx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_scaling()
#NYI #
#NYI # Write the <c:scaling> element.
#NYI #
#NYI sub _write_scaling {
#NYI 
#NYI     my $self     = shift;
#NYI     my $reverse  = shift;
#NYI     my $min      = shift;
#NYI     my $max      = shift;
#NYI     my $log_base = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:scaling' );
#NYI 
#NYI     # Write the c:logBase element.
#NYI     $self->_write_c_log_base( $log_base );
#NYI 
#NYI     # Write the c:orientation element.
#NYI     $self->_write_orientation( $reverse );
#NYI 
#NYI     # Write the c:max element.
#NYI     $self->_write_c_max( $max );
#NYI 
#NYI     # Write the c:min element.
#NYI     $self->_write_c_min( $min );
#NYI 
#NYI     $self->xml_end_tag( 'c:scaling' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_log_base()
#NYI #
#NYI # Write the <c:logBase> element.
#NYI #
#NYI sub _write_c_log_base {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:logBase', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_orientation()
#NYI #
#NYI # Write the <c:orientation> element.
#NYI #
#NYI sub _write_orientation {
#NYI 
#NYI     my $self    = shift;
#NYI     my $reverse = shift;
#NYI     my $val     = 'minMax';
#NYI 
#NYI     $val = 'maxMin' if $reverse;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:orientation', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_max()
#NYI #
#NYI # Write the <c:max> element.
#NYI #
#NYI sub _write_c_max {
#NYI 
#NYI     my $self = shift;
#NYI     my $max  = shift;
#NYI 
#NYI     return unless defined $max;
#NYI 
#NYI     my @attributes = ( 'val' => $max );
#NYI 
#NYI     $self->xml_empty_tag( 'c:max', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_min()
#NYI #
#NYI # Write the <c:min> element.
#NYI #
#NYI sub _write_c_min {
#NYI 
#NYI     my $self = shift;
#NYI     my $min  = shift;
#NYI 
#NYI     return unless defined $min;
#NYI 
#NYI     my @attributes = ( 'val' => $min );
#NYI 
#NYI     $self->xml_empty_tag( 'c:min', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_axis_pos()
#NYI #
#NYI # Write the <c:axPos> element.
#NYI #
#NYI sub _write_axis_pos {
#NYI 
#NYI     my $self    = shift;
#NYI     my $val     = shift;
#NYI     my $reverse = shift;
#NYI 
#NYI     if ( $reverse ) {
#NYI         $val = 'r' if $val eq 'l';
#NYI         $val = 't' if $val eq 'b';
#NYI     }
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:axPos', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_number_format()
#NYI #
#NYI # Write the <c:numberFormat> element. Note: It is assumed that if a user
#NYI # defined number format is supplied (i.e., non-default) then the sourceLinked
#NYI # attribute is 0. The user can override this if required.
#NYI #
#NYI sub _write_number_format {
#NYI 
#NYI     my $self          = shift;
#NYI     my $axis          = shift;
#NYI     my $format_code   = $axis->{_num_format};
#NYI     my $source_linked = 1;
#NYI 
#NYI     # Check if a user defined number format has been set.
#NYI     if ( $format_code ne $axis->{_defaults}->{num_format} ) {
#NYI         $source_linked = 0;
#NYI     }
#NYI 
#NYI     # User override of sourceLinked.
#NYI     if ( $axis->{_num_format_linked} ) {
#NYI         $source_linked = 1;
#NYI     }
#NYI 
#NYI     my @attributes = (
#NYI         'formatCode'   => $format_code,
#NYI         'sourceLinked' => $source_linked,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'c:numFmt', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_cat_number_format()
#NYI #
#NYI # Write the <c:numFmt> element. Special case handler for category axes which
#NYI # don't always have a number format.
#NYI #
#NYI sub _write_cat_number_format {
#NYI 
#NYI     my $self           = shift;
#NYI     my $axis           = shift;
#NYI     my $format_code    = $axis->{_num_format};
#NYI     my $source_linked  = 1;
#NYI     my $default_format = 1;
#NYI 
#NYI     # Check if a user defined number format has been set.
#NYI     if ( $format_code ne $axis->{_defaults}->{num_format} ) {
#NYI         $source_linked  = 0;
#NYI         $default_format = 0;
#NYI     }
#NYI 
#NYI     # User override of linkedSource.
#NYI     if ( $axis->{_num_format_linked} ) {
#NYI         $source_linked = 1;
#NYI     }
#NYI 
#NYI     # Skip if cat doesn't have a num format (unless it is non-default).
#NYI     if ( !$self->{_cat_has_num_fmt} && $default_format ) {
#NYI         return;
#NYI     }
#NYI 
#NYI     my @attributes = (
#NYI         'formatCode'   => $format_code,
#NYI         'sourceLinked' => $source_linked,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'c:numFmt', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_number_format()
#NYI #
#NYI # Write the <c:numberFormat> element for data labels.
#NYI #
#NYI sub _write_data_label_number_format {
#NYI 
#NYI     my $self          = shift;
#NYI     my $format_code   = shift;
#NYI     my $source_linked = 0;
#NYI 
#NYI     my @attributes = (
#NYI         'formatCode'   => $format_code,
#NYI         'sourceLinked' => $source_linked,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'c:numFmt', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_major_tick_mark()
#NYI #
#NYI # Write the <c:majorTickMark> element.
#NYI #
#NYI sub _write_major_tick_mark {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:majorTickMark', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_minor_tick_mark()
#NYI #
#NYI # Write the <c:minorTickMark> element.
#NYI #
#NYI sub _write_minor_tick_mark {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:minorTickMark', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tick_label_pos()
#NYI #
#NYI # Write the <c:tickLblPos> element.
#NYI #
#NYI sub _write_tick_label_pos {
#NYI 
#NYI     my $self = shift;
#NYI     my $val = shift || 'nextTo';
#NYI 
#NYI     if ( $val eq 'next_to' ) {
#NYI         $val = 'nextTo';
#NYI     }
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:tickLblPos', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_cross_axis()
#NYI #
#NYI # Write the <c:crossAx> element.
#NYI #
#NYI sub _write_cross_axis {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:crossAx', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_crosses()
#NYI #
#NYI # Write the <c:crosses> element.
#NYI #
#NYI sub _write_crosses {
#NYI 
#NYI     my $self = shift;
#NYI     my $val = shift || 'autoZero';
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:crosses', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_crosses_at()
#NYI #
#NYI # Write the <c:crossesAt> element.
#NYI #
#NYI sub _write_c_crosses_at {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:crossesAt', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_auto()
#NYI #
#NYI # Write the <c:auto> element.
#NYI #
#NYI sub _write_auto {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:auto', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_label_align()
#NYI #
#NYI # Write the <c:labelAlign> element.
#NYI #
#NYI sub _write_label_align {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 'ctr';
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:lblAlgn', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_label_offset()
#NYI #
#NYI # Write the <c:labelOffset> element.
#NYI #
#NYI sub _write_label_offset {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:lblOffset', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tick_lbl_skip()
#NYI #
#NYI # Write the <c:tickLblSkip> element.
#NYI #
#NYI sub _write_tick_lbl_skip {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:tickLblSkip', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tick_mark_skip()
#NYI #
#NYI # Write the <c:tickMarkSkip> element.
#NYI #
#NYI sub _write_tick_mark_skip {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:tickMarkSkip', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_major_gridlines()
#NYI #
#NYI # Write the <c:majorGridlines> element.
#NYI #
#NYI sub _write_major_gridlines {
#NYI 
#NYI     my $self      = shift;
#NYI     my $gridlines = shift;
#NYI 
#NYI     return unless $gridlines;
#NYI     return unless $gridlines->{_visible};
#NYI 
#NYI     if ( $gridlines->{_line}->{_defined} ) {
#NYI         $self->xml_start_tag( 'c:majorGridlines' );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $gridlines );
#NYI 
#NYI         $self->xml_end_tag( 'c:majorGridlines' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'c:majorGridlines' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_minor_gridlines()
#NYI #
#NYI # Write the <c:minorGridlines> element.
#NYI #
#NYI sub _write_minor_gridlines {
#NYI 
#NYI     my $self      = shift;
#NYI     my $gridlines = shift;
#NYI 
#NYI     return unless $gridlines;
#NYI     return unless $gridlines->{_visible};
#NYI 
#NYI     if ( $gridlines->{_line}->{_defined} ) {
#NYI         $self->xml_start_tag( 'c:minorGridlines' );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $gridlines );
#NYI 
#NYI         $self->xml_end_tag( 'c:minorGridlines' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'c:minorGridlines' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_cross_between()
#NYI #
#NYI # Write the <c:crossBetween> element.
#NYI #
#NYI sub _write_cross_between {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $val = shift || $self->{_cross_between};
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:crossBetween', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_major_unit()
#NYI #
#NYI # Write the <c:majorUnit> element.
#NYI #
#NYI sub _write_c_major_unit {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:majorUnit', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_minor_unit()
#NYI #
#NYI # Write the <c:minorUnit> element.
#NYI #
#NYI sub _write_c_minor_unit {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:minorUnit', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_major_time_unit()
#NYI #
#NYI # Write the <c:majorTimeUnit> element.
#NYI #
#NYI sub _write_c_major_time_unit {
#NYI 
#NYI     my $self = shift;
#NYI     my $val = shift || 'days';
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:majorTimeUnit', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_minor_time_unit()
#NYI #
#NYI # Write the <c:minorTimeUnit> element.
#NYI #
#NYI sub _write_c_minor_time_unit {
#NYI 
#NYI     my $self = shift;
#NYI     my $val = shift || 'days';
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:minorTimeUnit', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_legend()
#NYI #
#NYI # Write the <c:legend> element.
#NYI #
#NYI sub _write_legend {
#NYI 
#NYI     my $self          = shift;
#NYI     my $position      = $self->{_legend_position};
#NYI     my $font          = $self->{_legend_font};
#NYI     my @delete_series = ();
#NYI     my $overlay       = 0;
#NYI 
#NYI     if ( defined $self->{_legend_delete_series}
#NYI         && ref $self->{_legend_delete_series} eq 'ARRAY' )
#NYI     {
#NYI         @delete_series = @{ $self->{_legend_delete_series} };
#NYI     }
#NYI 
#NYI     if ( $position =~ s/^overlay_// ) {
#NYI         $overlay = 1;
#NYI     }
#NYI 
#NYI     my %allowed = (
#NYI         right  => 'r',
#NYI         left   => 'l',
#NYI         top    => 't',
#NYI         bottom => 'b',
#NYI     );
#NYI 
#NYI     return if $position eq 'none';
#NYI     return unless exists $allowed{$position};
#NYI 
#NYI     $position = $allowed{$position};
#NYI 
#NYI     $self->xml_start_tag( 'c:legend' );
#NYI 
#NYI     # Write the c:legendPos element.
#NYI     $self->_write_legend_pos( $position );
#NYI 
#NYI     # Remove series labels from the legend.
#NYI     for my $index ( @delete_series ) {
#NYI 
#NYI         # Write the c:legendEntry element.
#NYI         $self->_write_legend_entry( $index );
#NYI     }
#NYI 
#NYI     # Write the c:layout element.
#NYI     $self->_write_layout( $self->{_legend_layout}, 'legend' );
#NYI 
#NYI     # Write the c:txPr element.
#NYI     if ( $font ) {
#NYI         $self->_write_tx_pr( undef, $font );
#NYI     }
#NYI 
#NYI     # Write the c:overlay element.
#NYI     $self->_write_overlay() if $overlay;
#NYI 
#NYI     $self->xml_end_tag( 'c:legend' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_legend_pos()
#NYI #
#NYI # Write the <c:legendPos> element.
#NYI #
#NYI sub _write_legend_pos {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:legendPos', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_legend_entry()
#NYI #
#NYI # Write the <c:legendEntry> element.
#NYI #
#NYI sub _write_legend_entry {
#NYI 
#NYI     my $self  = shift;
#NYI     my $index = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:legendEntry' );
#NYI 
#NYI     # Write the c:idx element.
#NYI     $self->_write_idx( $index );
#NYI 
#NYI     # Write the c:delete element.
#NYI     $self->_write_delete( 1 );
#NYI 
#NYI     $self->xml_end_tag( 'c:legendEntry' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_overlay()
#NYI #
#NYI # Write the <c:overlay> element.
#NYI #
#NYI sub _write_overlay {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:overlay', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_plot_vis_only()
#NYI #
#NYI # Write the <c:plotVisOnly> element.
#NYI #
#NYI sub _write_plot_vis_only {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     # Ignore this element if we are plotting hidden data.
#NYI     return if $self->{_show_hidden_data};
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:plotVisOnly', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_print_settings()
#NYI #
#NYI # Write the <c:printSettings> element.
#NYI #
#NYI sub _write_print_settings {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:printSettings' );
#NYI 
#NYI     # Write the c:headerFooter element.
#NYI     $self->_write_header_footer();
#NYI 
#NYI     # Write the c:pageMargins element.
#NYI     $self->_write_page_margins();
#NYI 
#NYI     # Write the c:pageSetup element.
#NYI     $self->_write_page_setup();
#NYI 
#NYI     $self->xml_end_tag( 'c:printSettings' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_header_footer()
#NYI #
#NYI # Write the <c:headerFooter> element.
#NYI #
#NYI sub _write_header_footer {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_empty_tag( 'c:headerFooter' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_page_margins()
#NYI #
#NYI # Write the <c:pageMargins> element.
#NYI #
#NYI sub _write_page_margins {
#NYI 
#NYI     my $self   = shift;
#NYI     my $b      = 0.75;
#NYI     my $l      = 0.7;
#NYI     my $r      = 0.7;
#NYI     my $t      = 0.75;
#NYI     my $header = 0.3;
#NYI     my $footer = 0.3;
#NYI 
#NYI     my @attributes = (
#NYI         'b'      => $b,
#NYI         'l'      => $l,
#NYI         'r'      => $r,
#NYI         't'      => $t,
#NYI         'header' => $header,
#NYI         'footer' => $footer,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'c:pageMargins', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_page_setup()
#NYI #
#NYI # Write the <c:pageSetup> element.
#NYI #
#NYI sub _write_page_setup {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_empty_tag( 'c:pageSetup' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_auto_title_deleted()
#NYI #
#NYI # Write the <c:autoTitleDeleted> element.
#NYI #
#NYI sub _write_auto_title_deleted {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:autoTitleDeleted', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_title_rich()
#NYI #
#NYI # Write the <c:title> element for a rich string.
#NYI #
#NYI sub _write_title_rich {
#NYI 
#NYI     my $self    = shift;
#NYI     my $title   = shift;
#NYI     my $horiz   = shift;
#NYI     my $font    = shift;
#NYI     my $layout  = shift;
#NYI     my $overlay = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:title' );
#NYI 
#NYI     # Write the c:tx element.
#NYI     $self->_write_tx_rich( $title, $horiz, $font );
#NYI 
#NYI     # Write the c:layout element.
#NYI     $self->_write_layout( $layout, 'text' );
#NYI 
#NYI     # Write the c:overlay element.
#NYI     $self->_write_overlay() if $overlay;
#NYI 
#NYI     $self->xml_end_tag( 'c:title' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_title_formula()
#NYI #
#NYI # Write the <c:title> element for a rich string.
#NYI #
#NYI sub _write_title_formula {
#NYI 
#NYI     my $self    = shift;
#NYI     my $title   = shift;
#NYI     my $data_id = shift;
#NYI     my $horiz   = shift;
#NYI     my $font    = shift;
#NYI     my $layout  = shift;
#NYI     my $overlay = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:title' );
#NYI 
#NYI     # Write the c:tx element.
#NYI     $self->_write_tx_formula( $title, $data_id );
#NYI 
#NYI     # Write the c:layout element.
#NYI     $self->_write_layout( $layout, 'text' );
#NYI 
#NYI     # Write the c:overlay element.
#NYI     $self->_write_overlay() if $overlay;
#NYI 
#NYI     # Write the c:txPr element.
#NYI     $self->_write_tx_pr( $horiz, $font );
#NYI 
#NYI     $self->xml_end_tag( 'c:title' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tx_rich()
#NYI #
#NYI # Write the <c:tx> element.
#NYI #
#NYI sub _write_tx_rich {
#NYI 
#NYI     my $self  = shift;
#NYI     my $title = shift;
#NYI     my $horiz = shift;
#NYI     my $font  = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:tx' );
#NYI 
#NYI     # Write the c:rich element.
#NYI     $self->_write_rich( $title, $horiz, $font );
#NYI 
#NYI     $self->xml_end_tag( 'c:tx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tx_value()
#NYI #
#NYI # Write the <c:tx> element with a simple value such as for series names.
#NYI #
#NYI sub _write_tx_value {
#NYI 
#NYI     my $self  = shift;
#NYI     my $title = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:tx' );
#NYI 
#NYI     # Write the c:v element.
#NYI     $self->_write_v( $title );
#NYI 
#NYI     $self->xml_end_tag( 'c:tx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tx_formula()
#NYI #
#NYI # Write the <c:tx> element.
#NYI #
#NYI sub _write_tx_formula {
#NYI 
#NYI     my $self    = shift;
#NYI     my $title   = shift;
#NYI     my $data_id = shift;
#NYI     my $data;
#NYI 
#NYI     if ( defined $data_id ) {
#NYI         $data = $self->{_formula_data}->[$data_id];
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'c:tx' );
#NYI 
#NYI     # Write the c:strRef element.
#NYI     $self->_write_str_ref( $title, $data, 'str' );
#NYI 
#NYI     $self->xml_end_tag( 'c:tx' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_rich()
#NYI #
#NYI # Write the <c:rich> element.
#NYI #
#NYI sub _write_rich {
#NYI 
#NYI     my $self     = shift;
#NYI     my $title    = shift;
#NYI     my $horiz    = shift;
#NYI     my $rotation = undef;
#NYI     my $font     = shift;
#NYI 
#NYI     if ( $font && exists $font->{_rotation} ) {
#NYI         $rotation = $font->{_rotation};
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'c:rich' );
#NYI 
#NYI     # Write the a:bodyPr element.
#NYI     $self->_write_a_body_pr( $rotation, $horiz );
#NYI 
#NYI     # Write the a:lstStyle element.
#NYI     $self->_write_a_lst_style();
#NYI 
#NYI     # Write the a:p element.
#NYI     $self->_write_a_p_rich( $title, $font );
#NYI 
#NYI     $self->xml_end_tag( 'c:rich' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_body_pr()
#NYI #
#NYI # Write the <a:bodyPr> element.
#NYI sub _write_a_body_pr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $rot   = shift;
#NYI     my $horiz = shift;
#NYI 
#NYI     my @attributes = ();
#NYI 
#NYI     if ( !defined $rot && $horiz ) {
#NYI         $rot = -5400000;
#NYI     }
#NYI 
#NYI     push @attributes, ( 'rot' => $rot ) if defined $rot;
#NYI     push @attributes, ( 'vert' => 'horz' ) if $horiz;
#NYI 
#NYI     $self->xml_empty_tag( 'a:bodyPr', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_lst_style()
#NYI #
#NYI # Write the <a:lstStyle> element.
#NYI #
#NYI sub _write_a_lst_style {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_empty_tag( 'a:lstStyle' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_p_rich()
#NYI #
#NYI # Write the <a:p> element for rich string titles.
#NYI #
#NYI sub _write_a_p_rich {
#NYI 
#NYI     my $self  = shift;
#NYI     my $title = shift;
#NYI     my $font  = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:p' );
#NYI 
#NYI     # Write the a:pPr element.
#NYI     $self->_write_a_p_pr_rich( $font );
#NYI 
#NYI     # Write the a:r element.
#NYI     $self->_write_a_r( $title, $font );
#NYI 
#NYI     $self->xml_end_tag( 'a:p' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_p_formula()
#NYI #
#NYI # Write the <a:p> element for formula titles.
#NYI #
#NYI sub _write_a_p_formula {
#NYI 
#NYI     my $self = shift;
#NYI     my $font = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:p' );
#NYI 
#NYI     # Write the a:pPr element.
#NYI     $self->_write_a_p_pr_formula( $font );
#NYI 
#NYI     # Write the a:endParaRPr element.
#NYI     $self->_write_a_end_para_rpr();
#NYI 
#NYI     $self->xml_end_tag( 'a:p' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_p_pr_rich()
#NYI #
#NYI # Write the <a:pPr> element for rich string titles.
#NYI #
#NYI sub _write_a_p_pr_rich {
#NYI 
#NYI     my $self = shift;
#NYI     my $font = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:pPr' );
#NYI 
#NYI     # Write the a:defRPr element.
#NYI     $self->_write_a_def_rpr( $font );
#NYI 
#NYI     $self->xml_end_tag( 'a:pPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_p_pr_formula()
#NYI #
#NYI # Write the <a:pPr> element for formula titles.
#NYI #
#NYI sub _write_a_p_pr_formula {
#NYI 
#NYI     my $self = shift;
#NYI     my $font = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:pPr' );
#NYI 
#NYI     # Write the a:defRPr element.
#NYI     $self->_write_a_def_rpr( $font );
#NYI 
#NYI     $self->xml_end_tag( 'a:pPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_def_rpr()
#NYI #
#NYI # Write the <a:defRPr> element.
#NYI #
#NYI sub _write_a_def_rpr {
#NYI 
#NYI     my $self      = shift;
#NYI     my $font      = shift;
#NYI     my $has_color = 0;
#NYI 
#NYI     my @style_attributes = $self->_get_font_style_attributes( $font );
#NYI     my @latin_attributes = $self->_get_font_latin_attributes( $font );
#NYI 
#NYI     $has_color = 1 if $font && $font->{_color};
#NYI 
#NYI     if ( @latin_attributes || $has_color ) {
#NYI         $self->xml_start_tag( 'a:defRPr', @style_attributes );
#NYI 
#NYI 
#NYI         if ( $has_color ) {
#NYI             $self->_write_a_solid_fill( { color => $font->{_color} } );
#NYI         }
#NYI 
#NYI         if ( @latin_attributes ) {
#NYI             $self->_write_a_latin( @latin_attributes );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'a:defRPr' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'a:defRPr', @style_attributes );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_end_para_rpr()
#NYI #
#NYI # Write the <a:endParaRPr> element.
#NYI #
#NYI sub _write_a_end_para_rpr {
#NYI 
#NYI     my $self = shift;
#NYI     my $lang = 'en-US';
#NYI 
#NYI     my @attributes = ( 'lang' => $lang );
#NYI 
#NYI     $self->xml_empty_tag( 'a:endParaRPr', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_r()
#NYI #
#NYI # Write the <a:r> element.
#NYI #
#NYI sub _write_a_r {
#NYI 
#NYI     my $self  = shift;
#NYI     my $title = shift;
#NYI     my $font  = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:r' );
#NYI 
#NYI     # Write the a:rPr element.
#NYI     $self->_write_a_r_pr( $font );
#NYI 
#NYI     # Write the a:t element.
#NYI     $self->_write_a_t( $title );
#NYI 
#NYI     $self->xml_end_tag( 'a:r' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_r_pr()
#NYI #
#NYI # Write the <a:rPr> element.
#NYI #
#NYI sub _write_a_r_pr {
#NYI 
#NYI     my $self      = shift;
#NYI     my $font      = shift;
#NYI     my $has_color = 0;
#NYI     my $lang      = 'en-US';
#NYI 
#NYI     my @style_attributes = $self->_get_font_style_attributes( $font );
#NYI     my @latin_attributes = $self->_get_font_latin_attributes( $font );
#NYI 
#NYI     $has_color = 1 if $font && $font->{_color};
#NYI 
#NYI     # Add the lang type to the attributes.
#NYI     @style_attributes = ( 'lang' => $lang, @style_attributes );
#NYI 
#NYI 
#NYI     if ( @latin_attributes || $has_color ) {
#NYI         $self->xml_start_tag( 'a:rPr', @style_attributes );
#NYI 
#NYI 
#NYI         if ( $has_color ) {
#NYI             $self->_write_a_solid_fill( { color => $font->{_color} } );
#NYI         }
#NYI 
#NYI         if ( @latin_attributes ) {
#NYI             $self->_write_a_latin( @latin_attributes );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'a:rPr' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'a:rPr', @style_attributes );
#NYI     }
#NYI 
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_t()
#NYI #
#NYI # Write the <a:t> element.
#NYI #
#NYI sub _write_a_t {
#NYI 
#NYI     my $self  = shift;
#NYI     my $title = shift;
#NYI 
#NYI     $self->xml_data_element( 'a:t', $title );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_tx_pr()
#NYI #
#NYI # Write the <c:txPr> element.
#NYI #
#NYI sub _write_tx_pr {
#NYI 
#NYI     my $self     = shift;
#NYI     my $horiz    = shift;
#NYI     my $font     = shift;
#NYI     my $rotation = undef;
#NYI 
#NYI     if ( $font && exists $font->{_rotation} ) {
#NYI         $rotation = $font->{_rotation};
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'c:txPr' );
#NYI 
#NYI     # Write the a:bodyPr element.
#NYI     $self->_write_a_body_pr( $rotation, $horiz );
#NYI 
#NYI     # Write the a:lstStyle element.
#NYI     $self->_write_a_lst_style();
#NYI 
#NYI     # Write the a:p element.
#NYI     $self->_write_a_p_formula( $font );
#NYI 
#NYI     $self->xml_end_tag( 'c:txPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_marker()
#NYI #
#NYI # Write the <c:marker> element.
#NYI #
#NYI sub _write_marker {
#NYI 
#NYI     my $self = shift;
#NYI     my $marker = shift || $self->{_default_marker};
#NYI 
#NYI     return unless $marker;
#NYI     return if $marker->{automatic};
#NYI 
#NYI     $self->xml_start_tag( 'c:marker' );
#NYI 
#NYI     # Write the c:symbol element.
#NYI     $self->_write_symbol( $marker->{type} );
#NYI 
#NYI     # Write the c:size element.
#NYI     my $size = $marker->{size};
#NYI     $self->_write_marker_size( $size ) if $size;
#NYI 
#NYI     # Write the c:spPr element.
#NYI     $self->_write_sp_pr( $marker );
#NYI 
#NYI     $self->xml_end_tag( 'c:marker' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_marker_size()
#NYI #
#NYI # Write the <c:size> element.
#NYI #
#NYI sub _write_marker_size {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:size', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_symbol()
#NYI #
#NYI # Write the <c:symbol> element.
#NYI #
#NYI sub _write_symbol {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:symbol', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_sp_pr()
#NYI #
#NYI # Write the <c:spPr> element.
#NYI #
#NYI sub _write_sp_pr {
#NYI 
#NYI     my $self   = shift;
#NYI     my $series = shift;
#NYI 
#NYI     if (    !$series->{_line}->{_defined}
#NYI         and !$series->{_fill}->{_defined}
#NYI         and !$series->{_pattern}
#NYI         and !$series->{_gradient} )
#NYI     {
#NYI         return;
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_start_tag( 'c:spPr' );
#NYI 
#NYI     # Write the fill elements for solid charts such as pie/doughnut and bar.
#NYI     if ( $series->{_fill}->{_defined} ) {
#NYI 
#NYI         if ( $series->{_fill}->{none} ) {
#NYI 
#NYI             # Write the a:noFill element.
#NYI             $self->_write_a_no_fill();
#NYI         }
#NYI         else {
#NYI             # Write the a:solidFill element.
#NYI             $self->_write_a_solid_fill( $series->{_fill} );
#NYI         }
#NYI     }
#NYI 
#NYI     if ( $series->{_pattern} ) {
#NYI 
#NYI         # Write the a:pattFill element.
#NYI         $self->_write_a_patt_fill( $series->{_pattern} );
#NYI     }
#NYI 
#NYI     if ( $series->{_gradient} ) {
#NYI 
#NYI         # Write the a:gradFill element.
#NYI         $self->_write_a_grad_fill( $series->{_gradient} );
#NYI     }
#NYI 
#NYI 
#NYI     # Write the a:ln element.
#NYI     if ( $series->{_line}->{_defined} ) {
#NYI         $self->_write_a_ln( $series->{_line} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:spPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_ln()
#NYI #
#NYI # Write the <a:ln> element.
#NYI #
#NYI sub _write_a_ln {
#NYI 
#NYI     my $self       = shift;
#NYI     my $line       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     # Add the line width as an attribute.
#NYI     if ( my $width = $line->{width} ) {
#NYI 
#NYI         # Round width to nearest 0.25, like Excel.
#NYI         $width = int( ( $width + 0.125 ) * 4 ) / 4;
#NYI 
#NYI         # Convert to internal units.
#NYI         $width = int( 0.5 + ( 12700 * $width ) );
#NYI 
#NYI         @attributes = ( 'w' => $width );
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'a:ln', @attributes );
#NYI 
#NYI     # Write the line fill.
#NYI     if ( $line->{none} ) {
#NYI 
#NYI         # Write the a:noFill element.
#NYI         $self->_write_a_no_fill();
#NYI     }
#NYI     elsif ( $line->{color} ) {
#NYI 
#NYI         # Write the a:solidFill element.
#NYI         $self->_write_a_solid_fill( $line );
#NYI     }
#NYI 
#NYI     # Write the line/dash type.
#NYI     if ( my $type = $line->{dash_type} ) {
#NYI 
#NYI         # Write the a:prstDash element.
#NYI         $self->_write_a_prst_dash( $type );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'a:ln' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_no_fill()
#NYI #
#NYI # Write the <a:noFill> element.
#NYI #
#NYI sub _write_a_no_fill {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_empty_tag( 'a:noFill' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_solid_fill()
#NYI #
#NYI # Write the <a:solidFill> element.
#NYI #
#NYI sub _write_a_solid_fill {
#NYI 
#NYI     my $self = shift;
#NYI     my $fill = shift;
#NYI 
#NYI     $self->xml_start_tag( 'a:solidFill' );
#NYI 
#NYI     if ( $fill->{color} ) {
#NYI 
#NYI         my $color = $self->_get_color( $fill->{color} );
#NYI 
#NYI         # Write the a:srgbClr element.
#NYI         $self->_write_a_srgb_clr( $color, $fill->{transparency} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'a:solidFill' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_srgb_clr()
#NYI #
#NYI # Write the <a:srgbClr> element.
#NYI #
#NYI sub _write_a_srgb_clr {
#NYI 
#NYI     my $self         = shift;
#NYI     my $color        = shift;
#NYI     my $transparency = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $color );
#NYI 
#NYI     if ( $transparency ) {
#NYI         $self->xml_start_tag( 'a:srgbClr', @attributes );
#NYI 
#NYI         # Write the a:alpha element.
#NYI         $self->_write_a_alpha( $transparency );
#NYI 
#NYI         $self->xml_end_tag( 'a:srgbClr' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'a:srgbClr', @attributes );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_alpha()
#NYI #
#NYI # Write the <a:alpha> element.
#NYI #
#NYI sub _write_a_alpha {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     $val = ( 100 - int( $val ) ) * 1000;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'a:alpha', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_prst_dash()
#NYI #
#NYI # Write the <a:prstDash> element.
#NYI #
#NYI sub _write_a_prst_dash {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'a:prstDash', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_trendline()
#NYI #
#NYI # Write the <c:trendline> element.
#NYI #
#NYI sub _write_trendline {
#NYI 
#NYI     my $self      = shift;
#NYI     my $trendline = shift;
#NYI 
#NYI     return unless $trendline;
#NYI 
#NYI     $self->xml_start_tag( 'c:trendline' );
#NYI 
#NYI     # Write the c:name element.
#NYI     $self->_write_name( $trendline->{name} );
#NYI 
#NYI     # Write the c:spPr element.
#NYI     $self->_write_sp_pr( $trendline );
#NYI 
#NYI     # Write the c:trendlineType element.
#NYI     $self->_write_trendline_type( $trendline->{type} );
#NYI 
#NYI     # Write the c:order element for polynomial trendlines.
#NYI     if ( $trendline->{type} eq 'poly' ) {
#NYI         $self->_write_trendline_order( $trendline->{order} );
#NYI     }
#NYI 
#NYI     # Write the c:period element for moving average trendlines.
#NYI     if ( $trendline->{type} eq 'movingAvg' ) {
#NYI         $self->_write_period( $trendline->{period} );
#NYI     }
#NYI 
#NYI     # Write the c:forward element.
#NYI     $self->_write_forward( $trendline->{forward} );
#NYI 
#NYI     # Write the c:backward element.
#NYI     $self->_write_backward( $trendline->{backward} );
#NYI 
#NYI     if ( defined $trendline->{intercept} ) {
#NYI         # Write the c:intercept element.
#NYI         $self->_write_intercept( $trendline->{intercept} );
#NYI     }
#NYI 
#NYI     if ($trendline->{display_r_squared}) {
#NYI         # Write the c:dispRSqr element.
#NYI         $self->_write_disp_rsqr();
#NYI     }
#NYI 
#NYI     if ($trendline->{display_equation}) {
#NYI         # Write the c:dispEq element.
#NYI         $self->_write_disp_eq();
#NYI 
#NYI         # Write the c:trendlineLbl element.
#NYI         $self->_write_trendline_lbl();
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:trendline' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_trendline_type()
#NYI #
#NYI # Write the <c:trendlineType> element.
#NYI #
#NYI sub _write_trendline_type {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:trendlineType', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_name()
#NYI #
#NYI # Write the <c:name> element.
#NYI #
#NYI sub _write_name {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     return unless defined $data;
#NYI 
#NYI     $self->xml_data_element( 'c:name', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_trendline_order()
#NYI #
#NYI # Write the <c:order> element.
#NYI #
#NYI sub _write_trendline_order {
#NYI 
#NYI     my $self = shift;
#NYI     my $val = defined $_[0] ? $_[0] : 2;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:order', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_period()
#NYI #
#NYI # Write the <c:period> element.
#NYI #
#NYI sub _write_period {
#NYI 
#NYI     my $self = shift;
#NYI     my $val = defined $_[0] ? $_[0] : 2;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:period', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_forward()
#NYI #
#NYI # Write the <c:forward> element.
#NYI #
#NYI sub _write_forward {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:forward', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_backward()
#NYI #
#NYI # Write the <c:backward> element.
#NYI #
#NYI sub _write_backward {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return unless $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:backward', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_intercept()
#NYI #
#NYI # Write the <c:intercept> element.
#NYI #
#NYI sub _write_intercept {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:intercept', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_disp_eq()
#NYI #
#NYI # Write the <c:dispEq> element.
#NYI #
#NYI sub _write_disp_eq {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:dispEq', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_disp_rsqr()
#NYI #
#NYI # Write the <c:dispRSqr> element.
#NYI #
#NYI sub _write_disp_rsqr {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:dispRSqr', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_trendline_lbl()
#NYI #
#NYI # Write the <c:trendlineLbl> element.
#NYI #
#NYI sub _write_trendline_lbl {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     $self->xml_start_tag( 'c:trendlineLbl' );
#NYI 
#NYI     # Write the c:layout element.
#NYI     $self->_write_layout();
#NYI 
#NYI     # Write the c:numFmt element.
#NYI     $self->_write_trendline_num_fmt();
#NYI 
#NYI     $self->xml_end_tag( 'c:trendlineLbl' );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_trendline_num_fmt()
#NYI #
#NYI # Write the <c:numFmt> element.
#NYI #
#NYI sub _write_trendline_num_fmt {
#NYI 
#NYI     my $self          = shift;
#NYI     my $format_code   = 'General';
#NYI     my $source_linked = 0;
#NYI 
#NYI     my @attributes = (
#NYI         'formatCode'   => $format_code,
#NYI         'sourceLinked' => $source_linked,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'c:numFmt', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_hi_low_lines()
#NYI #
#NYI # Write the <c:hiLowLines> element.
#NYI #
#NYI sub _write_hi_low_lines {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $hi_low_lines = $self->{_hi_low_lines};
#NYI 
#NYI     return unless $hi_low_lines;
#NYI 
#NYI     if ( $hi_low_lines->{_line}->{_defined} ) {
#NYI 
#NYI         $self->xml_start_tag( 'c:hiLowLines' );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $hi_low_lines );
#NYI 
#NYI         $self->xml_end_tag( 'c:hiLowLines' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'c:hiLowLines' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI #############################################################################
#NYI #
#NYI # _write_drop_lines()
#NYI #
#NYI # Write the <c:dropLines> element.
#NYI #
#NYI sub _write_drop_lines {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my $drop_lines = $self->{_drop_lines};
#NYI 
#NYI     return unless $drop_lines;
#NYI 
#NYI     if ( $drop_lines->{_line}->{_defined} ) {
#NYI 
#NYI         $self->xml_start_tag( 'c:dropLines' );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $drop_lines );
#NYI 
#NYI         $self->xml_end_tag( 'c:dropLines' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'c:dropLines' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_overlap()
#NYI #
#NYI # Write the <c:overlap> element.
#NYI #
#NYI sub _write_overlap {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return if !defined $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:overlap', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_num_cache()
#NYI #
#NYI # Write the <c:numCache> element.
#NYI #
#NYI sub _write_num_cache {
#NYI 
#NYI     my $self  = shift;
#NYI     my $data  = shift;
#NYI     my $count = @$data;
#NYI 
#NYI     $self->xml_start_tag( 'c:numCache' );
#NYI 
#NYI     # Write the c:formatCode element.
#NYI     $self->_write_format_code( 'General' );
#NYI 
#NYI     # Write the c:ptCount element.
#NYI     $self->_write_pt_count( $count );
#NYI 
#NYI     for my $i ( 0 .. $count - 1 ) {
#NYI         my $token = $data->[$i];
#NYI 
#NYI         # Write non-numeric data as 0.
#NYI         if ( defined $token
#NYI             && $token !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/ )
#NYI         {
#NYI             $token = 0;
#NYI         }
#NYI 
#NYI         # Write the c:pt element.
#NYI         $self->_write_pt( $i, $token );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:numCache' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_str_cache()
#NYI #
#NYI # Write the <c:strCache> element.
#NYI #
#NYI sub _write_str_cache {
#NYI 
#NYI     my $self  = shift;
#NYI     my $data  = shift;
#NYI     my $count = @$data;
#NYI 
#NYI     $self->xml_start_tag( 'c:strCache' );
#NYI 
#NYI     # Write the c:ptCount element.
#NYI     $self->_write_pt_count( $count );
#NYI 
#NYI     for my $i ( 0 .. $count - 1 ) {
#NYI 
#NYI         # Write the c:pt element.
#NYI         $self->_write_pt( $i, $data->[$i] );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:strCache' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_format_code()
#NYI #
#NYI # Write the <c:formatCode> element.
#NYI #
#NYI sub _write_format_code {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'c:formatCode', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_pt_count()
#NYI #
#NYI # Write the <c:ptCount> element.
#NYI #
#NYI sub _write_pt_count {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:ptCount', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_pt()
#NYI #
#NYI # Write the <c:pt> element.
#NYI #
#NYI sub _write_pt {
#NYI 
#NYI     my $self  = shift;
#NYI     my $idx   = shift;
#NYI     my $value = shift;
#NYI 
#NYI     return if !defined $value;
#NYI 
#NYI     my @attributes = ( 'idx' => $idx );
#NYI 
#NYI     $self->xml_start_tag( 'c:pt', @attributes );
#NYI 
#NYI     # Write the c:v element.
#NYI     $self->_write_v( $value );
#NYI 
#NYI     $self->xml_end_tag( 'c:pt' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_v()
#NYI #
#NYI # Write the <c:v> element.
#NYI #
#NYI sub _write_v {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'c:v', $data );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_protection()
#NYI #
#NYI # Write the <c:protection> element.
#NYI #
#NYI sub _write_protection {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return unless $self->{_protection};
#NYI 
#NYI     $self->xml_empty_tag( 'c:protection' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_d_pt()
#NYI #
#NYI # Write the <c:dPt> elements.
#NYI #
#NYI sub _write_d_pt {
#NYI 
#NYI     my $self   = shift;
#NYI     my $points = shift;
#NYI     my $index  = -1;
#NYI 
#NYI     return unless $points;
#NYI 
#NYI     for my $point ( @$points ) {
#NYI 
#NYI         $index++;
#NYI         next unless $point;
#NYI 
#NYI         $self->_write_d_pt_point( $index, $point );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_d_pt_point()
#NYI #
#NYI # Write an individual <c:dPt> element.
#NYI #
#NYI sub _write_d_pt_point {
#NYI 
#NYI     my $self   = shift;
#NYI     my $index = shift;
#NYI     my $point = shift;
#NYI 
#NYI         $self->xml_start_tag( 'c:dPt' );
#NYI 
#NYI         # Write the c:idx element.
#NYI         $self->_write_idx( $index );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $point );
#NYI 
#NYI         $self->xml_end_tag( 'c:dPt' );
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_d_lbls()
#NYI #
#NYI # Write the <c:dLbls> element.
#NYI #
#NYI sub _write_d_lbls {
#NYI 
#NYI     my $self   = shift;
#NYI     my $labels = shift;
#NYI 
#NYI     return unless $labels;
#NYI 
#NYI     $self->xml_start_tag( 'c:dLbls' );
#NYI 
#NYI     # Write the c:numFmt element.
#NYI     if ( $labels->{num_format} ) {
#NYI         $self->_write_data_label_number_format( $labels->{num_format} );
#NYI     }
#NYI 
#NYI     # Write the data label font elements.
#NYI     if ($labels->{font} ) {
#NYI         $self->_write_axis_font( $labels->{font} );
#NYI     }
#NYI 
#NYI     # Write the c:dLblPos element.
#NYI     $self->_write_d_lbl_pos( $labels->{position} ) if $labels->{position};
#NYI 
#NYI     # Write the c:showLegendKey element.
#NYI     $self->_write_show_legend_key() if $labels->{legend_key};
#NYI 
#NYI     # Write the c:showVal element.
#NYI     $self->_write_show_val() if $labels->{value};
#NYI 
#NYI     # Write the c:showCatName element.
#NYI     $self->_write_show_cat_name() if $labels->{category};
#NYI 
#NYI     # Write the c:showSerName element.
#NYI     $self->_write_show_ser_name() if $labels->{series_name};
#NYI 
#NYI     # Write the c:showPercent element.
#NYI     $self->_write_show_percent() if $labels->{percentage};
#NYI 
#NYI     # Write the c:separator element.
#NYI     $self->_write_separator($labels->{separator}) if $labels->{separator};
#NYI 
#NYI     # Write the c:showLeaderLines element.
#NYI     $self->_write_show_leader_lines() if $labels->{leader_lines};
#NYI 
#NYI     $self->xml_end_tag( 'c:dLbls' );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_legend_key()
#NYI #
#NYI # Write the <c:showLegendKey> element.
#NYI #
#NYI sub _write_show_legend_key {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showLegendKey', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_val()
#NYI #
#NYI # Write the <c:showVal> element.
#NYI #
#NYI sub _write_show_val {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showVal', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_cat_name()
#NYI #
#NYI # Write the <c:showCatName> element.
#NYI #
#NYI sub _write_show_cat_name {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showCatName', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_ser_name()
#NYI #
#NYI # Write the <c:showSerName> element.
#NYI #
#NYI sub _write_show_ser_name {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showSerName', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_percent()
#NYI #
#NYI # Write the <c:showPercent> element.
#NYI #
#NYI sub _write_show_percent {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showPercent', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_separator()
#NYI #
#NYI # Write the <c:separator> element.
#NYI #
#NYI sub _write_separator {
#NYI 
#NYI     my $self = shift;
#NYI     my $data = shift;
#NYI 
#NYI     $self->xml_data_element( 'c:separator', $data );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_leader_lines()
#NYI #
#NYI # Write the <c:showLeaderLines> element.
#NYI #
#NYI sub _write_show_leader_lines {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = 1;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showLeaderLines', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_d_lbl_pos()
#NYI #
#NYI # Write the <c:dLblPos> element.
#NYI #
#NYI sub _write_d_lbl_pos {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:dLblPos', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_delete()
#NYI #
#NYI # Write the <c:delete> element.
#NYI #
#NYI sub _write_delete {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:delete', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_invert_if_negative()
#NYI #
#NYI # Write the <c:invertIfNegative> element.
#NYI #
#NYI sub _write_c_invert_if_negative {
#NYI 
#NYI     my $self   = shift;
#NYI     my $invert = shift;
#NYI     my $val    = 1;
#NYI 
#NYI     return unless $invert;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:invertIfNegative', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_axis_font()
#NYI #
#NYI # Write the axis font elements.
#NYI #
#NYI sub _write_axis_font {
#NYI 
#NYI     my $self = shift;
#NYI     my $font = shift;
#NYI 
#NYI     return unless $font;
#NYI 
#NYI     $self->xml_start_tag( 'c:txPr' );
#NYI     $self->_write_a_body_pr($font->{_rotation});
#NYI     $self->_write_a_lst_style();
#NYI     $self->xml_start_tag( 'a:p' );
#NYI 
#NYI     $self->_write_a_p_pr_rich( $font );
#NYI 
#NYI     $self->_write_a_end_para_rpr();
#NYI     $self->xml_end_tag( 'a:p' );
#NYI     $self->xml_end_tag( 'c:txPr' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_latin()
#NYI #
#NYI # Write the <a:latin> element.
#NYI #
#NYI sub _write_a_latin {
#NYI 
#NYI     my $self       = shift;
#NYI     my @attributes = @_;
#NYI 
#NYI     $self->xml_empty_tag( 'a:latin', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_d_table()
#NYI #
#NYI # Write the <c:dTable> element.
#NYI #
#NYI sub _write_d_table {
#NYI 
#NYI     my $self  = shift;
#NYI     my $table = $self->{_table};
#NYI 
#NYI     return if !$table;
#NYI 
#NYI     $self->xml_start_tag( 'c:dTable' );
#NYI 
#NYI     if ( $table->{_horizontal} ) {
#NYI 
#NYI         # Write the c:showHorzBorder element.
#NYI         $self->_write_show_horz_border();
#NYI     }
#NYI 
#NYI     if ( $table->{_vertical} ) {
#NYI 
#NYI         # Write the c:showVertBorder element.
#NYI         $self->_write_show_vert_border();
#NYI     }
#NYI 
#NYI     if ( $table->{_outline} ) {
#NYI 
#NYI         # Write the c:showOutline element.
#NYI         $self->_write_show_outline();
#NYI     }
#NYI 
#NYI     if ( $table->{_show_keys} ) {
#NYI 
#NYI         # Write the c:showKeys element.
#NYI         $self->_write_show_keys();
#NYI     }
#NYI 
#NYI     if ( $table->{_font} ) {
#NYI         # Write the table font.
#NYI         $self->_write_tx_pr( undef, $table->{_font} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:dTable' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_horz_border()
#NYI #
#NYI # Write the <c:showHorzBorder> element.
#NYI #
#NYI sub _write_show_horz_border {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showHorzBorder', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_vert_border()
#NYI #
#NYI # Write the <c:showVertBorder> element.
#NYI #
#NYI sub _write_show_vert_border {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showVertBorder', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_outline()
#NYI #
#NYI # Write the <c:showOutline> element.
#NYI #
#NYI sub _write_show_outline {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showOutline', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_show_keys()
#NYI #
#NYI # Write the <c:showKeys> element.
#NYI #
#NYI sub _write_show_keys {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:showKeys', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_error_bars()
#NYI #
#NYI # Write the X and Y error bars.
#NYI #
#NYI sub _write_error_bars {
#NYI 
#NYI     my $self       = shift;
#NYI     my $error_bars = shift;
#NYI 
#NYI     return unless $error_bars;
#NYI 
#NYI     if ( $error_bars->{_x_error_bars} ) {
#NYI         $self->_write_err_bars( 'x', $error_bars->{_x_error_bars} );
#NYI     }
#NYI 
#NYI     if ( $error_bars->{_y_error_bars} ) {
#NYI         $self->_write_err_bars( 'y', $error_bars->{_y_error_bars} );
#NYI     }
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_err_bars()
#NYI #
#NYI # Write the <c:errBars> element.
#NYI #
#NYI sub _write_err_bars {
#NYI 
#NYI     my $self       = shift;
#NYI     my $direction  = shift;
#NYI     my $error_bars = shift;
#NYI 
#NYI     return unless $error_bars;
#NYI 
#NYI     $self->xml_start_tag( 'c:errBars' );
#NYI 
#NYI     # Write the c:errDir element.
#NYI     $self->_write_err_dir( $direction );
#NYI 
#NYI     # Write the c:errBarType element.
#NYI     $self->_write_err_bar_type( $error_bars->{_direction} );
#NYI 
#NYI     # Write the c:errValType element.
#NYI     $self->_write_err_val_type( $error_bars->{_type} );
#NYI 
#NYI     if ( !$error_bars->{_endcap} ) {
#NYI 
#NYI         # Write the c:noEndCap element.
#NYI         $self->_write_no_end_cap();
#NYI     }
#NYI 
#NYI     if ( $error_bars->{_type} eq 'stdErr' ) {
#NYI 
#NYI         # Don't need to write a c:errValType tag.
#NYI     }
#NYI     elsif ( $error_bars->{_type} eq 'cust' ) {
#NYI 
#NYI         # Write the custom error tags.
#NYI         $self->_write_custom_error( $error_bars );
#NYI     }
#NYI     else {
#NYI         # Write the c:val element.
#NYI         $self->_write_error_val( $error_bars->{_value} );
#NYI     }
#NYI 
#NYI     # Write the c:spPr element.
#NYI     $self->_write_sp_pr( $error_bars );
#NYI 
#NYI     $self->xml_end_tag( 'c:errBars' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_err_dir()
#NYI #
#NYI # Write the <c:errDir> element.
#NYI #
#NYI sub _write_err_dir {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:errDir', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_err_bar_type()
#NYI #
#NYI # Write the <c:errBarType> element.
#NYI #
#NYI sub _write_err_bar_type {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:errBarType', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_err_val_type()
#NYI #
#NYI # Write the <c:errValType> element.
#NYI #
#NYI sub _write_err_val_type {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:errValType', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_no_end_cap()
#NYI #
#NYI # Write the <c:noEndCap> element.
#NYI #
#NYI sub _write_no_end_cap {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:noEndCap', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_error_val()
#NYI #
#NYI # Write the <c:val> element for error bars.
#NYI #
#NYI sub _write_error_val {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:val', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_custom_error()
#NYI #
#NYI # Write the custom error bars tags.
#NYI #
#NYI sub _write_custom_error {
#NYI 
#NYI     my $self       = shift;
#NYI     my $error_bars = shift;
#NYI 
#NYI     if ( $error_bars->{_plus_values} ) {
#NYI 
#NYI         # Write the c:plus element.
#NYI         $self->xml_start_tag( 'c:plus' );
#NYI 
#NYI         if ( ref $error_bars->{_plus_values} eq 'ARRAY' ) {
#NYI             $self->_write_num_lit( $error_bars->{_plus_values} );
#NYI         }
#NYI         else {
#NYI             $self->_write_num_ref( $error_bars->{_plus_values},
#NYI                 $error_bars->{_plus_data}, 'num' );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'c:plus' );
#NYI     }
#NYI 
#NYI     if ( $error_bars->{_minus_values} ) {
#NYI 
#NYI         # Write the c:minus element.
#NYI         $self->xml_start_tag( 'c:minus' );
#NYI 
#NYI         if ( ref $error_bars->{_minus_values} eq 'ARRAY' ) {
#NYI             $self->_write_num_lit( $error_bars->{_minus_values} );
#NYI         }
#NYI         else {
#NYI             $self->_write_num_ref( $error_bars->{_minus_values},
#NYI                 $error_bars->{_minus_data}, 'num' );
#NYI         }
#NYI 
#NYI         $self->xml_end_tag( 'c:minus' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_num_lit()
#NYI #
#NYI # Write the <c:numLit> element for literal number list elements.
#NYI #
#NYI sub _write_num_lit {
#NYI 
#NYI     my $self = shift;
#NYI     my $data  = shift;
#NYI     my $count = @$data;
#NYI 
#NYI 
#NYI     # Write the c:numLit element.
#NYI     $self->xml_start_tag( 'c:numLit' );
#NYI 
#NYI     # Write the c:formatCode element.
#NYI     $self->_write_format_code( 'General' );
#NYI 
#NYI     # Write the c:ptCount element.
#NYI     $self->_write_pt_count( $count );
#NYI 
#NYI     for my $i ( 0 .. $count - 1 ) {
#NYI         my $token = $data->[$i];
#NYI 
#NYI         # Write non-numeric data as 0.
#NYI         if ( defined $token
#NYI             && $token !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/ )
#NYI         {
#NYI             $token = 0;
#NYI         }
#NYI 
#NYI         # Write the c:pt element.
#NYI         $self->_write_pt( $i, $token );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:numLit' );
#NYI 
#NYI 
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_up_down_bars()
#NYI #
#NYI # Write the <c:upDownBars> element.
#NYI #
#NYI sub _write_up_down_bars {
#NYI 
#NYI     my $self         = shift;
#NYI     my $up_down_bars = $self->{_up_down_bars};
#NYI 
#NYI     return unless $up_down_bars;
#NYI 
#NYI     $self->xml_start_tag( 'c:upDownBars' );
#NYI 
#NYI     # Write the c:gapWidth element.
#NYI     $self->_write_gap_width( 150 );
#NYI 
#NYI     # Write the c:upBars element.
#NYI     $self->_write_up_bars( $up_down_bars->{_up} );
#NYI 
#NYI     # Write the c:downBars element.
#NYI     $self->_write_down_bars( $up_down_bars->{_down} );
#NYI 
#NYI     $self->xml_end_tag( 'c:upDownBars' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_gap_width()
#NYI #
#NYI # Write the <c:gapWidth> element.
#NYI #
#NYI sub _write_gap_width {
#NYI 
#NYI     my $self = shift;
#NYI     my $val  = shift;
#NYI 
#NYI     return if !defined $val;
#NYI 
#NYI     my @attributes = ( 'val' => $val );
#NYI 
#NYI     $self->xml_empty_tag( 'c:gapWidth', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_up_bars()
#NYI #
#NYI # Write the <c:upBars> element.
#NYI #
#NYI sub _write_up_bars {
#NYI 
#NYI     my $self   = shift;
#NYI     my $format = shift;
#NYI 
#NYI     if ( $format->{_line}->{_defined} || $format->{_fill}->{_defined} ) {
#NYI 
#NYI         $self->xml_start_tag( 'c:upBars' );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $format );
#NYI 
#NYI         $self->xml_end_tag( 'c:upBars' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'c:upBars' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_down_bars()
#NYI #
#NYI # Write the <c:downBars> element.
#NYI #
#NYI sub _write_down_bars {
#NYI 
#NYI     my $self   = shift;
#NYI     my $format = shift;
#NYI 
#NYI     if ( $format->{_line}->{_defined} || $format->{_fill}->{_defined} ) {
#NYI 
#NYI         $self->xml_start_tag( 'c:downBars' );
#NYI 
#NYI         # Write the c:spPr element.
#NYI         $self->_write_sp_pr( $format );
#NYI 
#NYI         $self->xml_end_tag( 'c:downBars' );
#NYI     }
#NYI     else {
#NYI         $self->xml_empty_tag( 'c:downBars' );
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_c_smooth()
#NYI #
#NYI # Write the <c:smooth> element.
#NYI #
#NYI sub _write_c_smooth {
#NYI 
#NYI     my $self    = shift;
#NYI     my $smooth  = shift;
#NYI 
#NYI     return unless $smooth;
#NYI 
#NYI     my @attributes = ( 'val' => 1 );
#NYI 
#NYI     $self->xml_empty_tag( 'c:smooth', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_disp_units()
#NYI #
#NYI # Write the <c:dispUnits> element.
#NYI #
#NYI sub _write_disp_units {
#NYI 
#NYI     my $self    = shift;
#NYI     my $units   = shift;
#NYI     my $display = shift;
#NYI 
#NYI     return if not $units;
#NYI 
#NYI     my @attributes = ( 'val' => $units );
#NYI 
#NYI     $self->xml_start_tag( 'c:dispUnits' );
#NYI 
#NYI     $self->xml_empty_tag( 'c:builtInUnit', @attributes );
#NYI 
#NYI     if ( $display ) {
#NYI         $self->xml_start_tag( 'c:dispUnitsLbl' );
#NYI         $self->xml_empty_tag( 'c:layout' );
#NYI         $self->xml_end_tag( 'c:dispUnitsLbl' );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'c:dispUnits' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_grad_fill()
#NYI #
#NYI # Write the <a:gradFill> element.
#NYI #
#NYI sub _write_a_grad_fill {
#NYI 
#NYI     my $self     = shift;
#NYI     my $gradient = shift;
#NYI 
#NYI 
#NYI     my @attributes = (
#NYI         'flip'         => 'none',
#NYI         'rotWithShape' => 1,
#NYI     );
#NYI 
#NYI 
#NYI     if ( $gradient->{_type} eq 'linear' ) {
#NYI         @attributes = ();
#NYI     }
#NYI 
#NYI     $self->xml_start_tag( 'a:gradFill', @attributes );
#NYI 
#NYI     # Write the a:gsLst element.
#NYI     $self->_write_a_gs_lst( $gradient );
#NYI 
#NYI     if ( $gradient->{_type} eq 'linear' ) {
#NYI         # Write the a:lin element.
#NYI         $self->_write_a_lin( $gradient->{_angle} );
#NYI     }
#NYI     else {
#NYI         # Write the a:path element.
#NYI         $self->_write_a_path( $gradient->{_type} );
#NYI 
#NYI         # Write the a:tileRect element.
#NYI         $self->_write_a_tile_rect( $gradient->{_type} );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'a:gradFill' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_gs_lst()
#NYI #
#NYI # Write the <a:gsLst> element.
#NYI #
#NYI sub _write_a_gs_lst {
#NYI 
#NYI     my $self      = shift;
#NYI     my $gradient  = shift;
#NYI     my $positions = $gradient->{_positions};
#NYI     my $colors    = $gradient->{_colors};
#NYI 
#NYI     $self->xml_start_tag( 'a:gsLst' );
#NYI 
#NYI     for my $i ( 0 .. @$colors -1 ) {
#NYI 
#NYI         my $pos = int($positions->[$i] * 1000);
#NYI 
#NYI         my @attributes = ( 'pos' => $pos );
#NYI         $self->xml_start_tag( 'a:gs', @attributes );
#NYI 
#NYI         my $color = $self->_get_color( $colors->[$i] );
#NYI 
#NYI         # Write the a:srgbClr element.
#NYI         # TODO: Wait for a feature request to support transparency.
#NYI         $self->_write_a_srgb_clr( $color );
#NYI 
#NYI         $self->xml_end_tag( 'a:gs' );
#NYI     }
#NYI 
#NYI     $self->xml_end_tag( 'a:gsLst' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_lin()
#NYI #
#NYI # Write the <a:lin> element.
#NYI #
#NYI sub _write_a_lin {
#NYI 
#NYI     my $self   = shift;
#NYI     my $angle  = shift;
#NYI     my $scaled = 0;
#NYI 
#NYI     $angle = int( 60000 * $angle );
#NYI 
#NYI     my @attributes = (
#NYI         'ang'    => $angle,
#NYI         'scaled' => $scaled,
#NYI     );
#NYI 
#NYI     $self->xml_empty_tag( 'a:lin', @attributes );
#NYI }
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_path()
#NYI #
#NYI # Write the <a:path> element.
#NYI #
#NYI sub _write_a_path {
#NYI 
#NYI     my $self = shift;
#NYI     my $type = shift;
#NYI 
#NYI 
#NYI     my @attributes = ( 'path' => $type );
#NYI 
#NYI     $self->xml_start_tag( 'a:path', @attributes );
#NYI 
#NYI     # Write the a:fillToRect element.
#NYI     $self->_write_a_fill_to_rect( $type );
#NYI 
#NYI     $self->xml_end_tag( 'a:path' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_fill_to_rect()
#NYI #
#NYI # Write the <a:fillToRect> element.
#NYI #
#NYI sub _write_a_fill_to_rect {
#NYI 
#NYI     my $self       = shift;
#NYI     my $type       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     if ( $type eq 'shape' ) {
#NYI         @attributes = (
#NYI             'l' => 50000,
#NYI             't' => 50000,
#NYI             'r' => 50000,
#NYI             'b' => 50000,
#NYI         );
#NYI 
#NYI     }
#NYI     else {
#NYI         @attributes = (
#NYI             'l' => 100000,
#NYI             't' => 100000,
#NYI         );
#NYI     }
#NYI 
#NYI 
#NYI     $self->xml_empty_tag( 'a:fillToRect', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_tile_rect()
#NYI #
#NYI # Write the <a:tileRect> element.
#NYI #
#NYI sub _write_a_tile_rect {
#NYI 
#NYI     my $self       = shift;
#NYI     my $type       = shift;
#NYI     my @attributes = ();
#NYI 
#NYI     if ( $type eq 'shape' ) {
#NYI         @attributes = ();
#NYI     }
#NYI     else {
#NYI         @attributes = (
#NYI             'r' => -100000,
#NYI             'b' => -100000,
#NYI         );
#NYI     }
#NYI 
#NYI     $self->xml_empty_tag( 'a:tileRect', @attributes );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_patt_fill()
#NYI #
#NYI # Write the <a:pattFill> element.
#NYI #
#NYI sub _write_a_patt_fill {
#NYI 
#NYI     my $self     = shift;
#NYI     my $pattern  = shift;
#NYI 
#NYI     my @attributes = ( 'prst' => $pattern->{pattern} );
#NYI 
#NYI     $self->xml_start_tag( 'a:pattFill', @attributes );
#NYI 
#NYI     # Write the a:fgClr element.
#NYI     $self->_write_a_fg_clr( $pattern->{fg_color} );
#NYI 
#NYI     # Write the a:bgClr element.
#NYI     $self->_write_a_bg_clr( $pattern->{bg_color} );
#NYI 
#NYI     $self->xml_end_tag( 'a:pattFill' );
#NYI }
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_fg_clr()
#NYI #
#NYI # Write the <a:fgClr> element.
#NYI #
#NYI sub _write_a_fg_clr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $color = shift;
#NYI 
#NYI     $color = $self->_get_color( $color );
#NYI 
#NYI     $self->xml_start_tag( 'a:fgClr' );
#NYI 
#NYI     # Write the a:srgbClr element.
#NYI     $self->_write_a_srgb_clr( $color );
#NYI 
#NYI     $self->xml_end_tag( 'a:fgClr' );
#NYI }
#NYI 
#NYI 
#NYI 
#NYI ##############################################################################
#NYI #
#NYI # _write_a_bg_clr()
#NYI #
#NYI # Write the <a:bgClr> element.
#NYI #
#NYI sub _write_a_bg_clr {
#NYI 
#NYI     my $self  = shift;
#NYI     my $color = shift;
#NYI 
#NYI     $color = $self->_get_color( $color );
#NYI 
#NYI     $self->xml_start_tag( 'a:bgClr' );
#NYI 
#NYI     # Write the a:srgbClr element.
#NYI     $self->_write_a_srgb_clr( $color );
#NYI 
#NYI     $self->xml_end_tag( 'a:bgClr' );
#NYI }
#NYI 
#NYI 
#NYI 1;
#NYI 
#NYI __END__
#NYI 
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Chart - A class for writing Excel Charts.
#NYI 
#NYI =head1 SYNOPSIS
#NYI 
#NYI To create a simple Excel file with a chart using Excel::Writer::XLSX:
#NYI 
#NYI     #!/usr/bin/perl
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'chart.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # Add the worksheet data the chart refers to.
#NYI     my $data = [
#NYI         [ 'Category', 2, 3, 4, 5, 6, 7 ],
#NYI         [ 'Value',    1, 4, 5, 2, 1, 5 ],
#NYI 
#NYI     ];
#NYI 
#NYI     $worksheet->write( 'A1', $data );
#NYI 
#NYI     # Add a worksheet chart.
#NYI     my $chart = $workbook->add_chart( type => 'column' );
#NYI 
#NYI     # Configure the chart.
#NYI     $chart->add_series(
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$B$2:$B$7',
#NYI     );
#NYI 
#NYI     __END__
#NYI 
#NYI 
#NYI =head1 DESCRIPTION
#NYI 
#NYI The C<Chart> module is an abstract base class for modules that implement charts in L<Excel::Writer::XLSX>. The information below is applicable to all of the available subclasses.
#NYI 
#NYI The C<Chart> module isn't used directly. A chart object is created via the Workbook C<add_chart()> method where the chart type is specified:
#NYI 
#NYI     my $chart = $workbook->add_chart( type => 'column' );
#NYI 
#NYI Currently the supported chart types are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<area>
#NYI 
#NYI Creates an Area (filled line) style chart. See L<Excel::Writer::XLSX::Chart::Area>.
#NYI 
#NYI =item * C<bar>
#NYI 
#NYI Creates a Bar style (transposed histogram) chart. See L<Excel::Writer::XLSX::Chart::Bar>.
#NYI 
#NYI =item * C<column>
#NYI 
#NYI Creates a column style (histogram) chart. See L<Excel::Writer::XLSX::Chart::Column>.
#NYI 
#NYI =item * C<line>
#NYI 
#NYI Creates a Line style chart. See L<Excel::Writer::XLSX::Chart::Line>.
#NYI 
#NYI =item * C<pie>
#NYI 
#NYI Creates a Pie style chart. See L<Excel::Writer::XLSX::Chart::Pie>.
#NYI 
#NYI =item * C<doughnut>
#NYI 
#NYI Creates a Doughnut style chart. See L<Excel::Writer::XLSX::Chart::Doughnut>.
#NYI 
#NYI =item * C<scatter>
#NYI 
#NYI Creates a Scatter style chart. See L<Excel::Writer::XLSX::Chart::Scatter>.
#NYI 
#NYI =item * C<stock>
#NYI 
#NYI Creates a Stock style chart. See L<Excel::Writer::XLSX::Chart::Stock>.
#NYI 
#NYI =item * C<radar>
#NYI 
#NYI Creates a Radar style chart. See L<Excel::Writer::XLSX::Chart::Radar>.
#NYI 
#NYI =back
#NYI 
#NYI Chart subtypes are also supported in some cases:
#NYI 
#NYI     $workbook->add_chart( type => 'bar', subtype => 'stacked' );
#NYI 
#NYI The currently available subtypes are:
#NYI 
#NYI     area
#NYI         stacked
#NYI         percent_stacked
#NYI 
#NYI     bar
#NYI         stacked
#NYI         percent_stacked
#NYI 
#NYI     column
#NYI         stacked
#NYI         percent_stacked
#NYI 
#NYI     scatter
#NYI         straight_with_markers
#NYI         straight
#NYI         smooth_with_markers
#NYI         smooth
#NYI 
#NYI     radar
#NYI         with_markers
#NYI         filled
#NYI 
#NYI More charts and sub-types will be supported in time. See the L</TODO> section.
#NYI 
#NYI 
#NYI =head1 CHART METHODS
#NYI 
#NYI Methods that are common to all chart types are documented below. See the documentation for each of the above chart modules for chart specific information.
#NYI 
#NYI =head2 add_series()
#NYI 
#NYI In an Excel chart a "series" is a collection of information such as values, X axis labels and the formatting that define which data is plotted.
#NYI 
#NYI With an Excel::Writer::XLSX chart object the C<add_series()> method is used to set the properties for a series:
#NYI 
#NYI     $chart->add_series(
#NYI         categories => '=Sheet1!$A$2:$A$10', # Optional.
#NYI         values     => '=Sheet1!$B$2:$B$10', # Required.
#NYI         line       => { color => 'blue' },
#NYI     );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<values>
#NYI 
#NYI This is the most important property of a series and must be set for every chart object. It links the chart with the worksheet data that it displays. A formula or array ref can be used for the data range, see below.
#NYI 
#NYI =item * C<categories>
#NYI 
#NYI This sets the chart category labels. The category is more or less the same as the X axis. In most chart types the C<categories> property is optional and the chart will just assume a sequential series from C<1 .. n>.
#NYI 
#NYI =item * C<name>
#NYI 
#NYI Set the name for the series. The name is displayed in the chart legend and in the formula bar. The name property is optional and if it isn't supplied it will default to C<Series 1 .. n>.
#NYI 
#NYI =item * C<line>
#NYI 
#NYI Set the properties of the series line type such as colour and width. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<border>
#NYI 
#NYI Set the border properties of the series such as colour and style. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<fill>
#NYI 
#NYI Set the fill properties of the series such as colour. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<pattern>
#NYI 
#NYI Set the pattern properties of the series. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<gradien>
#NYI 
#NYI Set the gradient properties of the series. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<marker>
#NYI 
#NYI Set the properties of the series marker such as style and colour. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<trendline>
#NYI 
#NYI Set the properties of the series trendline such as linear, polynomial and moving average types. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<smooth>
#NYI 
#NYI The C<smooth> option is used to set the smooth property of a line series. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<y_error_bars>
#NYI 
#NYI Set vertical error bounds for a chart series. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<x_error_bars>
#NYI 
#NYI Set horizontal error bounds for a chart series. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<data_labels>
#NYI 
#NYI Set data labels for the series. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<points>
#NYI 
#NYI Set properties for individual points in a series. See the L</SERIES OPTIONS> section below.
#NYI 
#NYI =item * C<invert_if_negative>
#NYI 
#NYI Invert the fill colour for negative values. Usually only applicable to column and bar charts.
#NYI 
#NYI =item * C<overlap>
#NYI 
#NYI Set the overlap between series in a Bar/Column chart. The range is +/- 100. Default is 0.
#NYI 
#NYI     overlap => 20,
#NYI 
#NYI Note, it is only necessary to apply this property to one series of the chart.
#NYI 
#NYI =item * C<gap>
#NYI 
#NYI Set the gap between series in a Bar/Column chart. The range is 0 to 500. Default is 150.
#NYI 
#NYI     gap => 200,
#NYI 
#NYI Note, it is only necessary to apply this property to one series of the chart.
#NYI 
#NYI =back
#NYI 
#NYI The C<categories> and C<values> can take either a range formula such as C<=Sheet1!$A$2:$A$7> or, more usefully when generating the range programmatically, an array ref with zero indexed row/column values:
#NYI 
#NYI      [ $sheetname, $row_start, $row_end, $col_start, $col_end ]
#NYI 
#NYI The following are equivalent:
#NYI 
#NYI     $chart->add_series( categories => '=Sheet1!$A$2:$A$7'      ); # Same as ...
#NYI     $chart->add_series( categories => [ 'Sheet1', 1, 6, 0, 0 ] ); # Zero-indexed.
#NYI 
#NYI You can add more than one series to a chart. In fact, some chart types such as C<stock> require it. The series numbering and order in the Excel chart will be the same as the order in which they are added in Excel::Writer::XLSX.
#NYI 
#NYI     # Add the first series.
#NYI     $chart->add_series(
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$B$2:$B$7',
#NYI         name       => 'Test data series 1',
#NYI     );
#NYI 
#NYI     # Add another series. Same categories. Different range values.
#NYI     $chart->add_series(
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$C$2:$C$7',
#NYI         name       => 'Test data series 2',
#NYI     );
#NYI 
#NYI It is also possible to specify non-contiguous ranges:
#NYI 
#NYI     $chart->add_series(
#NYI         categories      => '=(Sheet1!$A$1:$A$9,Sheet1!$A$14:$A$25)',
#NYI         values          => '=(Sheet1!$B$1:$B$9,Sheet1!$B$14:$B$25)',
#NYI     );
#NYI 
#NYI 
#NYI =head2 set_x_axis()
#NYI 
#NYI The C<set_x_axis()> method is used to set properties of the X axis.
#NYI 
#NYI     $chart->set_x_axis( name => 'Quarterly results' );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI     name
#NYI     name_font
#NYI     name_layout
#NYI     num_font
#NYI     num_format
#NYI     line
#NYI     fill
#NYI     pattern
#NYI     gradient
#NYI     min
#NYI     max
#NYI     minor_unit
#NYI     major_unit
#NYI     interval_unit
#NYI     interval_tick
#NYI     crossing
#NYI     reverse
#NYI     position_axis
#NYI     log_base
#NYI     label_position
#NYI     major_gridlines
#NYI     minor_gridlines
#NYI     visible
#NYI     date_axis
#NYI     text_axis
#NYI     minor_unit_type
#NYI     major_unit_type
#NYI     minor_tick_mark
#NYI     major_tick_mark
#NYI     display_units
#NYI     display_units_visible
#NYI 
#NYI These are explained below. Some properties are only applicable to value or category axes, as indicated. See L<Value and Category Axes> for an explanation of Excel's distinction between the axis types.
#NYI 
#NYI =over
#NYI 
#NYI =item * C<name>
#NYI 
#NYI 
#NYI Set the name (title or caption) for the axis. The name is displayed below the X axis. The C<name> property is optional. The default is to have no axis name. (Applicable to category and value axes).
#NYI 
#NYI     $chart->set_x_axis( name => 'Quarterly results' );
#NYI 
#NYI The name can also be a formula such as C<=Sheet1!$A$1>.
#NYI 
#NYI =item * C<name_font>
#NYI 
#NYI Set the font properties for the axis title. (Applicable to category and value axes).
#NYI 
#NYI     $chart->set_x_axis( name_font => { name => 'Arial', size => 10 } );
#NYI 
#NYI =item * C<name_layout>
#NYI 
#NYI Set the C<(x, y)> position of the axis caption in chart relative units. (Applicable to category and value axes).
#NYI 
#NYI     $chart->set_x_axis(
#NYI         name        => 'X axis',
#NYI         name_layout => {
#NYI             x => 0.34,
#NYI             y => 0.85,
#NYI         }
#NYI     );
#NYI 
#NYI See the L</CHART LAYOUT> section below.
#NYI 
#NYI =item * C<num_font>
#NYI 
#NYI Set the font properties for the axis numbers. (Applicable to category and value axes).
#NYI 
#NYI     $chart->set_x_axis( num_font => { bold => 1, italic => 1 } );
#NYI 
#NYI See the L</CHART FONTS> section below.
#NYI 
#NYI =item * C<num_format>
#NYI 
#NYI Set the number format for the axis. (Applicable to category and value axes).
#NYI 
#NYI     $chart->set_x_axis( num_format => '#,##0.00' );
#NYI     $chart->set_y_axis( num_format => '0.00%'    );
#NYI 
#NYI The number format is similar to the Worksheet Cell Format C<num_format> apart from the fact that a format index cannot be used. The explicit format string must be used as shown above. See L<Excel::Writer::XLSX/set_num_format()> for more information.
#NYI 
#NYI =item * C<line>
#NYI 
#NYI Set the properties of the axis line type such as colour and width. See the L</CHART FORMATTING> section below.
#NYI 
#NYI     $chart->set_x_axis( line => { none => 1 });
#NYI 
#NYI 
#NYI =item * C<fill>
#NYI 
#NYI Set the fill properties of the axis such as colour. See the L</CHART FORMATTING> section below. Note, in Excel the axis fill is applied to the area of the numbers of the axis and not to the area of the axis bounding box. That background is set from the chartarea fill.
#NYI 
#NYI =item * C<pattern>
#NYI 
#NYI Set the pattern properties of the axis such as colour. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<gradient>
#NYI 
#NYI Set the gradient properties of the axis such as colour. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<min>
#NYI 
#NYI Set the minimum value for the axis range. (Applicable to value axes only.)
#NYI 
#NYI     $chart->set_x_axis( min => 20 );
#NYI 
#NYI =item * C<max>
#NYI 
#NYI Set the maximum value for the axis range. (Applicable to value axes only.)
#NYI 
#NYI     $chart->set_x_axis( max => 80 );
#NYI 
#NYI =item * C<minor_unit>
#NYI 
#NYI Set the increment of the minor units in the axis range. (Applicable to value axes only.)
#NYI 
#NYI     $chart->set_x_axis( minor_unit => 0.4 );
#NYI 
#NYI =item * C<major_unit>
#NYI 
#NYI Set the increment of the major units in the axis range. (Applicable to value axes only.)
#NYI 
#NYI     $chart->set_x_axis( major_unit => 2 );
#NYI 
#NYI =item * C<interval_unit>
#NYI 
#NYI Set the interval unit for a category axis. (Applicable to category axes only.)
#NYI 
#NYI     $chart->set_x_axis( interval_unit => 2 );
#NYI 
#NYI =item * C<interval_tick>
#NYI 
#NYI Set the tick interval for a category axis. (Applicable to category axes only.)
#NYI 
#NYI     $chart->set_x_axis( interval_tick => 4 );
#NYI 
#NYI =item * C<crossing>
#NYI 
#NYI Set the position where the y axis will cross the x axis. (Applicable to category and value axes.)
#NYI 
#NYI The C<crossing> value can either be the string C<'max'> to set the crossing at the maximum axis value or a numeric value.
#NYI 
#NYI     $chart->set_x_axis( crossing => 3 );
#NYI     # or
#NYI     $chart->set_x_axis( crossing => 'max' );
#NYI 
#NYI B<For category axes the numeric value must be an integer> to represent the category number that the axis crosses at. For value axes it can have any value associated with the axis.
#NYI 
#NYI If crossing is omitted (the default) the crossing will be set automatically by Excel based on the chart data.
#NYI 
#NYI =item * C<position_axis>
#NYI 
#NYI Position the axis on or between the axis tick marks. (Applicable to category axes only.)
#NYI 
#NYI There are two allowable values C<on_tick> and C<between>:
#NYI 
#NYI     $chart->set_x_axis( position_axis => 'on_tick' );
#NYI     $chart->set_x_axis( position_axis => 'between' );
#NYI 
#NYI =item * C<reverse>
#NYI 
#NYI Reverse the order of the axis categories or values. (Applicable to category and value axes.)
#NYI 
#NYI     $chart->set_x_axis( reverse => 1 );
#NYI 
#NYI =item * C<log_base>
#NYI 
#NYI Set the log base of the axis range. (Applicable to value axes only.)
#NYI 
#NYI     $chart->set_x_axis( log_base => 10 );
#NYI 
#NYI =item * C<label_position>
#NYI 
#NYI Set the "Axis labels" position for the axis. The following positions are available:
#NYI 
#NYI     next_to (the default)
#NYI     high
#NYI     low
#NYI     none
#NYI 
#NYI =item * C<major_gridlines>
#NYI 
#NYI Configure the major gridlines for the axis. The available properties are:
#NYI 
#NYI     visible
#NYI     line
#NYI 
#NYI For example:
#NYI 
#NYI     $chart->set_x_axis(
#NYI         major_gridlines => {
#NYI             visible => 1,
#NYI             line    => { color => 'red', width => 1.25, dash_type => 'dash' }
#NYI         }
#NYI     );
#NYI 
#NYI The C<visible> property is usually on for the X-axis but it depends on the type of chart.
#NYI 
#NYI The C<line> property sets the gridline properties such as colour and width. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<minor_gridlines>
#NYI 
#NYI This takes the same options as C<major_gridlines> above.
#NYI 
#NYI The minor gridline C<visible> property is off by default for all chart types.
#NYI 
#NYI =item * C<visible>
#NYI 
#NYI Configure the visibility of the axis.
#NYI 
#NYI     $chart->set_x_axis( visible => 0 );
#NYI 
#NYI 
#NYI =item * C<date_axis>
#NYI 
#NYI This option is used to treat a category axis with date or time data as a Date Axis. (Applicable to category axes only.)
#NYI 
#NYI     $chart->set_x_axis( date_axis => 1 );
#NYI 
#NYI This option also allows you to set C<max> and C<min> values for a category axis which isn't allowed by Excel for non-date category axes.
#NYI 
#NYI See L<Date Category Axes> for more details.
#NYI 
#NYI =item * C<text_axis>
#NYI 
#NYI This option is used to treat a category axis explicitly as a Text Axis. (Applicable to category axes only.)
#NYI 
#NYI     $chart->set_x_axis( text_axis => 1 );
#NYI 
#NYI 
#NYI =item * C<minor_unit_type>
#NYI 
#NYI For C<date_axis> axes, see above, this option is used to set the type of the minor units. (Applicable to date category axes only.)
#NYI 
#NYI     $chart->set_x_axis(
#NYI         date_axis         => 1,
#NYI         minor_unit        => 4,
#NYI         minor_unit_type   => 'months',
#NYI     );
#NYI 
#NYI The allowable values for this option are C<days>, C<months> and C<years>.
#NYI 
#NYI =item * C<major_unit_type>
#NYI 
#NYI Same as C<minor_unit_type>, see above, but for major axes unit types.
#NYI 
#NYI More than one property can be set in a call to C<set_x_axis()>:
#NYI 
#NYI     $chart->set_x_axis(
#NYI         name => 'Quarterly results',
#NYI         min  => 10,
#NYI         max  => 80,
#NYI     );
#NYI 
#NYI =item * C<major_tick_mark>
#NYI 
#NYI Set the axis major tick mark type to one of the following values:
#NYI 
#NYI     none
#NYI     inside
#NYI     outside
#NYI     cross   (inside and outside)
#NYI 
#NYI For example:
#NYI 
#NYI     $chart->set_x_axis( major_tick_mark => 'none',
#NYI                         minor_tick_mark => 'inside' );
#NYI 
#NYI =item * C<minor_tick_mark>
#NYI 
#NYI Set the axis minor tick mark type. Same as C<major_tick_mark>, see above.
#NYI 
#NYI =item * C<display_units>
#NYI 
#NYI Set the display units for the axis. This can be useful if the axis numbers are very large but you don't want to represent them in scientific notation. (Applicable to value axes only.) The available display units are:
#NYI 
#NYI     hundreds
#NYI     thousands
#NYI     ten_thousands
#NYI     hundred_thousands
#NYI     millions
#NYI     ten_millions
#NYI     hundred_millions
#NYI     billions
#NYI     trillions
#NYI 
#NYI Example:
#NYI 
#NYI     $chart->set_x_axis( display_units => 'thousands' )
#NYI     $chart->set_y_axis( display_units => 'millions' )
#NYI 
#NYI 
#NYI * C<display_units_visible>
#NYI 
#NYI Control the visibility of the display units turned on by the previous option. This option is on by default. (Applicable to value axes only.)::
#NYI 
#NYI     $chart->set_x_axis( display_units         => 'thousands',
#NYI                         display_units_visible => 0 )
#NYI 
#NYI =back
#NYI 
#NYI =head2 set_y_axis()
#NYI 
#NYI The C<set_y_axis()> method is used to set properties of the Y axis. The properties that can be set are the same as for C<set_x_axis>, see above.
#NYI 
#NYI 
#NYI =head2 set_x2_axis()
#NYI 
#NYI The C<set_x2_axis()> method is used to set properties of the secondary X axis.
#NYI The properties that can be set are the same as for C<set_x_axis>, see above.
#NYI The default properties for this axis are:
#NYI 
#NYI     label_position => 'none',
#NYI     crossing       => 'max',
#NYI     visible        => 0,
#NYI 
#NYI 
#NYI =head2 set_y2_axis()
#NYI 
#NYI The C<set_y2_axis()> method is used to set properties of the secondary Y axis.
#NYI The properties that can be set are the same as for C<set_x_axis>, see above.
#NYI The default properties for this axis are:
#NYI 
#NYI     major_gridlines => { visible => 0 }
#NYI 
#NYI 
#NYI =head2 combine()
#NYI 
#NYI The chart C<combine()> method is used to combine two charts of different
#NYI types, for example a column and line chart:
#NYI 
#NYI     my $column_chart = $workbook->add_chart( type => 'column', embedded => 1 );
#NYI 
#NYI     # Configure the data series for the primary chart.
#NYI     $column_chart->add_series(...);
#NYI 
#NYI     # Create a new column chart. This will use this as the secondary chart.
#NYI     my $line_chart = $workbook->add_chart( type => 'line', embedded => 1 );
#NYI 
#NYI     # Configure the data series for the secondary chart.
#NYI     $line_chart->add_series(...);
#NYI 
#NYI     # Combine the charts.
#NYI     $column_chart->combine( $line_chart );
#NYI 
#NYI See L<Combined Charts> for more details.
#NYI 
#NYI 
#NYI =head2 set_size()
#NYI 
#NYI The C<set_size()> method is used to set the dimensions of the chart. The size properties that can be set are:
#NYI 
#NYI      width
#NYI      height
#NYI      x_scale
#NYI      y_scale
#NYI      x_offset
#NYI      y_offset
#NYI 
#NYI The C<width> and C<height> are in pixels. The default chart width is 480 pixels and the default height is 288 pixels. The size of the chart can be modified by setting the C<width> and C<height> or by setting the C<x_scale> and C<y_scale>:
#NYI 
#NYI     $chart->set_size( width => 720, height => 576 );
#NYI 
#NYI     # Same as:
#NYI 
#NYI     $chart->set_size( x_scale => 1.5, y_scale => 2 );
#NYI 
#NYI The C<x_offset> and C<y_offset> position the top left corner of the chart in the cell that it is inserted into.
#NYI 
#NYI 
#NYI Note: the C<x_scale>, C<y_scale>, C<x_offset> and C<y_offset> parameters can also be set via the C<insert_chart()> method:
#NYI 
#NYI     $worksheet->insert_chart( 'E2', $chart, 2, 4, 1.5, 2 );
#NYI 
#NYI 
#NYI =head2 set_title()
#NYI 
#NYI The C<set_title()> method is used to set properties of the chart title.
#NYI 
#NYI     $chart->set_title( name => 'Year End Results' );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<name>
#NYI 
#NYI Set the name (title) for the chart. The name is displayed above the chart. The name can also be a formula such as C<=Sheet1!$A$1>. The name property is optional. The default is to have no chart title.
#NYI 
#NYI =item * C<name_font>
#NYI 
#NYI Set the font properties for the chart title. See the L</CHART FONTS> section below.
#NYI 
#NYI =item * C<overlay>
#NYI 
#NYI Allow the title to be overlaid on the chart. Generally used with the layout property below.
#NYI 
#NYI =item * C<layout>
#NYI 
#NYI Set the C<(x, y)> position of the title in chart relative units:
#NYI 
#NYI     $chart->set_title(
#NYI         name    => 'Title',
#NYI         overlay => 1,
#NYI         layout  => {
#NYI             x => 0.42,
#NYI             y => 0.14,
#NYI         }
#NYI     );
#NYI 
#NYI See the L</CHART LAYOUT> section below.
#NYI 
#NYI =item * C<none>
#NYI 
#NYI By default Excel adds an automatic chart title to charts with a single series and a user defined series name. The C<none> option turns this default title off. It also turns off all other C<set_title()> options.
#NYI 
#NYI     $chart->set_title( none => 1 );
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI =head2 set_legend()
#NYI 
#NYI The C<set_legend()> method is used to set properties of the chart legend.
#NYI 
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<none>
#NYI 
#NYI The C<none> option turns off the chart legend. In Excel chart legends are on by default:
#NYI 
#NYI     $chart->set_legend( none => 1 );
#NYI 
#NYI Note, for backward compatibility, it is also possible to turn off the legend via the C<position> property:
#NYI 
#NYI     $chart->set_legend( position => 'none' );
#NYI 
#NYI =item * C<position>
#NYI 
#NYI Set the position of the chart legend.
#NYI 
#NYI     $chart->set_legend( position => 'bottom' );
#NYI 
#NYI The default legend position is C<right>. The available positions are:
#NYI 
#NYI     top
#NYI     bottom
#NYI     left
#NYI     right
#NYI     overlay_left
#NYI     overlay_right
#NYI     none
#NYI 
#NYI =item * C<layout>
#NYI 
#NYI Set the C<(x, y)> position of the legend in chart relative units:
#NYI 
#NYI     $chart->set_legend(
#NYI         layout => {
#NYI             x      => 0.80,
#NYI             y      => 0.37,
#NYI             width  => 0.12,
#NYI             height => 0.25,
#NYI         }
#NYI     );
#NYI 
#NYI See the L</CHART LAYOUT> section below.
#NYI 
#NYI 
#NYI =item * C<delete_series>
#NYI 
#NYI This allows you to remove 1 or more series from the legend (the series will still display on the chart). This property takes an array ref as an argument and the series are zero indexed:
#NYI 
#NYI     # Delete/hide series index 0 and 2 from the legend.
#NYI     $chart->set_legend( delete_series => [0, 2] );
#NYI 
#NYI =item * C<font>
#NYI 
#NYI Set the font properties of the chart legend:
#NYI 
#NYI     $chart->set_legend( font => { bold => 1, italic => 1 } );
#NYI 
#NYI See the L</CHART FONTS> section below.
#NYI 
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI =head2 set_chartarea()
#NYI 
#NYI The C<set_chartarea()> method is used to set the properties of the chart area.
#NYI 
#NYI     $chart->set_chartarea(
#NYI         border => { none  => 1 },
#NYI         fill   => { color => 'red' }
#NYI     );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<border>
#NYI 
#NYI Set the border properties of the chartarea such as colour and style. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<fill>
#NYI 
#NYI Set the fill properties of the chartarea such as colour. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<pattern>
#NYI 
#NYI Set the pattern fill properties of the chartarea. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<gradient>
#NYI 
#NYI Set the gradient fill properties of the chartarea. See the L</CHART FORMATTING> section below.
#NYI 
#NYI 
#NYI =back
#NYI 
#NYI =head2 set_plotarea()
#NYI 
#NYI The C<set_plotarea()> method is used to set properties of the plot area of a chart.
#NYI 
#NYI     $chart->set_plotarea(
#NYI         border => { color => 'yellow', width => 1, dash_type => 'dash' },
#NYI         fill   => { color => '#92D050' }
#NYI     );
#NYI 
#NYI The properties that can be set are:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<border>
#NYI 
#NYI Set the border properties of the plotarea such as colour and style. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<fill>
#NYI 
#NYI Set the fill properties of the plotarea such as colour. See the L</CHART FORMATTING> section below.
#NYI 
#NYI 
#NYI =item * C<pattern>
#NYI 
#NYI Set the pattern fill properties of the plotarea. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<gradient>
#NYI 
#NYI Set the gradient fill properties of the plotarea. See the L</CHART FORMATTING> section below.
#NYI 
#NYI =item * C<layout>
#NYI 
#NYI Set the C<(x, y)> position of the plotarea in chart relative units:
#NYI 
#NYI     $chart->set_plotarea(
#NYI         layout => {
#NYI             x      => 0.35,
#NYI             y      => 0.26,
#NYI             width  => 0.62,
#NYI             height => 0.50,
#NYI         }
#NYI     );
#NYI 
#NYI See the L</CHART LAYOUT> section below.
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI =head2 set_style()
#NYI 
#NYI The C<set_style()> method is used to set the style of the chart to one of the 42 built-in styles available on the 'Design' tab in Excel:
#NYI 
#NYI     $chart->set_style( 4 );
#NYI 
#NYI The default style is 2.
#NYI 
#NYI 
#NYI =head2 set_table()
#NYI 
#NYI The C<set_table()> method adds a data table below the horizontal axis with the data used to plot the chart.
#NYI 
#NYI     $chart->set_table();
#NYI 
#NYI The available options, with default values are:
#NYI 
#NYI     vertical   => 1    # Display vertical lines in the table.
#NYI     horizontal => 1    # Display horizontal lines in the table.
#NYI     outline    => 1    # Display an outline in the table.
#NYI     show_keys  => 0    # Show the legend keys with the table data.
#NYI     font       => {}   # Standard chart font properties.
#NYI 
#NYI The data table can only be shown with Bar, Column, Line, Area and stock charts. For font properties see the L</CHART FONTS> section below.
#NYI 
#NYI 
#NYI =head2 set_up_down_bars
#NYI 
#NYI The C<set_up_down_bars()> method adds Up-Down bars to Line charts to indicate the difference between the first and last data series.
#NYI 
#NYI     $chart->set_up_down_bars();
#NYI 
#NYI It is possible to format the up and down bars to add C<fill>, C<pattern>, C<gradient> and C<border> properties if required. See the L</CHART FORMATTING> section below.
#NYI 
#NYI     $chart->set_up_down_bars(
#NYI         up   => { fill => { color => 'green' } },
#NYI         down => { fill => { color => 'red' } },
#NYI     );
#NYI 
#NYI Up-down bars can only be applied to Line charts and to Stock charts (by default).
#NYI 
#NYI 
#NYI =head2 set_drop_lines
#NYI 
#NYI The C<set_drop_lines()> method adds Drop Lines to charts to show the Category value of points in the data.
#NYI 
#NYI     $chart->set_drop_lines();
#NYI 
#NYI It is possible to format the Drop Line C<line> properties if required. See the L</CHART FORMATTING> section below.
#NYI 
#NYI     $chart->set_drop_lines( line => { color => 'red', dash_type => 'square_dot' } );
#NYI 
#NYI Drop Lines are only available in Line, Area and Stock charts.
#NYI 
#NYI 
#NYI =head2 set_high_low_lines
#NYI 
#NYI The C<set_high_low_lines()> method adds High-Low lines to charts to show the maximum and minimum values of points in a Category.
#NYI 
#NYI     $chart->set_high_low_lines();
#NYI 
#NYI It is possible to format the High-Low Line C<line> properties if required. See the L</CHART FORMATTING> section below.
#NYI 
#NYI     $chart->set_high_low_lines( line => { color => 'red' } );
#NYI 
#NYI High-Low Lines are only available in Line and Stock charts.
#NYI 
#NYI 
#NYI =head2 show_blanks_as()
#NYI 
#NYI The C<show_blanks_as()> method controls how blank data is displayed in a chart.
#NYI 
#NYI     $chart->show_blanks_as( 'span' );
#NYI 
#NYI The available options are:
#NYI 
#NYI         gap    # Blank data is shown as a gap. The default.
#NYI         zero   # Blank data is displayed as zero.
#NYI         span   # Blank data is connected with a line.
#NYI 
#NYI 
#NYI =head2 show_hidden_data()
#NYI 
#NYI Display data in hidden rows or columns on the chart.
#NYI 
#NYI     $chart->show_hidden_data();
#NYI 
#NYI 
#NYI =head1 SERIES OPTIONS
#NYI 
#NYI This section details the following properties of C<add_series()> in more detail:
#NYI 
#NYI     marker
#NYI     trendline
#NYI     y_error_bars
#NYI     x_error_bars
#NYI     data_labels
#NYI     points
#NYI     smooth
#NYI 
#NYI =head2 Marker
#NYI 
#NYI The marker format specifies the properties of the markers used to distinguish series on a chart. In general only Line and Scatter chart types and trendlines use markers.
#NYI 
#NYI The following properties can be set for C<marker> formats in a chart.
#NYI 
#NYI     type
#NYI     size
#NYI     border
#NYI     fill
#NYI     pattern
#NYI     gradient
#NYI 
#NYI The C<type> property sets the type of marker that is used with a series.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         marker     => { type => 'diamond' },
#NYI     );
#NYI 
#NYI The following C<type> properties can be set for C<marker> formats in a chart. These are shown in the same order as in the Excel format dialog.
#NYI 
#NYI     automatic
#NYI     none
#NYI     square
#NYI     diamond
#NYI     triangle
#NYI     x
#NYI     star
#NYI     short_dash
#NYI     long_dash
#NYI     circle
#NYI     plus
#NYI 
#NYI The C<automatic> type is a special case which turns on a marker using the default marker style for the particular series number.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         marker     => { type => 'automatic' },
#NYI     );
#NYI 
#NYI If C<automatic> is on then other marker properties such as size, border or fill cannot be set.
#NYI 
#NYI The C<size> property sets the size of the marker and is generally used in conjunction with C<type>.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         marker     => { type => 'diamond', size => 7 },
#NYI     );
#NYI 
#NYI Nested C<border> and C<fill> properties can also be set for a marker. See the L</CHART FORMATTING> section below.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         marker     => {
#NYI             type    => 'square',
#NYI             size    => 5,
#NYI             border  => { color => 'red' },
#NYI             fill    => { color => 'yellow' },
#NYI         },
#NYI     );
#NYI 
#NYI 
#NYI =head2 Trendline
#NYI 
#NYI A trendline can be added to a chart series to indicate trends in the data such as a moving average or a polynomial fit.
#NYI 
#NYI The following properties can be set for trendlines in a chart series.
#NYI 
#NYI     type
#NYI     order               (for polynomial trends)
#NYI     period              (for moving average)
#NYI     forward             (for all except moving average)
#NYI     backward            (for all except moving average)
#NYI     name
#NYI     line
#NYI     intercept           (for exponential, linear and polynomial only)
#NYI     display_equation    (for all except moving average)
#NYI     display_r_squared   (for all except moving average)
#NYI 
#NYI 
#NYI The C<type> property sets the type of trendline in the series.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         trendline  => { type => 'linear' },
#NYI     );
#NYI 
#NYI The available C<trendline> types are:
#NYI 
#NYI     exponential
#NYI     linear
#NYI     log
#NYI     moving_average
#NYI     polynomial
#NYI     power
#NYI 
#NYI A C<polynomial> trendline can also specify the C<order> of the polynomial. The default value is 2.
#NYI 
#NYI     $chart->add_series(
#NYI         values    => '=Sheet1!$B$1:$B$5',
#NYI         trendline => {
#NYI             type  => 'polynomial',
#NYI             order => 3,
#NYI         },
#NYI     );
#NYI 
#NYI A C<moving_average> trendline can also specify the C<period> of the moving average. The default value is 2.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         trendline  => {
#NYI             type   => 'moving_average',
#NYI             period => 3,
#NYI         },
#NYI     );
#NYI 
#NYI The C<forward> and C<backward> properties set the forecast period of the trendline.
#NYI 
#NYI     $chart->add_series(
#NYI         values    => '=Sheet1!$B$1:$B$5',
#NYI         trendline => {
#NYI             type     => 'linear',
#NYI             forward  => 0.5,
#NYI             backward => 0.5,
#NYI         },
#NYI     );
#NYI 
#NYI The C<name> property sets an optional name for the trendline that will appear in the chart legend. If it isn't specified the Excel default name will be displayed. This is usually a combination of the trendline type and the series name.
#NYI 
#NYI     $chart->add_series(
#NYI         values    => '=Sheet1!$B$1:$B$5',
#NYI         trendline => {
#NYI             type => 'linear',
#NYI             name => 'Interpolated trend',
#NYI         },
#NYI     );
#NYI 
#NYI The C<intercept> property sets the point where the trendline crosses the Y (value) axis:
#NYI 
#NYI     $chart->add_series(
#NYI         values    => '=Sheet1!$B$1:$B$5',
#NYI         trendline => {
#NYI             type      => 'linear',
#NYI             intercept => 0.8,
#NYI         },
#NYI     );
#NYI 
#NYI 
#NYI The C<display_equation> property displays the trendline equation on the chart.
#NYI 
#NYI     $chart->add_series(
#NYI         values    => '=Sheet1!$B$1:$B$5',
#NYI         trendline => {
#NYI             type             => 'linear',
#NYI             display_equation => 1,
#NYI         },
#NYI     );
#NYI 
#NYI The C<display_r_squared> property displays the R squared value of the trendline on the chart.
#NYI 
#NYI     $chart->add_series(
#NYI         values    => '=Sheet1!$B$1:$B$5',
#NYI         trendline => {
#NYI             type              => 'linear',
#NYI             display_r_squared => 1
#NYI         },
#NYI     );
#NYI 
#NYI 
#NYI Several of these properties can be set in one go:
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         trendline  => {
#NYI             type              => 'polynomial',
#NYI             name              => 'My trend name',
#NYI             order             => 2,
#NYI             forward           => 0.5,
#NYI             backward          => 0.5,
#NYI             intercept         => 1.5,
#NYI             display_equation  => 1,
#NYI             display_r_squared => 1,
#NYI             line              => {
#NYI                 color     => 'red',
#NYI                 width     => 1,
#NYI                 dash_type => 'long_dash',
#NYI             }
#NYI         },
#NYI     );
#NYI 
#NYI Trendlines cannot be added to series in a stacked chart or pie chart, radar chart, doughnut or (when implemented) to 3D, or surface charts.
#NYI 
#NYI =head2 Error Bars
#NYI 
#NYI Error bars can be added to a chart series to indicate error bounds in the data. The error bars can be vertical C<y_error_bars> (the most common type) or horizontal C<x_error_bars> (for Bar and Scatter charts only).
#NYI 
#NYI The following properties can be set for error bars in a chart series.
#NYI 
#NYI     type
#NYI     value        (for all types except standard error and custom)
#NYI     plus_values  (for custom only)
#NYI     minus_values (for custom only)
#NYI     direction
#NYI     end_style
#NYI     line
#NYI 
#NYI The C<type> property sets the type of error bars in the series.
#NYI 
#NYI     $chart->add_series(
#NYI         values       => '=Sheet1!$B$1:$B$5',
#NYI         y_error_bars => { type => 'standard_error' },
#NYI     );
#NYI 
#NYI The available error bars types are available:
#NYI 
#NYI     fixed
#NYI     percentage
#NYI     standard_deviation
#NYI     standard_error
#NYI     custom
#NYI 
#NYI All error bar types, except for C<standard_error> and C<custom> must also have a value associated with it for the error bounds:
#NYI 
#NYI     $chart->add_series(
#NYI         values       => '=Sheet1!$B$1:$B$5',
#NYI         y_error_bars => {
#NYI             type  => 'percentage',
#NYI             value => 5,
#NYI         },
#NYI     );
#NYI 
#NYI The C<custom> error bar type must specify C<plus_values> and C<minus_values> which should either by a C<Sheet1!$A$1:$A$5> type range formula or an arrayref of
#NYI values:
#NYI 
#NYI     $chart->add_series(
#NYI         categories   => '=Sheet1!$A$1:$A$5',
#NYI         values       => '=Sheet1!$B$1:$B$5',
#NYI         y_error_bars => {
#NYI             type         => 'custom',
#NYI             plus_values  => '=Sheet1!$C$1:$C$5',
#NYI             minus_values => '=Sheet1!$D$1:$D$5',
#NYI         },
#NYI     );
#NYI 
#NYI     # or
#NYI 
#NYI 
#NYI     $chart->add_series(
#NYI         categories   => '=Sheet1!$A$1:$A$5',
#NYI         values       => '=Sheet1!$B$1:$B$5',
#NYI         y_error_bars => {
#NYI             type         => 'custom',
#NYI             plus_values  => [1, 1, 1, 1, 1],
#NYI             minus_values => [2, 2, 2, 2, 2],
#NYI         },
#NYI     );
#NYI 
#NYI Note, as in Excel the items in the C<minus_values> do not need to be negative.
#NYI 
#NYI The C<direction> property sets the direction of the error bars. It should be one of the following:
#NYI 
#NYI     plus    # Positive direction only.
#NYI     minus   # Negative direction only.
#NYI     both    # Plus and minus directions, The default.
#NYI 
#NYI The C<end_style> property sets the style of the error bar end cap. The options are 1 (the default) or 0 (for no end cap):
#NYI 
#NYI     $chart->add_series(
#NYI         values       => '=Sheet1!$B$1:$B$5',
#NYI         y_error_bars => {
#NYI             type      => 'fixed',
#NYI             value     => 2,
#NYI             end_style => 0,
#NYI             direction => 'minus'
#NYI         },
#NYI     );
#NYI 
#NYI 
#NYI 
#NYI =head2 Data Labels
#NYI 
#NYI Data labels can be added to a chart series to indicate the values of the plotted data points.
#NYI 
#NYI The following properties can be set for C<data_labels> formats in a chart.
#NYI 
#NYI     value
#NYI     category
#NYI     series_name
#NYI     position
#NYI     percentage
#NYI     leader_lines
#NYI     separator
#NYI     legend_key
#NYI     num_format
#NYI     font
#NYI 
#NYI The C<value> property turns on the I<Value> data label for a series.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { value => 1 },
#NYI     );
#NYI 
#NYI The C<category> property turns on the I<Category Name> data label for a series.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { category => 1 },
#NYI     );
#NYI 
#NYI 
#NYI The C<series_name> property turns on the I<Series Name> data label for a series.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { series_name => 1 },
#NYI     );
#NYI 
#NYI The C<position> property is used to position the data label for a series.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { value => 1, position => 'center' },
#NYI     );
#NYI 
#NYI In Excel the data label positions vary for different chart types. The allowable positions are:
#NYI 
#NYI     |  Position     |  Line     |  Bar      |  Pie      |  Area     |
#NYI     |               |  Scatter  |  Column   |  Doughnut |  Radar    |
#NYI     |               |  Stock    |           |           |           |
#NYI     |---------------|-----------|-----------|-----------|-----------|
#NYI     |  center       |  Yes      |  Yes      |  Yes      |  Yes*     |
#NYI     |  right        |  Yes*     |           |           |           |
#NYI     |  left         |  Yes      |           |           |           |
#NYI     |  above        |  Yes      |           |           |           |
#NYI     |  below        |  Yes      |           |           |           |
#NYI     |  inside_base  |           |  Yes      |           |           |
#NYI     |  inside_end   |           |  Yes      |  Yes      |           |
#NYI     |  outside_end  |           |  Yes*     |  Yes      |           |
#NYI     |  best_fit     |           |           |  Yes*     |           |
#NYI 
#NYI Note: The * indicates the default position for each chart type in Excel, if a position isn't specified.
#NYI 
#NYI The C<percentage> property is used to turn on the display of data labels as a I<Percentage> for a series. It is mainly used for pie and doughnut charts.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { percentage => 1 },
#NYI     );
#NYI 
#NYI The C<leader_lines> property is used to turn on  I<Leader Lines> for the data label for a series. It is mainly used for pie charts.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { value => 1, leader_lines => 1 },
#NYI     );
#NYI 
#NYI Note: Even when leader lines are turned on they aren't automatically visible in Excel or Excel::Writer::XLSX. Due to an Excel limitation (or design) leader lines only appear if the data label is moved manually or if the data labels are very close and need to be adjusted automatically.
#NYI 
#NYI The C<separator> property is used to change the separator between multiple data label items:
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { percentage => 1 },
#NYI         data_labels => { value => 1, category => 1, separator => "\n" },
#NYI     );
#NYI 
#NYI The separator value must be one of the following strings:
#NYI 
#NYI             ','
#NYI             ';'
#NYI             '.'
#NYI             "\n"
#NYI             ' '
#NYI 
#NYI The C<legend_key> property is used to turn on  I<Legend Key> for the data label for a series:
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$B$1:$B$5',
#NYI         data_labels => { value => 1, legend_key => 1 },
#NYI     );
#NYI 
#NYI 
#NYI The C<num_format> property is used to set the number format for the data labels.
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$A$1:$A$5',
#NYI         data_labels => { value => 1, num_format => '#,##0.00' },
#NYI     );
#NYI 
#NYI The number format is similar to the Worksheet Cell Format C<num_format> apart from the fact that a format index cannot be used. The explicit format string must be used as shown above. See L<Excel::Writer::XLSX/set_num_format()> for more information.
#NYI 
#NYI The C<font> property is used to set the font properties of the data labels in a series:
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$A$1:$A$5',
#NYI         data_labels => {
#NYI             value => 1,
#NYI             font  => { name => 'Consolas' }
#NYI         },
#NYI     );
#NYI 
#NYI The C<font> property is also used to rotate the data labels in a series:
#NYI 
#NYI     $chart->add_series(
#NYI         values      => '=Sheet1!$A$1:$A$5',
#NYI         data_labels => {
#NYI             value => 1,
#NYI             font  => { rotation => 45 }
#NYI         },
#NYI     );
#NYI 
#NYI See the L</CHART FONTS> section below.
#NYI 
#NYI 
#NYI =head2 Points
#NYI 
#NYI In general formatting is applied to an entire series in a chart. However, it is occasionally required to format individual points in a series. In particular this is required for Pie and Doughnut charts where each segment is represented by a point.
#NYI 
#NYI In these cases it is possible to use the C<points> property of C<add_series()>:
#NYI 
#NYI     $chart->add_series(
#NYI         values => '=Sheet1!$A$1:$A$3',
#NYI         points => [
#NYI             { fill => { color => '#FF0000' } },
#NYI             { fill => { color => '#CC0000' } },
#NYI             { fill => { color => '#990000' } },
#NYI         ],
#NYI     );
#NYI 
#NYI The C<points> property takes an array ref of format options (see the L</CHART FORMATTING> section below). To assign default properties to points in a series pass C<undef> values in the array ref:
#NYI 
#NYI     # Format point 3 of 3 only.
#NYI     $chart->add_series(
#NYI         values => '=Sheet1!$A$1:$A$3',
#NYI         points => [
#NYI             undef,
#NYI             undef,
#NYI             { fill => { color => '#990000' } },
#NYI         ],
#NYI     );
#NYI 
#NYI     # Format the first point only.
#NYI     $chart->add_series(
#NYI         values => '=Sheet1!$A$1:$A$3',
#NYI         points => [ { fill => { color => '#FF0000' } } ],
#NYI     );
#NYI 
#NYI =head2 Smooth
#NYI 
#NYI The C<smooth> option is used to set the smooth property of a line series. It is only applicable to the C<Line> and C<Scatter> chart types.
#NYI 
#NYI     $chart->add_series( values => '=Sheet1!$C$1:$C$5',
#NYI                         smooth => 1 );
#NYI 
#NYI 
#NYI =head1 CHART FORMATTING
#NYI 
#NYI The following chart formatting properties can be set for any chart object that they apply to (and that are supported by Excel::Writer::XLSX) such as chart lines, column fill areas, plot area borders, markers, gridlines and other chart elements documented above.
#NYI 
#NYI     line
#NYI     border
#NYI     fill
#NYI     pattern
#NYI     gradient
#NYI 
#NYI Chart formatting properties are generally set using hash refs.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => { color => 'blue' },
#NYI     );
#NYI 
#NYI In some cases the format properties can be nested. For example a C<marker> may contain C<border> and C<fill> sub-properties.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => { color => 'blue' },
#NYI         marker     => {
#NYI             type    => 'square',
#NYI             size    => 5,
#NYI             border  => { color => 'red' },
#NYI             fill    => { color => 'yellow' },
#NYI         },
#NYI     );
#NYI 
#NYI =head2 Line
#NYI 
#NYI The line format is used to specify properties of line objects that appear in a chart such as a plotted line on a chart or a border.
#NYI 
#NYI The following properties can be set for C<line> formats in a chart.
#NYI 
#NYI     none
#NYI     color
#NYI     width
#NYI     dash_type
#NYI 
#NYI 
#NYI The C<none> property is uses to turn the C<line> off (it is always on by default except in Scatter charts). This is useful if you wish to plot a series with markers but without a line.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => { none => 1 },
#NYI     );
#NYI 
#NYI 
#NYI The C<color> property sets the color of the C<line>.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => { color => 'red' },
#NYI     );
#NYI 
#NYI The available colours are shown in the main L<Excel::Writer::XLSX> documentation. It is also possible to set the colour of a line with a HTML style RGB colour:
#NYI 
#NYI     $chart->add_series(
#NYI         line       => { color => '#FF0000' },
#NYI     );
#NYI 
#NYI 
#NYI The C<width> property sets the width of the C<line>. It should be specified in increments of 0.25 of a point as in Excel.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => { width => 3.25 },
#NYI     );
#NYI 
#NYI The C<dash_type> property sets the dash style of the line.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => { dash_type => 'dash_dot' },
#NYI     );
#NYI 
#NYI The following C<dash_type> values are available. They are shown in the order that they appear in the Excel dialog.
#NYI 
#NYI     solid
#NYI     round_dot
#NYI     square_dot
#NYI     dash
#NYI     dash_dot
#NYI     long_dash
#NYI     long_dash_dot
#NYI     long_dash_dot_dot
#NYI 
#NYI The default line style is C<solid>.
#NYI 
#NYI More than one C<line> property can be specified at a time:
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         line       => {
#NYI             color     => 'red',
#NYI             width     => 1.25,
#NYI             dash_type => 'square_dot',
#NYI         },
#NYI     );
#NYI 
#NYI =head2 Border
#NYI 
#NYI The C<border> property is a synonym for C<line>.
#NYI 
#NYI It can be used as a descriptive substitute for C<line> in chart types such as Bar and Column that have a border and fill style rather than a line style. In general chart objects with a C<border> property will also have a fill property.
#NYI 
#NYI 
#NYI =head2 Solid Fill
#NYI 
#NYI The fill format is used to specify filled areas of chart objects such as the interior of a column or the background of the chart itself.
#NYI 
#NYI The following properties can be set for C<fill> formats in a chart.
#NYI 
#NYI     none
#NYI     color
#NYI     transparency
#NYI 
#NYI The C<none> property is used to turn the C<fill> property off (it is generally on by default).
#NYI 
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         fill       => { none => 1 },
#NYI     );
#NYI 
#NYI The C<color> property sets the colour of the C<fill> area.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         fill       => { color => 'red' },
#NYI     );
#NYI 
#NYI The available colours are shown in the main L<Excel::Writer::XLSX> documentation. It is also possible to set the colour of a fill with a HTML style RGB colour:
#NYI 
#NYI     $chart->add_series(
#NYI         fill       => { color => '#FF0000' },
#NYI     );
#NYI 
#NYI The C<transparency> property sets the transparency of the solid fill color in the integer range 1 - 100:
#NYI 
#NYI     $chart->set_chartarea( fill => { color => 'yellow', transparency => 75 } );
#NYI 
#NYI The C<fill> format is generally used in conjunction with a C<border> format which has the same properties as a C<line> format.
#NYI 
#NYI     $chart->add_series(
#NYI         values     => '=Sheet1!$B$1:$B$5',
#NYI         border     => { color => 'red' },
#NYI         fill       => { color => 'yellow' },
#NYI     );
#NYI 
#NYI 
#NYI 
#NYI =head2 Pattern Fill
#NYI 
#NYI The pattern fill format is used to specify pattern filled areas of chart objects such as the interior of a column or the background of the chart itself.
#NYI 
#NYI The following properties can be set for C<pattern> fill formats in a chart:
#NYI 
#NYI     pattern:   the pattern to be applied (required)
#NYI     fg_color:  the foreground color of the pattern (required)
#NYI     bg_color:  the background color (optional, defaults to white)
#NYI 
#NYI 
#NYI For example:
#NYI 
#NYI     $chart->set_plotarea(
#NYI         pattern => {
#NYI             pattern  => 'percent_5',
#NYI             fg_color => 'red',
#NYI             bg_color => 'yellow',
#NYI         }
#NYI     );
#NYI 
#NYI The following patterns can be applied:
#NYI 
#NYI     percent_5
#NYI     percent_10
#NYI     percent_20
#NYI     percent_25
#NYI     percent_30
#NYI     percent_40
#NYI     percent_50
#NYI     percent_60
#NYI     percent_70
#NYI     percent_75
#NYI     percent_80
#NYI     percent_90
#NYI     light_downward_diagonal
#NYI     light_upward_diagonal
#NYI     dark_downward_diagonal
#NYI     dark_upward_diagonal
#NYI     wide_downward_diagonal
#NYI     wide_upward_diagonal
#NYI     light_vertical
#NYI     light_horizontal
#NYI     narrow_vertical
#NYI     narrow_horizontal
#NYI     dark_vertical
#NYI     dark_horizontal
#NYI     dashed_downward_diagonal
#NYI     dashed_upward_diagonal
#NYI     dashed_horizontal
#NYI     dashed_vertical
#NYI     small_confetti
#NYI     large_confetti
#NYI     zigzag
#NYI     wave
#NYI     diagonal_brick
#NYI     horizontal_brick
#NYI     weave
#NYI     plaid
#NYI     divot
#NYI     dotted_grid
#NYI     dotted_diamond
#NYI     shingle
#NYI     trellis
#NYI     sphere
#NYI     small_grid
#NYI     large_grid
#NYI     small_check
#NYI     large_check
#NYI     outlined_diamond
#NYI     solid_diamond
#NYI 
#NYI 
#NYI The foreground color, C<fg_color>, is a required parameter and can be a Html style C<#RRGGBB> string or a limited number of named colors. The available colours are shown in the main L<Excel::Writer::XLSX> documentation.
#NYI 
#NYI The background color, C<bg_color>, is optional and defaults to black.
#NYI 
#NYI If a pattern fill is used on a chart object it overrides the solid fill properties of the object.
#NYI 
#NYI 
#NYI =head2 Gradient Fill
#NYI 
#NYI The gradient fill format is used to specify gradient filled areas of chart objects such as the interior of a column or the background of the chart itself.
#NYI 
#NYI 
#NYI The following properties can be set for C<gradient> fill formats in a chart:
#NYI 
#NYI     colors:    a list of colors
#NYI     positions: an optional list of positions for the colors
#NYI     type:      the optional type of gradient fill
#NYI     angle:     the optional angle of the linear fill
#NYI 
#NYI The C<colors> property sets a list of colors that define the C<gradient>:
#NYI 
#NYI     $chart->set_plotarea(
#NYI         gradient => { colors => [ '#DDEBCF', '#9CB86E', '#156B13' ] }
#NYI     );
#NYI 
#NYI Excel allows between 2 and 10 colors in a gradient but it is unlikely that you will require more than 2 or 3.
#NYI 
#NYI As with solid or pattern fill it is also possible to set the colors of a gradient with a Html style C<#RRGGBB> string or a limited number of named colors. The available colours are shown in the main L<Excel::Writer::XLSX> documentation:
#NYI 
#NYI     $chart->add_series(
#NYI         values   => '=Sheet1!$A$1:$A$5',
#NYI         gradient => { colors => [ 'red', 'green' ] }
#NYI     );
#NYI 
#NYI The C<positions> defines an optional list of positions, between 0 and 100, of
#NYI where the colors in the gradient are located. Default values are provided for
#NYI C<colors> lists of between 2 and 4 but they can be specified if required:
#NYI 
#NYI     $chart->add_series(
#NYI         values   => '=Sheet1!$A$1:$A$5',
#NYI         gradient => {
#NYI             colors    => [ '#DDEBCF', '#156B13' ],
#NYI             positions => [ 10,        90 ],
#NYI         }
#NYI     );
#NYI 
#NYI The C<type> property can have one of the following values:
#NYI 
#NYI     linear        (the default)
#NYI     radial
#NYI     rectangular
#NYI     path
#NYI 
#NYI For example:
#NYI 
#NYI     $chart->add_series(
#NYI         values   => '=Sheet1!$A$1:$A$5',
#NYI         gradient => {
#NYI             colors => [ '#DDEBCF', '#9CB86E', '#156B13' ],
#NYI             type   => 'radial'
#NYI         }
#NYI     );
#NYI 
#NYI If C<type> isn't specified it defaults to C<linear>.
#NYI 
#NYI For a C<linear> fill the angle of the gradient can also be specified:
#NYI 
#NYI     $chart->add_series(
#NYI         values   => '=Sheet1!$A$1:$A$5',
#NYI         gradient => { colors => [ '#DDEBCF', '#9CB86E', '#156B13' ],
#NYI                       angle => 30 }
#NYI     );
#NYI 
#NYI The default angle is 90 degrees.
#NYI 
#NYI If gradient fill is used on a chart object it overrides the solid fill and pattern fill properties of the object.
#NYI 
#NYI 
#NYI 
#NYI 
#NYI =head1 CHART FONTS
#NYI 
#NYI The following font properties can be set for any chart object that they apply to (and that are supported by Excel::Writer::XLSX) such as chart titles, axis labels, axis numbering and data labels. They correspond to the equivalent Worksheet cell Format object properties. See L<Excel::Writer::XLSX/FORMAT_METHODS> for more information.
#NYI 
#NYI     name
#NYI     size
#NYI     bold
#NYI     italic
#NYI     underline
#NYI     rotation
#NYI     color
#NYI 
#NYI The following explains the available font properties:
#NYI 
#NYI =over
#NYI 
#NYI =item * C<name>
#NYI 
#NYI Set the font name:
#NYI 
#NYI     $chart->set_x_axis( num_font => { name => 'Arial' } );
#NYI 
#NYI =item * C<size>
#NYI 
#NYI Set the font size:
#NYI 
#NYI     $chart->set_x_axis( num_font => { name => 'Arial', size => 10 } );
#NYI 
#NYI =item * C<bold>
#NYI 
#NYI Set the font bold property, should be 0 or 1:
#NYI 
#NYI     $chart->set_x_axis( num_font => { bold => 1 } );
#NYI 
#NYI =item * C<italic>
#NYI 
#NYI Set the font italic property, should be 0 or 1:
#NYI 
#NYI     $chart->set_x_axis( num_font => { italic => 1 } );
#NYI 
#NYI =item * C<underline>
#NYI 
#NYI Set the font underline property, should be 0 or 1:
#NYI 
#NYI     $chart->set_x_axis( num_font => { underline => 1 } );
#NYI 
#NYI =item * C<rotation>
#NYI 
#NYI Set the font rotation in the range -90 to 90:
#NYI 
#NYI     $chart->set_x_axis( num_font => { rotation => 45 } );
#NYI 
#NYI This is useful for displaying large axis data such as dates in a more compact format.
#NYI 
#NYI =item * C<color>
#NYI 
#NYI Set the font color property. Can be a color index, a color name or HTML style RGB colour:
#NYI 
#NYI     $chart->set_x_axis( num_font => { color => 'red' } );
#NYI     $chart->set_y_axis( num_font => { color => '#92D050' } );
#NYI 
#NYI =back
#NYI 
#NYI Here is an example of Font formatting in a Chart program:
#NYI 
#NYI     # Format the chart title.
#NYI     $chart->set_title(
#NYI         name      => 'Sales Results Chart',
#NYI         name_font => {
#NYI             name  => 'Calibri',
#NYI             color => 'yellow',
#NYI         },
#NYI     );
#NYI 
#NYI     # Format the X-axis.
#NYI     $chart->set_x_axis(
#NYI         name      => 'Month',
#NYI         name_font => {
#NYI             name  => 'Arial',
#NYI             color => '#92D050'
#NYI         },
#NYI         num_font => {
#NYI             name  => 'Courier New',
#NYI             color => '#00B0F0',
#NYI         },
#NYI     );
#NYI 
#NYI     # Format the Y-axis.
#NYI     $chart->set_y_axis(
#NYI         name      => 'Sales (1000 units)',
#NYI         name_font => {
#NYI             name      => 'Century',
#NYI             underline => 1,
#NYI             color     => 'red'
#NYI         },
#NYI         num_font => {
#NYI             bold   => 1,
#NYI             italic => 1,
#NYI             color  => '#7030A0',
#NYI         },
#NYI     );
#NYI 
#NYI 
#NYI 
#NYI =head1 CHART LAYOUT
#NYI 
#NYI The position of the chart in the worksheet is controlled by the C<set_size()> method shown above.
#NYI 
#NYI It is also possible to change the layout of the following chart sub-objects:
#NYI 
#NYI     plotarea
#NYI     legend
#NYI     title
#NYI     x_axis caption
#NYI     y_axis caption
#NYI 
#NYI Here are some examples:
#NYI 
#NYI     $chart->set_plotarea(
#NYI         layout => {
#NYI             x      => 0.35,
#NYI             y      => 0.26,
#NYI             width  => 0.62,
#NYI             height => 0.50,
#NYI         }
#NYI     );
#NYI 
#NYI     $chart->set_legend(
#NYI         layout => {
#NYI             x      => 0.80,
#NYI             y      => 0.37,
#NYI             width  => 0.12,
#NYI             height => 0.25,
#NYI         }
#NYI     );
#NYI 
#NYI     $chart->set_title(
#NYI         name   => 'Title',
#NYI         layout => {
#NYI             x => 0.42,
#NYI             y => 0.14,
#NYI         }
#NYI     );
#NYI 
#NYI     $chart->set_x_axis(
#NYI         name        => 'X axis',
#NYI         name_layout => {
#NYI             x => 0.34,
#NYI             y => 0.85,
#NYI         }
#NYI     );
#NYI 
#NYI Note that it is only possible to change the width and height for the C<plotarea> and C<legend> objects. For the other text based objects the width and height are changed by the font dimensions.
#NYI 
#NYI The layout units must be a float in the range C<0 < x <= 1> and are expressed as a percentage of the chart dimensions as shown below:
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/layout.png" width="826" height="423" alt="Chart object layout." /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI From this the layout units are calculated as follows:
#NYI 
#NYI     layout:
#NYI         width  = w / W
#NYI         height = h / H
#NYI         x      = a / W
#NYI         y      = b / H
#NYI 
#NYI These units are slightly cumbersome but are required by Excel so that the chart object positions remain relative to each other if the chart is resized by the user.
#NYI 
#NYI Note that for C<plotarea> the origin is the top left corner in the plotarea itself and does not take into account the axes.
#NYI 
#NYI 
#NYI =head1 WORKSHEET METHODS
#NYI 
#NYI In Excel a chartsheet (i.e, a chart that isn't embedded) shares properties with data worksheets such as tab selection, headers, footers, margins, and print properties.
#NYI 
#NYI In Excel::Writer::XLSX you can set chartsheet properties using the same methods that are used for Worksheet objects.
#NYI 
#NYI The following Worksheet methods are also available through a non-embedded Chart object:
#NYI 
#NYI     get_name()
#NYI     activate()
#NYI     select()
#NYI     hide()
#NYI     set_first_sheet()
#NYI     protect()
#NYI     set_zoom()
#NYI     set_tab_color()
#NYI 
#NYI     set_landscape()
#NYI     set_portrait()
#NYI     set_paper()
#NYI     set_margins()
#NYI     set_header()
#NYI     set_footer()
#NYI 
#NYI See L<Excel::Writer::XLSX> for a detailed explanation of these methods.
#NYI 
#NYI =head1 EXAMPLE
#NYI 
#NYI Here is a complete example that demonstrates some of the available features when creating a chart.
#NYI 
#NYI     #!/usr/bin/perl
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'chart.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI     my $bold      = $workbook->add_format( bold => 1 );
#NYI 
#NYI     # Add the worksheet data that the charts will refer to.
#NYI     my $headings = [ 'Number', 'Batch 1', 'Batch 2' ];
#NYI     my $data = [
#NYI         [ 2,  3,  4,  5,  6,  7 ],
#NYI         [ 10, 40, 50, 20, 10, 50 ],
#NYI         [ 30, 60, 70, 50, 40, 30 ],
#NYI 
#NYI     ];
#NYI 
#NYI     $worksheet->write( 'A1', $headings, $bold );
#NYI     $worksheet->write( 'A2', $data );
#NYI 
#NYI     # Create a new chart object. In this case an embedded chart.
#NYI     my $chart = $workbook->add_chart( type => 'column', embedded => 1 );
#NYI 
#NYI     # Configure the first series.
#NYI     $chart->add_series(
#NYI         name       => '=Sheet1!$B$1',
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$B$2:$B$7',
#NYI     );
#NYI 
#NYI     # Configure second series. Note alternative use of array ref to define
#NYI     # ranges: [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
#NYI     $chart->add_series(
#NYI         name       => '=Sheet1!$C$1',
#NYI         categories => [ 'Sheet1', 1, 6, 0, 0 ],
#NYI         values     => [ 'Sheet1', 1, 6, 2, 2 ],
#NYI     );
#NYI 
#NYI     # Add a chart title and some axis labels.
#NYI     $chart->set_title ( name => 'Results of sample analysis' );
#NYI     $chart->set_x_axis( name => 'Test number' );
#NYI     $chart->set_y_axis( name => 'Sample length (mm)' );
#NYI 
#NYI     # Set an Excel chart style. Blue colors with white outline and shadow.
#NYI     $chart->set_style( 11 );
#NYI 
#NYI     # Insert the chart into the worksheet (with an offset).
#NYI     $worksheet->insert_chart( 'D2', $chart, 25, 10 );
#NYI 
#NYI     __END__
#NYI 
#NYI =begin html
#NYI 
#NYI <p>This will produce a chart that looks like this:</p>
#NYI 
#NYI <p><center><img src="http://jmcnamara.github.io/excel-writer-xlsx/images/examples/area1.jpg" width="527" height="320" alt="Chart example." /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI =head1 Value and Category Axes
#NYI 
#NYI Excel differentiates between a chart axis that is used for series B<categories> and an axis that is used for series B<values>.
#NYI 
#NYI In the example above the X axis is the category axis and each of the values is evenly spaced. The Y axis (in this case) is the value axis and points are displayed according to their value.
#NYI 
#NYI Since Excel treats the axes differently it also handles their formatting differently and exposes different properties for each.
#NYI 
#NYI As such some of C<Excel::Writer::XLSX> axis properties can be set for a value axis, some can be set for a category axis and some properties can be set for both.
#NYI 
#NYI For example the C<min> and C<max> properties can only be set for value axes and C<reverse> can be set for both. The type of axis that a property applies to is shown in the C<set_x_axis()> section of the documentation above.
#NYI 
#NYI Some charts such as C<Scatter> and C<Stock> have two value axes.
#NYI 
#NYI Date Axes are a special type of category axis which are explained below.
#NYI 
#NYI =head1 Date Category Axes
#NYI 
#NYI Date Category Axes are category axes that display time or date information. In Excel::Writer::XLSX Date Category Axes are set using the C<date_axis> option:
#NYI 
#NYI     $chart->set_x_axis( date_axis => 1 );
#NYI 
#NYI In general you should also specify a number format for a date axis although Excel will usually default to the same format as the data being plotted:
#NYI 
#NYI     $chart->set_x_axis(
#NYI         date_axis         => 1,
#NYI         num_format        => 'dd/mm/yyyy',
#NYI     );
#NYI 
#NYI Excel doesn't normally allow minimum and maximum values to be set for category axes. However, date axes are an exception. The C<min> and C<max> values should be set as Excel times or dates:
#NYI 
#NYI     $chart->set_x_axis(
#NYI         date_axis         => 1,
#NYI         min               => $worksheet->convert_date_time('2013-01-02T'),
#NYI         max               => $worksheet->convert_date_time('2013-01-09T'),
#NYI         num_format        => 'dd/mm/yyyy',
#NYI     );
#NYI 
#NYI For date axes it is also possible to set the type of the major and minor units:
#NYI 
#NYI     $chart->set_x_axis(
#NYI         date_axis         => 1,
#NYI         minor_unit        => 4,
#NYI         minor_unit_type   => 'months',
#NYI         major_unit        => 1,
#NYI         major_unit_type   => 'years',
#NYI         num_format        => 'dd/mm/yyyy',
#NYI     );
#NYI 
#NYI 
#NYI =head1 Secondary Axes
#NYI 
#NYI It is possible to add a secondary axis of the same type to a chart by setting the C<y2_axis> or C<x2_axis> property of the series:
#NYI 
#NYI     #!/usr/bin/perl
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'chart_secondary_axis.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # Add the worksheet data that the charts will refer to.
#NYI     my $data = [
#NYI         [ 2,  3,  4,  5,  6,  7 ],
#NYI         [ 10, 40, 50, 20, 10, 50 ],
#NYI 
#NYI     ];
#NYI 
#NYI     $worksheet->write( 'A1', $data );
#NYI 
#NYI     # Create a new chart object. In this case an embedded chart.
#NYI     my $chart = $workbook->add_chart( type => 'line', embedded => 1 );
#NYI 
#NYI     # Configure a series with a secondary axis
#NYI     $chart->add_series(
#NYI         values  => '=Sheet1!$A$1:$A$6',
#NYI         y2_axis => 1,
#NYI     );
#NYI 
#NYI     $chart->add_series(
#NYI         values => '=Sheet1!$B$1:$B$6',
#NYI     );
#NYI 
#NYI 
#NYI     # Insert the chart into the worksheet.
#NYI     $worksheet->insert_chart( 'D2', $chart );
#NYI 
#NYI     __END__
#NYI 
#NYI It is also possible to have a secondary, combined, chart either with a shared or secondary axis, see below.
#NYI 
#NYI =head1 Combined Charts
#NYI 
#NYI It is also possible to combine two different chart types, for example a column and line chart to create a Pareto chart using the Chart C<combine()> method:
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="https://raw.githubusercontent.com/jmcnamara/XlsxWriter/master/dev/docs/source/_images/chart_pareto.png" alt="Chart image." /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI Here is a simpler example:
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'chart_combined.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI     my $bold      = $workbook->add_format( bold => 1 );
#NYI 
#NYI     # Add the worksheet data that the charts will refer to.
#NYI     my $headings = [ 'Number', 'Batch 1', 'Batch 2' ];
#NYI     my $data = [
#NYI         [ 2,  3,  4,  5,  6,  7 ],
#NYI         [ 10, 40, 50, 20, 10, 50 ],
#NYI         [ 30, 60, 70, 50, 40, 30 ],
#NYI 
#NYI     ];
#NYI 
#NYI     $worksheet->write( 'A1', $headings, $bold );
#NYI     $worksheet->write( 'A2', $data );
#NYI 
#NYI     #
#NYI     # In the first example we will create a combined column and line chart.
#NYI     # They will share the same X and Y axes.
#NYI     #
#NYI 
#NYI     # Create a new column chart. This will use this as the primary chart.
#NYI     my $column_chart = $workbook->add_chart( type => 'column', embedded => 1 );
#NYI 
#NYI     # Configure the data series for the primary chart.
#NYI     $column_chart->add_series(
#NYI         name       => '=Sheet1!$B$1',
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$B$2:$B$7',
#NYI     );
#NYI 
#NYI     # Create a new column chart. This will use this as the secondary chart.
#NYI     my $line_chart = $workbook->add_chart( type => 'line', embedded => 1 );
#NYI 
#NYI     # Configure the data series for the secondary chart.
#NYI     $line_chart->add_series(
#NYI         name       => '=Sheet1!$C$1',
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$C$2:$C$7',
#NYI     );
#NYI 
#NYI     # Combine the charts.
#NYI     $column_chart->combine( $line_chart );
#NYI 
#NYI     # Add a chart title and some axis labels. Note, this is done via the
#NYI     # primary chart.
#NYI     $column_chart->set_title( name => 'Combined chart - same Y axis' );
#NYI     $column_chart->set_x_axis( name => 'Test number' );
#NYI     $column_chart->set_y_axis( name => 'Sample length (mm)' );
#NYI 
#NYI 
#NYI     # Insert the chart into the worksheet
#NYI     $worksheet->insert_chart( 'E2', $column_chart );
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="https://raw.githubusercontent.com/jmcnamara/XlsxWriter/master/dev/docs/source/_images/chart_combined1.png" alt="Chart image." /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI 
#NYI The secondary chart can also be placed on a secondary axis using the methods shown in the previous section.
#NYI 
#NYI In this case it is just necessary to add a C<y2_axis> parameter to the series and, if required, add a title using C<set_y2_axis()> B<of the secondary chart>. The following are the additions to the previous example to place the secondary chart on the secondary axis:
#NYI 
#NYI     ...
#NYI 
#NYI     $line_chart->add_series(
#NYI         name       => '=Sheet1!$C$1',
#NYI         categories => '=Sheet1!$A$2:$A$7',
#NYI         values     => '=Sheet1!$C$2:$C$7',
#NYI         y2_axis    => 1,
#NYI     );
#NYI 
#NYI     ...
#NYI 
#NYI     # Note: the y2 properites are on the secondary chart.
#NYI     $line_chart2->set_y2_axis( name => 'Target length (mm)' );
#NYI 
#NYI 
#NYI =begin html
#NYI 
#NYI <p><center><img src="https://raw.githubusercontent.com/jmcnamara/XlsxWriter/master/dev/docs/source/_images/chart_combined2.png" alt="Chart image." /></center></p>
#NYI 
#NYI =end html
#NYI 
#NYI 
#NYI The examples above use the concept of a I<primary> and I<secondary> chart. The primary chart is the chart that defines the primary X and Y axis. It is also used for setting all chart properties apart from the secondary data series. For example the chart title and axes properties should be set via the primary chart (except for the the secondary C<y2> axis properties which should be applied to the secondary chart).
#NYI 
#NYI See also C<chart_combined.pl> and C<chart_pareto.pl> examples in the distro for more detailed
#NYI examples.
#NYI 
#NYI There are some limitations on combined charts:
#NYI 
#NYI =over
#NYI 
#NYI =item * Pie charts cannot currently be combined.
#NYI 
#NYI =item * Scatter charts cannot currently be used as a primary chart but they can be used as a secondary chart.
#NYI 
#NYI =item * Bar charts can only combined secondary charts on a secondary axis. This is an Excel limitation.
#NYI 
#NYI =back
#NYI 
#NYI 
#NYI 
#NYI =head1 TODO
#NYI 
#NYI Chart features that are on the TODO list and will hopefully be added are:
#NYI 
#NYI =over
#NYI 
#NYI =item * Add more chart sub-types.
#NYI 
#NYI =item * Additional formatting options.
#NYI 
#NYI =item * More axis controls.
#NYI 
#NYI =item * 3D charts.
#NYI 
#NYI =item * Additional chart types.
#NYI 
#NYI =back
#NYI 
#NYI If you are interested in sponsoring a feature to have it implemented or expedited let me know.
#NYI 
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
