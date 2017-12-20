#NYI package Excel::Writer::XLSX::Shape;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # Shape - A class for writing Excel shapes.
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
#NYI use Exporter;
#NYI 
#NYI our @ISA     = qw(Exporter);
#NYI our $VERSION = '0.96';
#NYI our $AUTOLOAD;
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # new()
#NYI #
#NYI sub new {
#NYI 
#NYI     my $class      = shift;
#NYI     my $fh         = shift;
#NYI     my $self       = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );
#NYI 
#NYI     my %properties = @_;
#NYI 
#NYI     $self->{_name} = undef;
#NYI     $self->{_type} = 'rect';
#NYI 
#NYI     # Is a Connector shape. 1/0 Value is a hash lookup from type.
#NYI     $self->{_connect} = 0;
#NYI 
#NYI     # Is a Drawing. Always 0, since a single shape never fills an entire sheet.
#NYI     $self->{_drawing} = 0;
#NYI 
#NYI     # OneCell or Absolute: options to move and/or size with cells.
#NYI     $self->{_editAs} = '';
#NYI 
#NYI     # Auto-incremented, unless supplied by user.
#NYI     $self->{_id} = 0;
#NYI 
#NYI     # Shape text (usually centered on shape geometry).
#NYI     $self->{_text} = 0;
#NYI 
#NYI     # Shape stencil mode.  A copy (child) is created when inserted.
#NYI     # The link to parent is broken.
#NYI     $self->{_stencil} = 1;
#NYI 
#NYI     # Index to _shapes array when inserted.
#NYI     $self->{_element} = -1;
#NYI 
#NYI     # Shape ID of starting connection, if any.
#NYI     $self->{_start} = undef;
#NYI 
#NYI     # Shape vertex, starts at 0, numbered clockwise from 12 o'clock.
#NYI     $self->{_start_index} = undef;
#NYI 
#NYI     $self->{_end}       = undef;
#NYI     $self->{_end_index} = undef;
#NYI 
#NYI     # Number and size of adjustments for shapes (usually connectors).
#NYI     $self->{_adjustments} = [];
#NYI 
#NYI     # Start and end sides. t)op, b)ottom, l)eft, or r)ight.
#NYI     $self->{_start_side} = '';
#NYI     $self->{_end_side}   = '';
#NYI 
#NYI     # Flip shape Horizontally. eg. arrow left to arrow right.
#NYI     $self->{_flip_h} = 0;
#NYI 
#NYI     # Flip shape Vertically. eg. up arrow to down arrow.
#NYI     $self->{_flip_v} = 0;
#NYI 
#NYI     # shape rotation (in degrees 0-360).
#NYI     $self->{_rotation} = 0;
#NYI 
#NYI     # An alternate way to create a text box, because Excel allows it.
#NYI     # It is just a rectangle with text.
#NYI     $self->{_txBox} = 0;
#NYI 
#NYI     # Shape outline colour, or 0 for noFill (default black).
#NYI     $self->{_line} = '000000';
#NYI 
#NYI     # Line type: dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot.
#NYI     $self->{_line_type} = '';
#NYI 
#NYI     # Line weight (integer).
#NYI     $self->{_line_weight} = 1;
#NYI 
#NYI     # Shape fill colour, or 0 for noFill (default noFill).
#NYI     $self->{_fill} = 0;
#NYI 
#NYI     # Formatting for shape text, if any.
#NYI     $self->{_format} = {};
#NYI 
#NYI     # copy of colour palette table from Workbook.pm.
#NYI     $self->{_palette} = [];
#NYI 
#NYI     # Vertical alignment: t, ctr, b.
#NYI     $self->{_valign} = 'ctr';
#NYI 
#NYI     # Alignment: l, ctr, r, just
#NYI     $self->{_align} = 'ctr';
#NYI 
#NYI     $self->{_x_offset} = 0;
#NYI     $self->{_y_offset} = 0;
#NYI 
#NYI     # Scale factors, which also may be set when the shape is inserted.
#NYI     $self->{_scale_x} = 1;
#NYI     $self->{_scale_y} = 1;
#NYI 
#NYI     # Default size, which can be modified and/or scaled.
#NYI     $self->{_width}  = 50;
#NYI     $self->{_height} = 50;
#NYI 
#NYI     # Initial assignment. May be modified when prepared.
#NYI     $self->{_column_start} = 0;
#NYI     $self->{_row_start}    = 0;
#NYI     $self->{_x1}           = 0;
#NYI     $self->{_y1}           = 0;
#NYI     $self->{_column_end}   = 0;
#NYI     $self->{_row_end}      = 0;
#NYI     $self->{_x2}           = 0;
#NYI     $self->{_y2}           = 0;
#NYI     $self->{_x_abs}        = 0;
#NYI     $self->{_y_abs}        = 0;
#NYI 
#NYI     # Override default properties with passed arguments
#NYI     while ( my ( $key, $value ) = each( %properties ) ) {
#NYI 
#NYI         # Strip leading "-" from Tk style properties e.g. -color => 'red'.
#NYI         $key =~ s/^-//;
#NYI 
#NYI         # Add leading underscore "_" to internal hash keys, if not supplied.
#NYI         $key = "_" . $key unless $key =~ m/^_/;
#NYI 
#NYI         $self->{$key} = $value;
#NYI     }
#NYI 
#NYI     bless $self, $class;
#NYI     return $self;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_properties( name => 'Shape 1', type => 'rect' )
#NYI #
#NYI # Set shape properties.
#NYI #
#NYI sub set_properties {
#NYI 
#NYI     my $self       = shift;
#NYI     my %properties = @_;
#NYI 
#NYI     # Update properties with passed arguments.
#NYI     while ( my ( $key, $value ) = each( %properties ) ) {
#NYI 
#NYI         # Strip leading "-" from Tk style properties e.g. -color => 'red'.
#NYI         $key =~ s/^-//;
#NYI 
#NYI         # Add leading underscore "_" to internal hash keys, if not supplied.
#NYI         $key = "_" . $key unless $key =~ m/^_/;
#NYI 
#NYI         if ( !exists $self->{$key} ) {
#NYI             warn "Unknown shape property: $key. Property not set.\n";
#NYI             next;
#NYI         }
#NYI 
#NYI         $self->{$key} = $value;
#NYI     }
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # set_adjustment( adj1, adj2, adj3, ... )
#NYI #
#NYI # Set the shape adjustments array (as a reference).
#NYI #
#NYI sub set_adjustments {
#NYI 
#NYI     my $self = shift;
#NYI     $self->{_adjustments} = \@_;
#NYI }
#NYI 
#NYI 
#NYI ###############################################################################
#NYI #
#NYI # AUTOLOAD. Deus ex machina.
#NYI #
#NYI # Dynamically create set/get methods that aren't already defined.
#NYI #
#NYI sub AUTOLOAD {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     # Ignore calls to DESTROY.
#NYI     return if $AUTOLOAD =~ /::DESTROY$/;
#NYI 
#NYI     # Check for a valid method names, i.e. "set_xxx_Cy".
#NYI     $AUTOLOAD =~ /.*::(get|set)(\w+)/ or die "Unknown method: $AUTOLOAD\n";
#NYI 
#NYI     # Match the function (get or set) and attribute, i.e. "_xxx_yyy".
#NYI     my $gs        = $1;
#NYI     my $attribute = $2;
#NYI 
#NYI     # Check that the attribute exists.
#NYI     exists $self->{$attribute} or die "Unknown method: $AUTOLOAD\n";
#NYI 
#NYI     # The attribute value
#NYI     my $value;
#NYI 
#NYI     # set_property() pattern.
#NYI     # When a method is AUTOLOADED we store a new anonymous
#NYI     # sub in the appropriate slot in the symbol table. The speeds up subsequent
#NYI     # calls to the same method.
#NYI     #
#NYI     no strict 'refs';    # To allow symbol table hackery
#NYI 
#NYI     $value = $_[0];
#NYI     $value = 1 if not defined $value;    # The default value is always 1
#NYI 
#NYI     if ( $gs eq 'set' ) {
#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self  = shift;
#NYI             my $value = shift;
#NYI 
#NYI             $value = 1 if not defined $value;
#NYI             $self->{$attribute} = $value;
#NYI         };
#NYI 
#NYI         $self->{$attribute} = $value;
#NYI     }
#NYI     else {
#NYI         *{$AUTOLOAD} = sub {
#NYI             my $self = shift;
#NYI             return $self->{$attribute};
#NYI         };
#NYI 
#NYI         # Let AUTOLOAD return the attribute for the first invocation
#NYI         return $self->{$attribute};
#NYI     }
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
#NYI 1;
#NYI 
#NYI __END__
#NYI 
#NYI =head1 NAME
#NYI 
#NYI Shape - A class for creating Excel Drawing shapes
#NYI 
#NYI =head1 SYNOPSIS
#NYI 
#NYI To create a simple Excel file containing shapes using L<Excel::Writer::XLSX>:
#NYI 
#NYI     #!/usr/bin/perl
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'shape.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # Add a default rectangle shape.
#NYI     my $rect = $workbook->add_shape();
#NYI 
#NYI     # Add an ellipse with centered text.
#NYI     my $ellipse = $workbook->add_shape(
#NYI         type => 'ellipse',
#NYI         text => "Hello\nWorld"
#NYI     );
#NYI 
#NYI     # Add a plus shape.
#NYI     my $plus = $workbook->add_shape( type => 'plus');
#NYI 
#NYI     # Insert the shapes in the worksheet.
#NYI     $worksheet->insert_shape( 'B3', $rect );
#NYI     $worksheet->insert_shape( 'C3', $ellipse );
#NYI     $worksheet->insert_shape( 'D3', $plus );
#NYI 
#NYI 
#NYI =head1 DESCRIPTION
#NYI 
#NYI The C<Excel::Writer::XLSX::Shape> module is used to create Shape objects for L<Excel::Writer::XLSX>.
#NYI 
#NYI A Shape object is created via the Workbook C<add_shape()> method:
#NYI 
#NYI     my $shape_rect = $workbook->add_shape( type => 'rect' );
#NYI 
#NYI Once the object is created it can be inserted into a worksheet using the C<insert_shape()> method:
#NYI 
#NYI     $worksheet->insert_shape('A1', $shape_rect);
#NYI 
#NYI A Shape can be inserted multiple times if required.
#NYI 
#NYI     $worksheet->insert_shape('A1', $shape_rect);
#NYI     $worksheet->insert_shape('B2', $shape_rect, 20, 30);
#NYI 
#NYI 
#NYI =head1 METHODS
#NYI 
#NYI =head2 add_shape( %properties )
#NYI 
#NYI The C<add_shape()> Workbook method specifies the properties of the Shape in hash C<< property => value >> format:
#NYI 
#NYI     my $shape = $workbook->add_shape( %properties );
#NYI 
#NYI The available properties are shown below.
#NYI 
#NYI =head2 insert_shape( $row, $col, $shape, $x, $y, $scale_x, $scale_y )
#NYI 
#NYI The C<insert_shape()> Worksheet method sets the location and scale of the shape object within the worksheet.
#NYI 
#NYI     # Insert the shape into the worksheet.
#NYI     $worksheet->insert_shape( 'E2', $shape );
#NYI 
#NYI Using the cell location and the C<$x> and C<$y> cell offsets it is possible to position a shape anywhere on the canvas of a worksheet.
#NYI 
#NYI A more detailed explanation of the C<insert_shape()> method is given in the main L<Excel::Writer::XLSX> documentation.
#NYI 
#NYI 
#NYI =head1 SHAPE PROPERTIES
#NYI 
#NYI Any shape property can be queried or modified by the corresponding get/set method:
#NYI 
#NYI     my $ellipse = $workbook->add_shape( %properties );
#NYI     $ellipse->set_type( 'plus' );    # No longer an ellipse!
#NYI     my $type = $ellipse->get_type();  # Find out what it really is.
#NYI 
#NYI Multiple shape properties may also be modified in one go by using the C<set_properties()> method:
#NYI 
#NYI     $shape->set_properties( type => 'ellipse', text => 'Hello' );
#NYI 
#NYI The properties of a shape object that can be defined via C<add_shape()> are shown below.
#NYI 
#NYI =head2 name
#NYI 
#NYI Defines the name of the shape. This is an optional property and the shape will be given a default name if not supplied. The name is generally only used by Excel Macros to refer to the object.
#NYI 
#NYI =head2 type
#NYI 
#NYI Defines the type of the object such as C<rect>, C<ellipse> or C<triangle>:
#NYI 
#NYI     my $ellipse = $workbook->add_shape( type => 'ellipse' );
#NYI 
#NYI The default type is C<rect>.
#NYI 
#NYI The full list of available shapes is shown below.
#NYI 
#NYI See also the C<shapes_all.pl> program in the C<examples> directory of the distro. It creates an example workbook with all supported shapes labelled with their shape names.
#NYI 
#NYI 
#NYI =over 4
#NYI 
#NYI =item * Basic Shapes
#NYI 
#NYI     blockArc              can            chevron       cube          decagon
#NYI     diamond               dodecagon      donut         ellipse       funnel
#NYI     gear6                 gear9          heart         heptagon      hexagon
#NYI     homePlate             lightningBolt  line          lineInv       moon
#NYI     nonIsoscelesTrapezoid noSmoking      octagon       parallelogram pentagon
#NYI     pie                   pieWedge       plaque        rect          round1Rect
#NYI     round2DiagRect        round2SameRect roundRect     rtTriangle    smileyFace
#NYI     snip1Rect             snip2DiagRect  snip2SameRect snipRoundRect star10
#NYI     star12                star16         star24        star32        star4
#NYI     star5                 star6          star7         star8         sun
#NYI     teardrop              trapezoid      triangle
#NYI 
#NYI =item * Arrow Shapes
#NYI 
#NYI     bentArrow        bentUpArrow       circularArrow     curvedDownArrow
#NYI     curvedLeftArrow  curvedRightArrow  curvedUpArrow     downArrow
#NYI     leftArrow        leftCircularArrow leftRightArrow    leftRightCircularArrow
#NYI     leftRightUpArrow leftUpArrow       notchedRightArrow quadArrow
#NYI     rightArrow       stripedRightArrow swooshArrow       upArrow
#NYI     upDownArrow      uturnArrow
#NYI 
#NYI =item * Connector Shapes
#NYI 
#NYI     bentConnector2   bentConnector3   bentConnector4
#NYI     bentConnector5   curvedConnector2 curvedConnector3
#NYI     curvedConnector4 curvedConnector5 straightConnector1
#NYI 
#NYI =item * Callout Shapes
#NYI 
#NYI     accentBorderCallout1  accentBorderCallout2  accentBorderCallout3
#NYI     accentCallout1        accentCallout2        accentCallout3
#NYI     borderCallout1        borderCallout2        borderCallout3
#NYI     callout1              callout2              callout3
#NYI     cloudCallout          downArrowCallout      leftArrowCallout
#NYI     leftRightArrowCallout quadArrowCallout      rightArrowCallout
#NYI     upArrowCallout        upDownArrowCallout    wedgeEllipseCallout
#NYI     wedgeRectCallout      wedgeRoundRectCallout
#NYI 
#NYI =item * Flow Chart Shapes
#NYI 
#NYI     flowChartAlternateProcess  flowChartCollate        flowChartConnector
#NYI     flowChartDecision          flowChartDelay          flowChartDisplay
#NYI     flowChartDocument          flowChartExtract        flowChartInputOutput
#NYI     flowChartInternalStorage   flowChartMagneticDisk   flowChartMagneticDrum
#NYI     flowChartMagneticTape      flowChartManualInput    flowChartManualOperation
#NYI     flowChartMerge             flowChartMultidocument  flowChartOfflineStorage
#NYI     flowChartOffpageConnector  flowChartOnlineStorage  flowChartOr
#NYI     flowChartPredefinedProcess flowChartPreparation    flowChartProcess
#NYI     flowChartPunchedCard       flowChartPunchedTape    flowChartSort
#NYI     flowChartSummingJunction   flowChartTerminator
#NYI 
#NYI =item * Action Shapes
#NYI 
#NYI     actionButtonBackPrevious actionButtonBeginning actionButtonBlank
#NYI     actionButtonDocument     actionButtonEnd       actionButtonForwardNext
#NYI     actionButtonHelp         actionButtonHome      actionButtonInformation
#NYI     actionButtonMovie        actionButtonReturn    actionButtonSound
#NYI 
#NYI =item * Chart Shapes
#NYI 
#NYI Not to be confused with Excel Charts.
#NYI 
#NYI     chartPlus chartStar chartX
#NYI 
#NYI =item * Math Shapes
#NYI 
#NYI     mathDivide mathEqual mathMinus mathMultiply mathNotEqual mathPlus
#NYI 
#NYI =item * Stars and Banners
#NYI 
#NYI     arc            bevel          bracePair  bracketPair chord
#NYI     cloud          corner         diagStripe doubleWave  ellipseRibbon
#NYI     ellipseRibbon2 foldedCorner   frame      halfFrame   horizontalScroll
#NYI     irregularSeal1 irregularSeal2 leftBrace  leftBracket leftRightRibbon
#NYI     plus           ribbon         ribbon2    rightBrace  rightBracket
#NYI     verticalScroll wave
#NYI 
#NYI =item * Tab Shapes
#NYI 
#NYI     cornerTabs plaqueTabs squareTabs
#NYI 
#NYI =back
#NYI 
#NYI =head2 text
#NYI 
#NYI This property is used to make the shape act like a text box.
#NYI 
#NYI     my $rect = $workbook->add_shape( type => 'rect', text => "Hello\nWorld" );
#NYI 
#NYI The text is super-imposed over the shape. The text can be wrapped using the newline character C<\n>.
#NYI 
#NYI =head2 id
#NYI 
#NYI Identification number for internal identification. This number will be auto-assigned, if not assigned, or if it is a duplicate.
#NYI 
#NYI =head2 format
#NYI 
#NYI Workbook format for decorating the shape text (font family, size, and decoration).
#NYI 
#NYI =head2 start, start_index
#NYI 
#NYI Shape indices of the starting point for a connector and the index of the connection. Index numbers are zero-based, start from the top dead centre and are counted clockwise.
#NYI 
#NYI Indices are typically created for vertices and centre points of shapes. They are the blue connection points that appear when connection shapes are selected manually in Excel.
#NYI 
#NYI =head2 end, end_index
#NYI 
#NYI Same as above but for end points and end connections.
#NYI 
#NYI 
#NYI =head2 start_side, end_side
#NYI 
#NYI This is either the letter C<b> or C<r> for the bottom or right side of the shape to be connected to and from.
#NYI 
#NYI If the C<start>, C<start_index>, and C<start_side> parameters are defined for a connection shape, the shape will be auto located and linked to the starting and ending shapes respectively. This can be very useful for flow and organisation charts.
#NYI 
#NYI =head2 flip_h, flip_v
#NYI 
#NYI Set this value to 1, to flip the shape horizontally and/or vertically.
#NYI 
#NYI =head2 rotation
#NYI 
#NYI Shape rotation, in degrees, from 0 to 360.
#NYI 
#NYI =head2 line, fill
#NYI 
#NYI Shape colour for the outline and fill. Colours may be specified as a colour index, or in RGB format, i.e. C<AA00FF>.
#NYI 
#NYI See C<COLOURS IN EXCEL> in the main documentation for more information.
#NYI 
#NYI =head2 line_type
#NYI 
#NYI Line type for shape outline. The default is solid. The list of possible values is:
#NYI 
#NYI     dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot, solid
#NYI 
#NYI =head2 valign, align
#NYI 
#NYI Text alignment within the shape.
#NYI 
#NYI Vertical alignment can be:
#NYI 
#NYI     Setting     Meaning
#NYI     =======     =======
#NYI     t           Top
#NYI     ctr         Centre
#NYI     b           Bottom
#NYI 
#NYI Horizontal alignment can be:
#NYI 
#NYI     Setting     Meaning
#NYI     =======     =======
#NYI     l           Left
#NYI     r           Right
#NYI     ctr         Centre
#NYI     just        Justified
#NYI 
#NYI The default is to centre both horizontally and vertically.
#NYI 
#NYI =head2 scale_x, scale_y
#NYI 
#NYI Scale factor in x and y dimension, for scaling the shape width and height. The default value is 1.
#NYI 
#NYI Scaling may be set on the shape object or via C<insert_shape()>.
#NYI 
#NYI =head2 adjustments
#NYI 
#NYI Adjustment of shape vertices. Most shapes do not use this. For some shapes, there is a single adjustment to modify the geometry. For instance, the plus shape has one adjustment to control the width of the spokes.
#NYI 
#NYI Connectors can have a number of adjustments to control the shape routing. Typically, a connector will have 3 to 5 handles for routing the shape. The adjustment is in percent of the distance from the starting shape to the ending shape, alternating between the x and y dimension. Adjustments may be negative, to route the shape away from the endpoint.
#NYI 
#NYI =head2 stencil
#NYI 
#NYI Shapes work in stencil mode by default. That is, once a shape is inserted, its connection is separated from its master. The master shape may be modified after an instance is inserted, and only subsequent insertions will show the modifications.
#NYI 
#NYI This is helpful for Org charts, where an employee shape may be created once, and then the text of the shape is modified for each employee.
#NYI 
#NYI The C<insert_shape()> method returns a reference to the inserted shape (the child).
#NYI 
#NYI Stencil mode can be turned off, allowing for shape(s) to be modified after insertion. In this case the C<insert_shape()> method returns a reference to the inserted shape (the master). This is not very useful for inserting multiple shapes, since the x/y coordinates also gets modified.
#NYI 
#NYI =head1 TIPS
#NYI 
#NYI Use C<< $worksheet->hide_gridlines(2) >> to prepare a blank canvas without gridlines.
#NYI 
#NYI Shapes do not need to fit on one page. Excel will split a large drawing into multiple pages if required. Use the page break preview to show page boundaries superimposed on the drawing.
#NYI 
#NYI Connected shapes will auto-locate in Excel if you move either the starting shape or the ending shape separately. However, if you select both shapes (lasso or control-click), the connector will move with it, and the shape adjustments will not re-calculate.
#NYI 
#NYI =head1 EXAMPLE
#NYI 
#NYI     #!/usr/bin/perl
#NYI 
#NYI     use strict;
#NYI     use warnings;
#NYI     use Excel::Writer::XLSX;
#NYI 
#NYI     my $workbook  = Excel::Writer::XLSX->new( 'shape.xlsx' );
#NYI     my $worksheet = $workbook->add_worksheet();
#NYI 
#NYI     # Add a default rectangle shape.
#NYI     my $rect = $workbook->add_shape();
#NYI 
#NYI     # Add an ellipse with centered text.
#NYI     my $ellipse = $workbook->add_shape(
#NYI         type => 'ellipse',
#NYI         text => "Hello\nWorld"
#NYI     );
#NYI 
#NYI     # Add a plus shape.
#NYI     my $plus = $workbook->add_shape( type => 'plus');
#NYI 
#NYI     # Insert the shapes in the worksheet.
#NYI     $worksheet->insert_shape( 'B3', $rect );
#NYI     $worksheet->insert_shape( 'C3', $ellipse );
#NYI     $worksheet->insert_shape( 'D3', $plus );
#NYI 
#NYI 
#NYI See also the C<shapes_*.pl> program in the C<examples> directory of the distro.
#NYI 
#NYI =head1 TODO
#NYI 
#NYI =over 4
#NYI 
#NYI =item * Add shapes which have custom geometries.
#NYI 
#NYI =item * Provide better integration of workbook formats for shapes.
#NYI 
#NYI =item * Add further validation of shape properties to prevent creation of workbooks that will not open.
#NYI 
#NYI =item * Auto connect shapes that are not anchored to cell A1.
#NYI 
#NYI =item * Add automatic shape connection to shape vertices besides the object centre.
#NYI 
#NYI =item * Improve automatic shape connection to shapes with concave sides (e.g. chevron).
#NYI 
#NYI =back
#NYI 
#NYI =head1 AUTHOR
#NYI 
#NYI Dave Clarke dclarke@cpan.org
#NYI 
#NYI =head1 COPYRIGHT
#NYI 
#NYI (c) MM-MMXVII, John McNamara.
#NYI 
#NYI All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
