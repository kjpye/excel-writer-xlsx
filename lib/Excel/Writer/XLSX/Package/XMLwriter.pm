unit class Excel::Writer::XLSX::Package::XMLwriter;

###############################################################################
#
# XMLwriter - A base class for the Excel::Writer::XLSX writer classes.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation at end

use v6.c;

has $fh;

#NYI use Exporter;
#NYI use Carp;
#NYI use IO::File;
#NYI 
#NYI our @ISA     = qw(Exporter);
#NYI our $VERSION = '0.96';

#
# NOTE: this module is a light weight re-implementation of XML::Writer. See
# the Pod docs below for a full explanation. The methods  are implemented
# for speed rather than readability since they are used heavily in tight
# loops by Excel::Writer::XLSX.
#

###############################################################################
#
# BUILD()
#
# Constructor.

submethod BUILD(:$!fh) {
  # nothing explicit to do here
}


###############################################################################
#
# set_xml_writer()
#
# Set the XML writer filehandle for the object. This can either be done
# in the constructor (usually for testing since the file name isn't generally
# known at that stage) or later via this method.

method set_xml_writer($filename) {
  $!fh = $filename.IO.open: :w; # UTF8 is the default
}


###############################################################################
#
# xml_declaration()
#
# Write the XML declaration.
#
method xml_declaration {
    $!fh.print: qq[<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n];
}


###############################################################################
#
# xml_start_tag()
#
# Write an XML start tag with optional attributes.
#
method xml_start_tag($tag is copy, *%options) {

  for %options.kv -> $key, $value {
        $value .= escape_attributes;

        $tag ~= qq[ $key="$value"];
    }

    $!fh.print: "<{$tag}>";
}

###############################################################################
#
# xml_start_tag_unencoded()
#
# Write an XML start tag with optional, unencoded, attributes.
# This is a minor speed optimisation for elements that don't need encoding.

method xml_start_tag_unencoded($tag is copy, *%options) {

  for %options.kv -> $key, $value {
    $tag ~= qq[ $key="$value"];
  }

  $!fh.print: "<$tag>";
}

###############################################################################
#
# xml_end_tag()
#
# Write an XML end tag.

method xml_end_tag($tag) {
    $!fh.print: "</$tag>";
}

###############################################################################
#
# xml_empty_tag()
#
# Write an empty XML tag with optional attributes.
#
method xml_empty_tag($tag is copy, *%options) {

  for %options.kv -> $key, $value {
    $value .= escape_attributes;
    $tag ~= qq[ $key="$value"];
  }

  $!fh.print: "<$tag/>";
}

###############################################################################
#
# xml_empty_tag_unencoded()
#
# Write an empty XML tag with optional, unencoded, attributes.
# This is a minor speed optimisation for elements that don't need encoding.
#
method xml_empty_tag_unencoded($tag is copy, *%options) {

  for %options.kv -> $key, $value {
    $tag ~= qq[ $key="$value"];
  }
  $!fh.print: "<$tag/>";
}


###############################################################################
#
# xml_data_element()
#
# Write an XML element containing data with optional attributes.
# XML characters in the data are encoded.
#
method xml_data_element($tag is copy, $data is copy, *%options) {

  my $end_tag = $tag;

  for %options.kv -> $key, $value {
    $value .= escape_attributes;
    $tag ~= qq[ $key="$value"];
  }

  $data .= escape_data;

  $!fh.print: "<$tag>$data</$end_tag>";
}


###############################################################################
#
# xml_data_element_unencoded()
#
# Write an XML unencoded element containing data with optional attributes.
# This is a minor speed optimisation for elements that don't need encoding.
#
method xml_data_element_unencoded($tag is copy, $data, *%options) {

  my $end_tag = $tag;

  for %options.kv -> $key, $value {
    $tag ~= qq[ $key="$value"];
  }
  $!fh.print: "<$tag>$data</$end_tag>";
}


###############################################################################
#
# xml_string_element()
#
# Optimised tag writer for <c> cell string elements in the inner loop.
#
method xml_string_element($index, *%options) {

  my $attr  = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $!fh.print: "<c$attr t=\"s\"><v>{$index}</v></c>";
}


###############################################################################
#
# xml_si_element()
#
# Optimised tag writer for shared strings <si> elements.
#
method xml_si_element($string is copy, *%options) {

  my $attr   = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $string .= escape_data;

  $!fh.print: "<si><t$attr>$string</t></si>";
}


###############################################################################
#
# xml_rich_si_element()
#
# Optimised tag writer for shared strings <si> rich string elements.
#
method xml_rich_si_element($string) {
  $!fh.print: "<si>$string</si>";
}


###############################################################################
#
# xml_number_element()
#
# Optimised tag writer for <c> cell number elements in the inner loop.
#
method xml_number_element($number, *%options) {

  my $attr   = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }
  $!fh.print: "<c{$attr}><v>{$number}</v></c>";
}


###############################################################################
#
# xml_formula_element()
#
# Optimised tag writer for <c> cell formula elements in the inner loop.
#
method xml_formula_element($formula is copy, $result, *%options) {

  my $attr    = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $formula .= escape_data;

  $!fh.print: "<c$attr><f>$formula</f><v>$result</v></c>";
}


###############################################################################
#
# xml_inline_string()
#
# Optimised tag writer for inlineStr cell elements in the inner loop.
#
method xml_inline_string($string is copy, $preserve, *%options) {

  my $attr     = '';
  my $t_attr   = '';

  # Set the <t> attribute to preserve whitespace.
  $t_attr = ' xml:space="preserve"' if $preserve;

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $string .= escape_data;

  $!fh.print: "<c$attr t=\"inlineStr\"><is><t$t_attr>$string</t></is></c>";
}


###############################################################################
#
# xml_rich_inline_string()
#
# Optimised tag writer for rich inlineStr cell elements in the inner loop.
#
method xml_rich_inline_string($string, *%options) {

  my $attr   = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $!fh.print: "<c$attr t=\"inlineStr\"><is>$string</is></c>";
}


#NYI: use accessor for $!fh instead
#NYI ###############################################################################
#NYI #
#NYI # xml_get_fh()
#NYI #
#NYI # Return the output filehandle.
#NYI #
#NYI sub xml_get_fh {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return $self->{_fh};
#NYI }


###############################################################################
#
# _escape_attributes()
#
# Escape XML characters in attributes.
#
method escape_attributes($str is copy) {

  return $str if $str !~~ /<["&<>\n]>/;

  $str ~~ s:g/\&/&amp;/;
  $str ~~ s:g/\"/&quot;/;
  $str ~~ s:g/\</&lt;/;
  $str ~~ s:g/\>/&gt;/;
  $str ~~ s:g/\n/&#xA;/;

  return $str;
}


###############################################################################
#
# escape_data()
#
# Escape XML characters in data sections. Note, this is different from
# escape_attributes() in that double quotes are not escaped by Excel.
#
method escape_data($str is copy) {

  return $str if $str !~~ m/<[&<>]>/;

  $str ~~ s:g/\&/&amp;/;
  $str ~~ s:g/\</&lt;/;
  $str ~~ s:g/\>/&gt;/;

  return $str;
}

=begin pod

=head1 NAME

XMLwriter - A base class for the Excel::Writer::XLSX writer classes.

=head1 DESCRIPTION

This module is used by L<Excel::Writer::XLSX> for writing XML documents. It is a light weight re-implementation of L<XML::Writer>.

XMLwriter is approximately twice as fast as L<XML::Writer>. This speed is achieved at the expense of error and correctness checking. In addition not all of the L<XML::Writer> methods are implemented. As such, XMLwriter is not recommended for use outside of Excel::Writer::XLSX.

=head1 SEE ALSO

L<XML::Writer>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

(c) MM-MMXVII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
=end pod
