use v6.c+;

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

has $.fh is rw;

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
# set-xml-writer()
#
# Set the XML writer filehandle for the object. This can either be done
# in the constructor (usually for testing since the file name isn't generally
# known at that stage) or later via this method.

method set-xml-writer($filename) {
  $!fh = $filename.IO.open: :w; # UTF8 is the default
}


###############################################################################
#
# xml-declaration()
#
# Write the XML declaration.
#
method xml-declaration() {
    $!fh.print: qq[<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n];
}


###############################################################################
#
# xml-start-tag()
#
# Write an XML start tag with optional attributes.
#
method xml-start-tag($tag is copy, *%options) {

    note "xml-start-tag...";
    dd $!fh;
    dd $tag;
    dd %options;
    for %options.kv -> $key, $value is rw {
        $value = escape-attributes($value);

        $tag ~= qq[ $key="$value"];
    }
    $!fh.print: "<{$tag}>";
note "xml-start-tag printed \"$tag\"";
}

###############################################################################
#
# xml-start-tag-unencoded()
#
# Write an XML start tag with optional, unencoded, attributes.
# This is a minor speed optimisation for elements that don't need encoding.

method xml-start-tag-unencoded($tag is copy, *%options) {

  for %options.kv -> $key, $value {
    $tag ~= qq[ $key="$value"];
  }

  $!fh.print: "<{$tag}>";
}

###############################################################################
#
# xml-end-tag()
#
# Write an XML end tag.

method xml-end-tag($tag) {
    $!fh.print: "</{$tag}>";
}

###############################################################################
#
# xml-empty-tag()
#
# Write an empty XML tag with optional attributes.
#
method xml-empty-tag($tag is copy, *@options) {

  for @options -> $opt {
    my $value = escape-attributes($opt.value);
    my $key   = $opt.key;
    $tag     ~= qq[ $key="$value"];
  }

  $!fh.print: "<{$tag}/>";
}

###############################################################################
#
# xml-empty-tag-unencoded()
#
# Write an empty XML tag with optional, unencoded, attributes.
# This is a minor speed optimisation for elements that don't need encoding.
#
method xml-empty-tag-unencoded($tag is copy, *%options) {

  for %options.kv -> $key, $value {
    $tag ~= qq[ $key="$value"];
  }
  $!fh.print: "<{$tag}/>";
}


###############################################################################
#
# xml-data-element()
#
# Write an XML element containing data with optional attributes.
# XML characters in the data are encoded.
#
method xml-data-element($tag is copy, $data is copy, *%options) {

  my $end-tag = $tag;

  for %options.kv -> $key, $value {
    $value = escape-attributes($value);
    $tag ~= qq[ $key="$value"];
  }

  $data .= escape-data;

  $!fh.print: "<{$tag}>{$data}</{$end-tag}>";
}


###############################################################################
#
# xml-data-element-unencoded()
#
# Write an XML unencoded element containing data with optional attributes.
# This is a minor speed optimisation for elements that don't need encoding.
#
method xml-data-element-unencoded($tag is copy, $data, *%options) {

  my $end-tag = $tag;

  for %options.kv -> $key, $value {
    $tag ~= qq[ $key="$value"];
  }
  $!fh.print: "<$tag>$data</$end-tag>";
}


###############################################################################
#
# xml-string-element()
#
# Optimised tag writer for <c> cell string elements in the inner loop.
#
method xml-string-element($index, *%options) {

  my $attr  = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $!fh.print: "<c$attr t=\"s\"><v>{$index}</v></c>";
}


###############################################################################
#
# xml-si-element()
#
# Optimised tag writer for shared strings <si> elements.
#
method xml-si-element($string is copy, *%options) {

  my $attr   = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $string .= escape-data;

  $!fh.print: "<si><t$attr>$string</t></si>";
}


###############################################################################
#
# xml-rich-si-element()
#
# Optimised tag writer for shared strings <si> rich string elements.
#
method xml-rich-si-element($string) {
  $!fh.print: "<si>$string</si>";
}


###############################################################################
#
# xml-number-element()
#
# Optimised tag writer for <c> cell number elements in the inner loop.
#
method xml-number-element($number, *%options) {

  my $attr   = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }
  $!fh.print: "<c{$attr}><v>{$number}</v></c>";
}


###############################################################################
#
# xml-formula-element()
#
# Optimised tag writer for <c> cell formula elements in the inner loop.
#
method xml-formula-element($formula is copy, $result, *%options) {

  my $attr    = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $formula .= escape-data;

  $!fh.print: "<c$attr><f>$formula</f><v>$result</v></c>";
}


###############################################################################
#
# xml-inline-string()
#
# Optimised tag writer for inlineStr cell elements in the inner loop.
#
method xml-inline-string($string is copy, $preserve, *%options) {

  my $attr     = '';
  my $t-attr   = '';

  # Set the <t> attribute to preserve whitespace.
  $t-attr = ' xml:space="preserve"' if $preserve;

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $string .= escape-data;

  $!fh.print: "<c$attr t=\"inlineStr\"><is><t$t-attr>$string</t></is></c>";
}


###############################################################################
#
# xml-rich-inline-string()
#
# Optimised tag writer for rich inlineStr cell elements in the inner loop.
#
method xml-rich-inline-string($string, *%options) {

  my $attr   = '';

  for %options.kv -> $key, $value {
    $attr ~= qq[ $key="$value"];
  }

  $!fh.print: "<c$attr t=\"inlineStr\"><is>$string</is></c>";
}


#NYI: use accessor for $!fh instead
#NYI ###############################################################################
#NYI #
#NYI # xml-get-fh()
#NYI #
#NYI # Return the output filehandle.
#NYI #
#NYI sub xml-get-fh {
#NYI 
#NYI     my $self = shift;
#NYI 
#NYI     return $self->{_fh};
#NYI }


###############################################################################
#
# escape-attributes()
#
# Escape XML characters in attributes.
#
sub escape-attributes($str is copy) {

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
# escape-data()
#
# Escape XML characters in data sections. Note, this is different from
# escape-attributes() in that double quotes are not escaped by Excel.
#
method escape-data($str is copy) {

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
