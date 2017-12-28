unit class Excel::Writer::XLSX::Package::Comments;

###############################################################################
#
# Comments - A class for writing the Excel XLSX Comments files.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2017, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use v6.c;
use Excel::Writer::XLSX::Package::XMLwriter;
use Excel::Writer::XLSX::Utility;


#NYI our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
#NYI our $VERSION = '0.96';


###############################################################################
#
# Public and private API methods.
#
###############################################################################

has %!author_ids;

###############################################################################
#
# new()
#
# Constructor.
#
#NYI sub new {

#NYI     my $class = shift;
#NYI     my $fh    = shift;
#NYI     my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );

#NYI     $self->{_author_ids} = {};

#NYI     bless $self, $class;

#NYI     return $self;
#NYI }


###############################################################################
#
# _assemble_xml_file()
#
# Assemble and write the XML file.
#
method assemble_xml_file($comments-data) {
    self.xml_declaration;

    # Write the comments element.
    self.write_comments();

    # Write the authors element.
    self.write_authors( $comments-data );

    # Write the commentList element.
    self.write_comment_list( $comments-data );

    self.xml_end_tag( 'comments' );

    # Close the XML writer filehandle.
    self.xml_get_fh.close();
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# XML writing methods.
#
###############################################################################


##############################################################################
#
# _write_comments()
#
# Write the <comments> element.
#
method write_comments {
    my $xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    my @attributes = ( 'xmlns' => $xmlns );

    self.xml_start_tag( 'comments', @attributes );
}


##############################################################################
#
# _write_authors()
#
# Write the <authors> element.
#
method write_authors($comment-data) {
    my $author_count = 0;

    self.xml_start_tag( 'authors' );

    for $comment-data -> $comment {
        my $author = $comment[3];

        if $author.defined && ! %!author_ids{$author}.exists {

            # Store the author id.
            %!author_ids{$author} = $author_count++;

            # Write the author element.
            self.write_author( $author );
        }
    }

    self.xml_end_tag( 'authors' );
}


##############################################################################
#
# _write_author()
#
# Write the <author> element.
#
method write_author($data) {
    self.xml_data_element( 'author', $data );
}


##############################################################################
#
# _write_comment_list()
#
# Write the <commentList> element.
#
method write_comment_list($comment-data) {
    self.xml_start_tag( 'commentList' );

    for $comment-data -> $comment {
        my $row    = $comment[0];
        my $col    = $comment[1];
        my $text   = $comment[2];
        my $author = $comment[3];

        # Look up the author id.
        my $author-id;
        $author-id = %!author_ids{$author} if defined $author;

        # Write the comment element.
        self.write_comment( $row, $col, $text, $author-id );
    }

    self.xml_end_tag( 'commentList' );
}


##############################################################################
#
# _write_comment()
#
# Write the <comment> element.
#
method write_comment($row, $col, $text, $author-id) {
    my $ref       = xl-rowcol-to-cell( $row, $col );

    my @attributes = ( 'ref' => $ref );

    @attributes.push: ( 'authorId' => $author-id ) if defined $author-id;


    self.xml_start_tag( 'comment', @attributes );

    # Write the text element.
    self.write_text( $text );


    self.xml_end_tag( 'comment' );
}


##############################################################################
#
# _write_text()
#
# Write the <text> element.
#
method write_text($text) {
    self.xml_start_tag( 'text' );

    # Write the text r element.
    self.write_text_r( $text );

    self.xml_end_tag( 'text' );
}


##############################################################################
#
# _write_text_r()
#
# Write the <r> element.
#
method write_text_r($text) {
    self.xml_start_tag( 'r' );

    # Write the rPr element.
    self.write_r_pr();

    # Write the text r element.
    self.write_text_t( $text );

    self.xml_end_tag( 'r' );
}


##############################################################################
#
# _write_text_t()
#
# Write the text <t> element.
#
method write_text_t($text) {
    my @attributes = ();

    if $text ~~ /^\s/ || $text ~~ /\s$/ {
        push @attributes, ( 'xml:space' => 'preserve' );
    }

    self.xml_data_element( 't', $text, @attributes );
}


##############################################################################
#
# _write_r_pr()
#
# Write the <rPr> element.
#
method write_r_pr {
    self.xml_start_tag( 'rPr' );

    # Write the sz element.
    self.write_sz();

    # Write the color element.
    self.write_color();

    # Write the rFont element.
    self.write_r_font();

    # Write the family element.
    self.write_family();

    self.xml_end_tag( 'rPr' );
}


##############################################################################
#
# _write_sz()
#
# Write the <sz> element.
#
method write_sz {
    my $val  = 8;

    my @attributes = ( 'val' => $val );

    self.xml_empty_tag( 'sz', @attributes );
}


##############################################################################
#
# _write_color()
#
# Write the <color> element.
#
method write_color {
    my $indexed = 81;

    my @attributes = ( 'indexed' => $indexed );

    self.xml_empty_tag( 'color', @attributes );
}


##############################################################################
#
# _write_r_font()
#
# Write the <rFont> element.
#
method write_r_font {
    my $val  = 'Tahoma';

    my @attributes = ( 'val' => $val );

    self.xml_empty_tag( 'rFont', @attributes );
}


##############################################################################
#
# _write_family()
#
# Write the <family> element.
#
method write_family {
    my $val  = 2;

    my @attributes = ( 'val' => $val );

    self.xml_empty_tag( 'family', @attributes );
}


=begin pod


__END__

=pod

=head1 NAME

Comments - A class for writing the Excel XLSX Comments files.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

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
