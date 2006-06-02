package Excel::Template::Element::Image;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Element);

    use Excel::Template::Element;
}

1;
__END__

=head1 NAME

Excel::Template::Element::Image - Excel::Template::Element::Image

=head1 PURPOSE

To insert an image into the worksheet

=head1 NODE NAME

CELL

=head1 INHERITANCE

L<ELEMENT|Excel::Template::Element>

=head1 DEPENDENCIES

None

=head1 USAGE

  <image path="/Some/Full/Path"/>

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

Nothing

=cut
