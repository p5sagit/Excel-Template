package Excel::Template::Container::Worksheet;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Container);

    use Excel::Template::Container;
}

sub render
{
    my $self = shift;
    my ($context) = @_;

    $context->new_worksheet($self->{NAME});

    return $self->SUPER::render($context);
}

1;
__END__

=head1 NAME

Excel::Template::Container::Worksheet - Excel::Template::Container::Worksheet

=head1 PURPOSE

To provide a new worksheet.

=head1 NODE NAME

WORKSHEET

=head1 INHERITANCE

Excel::Template::Container

=head1 ATTRIBUTES

=over 4

=item * NAME

This is the name of the worksheet to be added.

=back 4

=head1 CHILDREN

None

=head1 EFFECTS

None

=head1 DEPENDENCIES

None

=head1 USAGE

  <worksheet name="My Taxes">
    ... Children here
  </worksheet>

In the above example, the children will be executed in the context of the
"My Taxes" worksheet.

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

ROW, CELL, FORMULA

=cut
