package Excel::Template::Element::Formula;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Element::Cell);

    use Excel::Template::Element::Cell;
}

sub get_text
{
    my $self = shift;
    my ($context) = @_;

    my $text = $self->SUPER::get_text($context);

# At this point, we must do back-reference dereferencing

    return $text;
}

sub render
{
    my $self = shift;
    my ($context) = @_;

    $context->active_worksheet->write_formula(
        (map { $context->get($self, $_) } qw(ROW COL)),
        $self->get_text($context),
    );

    return 1;
}

1;
__END__

=head1 NAME

Excel::Template::Element::Formula - Excel::Template::Element::Formula

=head1 PURPOSE

To write formulas to the worksheet

=head1 NODE NAME

FORMULA

=head1 INHERITANCE

Excel::Template::Element::Cell

=head1 ATTRIBUTES

=over 4

=item * TEXT

This is the formula to write to the cell. This can either be text or a parameter
with a dollar-sign in front of the parameter name.

=item * COL

Optionally, you can specify which column you want this cell to be in. It can be
either a number (zero-based) or an offset. See Excel::Template for more info on
offset-based numbering.

=back 4

There will be more parameters added, as features are added.

=head1 CHILDREN

None

=head1 EFFECTS

This will consume one column on the current row. 

=head1 DEPENDENCIES

None

=head1 USAGE

  <formula text="=(1 + 2)"/>
  <formula>=SUM(A1:A5)</formula>

  <formula text="$Param2"/>
  <formula>=(A1 + <var name="Param">)</formula>

In the above example, four formulas are written out. The first two have the
formula hard-coded. The second two have variables. The third and fourth items
have another thing that should be noted. If you have a formula where you want a
variable in the middle, you have to use the latter form. Variables within
parameters are the entire parameter's value.

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

ROW, VAR, CELL

=cut
