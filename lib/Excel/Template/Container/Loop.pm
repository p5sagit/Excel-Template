package Excel::Template::Container::Loop;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Container);

    use Excel::Template::Container;
}

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);

    if (exists $self->{MAXITERS} && $self->{MAXITERS} < 1)
    {
        die "<loop> MAXITERS must be greater than or equal to 1", $/;
    }
    else
    {
        $self->{MAXITERS} = 0;
    }

    return $self;
}

sub make_iterator
{
    my $self = shift;
    my ($context) = @_;

    return Excel::Template::Factory->create('ITERATOR',
        NAME     => $context->get($self, 'NAME'),
        MAXITERS => $context->get($self, 'MAXITERS'),
        CONTEXT  => $context,
    );
}

sub render
{
    my $self = shift;
    my ($context) = @_;

    unless ($self->{ITERATOR} && $self->{ITERATOR}->more_params)
    {
        $self->{ITERATOR} = $self->make_iterator($context);
    }
    my $iterator = $self->{ITERATOR};

    $iterator->enter_scope;

    while ($iterator->can_continue)
    {
        $iterator->next;

        unless ($self->iterate_over_children($context))
        {
            $iterator->back_up;
            last;
        }
    }

    $iterator->exit_scope;

    return 0 if $iterator->more_params;

    return 1;
}

sub total_of
{
    my $self = shift;
    my ($context, $attr) = @_;

    my $iterator = $self->make_iterator($context);

    my $total = 0;

    $iterator->enter_scope;
    while ($iterator->can_continue)
    {
        $iterator->next;
        $total += $self->SUPER::total_of($context, $attr);
    }
    $iterator->exit_scope;

    return $total;
}

sub max_of
{
    my $self = shift;
    my ($context, $attr) = @_;

    my $iterator = $self->make_iterator($context);

    my $max = $context->get($self, $attr);

    $iterator->enter_scope;
    while ($iterator->can_continue)
    {
        $iterator->next;
        my $v = $self->SUPER::max_of($context, $attr);

        $max = $v if $max < $v;
    }
    $iterator->exit_scope;

    return $max;
}

1;
__END__

=head1 NAME

Excel::Template::Container::Loop

=head1 PURPOSE

To provide looping

=head1 NODE NAME

LOOP

=head1 INHERITANCE

Excel::Template::Container

=head1 ATTRIBUTES

=over 4

=item * NAME

This is the name of the loop. It's used to identify within the parameter set
what variables to expose to the children nodes each iteration.

=back 4

=head1 CHILDREN

None

=head1 EFFECTS

None

=head1 DEPENDENCIES

None

=head1 USAGE

  <loop name="LOOPY">
    ... Children here ...
  </loop>

In the above example, the children nodes would have access to the LOOPY array
of hashes as parameters. Each iteration through the array would expose a
different hash of parameters to the children.

These loops work just like HTML::Template's loops. (I promise I'll give more
info here!)

There is one difference - I prefer using Perl-like scoping, so accessing of
variables outside the LOOP scope from within is perfectly acceptable. You can
also hide outside variables with inner values, if you desire, just like Perl.

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

=cut