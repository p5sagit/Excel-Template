package Excel::Template::Container::Conditional;

#GGG Convert <conditional> to be a special case of <switch>?

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Container);

    use Excel::Template::Container;
}

my %isOp = (
    '='  => '==',
    (map { $_ => $_ } ( '>', '<', '==', '!=', '>=', '<=' )),
    (map { $_ => $_ } ( 'gt', 'lt', 'eq', 'ne', 'ge', 'le' )),
);

sub conditional_passes
{
    my $self = shift;
    my ($context) = @_;

    my $name = $context->get($self, 'NAME');
    return 0 unless $name =~ /\S/;

    my $val = $context->param($name);
    $val = @{$val} while UNIVERSAL::isa($val, 'ARRAY');
    $val = ${$val} while UNIVERSAL::isa($val, 'SCALAR');

    my $value = $context->get($self, 'VALUE');
    if (defined $value)
    {
        my $op = $context->get($self, 'OP');
        $op = defined $op && exists $isOp{$op}
            ? $isOp{$op}
            : '==';

        # Force numerical context on both values;
        $value *= 1;
        $val *= 1;

        my $res;
        for ($op)
        {
            /^>$/  && do { $res = ($val > $value);  last };
            /^<$/  && do { $res = ($val < $value);  last };
            /^==$/ && do { $res = ($val == $value); last };
            /^!=$/ && do { $res = ($val != $value); last };
            /^>=$/ && do { $res = ($val >= $value); last };
            /^<=$/ && do { $res = ($val <= $value); last };
            /^gt$/ && do { $res = ($val gt $value); last };
            /^lt$/ && do { $res = ($val lt $value); last };
            /^eq$/ && do { $res = ($val eq $value); last };
            /^ne$/ && do { $res = ($val ne $value); last };
            /^ge$/ && do { $res = ($val ge $value); last };
            /^le$/ && do { $res = ($val le $value); last };

            die "Unknown operator in conditional resolve", $/;
        }

        return 0 unless $res;
    }
    elsif (my $is = uc $context->get($self, 'IS'))
    {
        my $istrue = $val && 1;
        if ($is eq 'TRUE')
        {
            return 0 unless $istrue;
        }
        else
        {
            warn "Conditional 'is' value was [$is], defaulting to 'FALSE'" . $/
                if $is ne 'FALSE';

            return 0 if $istrue;
        }
    }

    return 1;
}

sub render
{
    my $self = shift;
    my ($context) = @_;

    return 1 unless $self->conditional_passes($context);

    return $self->iterate_over_children($context);
}

sub max_of
{
    my $self = shift;
    my ($context, $attr) = @_;

    return 0 unless $self->conditional_passes($context);

    return $self->SUPER::max_of($context, $attr);
}

sub total_of
{
    my $self = shift;
    my ($context, $attr) = @_;

    return 0 unless $self->conditional_passes($context);

    return $self->SUPER::total_of($context, $attr);
}

1;
__END__

=head1 NAME

Excel::Template::Container::Conditional - Excel::Template::Container::Conditional

=head1 PURPOSE

To provide conditional execution of children nodes

=head1 NODE NAME

IF

=head1 INHERITANCE

Excel::Template::Container

=head1 ATTRIBUTES

=over 4

=item * NAME

This is the name of the parameter to be testing. It is resolved like any other
parameter. 

=item * VALUE

If VALUE is set, then a comparison operation is done. The value of NAME is
compared to VALUE using the value of OP.

=item * OP

If VALUE is set, then this is checked. If it isn't present, then '==' (numeric
equality) is assumed. OP must be one of the numeric comparison operators or the
string comparison operators. All 6 of each kind is supported.

=item * IS

If VALUE is not set, then IS is checked. IS is allowed to be either "TRUE" or
"FALSE". The boolean value of NAME is checked against IS.

=back 4

=head1 CHILDREN

None

=head1 EFFECTS

None

=head1 DEPENDENCIES

None

=head1 USAGE

  <if name="__ODD__" is="false">
    ... Children here
  </if>

In the above example, the children will be executed if the value of __ODD__
(which is set by the LOOP node) is false. So, for all even iterations.

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

LOOP

=cut