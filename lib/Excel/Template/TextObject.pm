package Excel::Template::TextObject;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Base);

    use Excel::Template::Base;

UNI_YES     use Unicode::String;
}

# This is a helper object. It is not instantiated by the user,
# nor does it represent an XML object. Rather, certain elements,
# such as <textbox>, can use this object to do text with variable
# substitutions.

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);

    $self->{STACK} = [] unless UNIVERSAL::isa($self->{STACK}, 'ARRAY');

    return $self;
}

sub resolve
{
    my $self = shift;
    my ($context) = @_;

UNI_YES    my $t = Unicode::String::utf8('');
UNI_NO    my $t = '';

    for my $tok (@{$self->{STACK}})
    {
        my $val = $tok;
        $val = $val->resolve($context)
            if Excel::Template::Factory::is_embedded( $val );

UNI_YES        $t .= Unicode::String::utf8("$val");
UNI_NO        $t .= $val;
    }

    return $t;
}

1;
__END__

=head1 NAME

Excel::Template::TextObject

=head1 PURPOSE

=head1 NODE NAME

=head1 INHERITANCE

=head1 ATTRIBUTES

=head1 CHILDREN

=head1 AFFECTS

=head1 DEPENDENCIES

=head1 USAGE

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

=cut