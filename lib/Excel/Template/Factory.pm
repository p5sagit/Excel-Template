package Excel::Template::Factory;

use strict;

BEGIN {
    use vars qw(%Manifest %isBuildable);
}

%Manifest = (

# These are the instantiable nodes
    'IF'        => 'Excel::Template::Container::Conditional',
    'LOOP'      => 'Excel::Template::Container::Loop',
    'ROW'       => 'Excel::Template::Container::Row',
    'SCOPE'     => 'Excel::Template::Container::Scope',
    'WORKBOOK'  => 'Excel::Template::Container::Workbook',
    'WORKSHEET' => 'Excel::Template::Container::Worksheet',

    'BACKREF'   => 'Excel::Template::Element::Backref',
    'CELL'      => 'Excel::Template::Element::Cell',
    'FORMULA'   => 'Excel::Template::Element::Formula',
    'RANGE'     => 'Excel::Template::Element::Range',
    'VAR'       => 'Excel::Template::Element::Var',

    'FORMAT'    => 'Excel::Template::Container::Format',

# These are all the Format short-cut objects
# They are also instantiable
    'BOLD'      => 'Excel::Template::Container::Bold',
    'HIDDEN'    => 'Excel::Template::Container::Hidden',
    'ITALIC'    => 'Excel::Template::Container::Italic',
    'LOCKED'    => 'Excel::Template::Container::Locked',
    'OUTLINE'   => 'Excel::Template::Container::Outline',
    'SHADOW'    => 'Excel::Template::Container::Shadow',
    'STRIKEOUT' => 'Excel::Template::Container::Strikeout',

# These are the helper objects
# They are also in here to make E::T::Factory::isa() work.
    'CONTEXT'    => 'Excel::Template::Context',
    'ITERATOR'   => 'Excel::Template::Iterator',
    'TEXTOBJECT' => 'Excel::Template::TextObject',

    'CONTAINER'  => 'Excel::Template::Container',
    'ELEMENT'    => 'Excel::Template::Element',

    'BASE'       => 'Excel::Template::Base',
);

%isBuildable = map { $_ => 1 } qw(
    BOLD
    CELL
    FORMAT
    FORMULA
    IF
    HIDDEN
    ITALIC
    LOCKED
    OUTLINE
    LOOP
    BACKREF
    RANGE
    ROW
    SCOPE
    SHADOW
    STRIKEOUT
    VAR
    WORKBOOK
    WORKSHEET
);

sub register
{
    my %params = @_;

    my @param_names = qw(name class isa);
    for (@param_names)
    {
        unless ($params{$_})
        {
            warn "$_ was not supplied to register()\n" if $^W;
            return 0;
        }
    }

    my $name = uc $params{name};
    if (exists $Manifest{$name})
    {
        warn "$params{name} already exists in the manifest.\n" if $^W;
        return 0;
    }

    my $isa = uc $params{isa};
    unless (exists $Manifest{$isa})
    {
        warn "$params{isa} does not exist in the manifest.\n" if $^W;
        return 0;
    }

    $Manifest{$name} = $params{class};
    $isBuildable{$name} = 1;

    {
        no strict 'refs';
        unshift @{"$params{class}::ISA"}, $Manifest{$isa};
    }

    return 1;
}

sub _create
{
    my $class = shift;
    my $name = uc shift;

    return unless exists $Manifest{$name};

    (my $filename = $Manifest{$name}) =~ s!::!/!g;
 
    eval {
        require "$filename.pm";
    }; if ($@) {
        die "Cannot find or compile PM file for '$name' ($filename)\n";
    }
 
    return $Manifest{$name}->new(@_);
}

sub _create_node
{
    my $class = shift;
    my $name = uc shift;

    return unless exists $isBuildable{$name};

    return $class->_create($name, @_);
}

sub isa
{
    return unless @_ >= 2;
    exists $Manifest{uc $_[1]}
        ? UNIVERSAL::isa($_[0], $Manifest{uc $_[1]})
        : UNIVERSAL::isa(@_)
}

sub is_embedded
{
    return unless @_ >= 1;

    isa( $_[0], $_ ) && return ~~1 for qw( VAR BACKREF RANGE );
    return;
}

1;
__END__

=head1 NAME

Excel::Template::Factory

=head1 PURPOSE

To provide a common way to instantiate Excel::Template nodes

=head1 USAGE

=head2 register()

Use this to register your own nodes.

Example forthcoming.

=head1 METHODS

=head2 isa

This is a customized isa() wrapper for syntactic sugar

=head2 is_embedded

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

=cut
