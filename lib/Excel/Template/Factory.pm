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

    'CELL'      => 'Excel::Template::Element::Cell',
    'FORMULA'   => 'Excel::Template::Element::Formula',
    'VAR'       => 'Excel::Template::Element::Var',

    'FORMAT'    => 'Excel::Template::Container::Format',

# These are all the Format short-cut objects
    'BOLD'      => 'Excel::Template::Container::Bold',
    'HIDDEN'    => 'Excel::Template::Container::Hidden',
    'ITALIC'    => 'Excel::Template::Container::Italic',
    'LOCKED'    => 'Excel::Template::Container::Locked',
    'OUTLINE'   => 'Excel::Template::Container::Outline',
    'SHADOW'    => 'Excel::Template::Container::Shadow',
    'STRIKEOUT' => 'Excel::Template::Container::Strikeout',

# These are the helper objects

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
    ITALIC
    OUTLINE
    LOOP
    ROW
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
            warn "$_ was not supplied to register()\n";
            return 0;
        }
    }

    my $name = uc $params{name};
    if (exists $Manifest{$name})
    {
        warn "$params{name} already exists in the manifest.\n";
        return 0;
    }

    my $isa = uc $params{isa};
    unless (exists $Manifest{$isa})
    {
        warn "$params{isa} does not exist in the manifest.\n";
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

sub create
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

sub create_node
{
    my $class = shift;
    my $name = uc shift;

    return unless exists $isBuildable{$name};

    return $class->create($name, @_);
}

sub isa
{
    return unless @_ >= 2;
    exists $Manifest{uc $_[1]}
        ? UNIVERSAL::isa($_[0], $Manifest{uc $_[1]})
        : UNIVERSAL::isa(@_)
}

1;
__END__

=head1 NAME

Excel::Template::Factory

=head1 PURPOSE

=head1 NODE NAME

=head1 INHERITANCE

=head1 ATTRIBUTES

=head1 CHILDREN

=head1 AFFECTS

=head1 DEPENDENCIES

=head1 USAGE

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

=cut
