package Excel::Template::Element::Cell;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Element);

    use Excel::Template::Element;
}

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);
                                                                                
    $self->{TXTOBJ} = Excel::Template::Factory->create('TEXTOBJECT');
                                                                                
    return $self;
}

sub get_text
{
    my $self = shift;
    my ($context) = @_;
                                                                                
    my $txt = $context->get($self, 'TEXT');
    if (defined $txt)
    {
        my $txt_obj = Excel::Template::Factory->create('TEXTOBJECT');
        push @{$txt_obj->{STACK}}, $txt;
        $txt = $txt_obj->resolve($context);
    }
    elsif ($self->{TXTOBJ})
    {
        $txt = $self->{TXTOBJ}->resolve($context)
    }
    else
    {
UNI_YES        $txt = Unicode::String::utf8('');
UNI_NO         $txt = '';
    }
                                                                                
    return $txt;
}

sub render
{
    my $self = shift;
    my ($context, $method) = @_;

    $method ||= 'write';

    my ($row, $col) = map { $context->get($self, $_) } qw(ROW COL);

    my $ref = uc $context->get( $self, 'REF' );
    if (defined $ref && length $ref)
    {
        $context->add_reference( $ref, $row, $col );
    }

    $context->active_worksheet->$method(
        $row, $col,
        $self->get_text($context),
        $context->active_format,
    );

    return 1;
}

sub deltas
{
    return {
        COL => +1,
    };
}

1;
__END__

=head1 NAME

Excel::Template::Element::Cell - Excel::Template::Element::Cell

=head1 PURPOSE

To actually write stuff to the worksheet

=head1 NODE NAME

CELL

=head1 INHERITANCE

Excel::Template::Element

=head1 ATTRIBUTES

=over 4

=item * TEXT

This is the text to write to the cell. This can either be text or a parameter
with a dollar-sign in front of the parameter name.

=item * COL

Optionally, you can specify which column you want this cell to be in. It can be
either a number (zero-based) or an offset. See Excel::Template for more info on
offset-based numbering.

=item * REF

Adds the current cell to the a list of cells that can be backreferenced.
This is useful when the current cell needs to be referenced by a
formula. See BACKREF and RANGE.

=back 4

There will be more parameters added, as features are added.

=head1 CHILDREN

Excel::Template::Element::Formula

=head1 EFFECTS

This will consume one column on the current row. 

=head1 DEPENDENCIES

None

=head1 USAGE

  <cell text="Some Text Here"/>
  <cell>Some other text here</cell>

  <cell text="$Param2"/>
  <cell>Some <var name="Param"> text here</cell>

In the above example, four cells are written out. The first two have text hard-
coded. The second two have variables. The third and fourth items have another
thing that should be noted. If you have text where you want a variable in the
middle, you have to use the latter form. Variables within parameters are the
entire parameter's value.

Please see Spreadsheet::WriteExcel for what constitutes a legal formula.

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

ROW, VAR, FORMULA

=cut
