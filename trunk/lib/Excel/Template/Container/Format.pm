package Excel::Template::Container::Format;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw( Excel::Template::Container );

    use Excel::Template::Container;
}

use Excel::Template::Format;

sub render
{
    my $self = shift;
    my ($context) = @_;

    my $old_format = $context->active_format;
    my $format = Excel::Template::Format->copy(
        $context, $old_format,

        %{$self},
    );
    $context->active_format($format);

    my $child_success = $self->iterate_over_children($context);

    $context->active_format($old_format);
}

1;
__END__

=head1 NAME

Excel::Template::Container::Format - Excel::Template::Container::Format

=head1 PURPOSE

To format all children according to the parameters

=head1 NODE NAME

FORMAT

=head1 INHERITANCE

Excel::Template::Container

=head1 ATTRIBUTES

=over 4

=item * bold

This will set bold to on or off, depending on the boolean value.

=item * hidden

This will set whether the cell is hidden to on or off, depending on the boolean
value. (q.v. BOLD tag)

=item * italic

This will set italic to on or off, depending on the boolean value. (q.v. ITALIC
tag)

=item * font_outline

This will set font_outline to on or off, depending on the boolean value. (q.v.
OUTLINE tag)

=item * font_shadow

This will set font_shadow to on or off, depending on the boolean value. (q.v.
SHADOW tag)

=item * font_strikeout

This will set font_strikeout to on or off, depending on the boolean value. (q.v.
STRIKEOUT tag)

=back 4

=head1 CHILDREN

None

=head1 EFFECTS

None

=head1 DEPENDENCIES

None

=head1 USAGE

  <format bold="1">
    ... Children here
  </format>

In the above example, the children will be displayed (if they are displaying
elements) in a bold format. All other formatting will remain the same and the
"bold"-ness will end at the end tag.

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

BOLD, HIDDEN, ITALIC, OUTLINE, SHADOW, STRIKEOUT

=cut
