package Excel::Template::Element::Image;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Element);

    use Excel::Template::Element;
}

sub render {
    my $self = shift;
    my ($context) = @_;

    my $path = $context->get( $self, 'PATH' );
    my ($row, $col, $offset, $scale) = map {
        $context->get($self, $_)
    } qw( ROW COL OFFSET SCALE );

    my @offsets = (0,0);
    if ( $offset =~ /^\s*([\d.]+)\s*,\s*([\d.]+)/ ) {
        @offsets = ($1,$2);
    }

    my @scales = (0,0);
    if ( $scale =~ /^\s*([\d.]+)\s*,\s*([\d.]+)/ ) {
        @scales = ($1,$2);
    }

    $context->active_worksheet->insert_image(
        $row, $col, $path, @offsets, @scales,
    );

    return 1;
}

sub deltas {
    return {
        COL => +1,
    };
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

  <image path="/Some/Full/Path" />
  <image path="/Some/Full/Path" offset="2,5" />
  <image path="/Some/Full/Path" scale="2,0.4" />
  <image path="/Some/Full/Path" offset="4,0" scale="0,2" />

Please see L<Spreadsheet::WriteExcel/> for more information about the offset and scaling options as well as any other restrictions that might be in place. This node does B<NOT> perform any sort of validation upon your parameters. You are assumed to know what you are doing.

=head1 AUTHOR

Rob Kinyon (rob.kinyon@gmail.com)

=head1 SEE ALSO

Nothing

=cut
