package Spreadsheet::WriteExcel;

use strict;

use mock;
use Spreadsheet::WriteExcel::Worksheet;

sub new {
    my $self = bless {
    }, shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::new( '@_' )";
    }

    return $self;
}

sub close {
    shift;
    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::close( '@_' )";
    }
}

sub add_worksheet {
    shift;
    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::add_worksheet( '@_' )";
    }
    return Spreadsheet::WriteExcel::Worksheet->new;
}

my $format_num = 1;
sub add_format {
    shift;
    my %x = @_;
    my @x = map { $_ => $x{$_} } sort keys %x;
    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::add_format( '@x' )";
    }
    return $format_num++;
}

1;
__END__
