package Spreadsheet::WriteExcel::Worksheet;

use strict;

use mock;

sub new {
    my $self = bless {
    }, shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::new( '@_' )";
    }

    return $self;
}

sub write_formula {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::write_formula( '@_' )";
    }
}

sub write {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::write( '@_' )";
    }
}

sub set_row {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::set_row( '@_' )";
    }
}

sub set_column {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::set_column( '@_' )";
    }
}

1;
__END__
