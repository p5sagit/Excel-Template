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

sub write_string {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::write_string( '@_' )";
    }
}

sub write_number {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::write_number( '@_' )";
    }
}

sub write_blank {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::write_blank( '@_' )";
    }
}

sub write_url {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::write_url( '@_' )";
    }
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

sub keep_leading_zeros {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::keep_leading_zeros( '@_' )";
    }
}

sub insert_bitmap {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::insert_bitmap( '@_' )";
    }
}

sub freeze_panes {
    my $self = shift;

    {
        local $" = "', '";
        push @mock::calls, __PACKAGE__ . "::freeze_panes( '@_' )";
    }
}

1;
__END__
