use strict;
use warnings;
$|++;

use Test::More tests => 5;

use lib 't';
use mock;
mock->reset;

my $CLASS = 'Excel::Template';
use_ok( $CLASS );

my $object = $CLASS->new(
    filename => 't/011.xml',
);
isa_ok( $object, $CLASS );

ok(
    $object->param( 
        loopy => [
            { int => 0, char => 'n' },
            { int => 0, char => 'y' },
            { int => 1, char => 'n' },
            { int => 1, char => 'y' },
        ],
    ),
    'Parameters set',
);

ok( $object->write_file( 'filename' ), 'Something returned' );

my @calls = mock->get_calls;
is( join( $/, @calls, '' ), <<__END_EXPECTED__, 'Calls match up' );
Spreadsheet::WriteExcel::new( 'filename' )
Spreadsheet::WriteExcel::add_format( '' )
Spreadsheet::WriteExcel::add_worksheet( 'conditional' )
Spreadsheet::WriteExcel::Worksheet::new( '' )
Spreadsheet::WriteExcel::Worksheet::write( '0', '0', 'not bool', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '0', '1', 'int', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '1', '0', 'not bool', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '1', '1', 'int', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '1', '2', 'char', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '2', '0', 'bool', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '3', '0', 'bool', '1' )
Spreadsheet::WriteExcel::Worksheet::write( '3', '1', 'char', '1' )
Spreadsheet::WriteExcel::close( '' )
__END_EXPECTED__
