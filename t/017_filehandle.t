use strict;

use Test::More tests => 4;

use lib 't';
use mock;
mock::reset;

my $CLASS = 'Excel::Template';
use_ok( $CLASS );

my $object = $CLASS->new(
    file => \*DATA,
);
isa_ok( $object, $CLASS );

ok( $object->write_file( 'filename' ), 'Something returned' );

my @calls = mock::get_calls;
is( join( $/, @calls, '' ), <<__END_EXPECTED__, 'Calls match up' );
Spreadsheet::WriteExcel::new( 'filename' )
Spreadsheet::WriteExcel::add_format( '' )
Spreadsheet::WriteExcel::add_worksheet( '' )
Spreadsheet::WriteExcel::Worksheet::new( '' )
Spreadsheet::WriteExcel::add_worksheet( 'foo' )
Spreadsheet::WriteExcel::Worksheet::new( '' )
Spreadsheet::WriteExcel::add_worksheet( '' )
Spreadsheet::WriteExcel::Worksheet::new( '' )
Spreadsheet::WriteExcel::close( '' )
__END_EXPECTED__

__DATA__
<workbook>
  <worksheet />
  <worksheet name="foo"/>
  <worksheet name="foo"/>
</workbook>
