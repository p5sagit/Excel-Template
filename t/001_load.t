# -*- perl -*-

# t/001_load.t - check module loading and create testing directory

use Test::More tests => 2;

BEGIN { use_ok( 'Excel::Template' ); }

my $object = Excel::Template->new ();
isa_ok ($object, 'Excel::Template');


