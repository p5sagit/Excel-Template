use strict;

use Test::More;

use lib 't';

eval "use Test::Distribution not => 'versions'";
plan skip_all => "Test::Distribution required for testing DISTRIBUTION coverage" if $@;
