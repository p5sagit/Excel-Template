use Module::Build;

use strict;

my $requires = {
    'perl'              => '5.6.0',
    'Test::More'        => 0.01,
    'XML::Parser'       => 0.01,
    'IO::Scalar'        => 0.01,
    'File::Basename'    => 0.01,
    'Spreadsheet::WriteExcel' => 0.42,
    'Spreadsheet::WriteExcel::Utility' => 0.01,
};

if ($] < 5.008) {
    eval { require Unicode::String; };
    if ($@) {
       print "Note: If you want to work with Unicode, you will need to install";
       print "the optional module Unicode::String and set USE_UNICODE to true.";
       my $resp = Module::Build->prompt( "Do you want to install it now?", "y");
       if ( $resp =~ /y/i ) {
           $requires->{'Unicode::String'} => '0.01';
       }
    } else {
        print "You have Unicode::String installed. If you warn to use Unicode,";
        print "You will need to set USE_UNICODE to true."
    }
}

my $build = Module::Build->new(
    module_name => 'Excel::Template',
    license => 'perl',
    requires => $requires,
    optional => {
    },
    build_requires => {
        'File::Path'      => '0.01',
        'File::Temp'      => '0.01',
        'Test::Deep'      => '0.095',
        'Test::More'      => '0.47',
        'Test::Exception' => '0.21',
    },
    create_makefile_pl => 'traditional',
    add_to_cleanup => [
        'META.yml', '*.bak', '*.gz', 'Makefile.PL', 'cover_db',
    ],
#    test_files => 't/??_*.t',
);

$build->create_build_script;
