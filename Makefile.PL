use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.

use strict;

my $prereqs = {
    'Spreadsheet::WriteExcel' => 0.42,
    'Test::More'              => 0.01,
    'XML::Parser'             => 0.01,
    'IO::Scalar'              => 0.01,
    'File::Basename'          => 0.01,
};

# The assumption is Perl 5.8.0 and greater doesn't need Unicode::String.

if ($] < 5.008)
{
    eval { require Unicode::String; };
    if ($@)
    {
        print "Note: If you want to work with Unicode, you will need to install";
        print "the optional module Unicode::String and set USE_UNICODE to true.";
    }
}

WriteMakefile(
    NAME         => 'Excel::Template',
    VERSION_FROM => 'lib/Excel/Template.pm', # finds $VERSION
    AUTHOR       => 'Rob Kinyon (rob.kinyon@gmail.com)',
    ABSTRACT     => 'Excel::Template',
    PREREQ_PM    => $prereqs,
);
