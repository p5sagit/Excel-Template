use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.

use strict;

my $prereqs = {
    'Test::Simple'            => 0.44,
    'Spreadsheet::WriteExcel' => 0.42,
    'XML::Parser'             => 0.01,
    'IO::Scalar'              => 0.01,
    'File::Basename'          => 0.01,
};

# The assumption is Perl 5.8.0 and greater doesn't need Unicode::String.

my $use_unicode = 0;
if ($] < 5.008)
{
    print "Do you want Unicode support? ";
    my $answer = <STDIN>;
    my $need_unicode = $answer =~ /^[Yy]/;

    if ($need_unicode)
    {
            $prereqs->{'Unicode::String'} = '0.01';
            $use_unicode = 1;
    }
}

my $pm_filter = $use_unicode
    ? 'perl -p -e "s!UNI_YES ! !g;s!UNI_NO  !\\#!g"'
    : 'perl -p -e "s!UNI_NO  ! !g;s!UNI_YES !\\#!g"';

WriteMakefile(
    NAME         => 'Excel::Template',
    VERSION_FROM => 'lib/Excel/Template.pm', # finds $VERSION
    AUTHOR       => 'Rob Kinyon (rob.kinyon@gmail.com)',
    ABSTRACT     => 'Excel::Template',
    PREREQ_PM    => $prereqs,
    PM_FILTER    => $pm_filter,
);
