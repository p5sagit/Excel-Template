use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME         => 'Excel::Template',
    VERSION_FROM => 'lib/Excel/Template.pm', # finds $VERSION
    AUTHOR       => 'Rob Kinyon (rkinyon@columbus.rr.com',
    ABSTRACT     => 'Excel::Template',
    PREREQ_PM    => {
                     'Test::Simple' => 0.44,
                     'Spreadsheet::WriteExcel' => 0.42,
                     'XML::Parser' => 0.01,
                     'IO::Scalar' => 0.01,
                     'IO::File' => 0.01,
                    },
);
