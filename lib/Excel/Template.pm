package Excel::Template;

use strict;

BEGIN {
    use Excel::Template::Base;
    use vars qw ($VERSION @ISA);

    $VERSION  = '0.18';
    @ISA      = qw( Excel::Template::Base );
}

use File::Basename;
use XML::Parser;
use IO::Scalar;

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);

    $self->parse_xml($self->{FILENAME})
        if defined $self->{FILENAME};

    my @renderer_classes = ( 'Spreadsheet::WriteExcel' );
    if (exists $self->{BIG_FILE} && $self->{BIG_FILE})
    {
        unshift @renderer_classes, 'Spreadsheet::WriteExcel::Big';
    }

    $self->{RENDERER} = undef;
    foreach my $class (@renderer_classes)
    {
        (my $filename = $class) =~ s!::!/!g;
        eval {
            require "$filename.pm";
            $class->import;
        };
        if ($@) {
            warn "Could not find or compile '$class'\n" if $^W;
        } else {
            $self->{RENDERER} = $class;
            last;
        }
    }

    defined $self->{RENDERER} ||
        die "Could not find a renderer class. Tried:\n\t" .
            join("\n\t", @renderer_classes) .
            "\n";

    return $self;
}

sub param
{
    my $self = shift;

    # Allow an arbitrary number of hashrefs, so long as they're the first things    # into param(). Put each one onto the end, de-referenced.
    push @_, %{shift @_} while UNIVERSAL::isa($_[0], 'HASH');

    (@_ % 2)
        && die __PACKAGE__, "->param() : Odd number of parameters to param()\n";

    my %params = @_;
    $params{uc $_} = delete $params{$_} for keys %params;
    @{$self->{PARAM_MAP}}{keys %params} = @params{keys %params};

    return !!1;
}

sub write_file
{
    my $self = shift;
    my ($filename) = @_;

    my $xls = $self->{RENDERER}->new($filename)
        || die "Cannot create XLS in '$filename': $!\n";

    $self->_prepare_output($xls);

    $xls->close;

    return !!1;
}

sub output
{
    my $self = shift;

    my $output;
    tie *XLS, 'IO::Scalar', \$output;

    $self->write_file(\*XLS);

    return $output;
}

sub parse_xml
{
    my $self = shift;
    my ($fname) = @_;

    my ($filename, $dirname) = fileparse($fname);
 
    my @stack;
    my $parser = XML::Parser->new(
        Base => $dirname,
        Handlers => {
            Start => sub {
                shift;

                my $name = uc shift;

                my $node = Excel::Template::Factory->create_node($name, @_);
                die "'$name' (@_) didn't make a node!\n" unless defined $node;

                if ( $node->isa( 'WORKBOOK' ) )
                {
                    push @{$self->{WORKBOOKS}}, $node;
                }
                elsif ( $node->is_embedded )
                {
                    return unless @stack;
                                                                                
                    if (exists $stack[-1]{TXTOBJ} &&
                        $stack[-1]{TXTOBJ}->isa('TEXTOBJECT'))
                    {
                        push @{$stack[-1]{TXTOBJ}{STACK}}, $node;
                    }
 
                }
                else
                {
                    push @{$stack[-1]{ELEMENTS}}, $node
                        if @stack;
                }
                push @stack, $node;
            },
            Char => sub {
                shift;
                return unless @stack;

                my $parent = $stack[-1];

                if (
                    exists $parent->{TXTOBJ}
                        &&
                    $parent->{TXTOBJ}->isa('TEXTOBJECT')
                ) {
                    push @{$parent->{TXTOBJ}{STACK}}, @_;
                }
            },
            End => sub {
                shift;
                return unless @stack;

                pop @stack if $stack[-1]->isa(uc $_[0]);
            },
        },
    );

    {
        open( INFILE, "<$fname" )
            || die "Cannot open '$fname' for reading: $!\n";

        $parser->parse(do { local $/ = undef; <INFILE> });

        close INFILE;
    }

    return !!1;
}
*parse = \&parse_xml;

sub _prepare_output
{
    my $self = shift;
    my ($xls) = @_;

    my $context = Excel::Template::Factory->create(
        'CONTEXT',

        XLS       => $xls,
        PARAM_MAP => [ $self->{PARAM_MAP} ],
    );

    $_->render($context) for @{$self->{WORKBOOKS}};

    return !!1;
}

sub register { shift; Excel::Template::Factory::register(@_) }

1;
__END__

=head1 NAME

Excel::Template - Excel::Template

=head1 SYNOPSIS

First, make a template. This is an XML file, describing the layout of the
spreadsheet.

For example, test.xml:

  <workbook>
      <worksheet name="tester">
          <cell text="$HOME"/>
          <cell text="$PATH"/>
      </worksheet>
  </workbook>

Now, create a small program to use it:

  #!/usr/bin/perl -w
  use Excel::Template

  # Create the Excel template
  my $template = Excel::Template->new(
      filename => 'test.xml',
  );

  # Add a few parameters
  $template->param(
      HOME => $ENV{HOME},
      PATH => $ENV{PATH},
  );

  $template->write_file('test.xls');

If everything worked, then you should have a spreadsheet in your work directory
that looks something like:

             A                B                C
    +----------------+----------------+----------------
  1 | /home/me       | /bin:/usr/bin  |
    +----------------+----------------+----------------
  2 |                |                |
    +----------------+----------------+----------------
  3 |                |                |

=head1 DESCRIPTION

This is a module used for templating Excel files. Its genesis came from the
need to use the same datastructure as HTML::Template, but provide Excel files
instead. The existing modules don't do the trick, as they require replication
of logic that's already been done within HTML::Template.

Currently, only a small subset of the planned features are supported. This is
meant to be a test of the waters, to see what features people actually want.

=head1 MOTIVATION

I do a lot of Perl/CGI for reporting purposes. In nearly every place I've been,
I've been asked for HTML, PDF, and Excel. HTML::Template provides the first, and
PDF::Template does the second pretty well. But, generating Excel was the
sticking point. I already had the data structure for the other templating
modules, but I just didn't have an easy mechanism to get that data structure
into an XLS file.

=head1 USAGE

=head2 new()

This creates a Excel::Template object. If passed a filename parameter, it will
parse the template in the given file. (You can also use the parse() method,
described below.)

=head2 param()

This method is exactly like HTML::Template's param() method. Although I will
be adding more to this section later, please see HTML::Template's description
for info right now.

=head2 parse() / parse_xml()

This method actually parses the template file. It can either be called
separately or through the new() call. It will die() if it runs into a situation
it cannot handle.

=head2 write_file()

Create the Excel file and write it to the specified filename, if possible. (This
is when the actual merging of the template and the parameters occurs.)

=head2 output()

It will act just like HTML::Template's output() method, returning the resultant
file as a stream, usually for output to the web. (This is when the actual
merging of the template and the parameters occurs.)

=head1 SUPPORTED NODES

This is just a list of nodes. See the other classes in this distro for more
details on specific parameters and the like.

Every node can set the ROW and COL parameters. These are the actual ROW/COL
values that the next CELL tag will write into.

=over 4

=item * WORKBOOK

=item * WORKSHEET

=item * IF

=item * LOOP

=item * ROW

=item * CELL

=item * FORMULA

=item * BOLD

=item * ITALIC

=back 4

=head1 BUGS

None, that I know of.

=head1 SUPPORT

This is currently beta-quality software. The featureset is extremely limited,
but I expect to be adding on to it very soon.

=head1 AUTHOR

    Rob Kinyon
    rob.kinyon@gmail.com

=head1 CONTRIBUTORS

There is a mailing list at http://groups-beta.google.com/group/ExcelTemplate

Robert Graff -

=over 4

=item * Finishing formats

=item * Fixing several bugs in worksheet naming

=back 4

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

=head1 SEE ALSO

perl(1), HTML::Template, Spreadsheet::WriteExcel.

=cut
