package Excel::Template;

use strict;

BEGIN {
    use Excel::Template::Base;
    use vars qw ($VERSION @ISA);

    $VERSION  = '0.22';
    @ISA      = qw( Excel::Template::Base );
}

use File::Basename;
use XML::Parser;
use IO::Scalar;

use constant RENDER_NML => 'normal';
use constant RENDER_BIG => 'big';
use constant RENDER_XML => 'xml';

my %renderers = (
    RENDER_NML, 'Spreadsheet::WriteExcel',
    RENDER_BIG, 'Spreadsheet::WriteExcel::Big',
    RENDER_XML, 'Spreadsheet::WriteExcelXML',
);

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);

    $self->{FILE} = $self->{FILENAME}
        if !defined $self->{FILE} && defined $self->{FILENAME};

    $self->parse_xml($self->{FILE})
        if defined $self->{FILE};

    my @renderer_classes = ( 'Spreadsheet::WriteExcel' );

    if (exists $self->{RENDERER} && $self->{RENDERER})
    {
        if (exists $renderers{ lc $self->{RENDERER} })
        {
            unshift @renderer_classes, $renderers{ lc $self->{RENDERER} };
        }
        elsif ($^W)
        {
            warn "'$self->{RENDERER}' is not recognized\n";
        }
    }
    elsif (exists $self->{BIG_FILE} && $self->{BIG_FILE})
    {
        warn "Use of BIG_FILE is deprecated.\n";
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

    $self->{USE_UNICODE} = ~~0
        if $] >= 5.008;

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

    return ~~1;
}

sub write_file
{
    my $self = shift;
    my ($filename) = @_;

    my $xls = $self->{RENDERER}->new($filename)
        || die "Cannot create XLS in '$filename': $!\n";

    $self->_prepare_output($xls);

    $xls->close;

    return ~~1;
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
    my ($file) = @_;

    my @stack;
    my @parms = (
        Handlers => {
            Start => sub {
                shift;

                my $name = uc shift;

                my $node = Excel::Template::Factory->_create_node($name, @_);
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

    if ( ref $file )
    {
        *INFILE = $file;
    }
    else
    {
        my ($filename, $dirname) = fileparse($file);
 
        push @parms, Base => $dirname;

        open( INFILE, "<$file" )
            || die "Cannot open '$file' for reading: $!\n";

    }

    my $parser = XML::Parser->new( @parms );
    $parser->parse(do { local $/ = undef; <INFILE> });

    unless ( ref $file )
    {
        close INFILE;
    }

    return ~~1;
}
*parse = \&parse_xml;

sub _prepare_output
{
    my $self = shift;
    my ($xls) = @_;

    my $context = Excel::Template::Factory->_create(
        'CONTEXT',

        XLS       => $xls,
        PARAM_MAP => [ $self->{PARAM_MAP} ],
        UNICODE   => $self->{UNICODE},
    );

#    print "@{$self->{WORKBOOKS}}\n";
    $_->render($context) for @{$self->{WORKBOOKS}};

    return ~~1;
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
  use Excel::Template;

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

=head1 MOTIVATION

I do a lot of Perl/CGI for reporting purposes. In nearly every place I've been,
I've been asked for HTML, PDF, and Excel. HTML::Template provides the first, and
PDF::Template does the second pretty well. But, generating Excel was the
sticking point. I already had the data structure for the other templating
modules, but I just didn't have an easy mechanism to get that data structure
into an XLS file.

=head1 USAGE

=head2 new()

This creates a Excel::Template object.

=head3 Parameters

=over 4

=item * FILE / FILENAME

Excel::Template will parse the template in the given file or filehandle automatically. (You can also use the parse() method, described below.)

If you want to use the __DATA__ section, you can do so by passing

  FILE => \*DATA

=item * RENDERER

The default rendering engine is Spreadsheet::WriteExcel. You may, if you choose, change that to another choice. The legal values are:

=over 4

=item * Excel::Template->RENDER_NML

This is the default of Spreadsheet::WriteExcel.

=item * Excel::Template->RENDER_BIG

This attempts to load Spreadsheet::WriteExcel::Big.

=item * Excel::Template->RENDER_XML

This attempts to load Spreadsheet::WriteExcelXML.

=back

=item * USE_UNICODE

This will use L<Unicode::String> to represent strings instead of Perl's internal string handling. You must already have L<Unicode::String> installed on your system.

The USE_UNICODE parameter will be ignored if you are using Perl 5.8 or higher as
Perl's internal string handling is unicode-aware.

NOTE: Certain older versions of L<OLE::Storage_Lite> and mod_perl clash for some
reason. Upgrading to the latest version of L<OLE::Storage_Lite> should fix the
problem.

=back

=head3 Deprecated

=over 4

=item * BIG_FILE

Instead, use RENDERER => Excel::Template->RENDER_BIG

=back

=head2 param()

This method is exactly like L<HTML::Template>'s param() method.

=head2 parse() / parse_xml()

This method actually parses the template file. It can either be called separately or through the new() call. It will die() if it runs into a situation it cannot handle.

If a filename is passed in (vs. a filehandle), the directory name will be passed in to XML::Parser as the I<Base> parameter. This will allow for XML directives to work as expected.

=head2 write_file()

Create the Excel file and write it to the specified filename, if possible. (This
is when the actual merging of the template and the parameters occurs.)

=head2 output()

It will act just like HTML::Template's output() method, returning the resultant
file as a stream, usually for output to the web. (This is when the actual
merging of the template and the parameters occurs.)

=head2 register()

This allows you to register a class as handling a node. q.v. L<Excel::Template::Factory> for more info.

=head1 SUPPORTED NODES

This is a partial list of nodes. See the other classes in this distro for more
details on specific parameters and the like.

Every node can set the ROW and COL parameters. These are the actual ROW/COL
values that the next CELL-type tag will write into.

=over 4

=item * L<WORKBOOK|Excel::Template::Container::Workbook>

This is the node representing the workbook. It is the parent for all other
nodes.

=item * L<WORKSHEET|Excel::Template::Container::Worksheet>

This is the node representing a given worksheet.

=item * L<IF|Excel::Template::Container::Conditional>

This node represents a conditional expression. Its children may or may not be
rendered. It behaves just like L<HTML::Template>'s TMPL_IF.

=item * L<LOOP|Excel::Template::Container::Loop>

This node represents a loop. It behaves just like L<HTML::Template>'s TMPL_LOOP.

=item * L<ROW|Excel::Template::Container::Row>

This node represents a row of data. This is the A in A1.

=item * L<FORMAT|Excel::Template::Container::Format>

This node varies the format for its children. All formatting options supported
in L<Spreadsheet::WriteExcel> are supported here. There are also a number of
formatting shortcuts, such as L<BOLD|Excel::Template::Container::Bold> and
L<ITALIC|Excel::Template::Container::Italic>.

=item * L<BACKREF|Excel::Template::Element::Backref>

This refers back to a cell previously named.

=item * L<CELL|Excel::Template::Element::Cell>

This is the actual cell in a spreadsheet.

=item * L<FORMULA|Excel::Template::Element::Formula>

This is a formula in a spreadsheet.

=item * L<RANGE|Excel::Template::Element::Range>

This is a BACKREF for a number of identically-named cells.

=item * L<VAR|Excel::Template::Element::Var>

This is a variable. It is generally used when the 'text' attribute isn't
sufficient.

=back

=head1 BUGS

None, that I know of.

=head1 SUPPORT

This is production quality software, used in several production web
applications.

=head1 AUTHOR

    Rob Kinyon (rob.kinyon@gmail.com)

=head1 CONTRIBUTORS

There is a mailing list at http://groups.google.com/group/ExcelTemplate or exceltemplate@googlegroups.com

=head2 Robert Graff

=over 4

=item * Finishing formats

=item * Fixing several bugs in worksheet naming

=back

=head1 TEST COVERAGE

I used Devel::Cover to test the coverage of my tests. Every release, I intend to improve these numbers.

Excel::Template is also part of the CPAN Kwalitee initiative, being one of the top 100 non-core modules downloaded from CPAN. If you wish to help out, please feel free to contribute tests, patches, and/or suggestions.

---------------------------- ------ ------ ------ ------ ------ ------ ------
File                           stmt branch   cond    sub    pod   time  total
---------------------------- ------ ------ ------ ------ ------ ------ ------
blib/lib/Excel/Template.pm     90.4   62.5   58.8   90.5  100.0   30.2   82.0
...ib/Excel/Template/Base.pm   83.3   50.0   66.7   75.0   88.9    8.2   80.0
...cel/Template/Container.pm   46.3   20.0   33.3   58.3   85.7    4.6   47.7
...emplate/Container/Bold.pm  100.0    n/a    n/a  100.0    0.0    0.5   95.0
.../Container/Conditional.pm   58.5   52.3   66.7   75.0   66.7    0.8   58.4
...plate/Container/Format.pm  100.0    n/a    n/a  100.0    0.0    0.7   96.6
...plate/Container/Hidden.pm  100.0    n/a    n/a  100.0    0.0    0.2   95.0
...plate/Container/Italic.pm  100.0    n/a    n/a  100.0    0.0    0.1   95.0
...plate/Container/Locked.pm  100.0    n/a    n/a  100.0    0.0    0.1   95.0
...emplate/Container/Loop.pm   55.6   40.0   50.0   77.8   75.0    0.5   56.6
...late/Container/Outline.pm   71.4    n/a    n/a   80.0    0.0    0.0   70.0
...Template/Container/Row.pm  100.0   75.0    n/a  100.0   50.0    0.3   93.8
...mplate/Container/Scope.pm  100.0    n/a    n/a  100.0    n/a    0.1  100.0
...plate/Container/Shadow.pm  100.0    n/a    n/a  100.0    0.0    0.1   95.0
...te/Container/Strikeout.pm  100.0    n/a    n/a  100.0    0.0    0.1   95.0
...ate/Container/Workbook.pm  100.0    n/a    n/a  100.0    n/a    1.0  100.0
...te/Container/Worksheet.pm   94.1   50.0    n/a  100.0    0.0    0.9   88.0
...Excel/Template/Context.pm   83.1   53.4   54.2   95.0   92.9   19.1   75.2
...Excel/Template/Element.pm  100.0    n/a    n/a  100.0    n/a    0.4  100.0
...mplate/Element/Backref.pm  100.0   50.0   33.3  100.0    0.0    0.1   87.1
.../Template/Element/Cell.pm   95.8   65.0   80.0  100.0   66.7    3.6   86.9
...mplate/Element/Formula.pm  100.0    n/a    n/a  100.0    0.0    0.3   94.1
...Template/Element/Range.pm  100.0   66.7    n/a  100.0   66.7    0.2   93.3
...l/Template/Element/Var.pm  100.0    n/a    n/a  100.0    0.0    0.1   94.1
...Excel/Template/Factory.pm   57.1   34.6    n/a   88.9  100.0   14.3   55.2
.../Excel/Template/Format.pm   98.3   81.2   33.3  100.0  100.0    8.9   93.2
...xcel/Template/Iterator.pm   85.2   70.6   70.6   84.6   87.5    1.9   80.4
...el/Template/TextObject.pm   92.9   62.5   33.3  100.0   50.0    2.7   83.0
Total                          83.1   56.6   58.3   91.1   98.7  100.0   78.8
---------------------------- ------ ------ ------ ------ ------ ------ ------

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

=head1 SEE ALSO

perl(1), HTML::Template, Spreadsheet::WriteExcel.

=cut
