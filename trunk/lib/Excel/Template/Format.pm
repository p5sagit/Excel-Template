package Excel::Template::Format;

use strict;

# This is the format repository. Spreadsheet::WriteExcel does not cache the
# known formats. So, it is very possible to continually add the same format
# over and over until you run out of RAM or addressability in the XLS file. In
# real life, less than 10-20 formats are used, and they're re-used in various
# places in the file. This provides a way of keeping track of already-allocated
# formats and making new formats based on old ones.

{
    my %_Formats;

    sub _assign { $_Formats{$_[0]} = $_[1]; $_Formats{$_[1]} = $_[0] }
#        my $key = shift;
#        my $format = shift;
#        $_Formats{$key} = $format;
#        $_Formats{$format} = $key;
#    }

    sub _retrieve_key { $_Formats{ $_[0] } }
#        my $format = shift;
#        return $_Formats{$format};
#    }

    *_retrieve_format = \&_retrieve_key;
#    sub _retrieve_format {
#        my $key = shift;
#        return $_Formats{$key};
#    }
}

{
    my @_boolean_formats = qw(
        bold italic locked hidden font_outline font_shadow font_strikeout
        text_wrap text_justlast shrink
    );

    my @_integer_formats = qw(
        size num_format underline rotation indent pattern border
        bottom top left right
    );

    my @_string_formats = qw(
        font color align valign bg_color fg_color border_color
        bottom_color top_color left_color right_color
    );

    sub _params_to_key
    {
        my %params = @_;
        $params{lc $_} = delete $params{$_} for keys %params;

        my @parts = (
            (map { !! $params{$_} } @_boolean_formats),
            (map { $params{$_} ? $params{$_} + 0 : '' } @_integer_formats),
            (map { $params{$_} || '' } @_string_formats),
        );

        return join( "\n", @parts );
    }

    sub _key_to_params
    {
        my ($key) = @_;

        my @key_parts = split /\n/, $key;

        my @boolean_parts = splice @key_parts, 0, scalar( @_boolean_formats );
        my @integer_parts = splice @key_parts, 0, scalar( @_integer_formats );
        my @string_parts  = splice @key_parts, 0, scalar( @_string_formats );

        my %params;
        $params{ $_boolean_formats[$_] } = !!1
            for grep { $boolean_parts[$_] } 0 .. $#_boolean_formats;

        $params{ $_integer_formats[$_] } = $integer_parts[$_]
            for grep { length $integer_parts[$_] } 0 .. $#_integer_formats;

        $params{ $_string_formats[$_] } = $string_parts[$_]
            for grep { $string_parts[$_] } 0 .. $#_string_formats;

        return %params;
    }

    sub copy
    {
        shift;
        my ($context, $old_fmt, %properties) = @_;

        defined(my $key = _retrieve_key($old_fmt))
            || die "Internal Error: Cannot find key for format '$old_fmt'!\n";

        my %params = _key_to_params($key);
        PROPERTY:
        while ( my ($prop, $value) = each %properties )
        {
            $prop = lc $prop;
            foreach (@_boolean_formats)
            {
                if ($prop eq $_) {
                    $params{$_} = ($value && $value !~ /false/i);
                    next PROPERTY;
                }
            }
            foreach (@_integer_formats, @_string_formats)
            {
                if ($prop eq $_) {
                    $params{$_} = $value;
                    next PROPERTY;
                }
            }
        }

        my $new_key = _params_to_key(%params);

        my $format = _retrieve_format($new_key);
        return $format if $format;

        $format = $context->{XLS}->add_format(%params);
        _assign($new_key, $format);
        return $format;
    }
}

sub blank_format
{
    shift;
    my ($context) = @_;

    my $blank_key = _params_to_key();

    my $format = _retrieve_format($blank_key);
    return $format if $format;

    $format = $context->{XLS}->add_format;
    _assign($blank_key, $format);
    return $format;
}

1;
__END__