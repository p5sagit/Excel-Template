package Excel::Template::Format;

use strict;

# This is the format repository. Spreadsheet::WriteExcel does not cache the
# known formats. So, it is very possible to continually add the same format
# over and over until you run out of RAM or addressability in the XLS file. In
# real life, less than 10-20 formats are used, and they're re-used in various
# places in the file. This provides a way of keeping track of already-allocated
# formats and making new formats based on old ones.

{
    # %_Parameters is a hash with the key being the format name and the value
    # being the index/length of the format in the bit-vector.
    my %_Formats = ( 
        bold   => [ 0, 1 ],
        italic => [ 1, 1 ],
        locked => [ 2, 1 ],
        hidden => [ 3, 1 ],
        font_outline   => [ 4, 1 ],
        font_shadow    => [ 5, 1 ],
        font_strikeout => [ 6, 1 ],
    );
 
    sub _params_to_vec
    {
        my %params = @_;
        $params{lc $_} = delete $params{$_} for keys %params;
 
        my $vec = '';
        vec( $vec, $_Formats{$_}[0], $_Formats{$_}[1] ) = ($params{$_} && 1)  
            for grep { exists $_Formats{$_} } 
                map { lc } keys %params;
 
        $vec;
    }
 
    sub _vec_to_params
    {
        my ($vec) = @_;
 
        my %params;
        while (my ($k, $v) = each %_Formats) 
        {
            next unless vec( $vec, $v->[0], $v->[1] );
            $params{$k} = 1;
        }
 
        %params;
    }
}

{
    my %_Formats;

    sub _assign {
        $_Formats{$_[0]} = $_[1] unless exists $_Formats{$_[0]};
        $_Formats{$_[1]} = $_[0] unless exists $_Formats{$_[1]};
    }

    sub _retrieve_vec    { ref($_[0]) ? ($_Formats{$_[0]}) : ($_[0]); }
    sub _retrieve_format { ref($_[0]) ? ($_[0]) : ($_Formats{$_[0]}); }
}

sub blank_format
{
    shift;
    my ($context) = @_;

    my $blank_vec = _params_to_vec();

    my $format = _retrieve_format($blank_vec);
    return $format if $format;

    $format = $context->{XLS}->add_format;
    _assign($blank_vec, $format);
    $format;
}

sub copy
{
    shift;
    my ($context, $old_format, %properties) = @_;

    defined(my $vec = _retrieve_vec($old_format))
        || die "Internal Error: Cannot find vector for format '$old_format'!\n";

    my $new_vec = _params_to_vec(%properties);

    $new_vec |= $vec;

    my $format = _retrieve_format($new_vec);
    return $format if $format;

    $format = $context->{XLS}->add_format(_vec_to_params($new_vec));
    _assign($new_vec, $format);
    $format;
}

1;
__END__

Category   Description       Property        Method Name          Implemented
--------   -----------       --------        -----------          -----------
Font       Font type         font            set_font()
           Font size         size            set_size()
           Font color        color           set_color()
           Bold              bold            set_bold()              YES
           Italic            italic          set_italic()            YES
           Underline         underline       set_underline()
           Strikeout         font_strikeout  set_font_strikeout()    YES
           Super/Subscript   font_script     set_font_script()
           Outline           font_outline    set_font_outline()      YES
           Shadow            font_shadow     set_font_shadow()       YES

Number     Numeric format    num_format      set_num_format()

Protection Lock cells        locked          set_locked()            YES
           Hide formulas     hidden          set_hidden()            YES

Alignment  Horizontal align  align           set_align()
           Vertical align    valign          set_align()
           Rotation          rotation        set_rotation()
           Text wrap         text_wrap       set_text_wrap()
           Justify last      text_justlast   set_text_justlast()
           Merge             merge           set_merge()

Pattern    Cell pattern      pattern         set_pattern()
           Background color  bg_color        set_bg_color()
           Foreground color  fg_color        set_fg_color()

Border     Cell border       border          set_border()
           Bottom border     bottom          set_bottom()
           Top border        top             set_top()
           Left border       left            set_left()
           Right border      right           set_right()
           Border color      border_color    set_border_color()
           Bottom color      bottom_color    set_bottom_color()
           Top color         top_color       set_top_color()
           Left color        left_color      set_left_color()
           Right color       right_color     set_right_color()
