#
#   Generates tbbtnent.h from tborder_raw.txt and msobtn.h
#
#   msobtn.h: List of toolbar command IDs (tcid) with the bitmap name in comments
#   tbbtnent.h: Icon index. Where to find the icon for a given id in a bstrip
#   tborder.txt: Order of bitmaps in the bstrips
#

use strict;

my $szEntDst = "tbbtnent.h";
my $dir;

sub mydie {
    close fileEntDst;
    unlink "${dir}${szEntDst}";
    print stderr "\n";
    die @_;
}

sub process_file {

# Parse arguments
    my $tborder;
    my $tcids;
    ($tborder, $tcids, $dir) = @_;

# Read the tborder input file and build up a mapping from name to index
    local *TborderInHandle;
    open(TborderInHandle, "$tborder") || mydie "Error: Cannot open file $tborder: $!";
    my $tbIndex = 0;
    my %Tborder;
    while( <TborderInHandle> )
    {
# Search for lines with only one word on them (skips comments or any line beginning
# with punctuation. The \s* allows for any trailing spaces.
# The second form allows for an optional decimal number (ignored)
        if( /^(\w+)\s*$/ || /^(\w+)\s+\d+$/ )
        {
            if( $Tborder{$1} )
            {
                print stderr "Warning: Duplicate entry in $tborder ($1) ignored\n";
            }
            else
            {
                $Tborder{$1} = $tbIndex++;
            }
        }
    }
    close(TborderInHandle);

# Open the tcids file
    local *FILEIN;
    open(FILEIN, "$tcids") || mydie "Error: Cannot open file $tcids: $!";

    if (defined($dir) && ${dir} ne "")
    {
        $dir .= "\\";
    }

    open(fileEntDst, "> ${dir}${szEntDst}") || mydie "Error: Cannot create file ${dir}${szEntDst}";

# Generate BtnTable header
    print fileEntDst "/*****************************************************************************\n";
    print fileEntDst "	tbbtnent.h: list of corresponding bitmaps (strip, index)\n\n";
    print fileEntDst "	Owner: chrismcb\n";
    print fileEntDst "	Copyright (c) 1994-2002 Microsoft Corporation\n";
    print fileEntDst "	Generated file!  Do not edit!\n\n";
    print fileEntDst "	WARNING!! Please read tbbtn.pl.\n";
    print fileEntDst "*****************************************************************************/\n\n";

    while (<FILEIN>) {

# The following regexp finds a line like:
#     #define msotcidFoo      1234 // FooImage,
# Where Foo is the command name and FooImage is the icon name (often the same)
        if( /^#define msotcid(\w+)\s+\d+\s+\/\/\s(\w+),/ )
        {
# Look up the name ($2) in the Tborder hash
            my $found = $Tborder{$2};
            if( !$found )
            {
                printf fileEntDst "\t0, // msotcid%s uses icon Nil\n", $1;
            }
            else
            {
                printf fileEntDst "\t%d, // msotcid%s uses icon %s\n", $found, $1, $2;
            }
        }
    }

# Return success code
    0;
}

sub usage
{
    printf STDERR "";
    printf STDERR "Usage: perl tbbtn.pl tborder.txt msobtn.h outdir\n";
    exit 1;
}

# main
{
    my $rc;

    if (($#ARGV == 0 && lc($ARGV[1]) =~ m,^/?$|^/h$|^-help$,) || $#ARGV != 2) {
        usage();
    }

    $rc = process_file @ARGV;

    exit $rc;
}
