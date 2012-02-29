#
#	Adds new icons found in %btndrp%\icons to tborder.txt
#

my %order;
my %lcorder;
my $tborderin = "$ENV{BTNDRP}\\tborder.txt";
my $tborderout = "$ENV{DESTDIR}\\tborder.txt";
my $szIcons = "$ENV{BTNDRP}\\icons\\";
my $indexmax = 0;

	open(file, "< $tborderin") || die "Cannot open file $tborderin";
	while (<file>)
		{
		chomp;
		my ($icon, $index);
		($icon, $index) = split/\t/;
#		print "$icon  $index\n";
		if (defined($icon) and not $icon =~ m|//|)
			{
			$order{$icon} = $index;
			$lcorder{lc $icon} = $icon;
			if ($index > $indexmax)
				{
				$indexmax = $index;
				}
			}
		}
	close file;
	my $files;

	opendir dirIcons, $szIcons;
	my @iconfiles = readdir dirIcons;
	closedir dirIcons;
	foreach $files (@iconfiles)
		{
		if ($files =~ m/(.bmp|.BMP)$/)
			{
			my $iconname = substr $files, 0, -6;
			if (!defined($lcorder{lc $iconname}))
				{
				$order{$iconname} = ++$indexmax;
				$lcorder{lc $iconname} = $iconname;
				}
			}
		}
	
	open(file, "> $tborderout") || die "Cannot open file $tborderin";
	print file "// tborder.txt\n";
	print file "//\n";
	print file "// The order the icons are found in the button strip.\n";
	print file "// The icon is the name of the tcid, is also the name of the bitmaps.\n";
	print file "// The order is the location in the bitmap. Bitmap # is index div bstrip length, and index into bstrip is index modulo bstrip length\n";
	print file "\n";
	print file "//Icon	Index\n";

	foreach $files (sort {$order{$a} <=> $order{$b}} keys %order)
		{
		print file "$files\t$order{$files}\n";
		}
	close file;
