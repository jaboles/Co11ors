#pragma once

/*****************************************************************************
	msohpal.h

	Owner: JohnBo
	Copyright (c) 1996 Microsoft Corporation

	The halftone palette
*****************************************************************************/

/*****************************************************************************
	Halftone dithering

	A better name for this would be "ordered dither color approximation".  A
	halftone (printers term) conventionally consists of a dot of varying size
	used to approximate a grey or color.  PostScript uses this internally to
	approximate grey levels on appropriate monochrome devices.  On the screen
	an ordered dither generally produces better results and does work even
	when the area filled is smaller than the dither pattern (important for
	brushes).

	The halftone palette is a const LOGPALETTE which matches that in Win95,
	the values in it are used to locate appropriate colors.
******************************************************************* JohnBo **/
typedef struct
	{
	WORD         palVersion;
	WORD         palNumEntries;
	PALETTEENTRY palPalEntry[256];
	}
MSOHPALHALFTONE;
extern MSOMACPUBDATA MSOHPALHALFTONE msovhpalHalftone;

/* Definitions for the Mac palette layout. */
#if MAC
	#define MAC_CUSTOM_START 227
	#define MAC_CUSTOM_END 240

#if DEBUG
	/* The following value controls whether or not we get a halftone palette
		on the Mac - this should go, but for the moment it is convenient while
		finding out how to handle the Mac. */
	extern MSOMACPUBDATA BOOL msovfMacHalftone;

	#define FUSEHALFTONE msovfMacHalftone
#else
	#define FUSEHALFTONE ((BOOL)FALSE)
#endif

#else

#define FUSEHALFTONE ((BOOL)TRUE)

#endif

//$[VSMSO_BEGIN] Moved here from msoesche.h
/*----------------------------------------------------------------------------
	HALFTONE PALETTES

	The Office Halftone palette fulfils two purposes, it supports halftone
	dithered brushes and it gives a "compromise" set of colors to use in an
	application which wants to make best use of color on a 256 color display
	yet needs to co-exist with other applications using the same screen (i.e.
	any Office application!)

	The palette works identically on both Win32 and WLM.  Under Win32 it has
	the same effect as CreateHalftonePalette() however it must NOT be deleted
	until the application terminates - only one palette is ever created per
	application instance (this significantly speeds palette realization, as
	the work only needs to be done once).

	The palette can be selected into any DC, but should normally only be
	selected into a screen DC.  It does no harm selecting the palette into a
	screen DC with other than 256 colors beyond altering the interpretation of
	PALETTERGB and PALETTEINDEX colors (PALETTERGB colors will be mapped
	through the halftone palette, so on a 65536 color display only 256
	will be available).  The palette must be used on only one device at a
	time, as with any palette, although it can be selected into multiple DCs
	for the same device (and can be used on compatible bitmaps).  This
	restriction does *not* apply to WLM programs running on Macs with multiple
	screens - WLM builds a single (application wide) "physical palette", the
	palette translation tables are built for that "physical palette" not for
	a particular screen (a Mac palette must work on any device - a window
	may appear on multiple devices at once!)

	The set of colors available on the Mac and on Win32 differ because of the
	WLM reserved colors (WLM reserves 17 colors out of the 256, the Mac
	halftone palette is carefully designed to ensure that these colors, if
	present, are accessible to PALETTERGB).
------------------------------------------------------------------- JohnBo -*/
/* Return the office halftone palette - this creates the palette if required
	and returns a handle onto it.  The palette must NEVER be deleted.  The
	routine should only fail if GDI resources are exhaused, once it has
	succeeded in a given application instance it will always succeed. */
MSOAPI_(HPALETTE) MsoHpalOffice();

MSOAPI_(void) MsoFreeHpalOffice(HMSOINST hmso);

/* Remove the halftone palette - this will cause it to be regenerated next
	time it is required.  The application must call this when the desktop
	colors change.  The app must then remove all its uses of the returned
	palette and delete it.  The routine returns NULL if the halftone palette
	has not been created.  This interface can also be used to flush the
	halftone palette to reclaim resources (e.g. the application might realize
	that halftone colors are no longer in use).

	NOTE: the API returns the old palette, it does not delete it - the caller
	must do that after ensuring that the palette is no longer selected in
	any HDC (a palette cannot be deleted while selected into an HDC). */
MSOAPI_(HPALETTE) MsoHpalRemove();

/* Create a new office halftone palette - this will be distinct from the
	palette created by MsoHpalOffice() and must be destroyed by the caller
	after use.  This is provided to support applications which need to use
	the palette on a Win32 device other than the screen.  Note that the palette
	contains the "VGA colors" and the "desktop" colors (the four system colors
	which can be changed by the user).  If the system (reserved) colors in the
	palette do not match the current system colors then the palette will
	realize incorrectly, although there is a reasonable chance that the
	reserved colors will match some of the other colors in the palette, the
	palette must be destroyed and recreated whenever the system colors
	changed. */
MSOAPI_(HPALETTE) MsoHpalCreate();
//$[VSMSO_END]

#if 0 //$[VSMSO]
#ifdef MAC
	/* The following accessor macros are defined for convenience on the Mac.
		Both expect r, g, b in the range 0..5 (and they do not check!)  These
		MUST be kept in sync with the actual palette (which is in gelhpal.cpp),
		otherwise everything will go wrong. */
	#define MACPAL(r,g,b) (215 - ((b) + ((g) + ((r)*6))*6))
	#define MACPALG(gray) (215 - (gray)*43)
#endif

/*----------------------------------------------------------------------------
	Halftone palette look-up table.  Go from r/g/b each in the range 0..5 to
	the correct entry in the palette, this is not required on the MAC because
	the entries are in order.
------------------------------------------------------------------- JohnBo -*/
#ifndef MAC
	extern BYTE _msovrgbPalette[6/*B*/][6/*G*/][6/*R*/];
#endif
typedef BYTE _RGBLookupType[6][6];
#define msovrgbPalette ((const _RGBLookupType *)_msovrgbPalette)

/*----------------------------------------------------------------------------
	Gray value look-up table.  The table is an array of indices into the
	halftone palette to select the correct entry for a particular gray level.
	There are different tables for the MAC and PC because of the different
	numbers of grays available.
------------------------------------------------------------------- JohnBo -*/
#ifdef MAC
	#define msocHpalGrays 17
#else
	#define msocHpalGrays 32
#endif
extern BYTE _msovrgbPaletteGrays[msocHpalGrays];
#define msovrgbPaletteGrays ((const BYTE *)_msovrgbPaletteGrays)


/* MsoHalftoneIndex

	The index of the halftone palette entry corresponding to the desired
	color.  The input values are 16.16 fixed point component values in the
	range 0..5.  On the PC use the lookup table, on the Mac use simple
	arithmetic.  NOTE the important range restrictions - 0..5 INCLUSIVE,
	the caller must scale the values to this range. */
inline BYTE MsoHalftoneIndex(UINT ir, UINT ig, UINT ib)
	{
	GELFatal0Tag(ir < 0x60000 && ig < 0x60000 && ib < 0x60000,
							"Halftone palette: range index error", ASSERTTAG('rgvr'));
	#ifdef MAC
		return ByteFromLong(MACPAL(ir>>16, ig>>16, ib>>16));
	#else
		return msovrgbPalette[ib>>16][ig>>16][ir>>16];
	#endif
	}


/* MsoHalftoneIndex29

	The index of the halftone palette entry corresponding to the desired
	color.  The input values are 3.29 unsigned fixed point component values
	in the range 0..5.  On the PC use the lookup table, on the Mac use simple
	arithmetic.  NOTE the important range restrictions - 0..5 INCLUSIVE,
	the caller must scale the values to this range. */
inline BYTE MsoHalftoneIndex29(UINT ir, UINT ig, UINT ib)
	{
	GELFatal0Tag(ir < (6U<<29) && ig < (6U<<29) && ib < (6U<<29),
							"Halftone palette29: range index error", ASSERTTAG('rgvs'));
	#ifdef MAC
		return ByteFromLong(MACPAL(ir>>29, ig>>29, ib>>29));
	#else
		return msovrgbPalette[ib>>29][ig>>29][ir>>29];
	#endif
	}


/* MsoBSystemIndex

	Return the index of a system or custom color entry if close enough to the
	given color, else return 0.  fBiLevel indicates that the best
	approximation will be bi level so a large slop should be used. */
BYTE MsoBSystemIndex(COLORREF col, BOOL fBiLevel);


/*****************************************************************************
	Halftone palette utilities
******************************************************************* JohnBo **/
/* Is the given DIB_RGB_COLORS palette a match for the Office halftone
	palette?  Is this returns true the Office palette can be selected in
	place of the given DIB palette, the number of entries must match those
	in the halftone palette too. */
MSOMACAPI_(BOOL) FPaletteMatchesHalftone(const RGBQUAD *pquad, int c);


/*****************************************************************************
	Palette related DC queries
******************************************************************* JohnBo **/
#if !MAC
/* Examine an HDC to see if it is appropriate to use the Office halftone
	palette on it. */
BOOL FShouldUseHpalHalftone(HDC hdc, DWORD dwType);

/* Examine an HDC to see if it *is* using the halftone palette *and* this
	is appropriate - this API deals with the problematic case of a true color
	HDC which has the halftone palette selected - the API returns FALSE
	because the caller should not map colors into the halftone palette. */
BOOL FDoesUseHpalHalftone(HDC hdc, DWORD dwType);
#endif

/*****************************************************************************
	Generate a "standard" 16 color palette.
******************************************************************* JohnBo **/
MSOMACAPI_(const PALETTEENTRY *) Ppe16(PALETTEENTRY rgpe[16]);


/*****************************************************************************
	Miscellaneous mathematical utilities employed when handling color values.
******************************************************************* JohnBo **/
/*----------------------------------------------------------------------------
	UReduce32, UReduce64

	Given a 16.16 value in the range 0..31 generate the value in the range
	0..255 corresponding to the integral part - this involves taking the 5
	bit integral value and expanding it to 8 bits in such a way that 0 maps
	to 0 and 31 maps to 255.

	Similarly for a 0..63 value
------------------------------------------------------------------- JohnBo -*/
inline BYTE UReduce32(UINT u)
	{
	u >>= 16;      // 5 bit value
	u |= (u << 5); // 10 bits
	return ByteFromLong(u >> 2); // reduce to 8
	}


inline BYTE UReduce64(UINT u)
	{
	u >>= 16;      // 6 bit value
	u |= (u << 6); // 12 bits
	return ByteFromLong(u >> 4); // reduce to 8
	}


/*----------------------------------------------------------------------------
	Expand an 8 bit value to a 16 bit value.
------------------------------------------------------------------- JohnBo -*/
inline USHORT UExpand16(BYTE u)
	{
	return USHORT(u | (u << 8));
	}


/*----------------------------------------------------------------------------
	Expand an 8 bit value to a 32 bit value.
------------------------------------------------------------------- JohnBo -*/
inline UINT UExpand32(BYTE u)
	{
	USHORT w = UExpand16(u);
	return w | (w << 16);
	}

#endif //$[VSMSO]
