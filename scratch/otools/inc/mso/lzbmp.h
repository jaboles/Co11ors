#pragma once

#ifndef _LZBMP
#define _LZBMP

#define LZBITMAP 998
#define RT_LZBITMAP MAKEINTRESOURCE(LZBITMAP)

#ifndef RC_INVOKED

#pragma warning(push)
#pragma warning(disable:4200) // nonstandard extension used : zero-sized array in struct/union

typedef struct _LZBITMAPHEADER {
   WORD dx;		// Width of the bitmap
   WORD dy;		// Height of the bitmap
   WORD type;		// Type of bitmap.  Right now, it must be zero
   BYTE data[0];  	// Compressed data in LZW format
} LZBITMAPHEADER;

#pragma warning(pop)

/*-----------------------------------------------------------------------
|	rgrgbqColorPalette
|
|	Standard windows bitmap color palette
---------------------------------------------------------------WESC----*/

#define RgbQuad(r, g, b) {b, g, r, 0}
const RGBQUAD rgrgbqColorPalette[] =
{
	RgbQuad(0x00, 0x00, 0x00),
	RgbQuad(0x80, 0x00, 0x00),
	RgbQuad(0x00, 0x80, 0x00),
	RgbQuad(0x80, 0x80, 0x00),
	RgbQuad(0x00, 0x00, 0x80),
	RgbQuad(0x80, 0x00, 0x80),
	RgbQuad(0x00, 0x80, 0x80),
	RgbQuad(0x80, 0x80, 0x80),

	RgbQuad(0xC0, 0xC0, 0xC0),
	RgbQuad(0xFF, 0x00, 0x00),
	RgbQuad(0x00, 0xFF, 0x00),
	RgbQuad(0xFF, 0xFF, 0x00),
	RgbQuad(0x00, 0x00, 0xFF),
	RgbQuad(0xFF, 0x00, 0xFF),
	RgbQuad(0x00, 0xFF, 0xFF),
	RgbQuad(0xFF, 0xFF, 0xFF)
};

const RGBQUAD rgrgbqBWPalette[] =
{
	RgbQuad(0x00, 0x00, 0x00),
	RgbQuad(0xFF, 0xFF, 0xFF)
};
#undef RgbQuad

#endif /* RC_INVOKED */

#endif
