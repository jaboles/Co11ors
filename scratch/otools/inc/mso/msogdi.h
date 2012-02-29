#pragma once
#ifndef MSOGDI_H
#define MSOGDI_H 1
#if 0 //$[VSMSO]
/*****************************************************************************
	msogdi.h

	Owner: JohnBo
	Security Owner: RyanHill
	Copyright (c) 2000 Microsoft Corporation

	GDI header file wrapper - provides support functionality for the GDI
	interfaces.
******************************************************************* JohnBo **/
// Pre-include this to avoid multiple definitions of malloc etc
#ifndef _INC_STDLIB
	#include <stdlib.h>
#endif

#ifndef _GDIPLUS_H
	#define DCR_USE_NEW_105760 1
	#define DCR_USE_NEW_127084 1
	#define DCR_USE_NEW_135429 1
	#define DCR_USE_NEW_140782 1
	#define DCR_USE_NEW_140855 1
	#define DCR_USE_NEW_140857 1
	#define DCR_USE_NEW_140861 1
	#define DCR_USE_NEW_145135 1
	#define DCR_USE_NEW_145138 1
	#define DCR_USE_NEW_145139 1
	#define DCR_USE_NEW_145804 1
	#define DCR_USE_NEW_146933 1
	#define DCR_USE_NEW_152154 1
	#define DCR_USE_NEW_175866 1
	#define DCR_USE_NEW_202903 1

	#include <../gdiplus/gdiplus.h>
#else
	#error gdiplus.h included before msogdi.h
#endif

/* Extra APIs only in the static library (not the DLL). */
namespace Gdiplus
	{
	namespace DllExports
		{
		// This will fix the GDI version of this method
		// to map Win32 GDI path clipping to XOR + PostScript
		extern "C" UINT WINGDIPAPI GdipGetWinMetaFileBits(
			HENHMETAFILE    hEmf,           // handle to the enhanced metafile
			UINT            bufferSize,     // buffer size
			LPBYTE          buffer,         // pointer to buffer
			INT             mapMode,        // mapping mode
			HDC             hdcRef          // handle to reference device context
			);
		};

	struct ImageInfo
	{
	GUID RawDataFormat;
	PixelFormat PixelFormat;
	UINT Width, Height;
	UINT TileWidth, TileHeight;
	double Xdpi, Ydpi;
	UINT Flags;
	};

	};

/* The following define wrapper classes to instantiate stuff on the
	stack, given a Foo it defines the stack class MSOFoo, initialized
	with a GpFoo*. */
struct MSOPEN : public Gdiplus::Pen
	{
	inline MSOPEN(Gdiplus::GpPen *ppen) : Gdiplus::Pen(ppen, Gdiplus::Ok) {}
	inline ~MSOPEN(void) { SetNativePen(NULL); }
   MSOPEN(const MSOPEN &);
   MSOPEN& operator=(const MSOPEN &);
	};

struct MSOBRUSH : public Gdiplus::Brush
	{
	inline MSOBRUSH(Gdiplus::GpBrush *pbrush) :
		Gdiplus::Brush(pbrush, Gdiplus::Ok) {}
	inline ~MSOBRUSH(void) { SetNativeBrush(NULL); }
   MSOBRUSH(const MSOBRUSH &);
   MSOBRUSH& operator=(const MSOBRUSH &);
	};

struct MSOGRAPHICSPATH : public Gdiplus::GraphicsPath
	{
	inline MSOGRAPHICSPATH(Gdiplus::GpPath *ppath) :
		Gdiplus::GraphicsPath(ppath)
		{}
	inline ~MSOGRAPHICSPATH(void) { SetNativePath(NULL); }
   MSOGRAPHICSPATH(const MSOGRAPHICSPATH &);
   MSOGRAPHICSPATH& operator=(const MSOGRAPHICSPATH &);
	};

struct MSOBITMAP : public Gdiplus::Bitmap
	{
	inline MSOBITMAP(Gdiplus::GpBitmap *pbmp) :
		Gdiplus::Bitmap(pbmp)
		{}
	inline ~MSOBITMAP(void) { SetNativeImage(NULL); }
   MSOBITMAP(const MSOBITMAP &);
   MSOBITMAP& operator=(const MSOBITMAP &);
	};

class MSOGRAPHICS : public Gdiplus::Graphics
	{
public:
	MSOGRAPHICS(const MSODC *pdc) :
		Gdiplus::Graphics(pdc ? pdc->Pgraphics() : NULL)
		{
		}

	~MSOGRAPHICS(void)
		{
		SetNativeGraphics(NULL);
		}

   MSOGRAPHICS(const MSOGRAPHICS &);
   MSOGRAPHICS& operator=(const MSOGRAPHICS &);
	};

#endif //$[VSMSO]
#else
#pragma message ("msogdi.h file included more than once.  Skipping.")
#endif
