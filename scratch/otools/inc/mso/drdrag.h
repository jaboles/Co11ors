#pragma once

/****************************************************************************

	DRDRAG.H

	Owner: DaveBu
	Copyright (c) 1994 Microsoft Corporation

*****************************************************************************/

#ifndef MSODRAG_H
#define MSODRAG_H

typedef int FADC; // FAnimateDrag Client
	/* This is supposed to contain any data the drag proc needs to
		differentiate different callers and/or caller conditions.  It
		can contain flags or an enumeration.  If this stuff get's too
		complicated, you'll have to hang a structure off the client data
		pointers at the end of the FADSE and/or FADS. */


typedef int FADE; // FAnimateDrag Event

#define fadeNil             0 /* Perhaps a modifier key changed. */
#define fadeMouseMove       1
#define fadeKeyDown         2
#define fadeKeyUp           3
#define fadeMouseDown       4
#define fadeMouseUp         5
#define fadeDoubleClick     6 /* This a second mouse down message. */
#define fadeSysKeyDown      7
#define fadeSysKeyUp        8

#define drgaError         0  // Aborted drag
#define drgaCancel        1  // Empty drag
#define drgaSuccess       2  // Normal successful termination
#define drgaOleCopy       3  // OLE drag-drop copy action
#define drgaOleMove       4  // OLE drag-drop move action
#define drgaOleLink       5  // OLE drag-drop link action
#define drgaOleLocalCopy  6  // OLE local drag-drop copy action
#define drgaOleLocalMove  7  // OLE local drag-drop move action
#define drgaOleLocalLink  8  // OLE local drag-drop link action


/* A DRAGB (DRAG Block) contains all the fields of the DRAG that are
	constant and shared for all elements of a drag and are public to
	the drag proc.  Note that any changes here will affect the
	implementation of the DRAG in drdrag.i. */

/********************** UpdateDragRc ***********************/

#define msofudrNil          0x0000
#define msofudrInside       0x0001
#define msofudrOutside      0x0002
#define msofudrSolid        0x0004
#define msofudrDotted       0x0008
#define msofudrFuzzy        0x0010
#define msofudrDiag         0x0020
#define msofudrMyBrush      0x0040
#define msofudrPreserveClip 0x0080
#define msofudrMyBrushDPx   0x0140
#define msofudrMyBrushPDxn  0x0240

#define MsoDrawDragRc(hdc, prc, dxp, dyp, hrgnClip, grfudr) \
	MsoUpdateDragRc(hdc, prc, dxp, dyp, NULL, 0, 0, hrgnClip, grfudr)

MSOMACAPI_(void) MsoUpdateDragRc(HDC, RECT *, LONG, LONG, RECT *, LONG, LONG, HRGN, int);
int MsoFXorDragRcIntoRgn(HRGN, RECT *, LONG, LONG);
void MsoXorDragRc(HDC, RECT *, LONG, LONG, DWORD);
MSOMACAPIX_(void) PatBltRect(HDC, RECT *, DWORD);

/* Raster Operations */
#define ROP_S             0x00CC0020L   // same as SRCCOPY
#define ROP_Sn            0x00330008L
#define ROP_PSo           0x00FC008AL
#define ROP_PSno          0x00F3022AL
#define ROP_DPSxx         0x00960169L
#define ROP_PDSxxn        0x00690145L
#define ROP_Pn            0x000f0001L
#define ROP_DPSnao        0x00B00E05L
#define ROP_DSnx          0x00990066L
#define ROP_DPo           0x00FA0089L
#define ROP_DPa           0x00A000C9L
#define ROP_SPno          0x00CF0224L
#define ROP_SDPanan       0x00B30CE8L
#define ROP_PSDPxhx       0x00B8074AL
#define ROP_DSa           0x008800C6L   // same as SRCAND
#define ROP_DSo           0x00EE0086L   // same as SRCPAINT
#define ROP_DSna          0x00220326L
#define ROP_PDxn          0x00A50065L
#define ROP_DPSoa         0x00A803A9L  // For Dithering Bitmaps
#define ROP_DPSDxhx       0x00CA0749L  // For Diming Bitmaps with hbrdither
#define ROP_DPSao         0x00EA02E9L
#define ROP_DPx			  0x005A0049L

#endif /* MSODRAG_H */
