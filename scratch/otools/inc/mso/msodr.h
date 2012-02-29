#pragma once
/****************************************************************************
        MSODR.H

        Owner: PeterEn
        Copyright (c) 1994 Microsoft Corporation

        Exported declarations for all Office Drawing Layer (Escher) stuff.
****************************************************************************/
#ifndef MSODR_H
#define MSODR_H

// AKadatch: add *_PTR types just in case obsolete WinNT.h used
#include <basetsd.h>

#include "msobp.h"

#define NewParse   FALSE

#if 0 //$[VSMSO]
interface IMsoReducedHTMLImport;
interface IMsoHTMLImport;
interface IMsoDrawingHTMLExportSite;
interface IMsoInkSurface;

/****************************************************************************
        These 2 structures used to live in drdrag.h but are ref'ed here now...
        TODO (peteren): Hey!!! This file used to be relatively neat and tidy!
        Who put this junky drag stuff right smack at the beginning instead of
        reasonably near something to which it's actually related?  Please
        put this where it actually belongs.
*****************************************************************************/

/* This structure is still present only to optimize the copying of data
        when only *ONE* DRGH is being dragged per DRAGE. This may be removed
        later. */
typedef struct _MSODRGPRC   /* This is the full DRGP structure to facilitate
                                blitting a DRGP and 2 points into it. */
        {
        RECT rchTrack;
        RECT rciTrack;
        RECT rcvTrack;
        POINT pthTrack;
        POINT ptiTrack;
        POINT ptvTrack;
        LONG    langle;
        ULONG fInvertedHorz : 1;
        ULONG fInvertedVert : 1;
        } MSODRGPRC;

/* This structure contains the essential elements needed by a DRAGE
        to perform drags per DRGH of the shape. There will be multiple
        elements of these only for multiple-vertex or multiple-adjust
        handle drags. - SreeramN */
typedef struct _MSODRGP /* DRag Ghost Position */
        {
        POINT pthTrack;
        POINT ptiTrack;
        POINT ptvTrack;
        } MSODRGP;
#endif //$[VSMSO]

/****************************************************************************
        Miscellaneous Rectangle, Region, GDI utilities.

        TODO peteren: Put these somewhere else!

        Note that all the inline functions in here are declared wih __inline
        instead of just inline.  That's so that they'll work when included
        in C code (instead of C++).
*****************************************************************************/


/****************************************************************************
        Miscellaneous GDI extensions
*****************************************************************************/

#define MsoDeletePen(hpen) \
        DeleteObject((HGDIOBJ)(HPEN)(hpen))
#define MsoSelectPen(hdc, hpen) \
        ((HPEN)SelectObject((hdc), (HGDIOBJ)(HPEN)(hpen)))
#define MsoGetStockPen(i) \
        ((HPEN)GetStockObject(i))
#define MsoDeleteBrush(hbr) \
        DeleteObject((HGDIOBJ)(HBRUSH)(hbr))
#define MsoSelectBrush(hdc, hbr) \
        ((HBRUSH)SelectObject((hdc), (HGDIOBJ)(HBRUSH)(hbr)))
#define MsoGetStockBrush(i) \
        ((HBRUSH)GetStockObject(i))
#define MsoDeleteRgn(hrgn) \
        DeleteObject((HGDIOBJ)(HRGN)(hrgn))
#define MsoDeletePalette(hpal) \
        DeleteObject((HGDIOBJ)(HPALETTE)(hpal))
#define MsoDeleteFont(hfont) \
        DeleteObject((HGDIOBJ)(HFONT)(hfont))
#define MsoSelectFont(hdc, hfont) \
        ((HFONT)SelectObject((hdc), (HGDIOBJ)(HFONT)(hfont)))
#define MsoGetStockFont(i) \
        ((HFONT)GetStockObject(i))
#define MsoDeleteBitmap(hbmp) \
        DeleteObject((HGDIOBJ)(HBITMAP)(hbmp))
#define MsoSelectBitmap(hdc, hbmp) \
        ((HBITMAP)SelectObject((hdc), (HGDIOBJ)(HBITMAP)(hbmp)))
#define MsoGetCurrentPen(hdc) \
        ((HPEN)GetCurrentObject(hdc, OBJ_PEN))
#define MsoGetCurrentBrush(hdc) \
        ((HBRUSH)GetCurrentObject(hdc, OBJ_BRUSH))

/****************************************************************************
        Coordinates, Points, Rectangles

        Some of the APIs we export here are really just thin wrappers for
        Win32 APIs.  Also we define here wrapper types (MSOPOINT, MSORECT)
        for Win32 types.  There are two reasons for doing this.  First (in
        at least the case of the functions) it lets us give them proper
        hungarian names.  Second (more pragmatically) it lets us reference and
        operate on these structures even from within Mac Excel, where
        the WLM definitions don't apply and RECT means an 8-byte MacOS Rect
        and UpdateRect means the MacOS function.
*****************************************************************************/

typedef RECT MSORECT;
#define cbMSORECT (sizeof(MSORECT))
typedef POINT MSOPOINT;
#define cbMSOPOINT (sizeof(MSOPOINT))
        /* Hey!  These types are primarily intended to let non-WLM-speaking code
                (like in Mac Excel) where RECT on the Mac is _not_ AfxRECT easily
                (that is, without ifdef's) declare Office-compatible structs.  We
                shouldn't go using these types inside Office except possibly as the
                types of arguments in exported functions.  Office should use normal
                Win32 types whereever possible. */

/* MsoOffsetPt moves a point by specified distances along the x and y
        axis.  It's the same as MsoTranslatePt except that it does it to a single
        point, in place. */
MSOMACAPI_(void) MsoOffsetPt(POINT *ppt, int dx, int dy);

/* MsoOffsetRc moves a rectangle by specified distances along the x and y
        axis.  It's the same as MsoTranslateRc except that it does it to a single
        rectangle, in place. */
MSOAPI_(void) MsoOffsetRc(MSORECT *prc, int dx, int dy);

/* MsoTranslatePt copies a point while moving along a specified vector. */
MSOMACAPI_(void) MsoTranslatePt(const POINT *pptSrc, POINT *pptDest,
        int dx, int dy);

/* MsoTranslateRc copies a rectangle while moving along a specified vector. */
MSOMACAPI_(void) MsoTranslateRc(const RECT *prcSrc, RECT *prcDest,
        int dx, int dy);

/* MsoDxOfRc returns the width of a rectangle. */
__inline int MsoDxOfRc(const MSORECT *prc)
        {return prc->right - prc->left;}

/* MsoDyOfRc returns the height of a rectangle. */
__inline int MsoDyOfRc(const MSORECT *prc)
        {return prc->bottom - prc->top;}

/* MsoEmptyRc makes a rectangle have zero area.  We used to just
        set prc->left to prc->right, but we occasionally ran into problems
        with people who need rectangles to have small (16-bit) coordinates.
        Hopefully setting the whole thing to zero's won't be a bit speed hit. */
__inline void MsoEmptyRc(MSORECT *prc)
        {MsoMemset(prc, 0, cbMSORECT);}

/* MsoFIsRcEmpty returns TRUE if a rectangle has zero or negative area. */
__inline BOOL MsoFIsRcEmpty(const MSORECT *prc)
        {
        return prc->right <= prc->left || prc->bottom <= prc->top;
        }

/* MsoFIsRcValid returns TRUE if a rectangle has zero or positive area. */
__inline BOOL MsoFIsRcValid(const MSORECT *prc)
        {
        return (prc->left <= prc->right && prc->top <= prc->bottom);
        }

/* MsoPrcSetRc sets the four coordinates of a rectangle and returns
        a pointer to the filled-out rectangle. */
MSOAPI_(MSORECT *) MsoPrcSetRc(MSORECT *prc, int xLeft, int yTop, int xRight,
        int yBottom);

/* MsoSetRc sets the four coordinates of a rectangle */
__inline void MsoSetRc(MSORECT *prc, int xLeft, int yTop, int xRight,
                int yBottom)
        {
        prc->left = xLeft;
        prc->top = yTop;
        prc->right = xRight;
        prc->bottom = yBottom;
        }

/* MsoPptSetPt sets the two coordiantes of a point and returns a pointer
        to the filled-out point. */
__inline MSOPOINT * MsoPptSetPt(MSOPOINT *ppt, int x, int y)
        {
        ppt->x = x;
        ppt->y = y;
        return ppt;
        }

/* MsoSetPt sets the two coordinates of a point */
__inline void MsoSetPt(MSOPOINT *ppt, int x, int y)
        {
        MsoPptSetPt(ppt, x, y);
        }

/* MsoFUnionRc unions prc1 and prc2 into prcDest, returning TRUE iff the
        result is a non-empty rectangle.  It's just the Win32 API
        UnionRect renamed to have nice hungarian and (more pragmatically) to
        let you easily (without ifdef's) operate on an MSORECT (WLM AfxRECT,
        actually) when sitting in Mac Excel where UnionRect references the
        MacOS function. */
__inline BOOL MsoFUnionRc(MSORECT *prcDest, const MSORECT *prc1,
                const MSORECT *prc2)
        {
        return UnionRect(prcDest, prc1, prc2);
        }

/* MsoFUnionRcLess sets prcDest to the largest rectangle that contains
        only area that's either in prc1 or prc2.  That's as opposed to
        MsoFUnionRc, which returns the smallest rectangle that contains
        all area that's in either prc1 or prc2.  Returns TRUE iff the
        resulting rectangle is non-empty. */
MSOMACAPI_(BOOL) MsoFUnionRcLess(MSORECT *prcDest, const MSORECT *prc1,
                const MSORECT *prc2);

/* MsoFIntersectRc intersects prc1 and prc2 into prcDest, returning TRUE iff
        the result is a non-empty rectangle.  It's just the Win32 API
        IntersectRect renamed to have nice hungarian and (more pragmatically) to
        let you easily (without ifdef's) operate on an MSORECT (WLM AfxRect,
        actually) when sitting in Mac Excel where RECT means something else. */
__inline BOOL MsoFIntersectRc(MSORECT *prcDest, const MSORECT *prc1,
                const MSORECT *prc2)
        {
        return IntersectRect(prcDest, prc1, prc2);
        }

/* MsoFEqRc compares two rectangles and returns TRUE iff they're identical. */
__inline BOOL MsoFEqRc(const MSORECT *prc1, const MSORECT *prc2)
        {
        return (MsoMemcmp(prc1, prc2, cbMSORECT) == 0);
        }

/* MsoFNeRc compares two rectangles and returns TRUE iff they're different. */
__inline BOOL MsoFNeRc(const MSORECT *prc1, const MSORECT *prc2)
        {
        return !MsoFEqRc(prc1, prc2);
        }

/* MsoFIsRcInRc returns TRUE iff *prcInside is inside *prcOutside, or more
        precisely, if every point in *prcInside is in *prcOutside.  An empty
        rectangle is in any rectangle and no non-empty rectangle is in an
        empty rectangle.  A rectangle is in itself. */
MSOMACAPI_(BOOL) MsoFIsRcInRc(const MSORECT *prcInside,
        const MSORECT *prcOutside);

/* MsoGetCenterPtOfRc returns the center of a rectangle.  We're careful
        to round in such a way that if you offset the rectangle (even moving
        it from negative to positive coordinates) you can just offset the
        old center point you got and get the same thing as if you called us
        with the new rectangle. */
MSOAPIMX_(void) MsoGetCenterPtOfRc(MSOPOINT *ppt, const MSORECT *prc);



/****************************************************************************
        Regions
*****************************************************************************/

typedef HRGN MSOHRGN;
        /* Hey! See comment by MSOPOINT about when it's appropriate to use this
                type.  Generally speaking it's only in non-WLM host code or the
                arguments of functions exported from office. */

#define msohrgnNil ((MSOHRGN)NULL)

/* MSORGNK -- ReGioN Kind. */
typedef int MSORGNK;
#define msorgnkError ((MSORGNK)(ERROR))
#define msorgnkEmpty ((MSORGNK)(NULLREGION))
#define msorgnkSimple ((MSORGNK)(SIMPLEREGION))
#define msorgnkComplex ((MSORGNK)(COMPLEXREGION))

#define MsoFCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, fnCombineMode) \
        (CombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, fnCombineMode) != ERROR)
#define MsoFCopyRgn(hrgnDest, hrgnSrc) \
        MsoFCombineRgn(hrgnDest, hrgnSrc, NA/*hrgnSrc2*/, RGN_COPY)
#define MsoFSubtractRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoFCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_DIFF)
#define MsoFIntersectRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoFCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_AND)
#define MsoFUnionRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoFCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_OR)
#define MsoFXorRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoFCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_XOR)

#define MsoRgnkCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, fnCombineMode) \
        ((MSORGNK)(CombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, fnCombineMode)))
#define MsoRgnkCopyRgn(hrgnDest, hrgnSrc) \
        MsoRgnkCombineRgn(hrgnDest, hrgnSrc, NA/*hrgnSrc2*/, RGN_COPY)
#define MsoRgnkSubtractRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoRgnkCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_DIFF)
#define MsoRgnkIntersectRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoRgnkCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_AND)
#define MsoRgnkUnionRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoRgnkCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_OR)
#define MsoRgnkXorRgn(hrgnDest, hrgnSrc1, hrgnSrc2) \
        MsoRgnkCombineRgn(hrgnDest, hrgnSrc1, hrgnSrc2, RGN_XOR)

#define MsoRgnkOffsetRgn(hrgn, dx, dy) \
        ((MSORGNK)(OffsetRgn(hrgn, dx, dy)))

#define MsoFOffsetRgn(hrgn, dx, dy) \
        (MsoRgnkOffsetRgn(hrgn, dx, dy) != msorgnkError)

#define MsoRgnkRgn(hrgn) MsoRgnkOffsetRgn(hrgn, 0/*dx*/, 0/*dy*/)

/* MsoFEqRgn returns TRUE iff two regions are identical */
__inline BOOL MsoFEqRgn(HRGN hrgn1, HRGN hrgn2)
{
        return EqualRgn(hrgn1, hrgn2);
}

/* MsoFCreateEmptyRgn creates an empty RGN. */
__inline BOOL MsoFCreateEmptyRgn(HRGN *phrgn)
{
        *phrgn = CreateRectRgn(0, 0, 0, 0);
        return (*phrgn != msohrgnNil);
}

/* MsoFEnsureRgn makes sure a region has beem allocated, allocating an empty
        region if necessary. */
__inline BOOL MsoFEnsureRgn(HRGN *phrgn)
{
        if (*phrgn != msohrgnNil)
                return TRUE;
        return MsoFCreateEmptyRgn(phrgn);
}

/* MsoFCreateCopyRgn makes a copy of an existing region. */
MSOAPIMX_(BOOL) MsoFCreateCopyRgn(HRGN *phrgnCopy, HRGN hrgnOriginal);

/* MsoFreeRgn frees a region pointed to by hrgn, checking for msohrgnNil.
        If you want the hrgn in question set to msohrgnNil as well use MsoFreeHrgn
        instead. */
__inline void MsoFreeRgn(HRGN hrgn)
{
        if (hrgn != msohrgnNil)
                MsoDeleteRgn(hrgn);
}

/* MsoFreeHrgn frees a region.  We check for msohrgnNil ahead of time and set
        the hrgn to msohrgnNil afterwards. */
__inline void MsoFreeHrgn(HRGN *phrgn)
{
        if (*phrgn != msohrgnNil)
                {
                MsoDeleteRgn(*phrgn);
                *phrgn = msohrgnNil;
                }
}

/* MsoFCreateRgnFromRc creates a RGN given a pointer to a RECT. */
__inline BOOL MsoFCreateRgnFromRc(HRGN *phrgn, const RECT *prc)
{
        *phrgn = CreateRectRgnIndirect(prc);
        return (*phrgn != msohrgnNil);
}

/* MsoSetRgnFromRc reduces a region to a specified rectangle. */
__inline void MsoSetRgnFromRc(HRGN hrgn, const RECT *prc)
{
        SetRectRgn(hrgn, prc->left, prc->top, prc->right, prc->bottom);
}

/* MsoFIsRgnEmpty returns TRUE if a region is empty (or if the passed-in
        hrgn is hrgnNil; this may seem weird but it's convienient). */
__inline BOOL MsoFIsRgnEmpty(HRGN hrgn)
{
        RECT rc;
        if (hrgn == msohrgnNil)
                return TRUE;
        GetRgnBox(hrgn, &rc);
        return (MsoFIsRcEmpty(&rc));
}

/* MsoEmptyRgn empties an existing region. */
__inline void MsoEmptyRgn(HRGN hrgn)
{
        SetRectRgn(hrgn, 0/*left*/, 0/*top*/, 0/*right*/, 0/*bottom*/);
}

/* MsoRgnkGetRgnBounds gets the bounding rectangle of a region, and
        returns an MSORGNK describing the region. */
__inline MSORGNK MsoRgnkGetRgnBounds(HRGN hrgn, MSORECT *prcBounds)
{
        if (hrgn == msohrgnNil)
                {
                MsoEmptyRc(prcBounds);
                return msorgnkEmpty;
                }
        return (MSORGNK)GetRgnBox(hrgn, prcBounds);
}

#if DEBUG
MSOAPI_(void) MsoDebugBlinkRc(HDC hdc, const RECT *prc);
MSOAPI_(void) MsoDebugBlinkRgn(HDC hdc, HRGN hrgn);
#define MsoDebugGdiFlush() GdiFlush()
#else // DEBUG
#define MsoDebugBlinkRc(hdc, prc)
#define MsoDebugBlinkRgn(hdc, hrgn)
#define MsoDebugGdiFlush()
#endif // DEBUG

__inline MSORGNK MsoRgnkExcludeClipRcl(HDC hdc, const RECT *prcl)
{
        return (MSORGNK)ExcludeClipRect(hdc, prcl->left, prcl->top,
                prcl->right, prcl->bottom);
}

__inline MSORGNK MsoRgnkIntersectClipRcl(HDC hdc, const RECT *prcl)
{
        return (MSORGNK)IntersectClipRect(hdc, prcl->left, prcl->top,
                prcl->right, prcl->bottom);
}

__inline MSORGNK MsoRgnkExcludeClipRgnd(HDC hdc, HRGN hrgnd)
{
        return (MSORGNK)ExtSelectClipRgn(hdc, hrgnd, RGN_DIFF);
}



#if 0 //$[VSMSO]
/****************************************************************************
        Extended Regions

        I've found tons of cases where regions are pain in that we're
        constantly querying them for their bounds and keeping track of
        rectangles along side them and other such crap.  Also there's tons
        of cases where for performance reasons you don't want to build
        up some incredibly complicated region, you just want to give up and
        use a rectangle.  Finally there's a ton of error checking code
        lying around because regions are always allocated instead of just
        being on the stack.

        MSORGNXs are supposed to solve all these problems.  But they're
        vaguely new and untested, so be wary and patient.

        All MSORGNXs must be initialized with some MsoInitRgnx*() API.
        All MSORGNXs must be free'd with MsoFreeRgnx().

        FUTURE peteren: Add debug checks to ensure all MSORGNXs freed.
*****************************************************************************/
#endif //$[VSMSO]

typedef int MSOT; // Tri-state (BELONGS SOMEWHERE MORE GENERAL)
#define msotTrue  ((MSOT)TRUE)
#define msotFalse ((MSOT)FALSE)
#define msotYes   msotTrue
#define msotNo    msotFalse
#define msotMaybe ((MSOT)2)

#if 0 //$[VSMSO]
typedef struct _MSORGNX
        {
        /* The bits in grfrngx do NOT describe the area of
                the region.  They set the rules for how this MSORGNX will
                behave or describe its history.  When you call
                MsoFCopyRgnxArea or MsoFMoveRgnxArea these bits area NOT
                copied. */
        union
                {
                ULONG grfrgnx;
                struct
                        {
                        // grfrgnxRules
                        ULONG nComplexityLim : 16;
                        ULONG fMore : 1;
                        ULONG fLess : 1;
                        ULONG fImprecise : 1;
                        ULONG : 5;
                        // grfrgnxHistory
                        ULONG fError : 1;
                        ULONG fSimplified : 1;
                        ULONG : 6;
                        };
                struct
                        {
                        ULONG grfrgnxRules : 24;
                        ULONG grfrgnxHistory : 8;
                        };
                };
        union
                {
                ULONG grfArea;
                struct
                        {
                        ULONG nComplexity : 16;
                        ULONG : 16;
                        };
                };
        MSORECT rcBounds;
        MSOHRGN hrgn;
        } MSORGNX;
#define cbMSORGNX (sizeof(MSORGNX))

#define msofrgnxMore       (1<<16)
#define msofrgnxLess       (1<<17)
#define msofrgnxImprecise  (1<<18)
#define msofrgnxError      (1<<19)
#define msofrngxSimplified (1<<20)

#define msogrfrgnxExact    (0)

/* --- APIs for getting and setting fields of MSORGNXs --- */

/* Call MsoSetRgnxGrfrgnx to set the rules and history of a MSORGNX.
        Pass in the maximum complexity you want (or 0 for no maximum
        complexity), OR in msofrgnxMore or msofrgnxLess if you want to
        specify the behavior when we reach maximum complexity or error. */
__inline void MsoSetRgnxGrfrgnx(MSORGNX *prgnx, ULONG grfrgnx)
        {
        prgnx->grfrgnx = grfrgnx;
        AssertInline(!(prgnx->fMore && prgnx->fLess));
        /* FUTURE peteren: Check against complexity and cleanup? */
        }

/* MsoGrfrgnxOfRgnx returns the flags of a MSORGNX */
__inline ULONG MsoGrfrgnxOfRgnx(const MSORGNX *prgnx)
        {return prgnx->grfrgnx;}

/* MsoPrcBoundsOfRgnx returns a pointer to the bounding rectangle
        of an MSORGNX */
__inline const MSORECT *MsoPrcBoundsOfRgnx(const MSORGNX *prgnx)
        {return &prgnx->rcBounds;}

/* MsoHrgnOfRgnx returns the hrgn from an MSORGNX (may be msohrgnNil) */
__inline MSOHRGN MsoHrgnOfRgnx(const MSORGNX *prgnx)
        {return prgnx->hrgn;}

/* MsoFIsRgnxEmpty returns TRUE iff *prgnx is empty */
__inline BOOL MsoFIsRgnxEmpty(const MSORGNX *prgnx)
        {return prgnx->nComplexity == 0;}

/* MsoFIsRgnxComplex returns TRUE iff *prgnx is comnplex */
__inline BOOL MsoFIsRgnxComplex(const MSORGNX *prgnx)
        {return prgnx->nComplexity > 1;}

/* MsoFIsRgnxSimple returns TRUE iff *prgnx is just its bounding rectangle. */
__inline BOOL MsoFIsRgnxSimple(const MSORGNX *prgnx)
        {return prgnx->nComplexity <= 1;}

/* MsoFRgnxError returns TRUE if an MSORGNX has had an error. */
__inline BOOL MsoFRgnxError(const MSORGNX *prgnx)
        {return prgnx->fError;}

/* MsoSetRgnxError marks an MSORGNX as having had an error. */
__inline void MsoSetRgnxError(MSORGNX *prgnx)
        {prgnx->fError = TRUE;}

/* MsoClearRgnxError clears the error in an MSORGNX. */
__inline void MsoClearRgnxError(MSORGNX *prgnx)
        {prgnx->fError = FALSE;}

/* MsoClearRgnxHistory clears all "history" state of the MSORGNX.
        MsoEmptyRgnx does this, although MsoEmptyRgnxArea does not. */
__inline void MsoClearRgnxHistory(MSORGNX *prgnx)
        {
        prgnx->fError = FALSE;
        prgnx->fSimplified = FALSE;
        }

/* MsoFEnsureRgnOfRgnx makes sure that a MSORGNX has a region in it.
        You can call MsoHrgnOfRgnx afterwards to get the region. */
__inline BOOL MsoFEnsureRgnOfRgnx(MSORGNX *prgnx)
        {
        return (prgnx->hrgn != msohrgnNil ||
                MsoFCreateRgnFromRc(&prgnx->hrgn, &prgnx->rcBounds));
        }


/* --- APIs for creating and freeing MSORGNXs --- */

/* MsoInitRgnx intializes an MSORGNX.  We promise that a MSORGNX
        full of zero's is initialized; this is convienient when initializing
        structures that contain MSORGNXs. */
__inline void MsoInitRgnx(MSORGNX *prgnx)
        {
        MsoMemset(prgnx, 0, cbMSORGNX);
        AssertInline(prgnx->nComplexityLim == 0);
        AssertInline(!prgnx->fMore);
        AssertInline(!prgnx->fLess);
        AssertInline(!prgnx->fImprecise);
        AssertInline(prgnx->nComplexity == 0);
        AssertInline(!prgnx->fError);
        AssertInline(!prgnx->fSimplified);
        AssertInline(MsoFIsRcEmpty(&prgnx->rcBounds));
        AssertInline(prgnx->hrgn == msohrgnNil);
        AssertInline(MsoFIsRgnxEmpty(prgnx));
        }

/* MsoInitRgnxPolicy initializes an MSORGNX and sets its policies. */
__inline void MsoInitRgnxGrfrgnx(MSORGNX *prgnx, ULONG grfrgnx)
        {
        MsoInitRgnx(prgnx);
        MsoSetRgnxGrfrgnx(prgnx, grfrgnx);
        }

/* MsoSetRgnxFromRc changes an initalized MSORGNX to a specified
        rectangle. */
MSOAPIX_(void) MsoSetRgnxFromRc(MSORGNX *prgnx, const RECT *prc);

/* MsoInitRgnxFromRc initalizes an MSORGNX from an MSORECT.  The MSORGNX
        should be previously uninitialized.   If you have an initalized
        MSORGNX you should just use MsoSetRgnxFromRc. */
__inline void MsoInitRgnxFromRc(MSORGNX *prgnx, const RECT *prc)
        {
        MsoInitRgnx(prgnx);
        MsoSetRgnxFromRc(prgnx, prc);
        }

/* MsoSetRgnxFromRgn changes an initalized MSORGNX to a specified
        MSOHRGN.  The MSORGNX will take over ownership of the MSOHRGN. */
MSOAPIX_(void) MsoSetRgnxFromRgn(MSORGNX *prgnx, MSOHRGN hrgn);

/* MsoInitRgnxFromRgn initalizes an MSORGNX from an MSOHRGN.  The MSORGNX
        should be previously uninitialized.  If you have an initalized
        MSORGNX you should call MsoSetRgnxFromRgn instead.  The MSORGNX
        will take over ownership of the MSOHRGN, so don't go freeing it. */
__inline void MsoInitRgnxFromRgn(MSORGNX *prgnx, MSOHRGN hrgn)
        {
        MsoInitRgnx(prgnx);
        MsoSetRgnxFromRgn(prgnx, hrgn);
        }

/* MsoFreeRgnx frees any allocations hanging out of an MSORGNX.  It also
        NULLs out the allocations so that it's safe to call it twice on
        the same MSORGNX.  After calling this the MSORGNX won't be internally
        consistant, so you shouldn't try to pass it to other MSORGNX APIs. */
__inline void MsoFreeRgnx(MSORGNX *prgnx)
        {MsoFreeHrgn(&prgnx->hrgn);}

/* MsoFIsRgnxFree returns TRUE if an MSORGNX does not need
        freeing. */
__inline BOOL MsoFIsRgnxFree(const MSORGNX *prgnx)
        {return (prgnx->hrgn == msohrgnNil);}

/* MsoFCopyRgnxCore copies an initalized MSORGNX over top of
        another initialized MSORGNX.  Use MsoFCopyRgnx or MsoFCopyRgnxArea
        instead. */
MSOAPIX_(BOOL) MsoFCopyRgnxCore(MSORGNX *prgnxDest, const MSORGNX *prgnxSrc,
        BOOL fOnlyArea);

/* MsoFCopyRgnx copies an MSORGNX from one initialized MSORGNX into
        another.  The rules and history of the source are also copied. */
__inline BOOL MsoFCopyRgnx(MSORGNX *prgnxDest, const MSORGNX *prgnxSrc)
        {return MsoFCopyRgnxCore(prgnxDest, prgnxSrc, FALSE/*fOnlyArea*/);}

/* MsoFCopyRgnx copies the area of one MSORGNX into another.  It
        does NOT change the rules or history of the destination MSORGNX. */
__inline BOOL MsoFCopyRgnxArea(MSORGNX *prgnxDest, const MSORGNX *prgnxSrc)
        {return MsoFCopyRgnxCore(prgnxDest, prgnxSrc, TRUE/*fOnlyArea*/);}

/* MsoMoveRgnxCore moves the contents of one MSORGNX into another.
        Both must be initialized ahead of time, and both will still need
        freeing afterwards.  prgnxSrc will be neatly empty afterwards.
        Use MsoMoveRgnx or MsoMoveRgnxArea instead. */
MSOAPIX_(void) MsoMoveRgnxCore(MSORGNX *prgnxDest, MSORGNX *prgnxSrc,
        BOOL fOnlyArea);

__inline void MsoMoveRgnx(MSORGNX *prgnxDest, MSORGNX *prgnxSrc)
        {MsoMoveRgnxCore(prgnxDest, prgnxSrc, FALSE/*fOnlyArea*/);}

__inline void MsoMoveRgnxArea(MSORGNX *prgnxDest, MSORGNX *prgnxSrc)
        {MsoMoveRgnxCore(prgnxDest, prgnxSrc, TRUE/*fOnlyArea*/);}


/* --- APIs for asking questions of MSORGNXs --- */

/* MsoFIsRcEntirelyInRgnx returns TRUE iff the specified rectangle
        is entirely inside the specified region. */
MSOAPIX_(BOOL) MsoFIsRcEntirelyInRgnx(const MSORECT *prc, MSORGNX *prgnx);

/* MsoTIsRcEnitrelyInRgnx does basically the same thing as
        MsoFIsRcEntirelyInRgnx, but it only does fast tests.  It returns
        msotYes if the rectangle is entirely within the region, msotNo if
        it's not, and msotMaybe if fast tests were not sufficient to tell. */
MSOAPIX_(MSOT) MsoTIsRcEntirelyInRgnx(const MSORECT *prc, MSORGNX *prgnx);

/* MsoFIsRgnxEntirelyInRgnx returns TRUE iff *prgnx1 is entirely
        inside *prgnx2. */
MSOAPIX_(BOOL) MsoFIsRgnxEntirelyInRgnx(MSORGNX *prgnx1, MSORGNX *prgnx2);

/* MsoTIsRgnxEntirelyInRgnx returns msotYes if *prgnx1 is entirely
        inside *prgnx2, msotNo if it isn't, and msotMaybe if doing only
        fast tests cannot tell. */
MSOAPIX_(MSOT) MsoTIsRgnxEntirelyInRgnx(MSORGNX *prgnx1, MSORGNX *prgnx2);

/* MsoTDoesRgnxIntersectRgnx returns msotYes if prgnx1 intersects prgnx2,
        msotNo if they don't intersect, and msotMaybe if it's impossible
        to tell quickly. */
MSOAPIX_(MSOT) MsoTDoesRgnxIntersectRgnx(MSORGNX *prgnx1, MSORGNX *prgnx2);

/* MsoTDoesRcIntersectRgnx returns msotYes if prc intersects prgnx,
        msotNo if they don't intersect, and msotMaybe if it's impossible
        to tell quickly. */
MSOAPIX_(MSOT) MsoTDoesRcIntersectRgnx(const RECT *prc, MSORGNX *prgnx);


/* --- APIs for manipulating MSORGNXs --- */

MSOAPIX_(void) MsoEmptyRgnxCore(MSORGNX *prgnx, BOOL fOnlyArea);

__inline void MsoEmptyRgnx(MSORGNX *prgnx)
        {MsoEmptyRgnxCore(prgnx, FALSE/*fOnlyArea*/);}

__inline void MsoEmptyRgnxArea(MSORGNX *prgnx)
        {MsoEmptyRgnxCore(prgnx, TRUE/*fOnlyArea*/);}

/* MsoFUnionRgnx is just MsoFUnionRgnx for MSORGNXs.  Returns TRUE
        on success. */
MSOAPIX_(BOOL) MsoFUnionRgnx(MSORGNX *prgnxDest, MSORGNX *prgnx1,
        MSORGNX *prgnx2);

/* MsoFSubtractRgnx is just MsoFSubtractRgn for MSORGNXs.  Returns TRUE on success. */
MSOAPIX_(BOOL) MsoFSubtractRgnx(MSORGNX *prgnxDest, MSORGNX *prgnx1,
        MSORGNX *prgnx2);

/* MsoFIntersectRgnx is just MsoFIntersectRgn for MSORGNXs.  Returns TRUE on success. */
MSOAPIX_(BOOL) MsoFIntersectRgnx(MSORGNX *prgnxDest, MSORGNX *prgnx1,
        MSORGNX *prgnx2);

/* MsoFAddRgnxToRgnx does just what it says.  Returns TRUE on success. */
__inline BOOL MsoFAddRgnxToRgnx(MSORGNX *prgnxDest, MSORGNX *prgnxAdd)
        {return MsoFUnionRgnx(prgnxDest, prgnxDest, prgnxAdd);}

/* MsoFSubtractRgnxFromRgnx does just what it says.  Returns TRUE on success. */
__inline BOOL MsoFSubtractRgnxFromRgnx(MSORGNX *prgnxDest, MSORGNX *prgnxSubtract)
        {return MsoFSubtractRgnx(prgnxDest, prgnxDest, prgnxSubtract);}

/* MsoFConfineRgnxWithinRgnx does just what it says.  Returns TRUE on success. */
__inline BOOL MsoFConfineRgnxWithinRgnx(MSORGNX *prgnxDest, MSORGNX *prgnxConfineWithin)
        {return MsoFIntersectRgnx(prgnxDest, prgnxDest, prgnxConfineWithin);}

/* MsoFChangeAddRgnxToRgnx is just like MsoFAddRgnxToRgnx except that
        instead of returning TRUE on success it returns TRUE if there's
        a possibility *prgnxDest changed.  This extra smarts means it's
        slightly slower than MsoFAddRgnxToRgnx. */
MSOAPIX_(BOOL) MsoFChangeAddRgnxToRgnx(MSORGNX *prgnxDest, MSORGNX *prgnxAdd);

/* MsoFChangeAddRcToRgnx is just like MsoFChangeAddRgnxToRgnx except
        that you add in a rectangle instead of a full MSORGNX. */
MSOAPIX_(BOOL) MsoFChangeAddRcToRgnx(MSORGNX *prgnxDest, MSORECT *prcAdd);

/* MsoFChangeSubtractRgnxFromRgnx is just like MsoFSubtractRgnxFromRgnx
        except that, like MsoFChangeAddRgnxToRgnx, it returns TRUE if
        prgnxDest might have changed. */
MSOAPIX_(BOOL) MsoFChangeSubtractRgnxFromRgnx(MSORGNX *prgnxDest, MSORGNX *prgnxSubtract);

/* MsoFAddRgnxToRgnxRef is very like MsoFAddRgnxToRgnx except it plays
        some funny games to avoid cloning regions until absolutely necessary. */
MSOAPIX_(BOOL) MsoFAddRgnxToRgnxRef(MSORGNX **pprgnxDest, MSORGNX *prgnxDest,
        MSORGNX *prgnxAdd);

/* MsoFSubtractRgnxFromRgnxRef is very like MsoFSubtractRgnxFromRgnx except
        it plays some funny games to avoid cloning regions until absolutely
        necessary. */
MSOAPIX_(BOOL) MsoFSubtractRgnxFromRgnxRef(MSORGNX **pprgnxDest,
        MSORGNX *prgnxDest, MSORGNX *prgnxSubtract);



/* --- Random left-over MSORGNX APIs --- */

#if DEBUG
MSOAPI_(void) MsoDebugBlinkRgnx(HDC hdc, const MSORGNX *prgnx);
#else // DEBUG
#define MsoDebugBlinkRgnx(hdc, prgnx)
#endif // DEBUG



/****************************************************************************
        Time-related functions
*****************************************************************************/


/*--------------------------------------------------------------------------
        MsoTicksNow

        Return a vaguely OS independent tick count, something close
        to sixtieths of a second.  Office uses MSOTICKS or TICKS as the
        hungarian for this sort of tick count.
----------------------------------------------------------------- PeterEn -*/
__inline int MsoTicksNow()
        {
        return GetTickCount()/*actually 1000ths of a second*/ >> 4;
        }


/****************************************************************************
        MSODC functions
*****************************************************************************/

/* MSOSDCB == Save DC Block */

#define msogrfsdcNil       0x00000000
#define msofsdcClipRgn     0x00000001
#define msofsdcViewportOrg 0x00000002
#define msofsdcPen         0x00000004
#define msofsdcRop         0x00000008
#define msofsdcBrush       0x00000010
#define msofsdcBkMode      0x00000020
#define msofsdcBkColor     0x00000040

typedef struct _MSOSDCB
        {
        ULONG grfsdc; // Which DC attributes have we saved.
        POINT ptwViewportOrg;
        HRGN hrgnClip;
        HPEN hpen;
        int rop;
        HBRUSH hbr;
        int bkmode;
        COLORREF clrBk;
        } MSOSDCB;
#define cbMSOSDCB (sizeof(MSOSDCB))

/* Allocate an MSOSDCB and pass it to MsoFSaveDc to save with minimum
        expense (space and time) a specified set of the attributes of a DC.
        You specify the attributes with a set of msofsdc* flags.  This function
        returns TRUE iff it succeeds. Iff it returns TRUE you _must_ pass the
        SDCB to MsoRestoreDc in order to avoid leaks.  If we ever need it I
        suppose we could add a function to destruct an SDCB.

        NOTE:  It doesn't make sense to call this function for WMF dc's, since
        no inquiry routines (GetFoo()) work and we are unable to save anything
        other than pens and brushes by selecting some bogus ones.  It is much
        more efficient to Select and save objects outside this routine.
        We will however return TRUE if you call MsoFSaveDc on a WMF. */
MSOAPIX_(BOOL) MsoFSaveDc(const MSODC *pdc, MSOSDCB *psdcb, ULONG grfsdc);

/* After calling MsoFSaveDc (and getting back TRUE) and using the DC you
        should call MsoRestoreDc to put things back the way they were. */
MSOAPIX_(void) MsoRestoreDc(MSOSDCB *psdcb, MSODC *pdc);










/****************************************************************************
        Begin stuff that actually belongs in msodr.h
*****************************************************************************/


#if 0 //$[VSMSO]
/* Include structures shared with the GEL. */
#ifndef MSOESCHE_H
        #include "msoesche.h"
#endif
#endif //$[VSMSO]

/* Forward declarations of stuff in this file. */
interface IMsoDisplayManager;
interface IMsoDisplayManagerSite;
interface IMsoDisplayElementSet;
interface IMsoDragSite;
interface IMsoDragProc;
interface IMsoDrawing;
interface IMsoDrawingSite;
interface IMsoDrawingGroup;
interface IMsoDrawingView;
interface IMsoDrawingViewSite;
interface IMsoDrawingGroup;
interface IMsoDrawingViewGroup;
interface IMsoDrawingSelection;
interface IMsoDrawingSelectionSite;
interface IMsoRule;
interface IMsoBitmapSurface;
interface IMsoHTMLExport;
struct _MSODGSLSI;
typedef struct _MSORHISD MSORHISD;
typedef struct _MSOHISD MSOHISD;
typedef struct _MSOETK MSOETK;
typedef struct _MSOCVS MSOCVS;

/****************************************************************************
        Exported types, constants, and enumerations
*****************************************************************************/

/* MSOHSP = Handle to a Shape.  Identifies a Shape relative to its
        parent Drawing.  This type is opaque to the outside world; inside Office
        we know it to be a pointer to a struct MSOSP, which we have typedef'd
        as SP.  This means inside Office we can treat MSOHSP's and SP *'s
        entirely interchangibly, and hopefully in the debugger everyone will
        be able to look inside MSOHSPs. */
typedef struct MSOSP *MSOHSP;
typedef MSOHSP PMSOSP;

#define msohspNil ((MSOHSP)NULL)

/* MSOHSPV = Handle to a Shape in a View.  This identifies a Shape relative
        to a Drawing View (DGV).  You can use an MSOHSP to identify a Shape
        in a DGV too, but it's slow. */
typedef int MSOHSPV;
#define msohspvNil ((MSOHSPV)-1)
#define msohspvBackground ((MSOHSPV)-2)

typedef void *MSOHUR;  // Handle to Undo Record.
#define msohurNil ((MSOHUR)NULL)
typedef void *MSOHSUR; // Handle to Shape Undo Record.
typedef void *MSOHRUR; // Rule undo record

typedef ULONG MSOID;  // generic object IDentifier
#define msoidNil ((MSOID)0)
typedef MSOID MSODGID;  // Drawing ID
#define msodgidNil ((MSODGID)0)
#define msodgidFirst ((MSODGID)1)
#define msodgidMax ((MSODGID)0x00010000) // High word reserved
typedef MSOID MSOSPID;  // Shape ID
typedef MSOID MSOSPIDF;  // Friendly Shape ID
#define msospidfNil ((MSOSPIDF)-1)
#define msospidNil ((MSOSPID)0)
#define msospidFirst ((MSOSPID)0x00000001)
#define msospidMax ((MSOSPID)0x10000000) // High three bits reserved
typedef MSOID MSOSPLID; // Shape Link ID
#define msosplidNil ((MSOSPLID)0)

typedef LONG MSOOID;   // Ole IDentifier.
#define msooidNil ((MSOOID)NULL)

typedef LONG MSOTXID; // TeXtbox ID
#define msotxidNull 0
#define msotxidNil 0

typedef ULONG MSORUID;

typedef struct _MSOPOINTF
        {
        FLOAT x;
        FLOAT y;
        } MSOPOINTF;

#if DEBUG
typedef struct _MSODGDB // DrawinG Debug parameter Block.
        {
        MSOHSP hsp;        // Used for msodmDgsWriteBePvAnchor, msodmDgvRcvOfHsp
        RECT rcv;          // Used for msodmDgvRcvOfHsp
        void *pvAnchor;    // Used for msodmDgsWriteBePvAnchor
        interface IMsoDrawing *pidg; // Used for msodmDgsWriteBePvAnchor
        BOOL fInUndo;      // Used for msodmDgsWriteBePvAnchor
        } MSODGDB;
#endif
#endif //$[VSMSO]

#define msoszOfficeDrawingClipboardFormat "Office Drawing Shape Format\0"
#define msoszOfficeDrawingDragDropFormat "Office Drawing Drag-Drop Format\0"
#define msoszOfficeDrawingHTMLFormat "HTML Format\0"
/* Pure bitmap formats plus the variants with embedded Office Art data. */
#define msoszGIFClipboardFormat WinMac("GIF\0", "GIFf")
#define msoszOfficeGIFClipboardFormat WinMac("GIF+Office Art\0", "ESGF")
#define msoszJFIFClipboardFormat WinMac("JFIF\0", "JPEG")
#define msoszOfficeJFIFClipboardFormat WinMac("JFIF+Office Art\0", "ESJF")
#define msoszPNGClipboardFormat WinMac("PNG\0", "PNGf")
#define msoszOfficePNGClipboardFormat WinMac("PNG+Office Art\0", "ESPF")
#define msoszOfficeEMFPLUSClipboardFormat WinMac("EMF Plus\0", "EMF+")
#define msoszCAGClipboardFormat WinMac("MSClipGallery\0", "XXXX")       /* No MAC format */
#define msoszInkClipboardFormat "Ink Serialized Format\0"

#define msoszRtf WinMac("Rich Text Format\0", "RTF ")
#define msoszEmbeddedObject WinMac("Embedded Object\0", "EOBJ")
#define msoszObject WinMac("Embed Source\0", "ESRC")
#define msoszObjDesc WinMac("Object Descriptor\0", "OBJD")
#define msoszLinkSrc WinMac("Link Source\0", "LNKS")
#define msoszLinkSrcDesc WinMac("Link Source Descriptor\0", "LKSD")
#define msocfMacShape 'ESPS'
#define msocfMacShapeDrag 'ESPD'
#define msocfMacRtf 'RTF '
#define msocfMacEmbeddedObject 'EOBJ'
#define msocfMacObject 'ESRC'
#define msocfMacObjDesc 'OBJD'
#define msocfMacLinkSrc 'LNKS'
#define msocfMacLinkSrcDesc 'LKSD'
#define msocfMacGIF 'GIFf'
#define msocfMacOfficeGIF 'ESGF'
#define msocfMacJFIF 'JPEG'
#define msocfMacOfficeJFIF 'ESJF'
#define msocfMacPNG 'PNGf'
#define msocfMacOfficePNG 'ESPF'

#if 0 //$[VSMSO]
typedef struct _MSOCSSBOX
        {
        LONG    left;
        LONG    top;
        LONG    right;
        LONG    bottom;

        // At present in the CSS spec units for box model properties are a "length" or a percentage of parent.
        // In principal it might be nice to distinguish pixels from other length designations (like in, cm, pt, etc.)
        // However to code currently just maps pixels into a corresponding length designation at import time.
        // To make this change we would stuff an enum for each instead of the boolean.  Also note:
        // the Percent is always at present false.
        union
                {
                ULONG grfposFlags;
                struct
                        {
                        ULONG   leftPercent:1;
                        ULONG   topPercent:1;
                        ULONG   rightPercent:1;
                        ULONG   bottomPercent:1;
                        };
                };
        } MSOCSSBOX;

typedef enum
        {
        msocsspositionStatic,
        msocsspositionAbsolute,
        msocsspositionRelative,
        msocsspositionFixed,
        } MSOCSSPOSITION;

typedef enum
        {
        msocssdisplayInline,
        msocssdisplayBlock,
        msocssdisplayList_item,
        msocssdisplayRun_in,
        msocssdisplayCompact,
        msocssdisplayMarker,
        msocssdisplayTable,
        msocssdisplayInline_table,
        msocssdisplayTable_row_group,
        msocssdisplayTable_header_group,
        msocssdisplayTable_footer_group,
        msocssdisplayTable_row,
        msocssdisplayTable_column_group,
        msocssdisplayTable_column,
        msocssdisplayTable_cell,
        msocssdisplayTable_caption,
        msocssdisplayNone,
        } MSOCSSDISPLAY;

typedef enum
        {
        msocssoverflowVisible,
        msocssoverflowAuto,
        msocssoverflowScroll,
        msocssoverflowHidden
        } MSOCSSOVERFLOW;

typedef enum
        {
        msocssfloatLeft,
        msocssfloatRight,
        msocssfloatNone
        } MSOCSSFLOAT;

// MSOPH -- position horizontal
typedef enum
        {
        msophAbs = 0, msophMin = msophAbs,
        msophLeft,
        msophCenter,
        msophRight,
        msophInside,
        msophOutside,
        msophMax,
        } MSOPH;

typedef enum
        {
        msotxalignCenter = 0,
        msotxalignLeft,
        msotxalignRight,
        } MSOTXALIGN;

// This structure contains full CSS Box model information on a shape.
typedef struct _MSOSPCSSANCHOR
        {
        MSOCSSBOX rcContent;                    // CSS:left, top, right, bottom
        MSOCSSBOX rcPadding;                    // CSS:padding-left, padding-top, padding-right, padding-bottom
        MSOCSSBOX rcBorders;                    // CSS:border-left, border-top, border-right, border-bottom
        MSOCSSBOX rcBorderWidths;       // Not a box, instead this is the width of each border
                                        // (CSS:border-left-width, border-top-width, border-right-width, border-bottom-width)
        MSOCSSBOX rcMargins;                    // CSS:margin-left, margin-top, margin-right, margin-bottom

        MSOCSSPOSITION   position; // CSS:position
        MSOCSSFLOAT      cssFloat; // CSS:float
        MSOCSSDISPLAY    display;  // CSS:display
        MSOCSSOVERFLOW   overflow; // CSS:overflow

        MSOCSSBOX                 rcMsoWrapDistance;     // word inter-op CSS
        MSOPH            msoPositionHorizantal; // word inter-op CSS
        LONG             zIndex;                // VML CSS: zIndex of shape.

#if FUTURE
        MSOCSSDIR  dir;         // Bidi direction.
        MSOCSSBOX        rcClip;
#endif
        } MSOSPCSSANCHOR;


/* MSODGSAK = Drawing Shape Anchor Kind */
typedef enum
        {
        msodgsakNil = 0,
        msodgsakNormal,
        msodgsakUndo,
        } MSODGSAK;

/* MSOSAC = Shape Anchor Change */
#define msogrfsacNil     0x00000000
#define msofsacDoingEdit 0x00000001
#define msofsacDoingUndo 0x00000002
#define msofsacDoingRedo 0x00000004
#define msofsacDontUndo  0x00000008
#define msofsacDontNotify 0x00000010
#define msofsacDontReanchor 0x00000020

/* MSOCVRK -- Coverage Region Kind */
/* The kind of region that we want computed in the RgnkUnionCoverage
        method of a DGV.  A pixel may either be drawn, not drawn or transparent -
        a transparent pixel is one whose value depends on the pixel which was
        there before (typically "additively" - some of the previous value "shows
        through").  In practice transparency is simulated in Escher by drawing
        some pixels but not others - this could be represented exactly in a region
        but the resultant region would be very complex so where shapes simulate a
        "transparent" effect in this way the transparent region which the shape
        claims includes tis pseudo-transparent area.

        The description of the three regions below deliberately includes a double
        negative - "the region covers no pixels that do not draw" means that all
        of the pixels in this region will be drawn by the shape, pixels outside
        the region may also be drawn! */
typedef enum
        {
        msocvrkModified,   // The region covers all the pixels that draw (and more)
        msocvrkOpaque,     // The region covers no pixels that do not draw
        msocvrkTransparent,// The region covers all pixels that are "transparent"
        msocvrkFastOpaque, // Like opaque, but faster with fewer pixels
        } MSOCVRK;

/* MSODGTU = Drawing Tool Use.
        These are an enumeration of the various things a user might do with
        an MSODGT.  Not all MSODGTUs apply to all MSODGTs. */
typedef enum
        {
        msodgtuNil = 0,
        msodgtuNormal,
        msodgtuClickOffShape, msodgtuCancelWithClick = msodgtuClickOffShape,/*TODO peteren: Remove uses of old name, "cancel" is misleading. */
        msodgtuCancelDuringDrag,
        msodgtuError,
        msodgtuClickOnShape,
        } MSODGTU;

/* MSOPRE = Shape Property Preset */
typedef enum
        {
        msopreNone = -1,
        msopreFirst = 0, msopreMin = msopreFirst, msopreMinLessOne = msopreMin - 1,

        msopreShadowMin, msopreShadowMinLessOne = msopreShadowMin - 1,
        msopreShadow1 = 0,
        msopreShadow2,
        msopreShadow3,
        msopreShadow4,
        msopreShadow5,
        msopreShadow6,
        msopreShadow7,
        msopreShadow8,
        msopreShadow9,
        msopreShadow10,
        msopreShadow11,
        msopreShadow12,
        msopreShadow13,
        msopreShadow14,
        msopreShadow15,
        msopreShadow16,
        msopreShadow17,
        msopreShadow18,
        msopreShadow19,
        msopreShadow20,
        msopreShadowMax, msopreShadowLast = msopreShadowMax - 1,

        // These Text Effect presets are for the Insert Text Effect dialog
        // not the text effect envelop swatch.
        msopreTextEffectMin, msopreTextEffectMinLessOne = msopreTextEffectMin - 1,
        msopreTextEffect1,
        msopreTextEffect2,
        msopreTextEffect3,
        msopreTextEffect4,
        msopreTextEffect5,
        msopreTextEffect6,
        msopreTextEffect7,
        msopreTextEffect8,
        msopreTextEffect9,
        msopreTextEffect10,
        msopreTextEffect11,
        msopreTextEffect12,
        msopreTextEffect13,
        msopreTextEffect14,
        msopreTextEffect15,
        msopreTextEffect16,
        msopreTextEffect17,
        msopreTextEffect18,
        msopreTextEffect19,
        msopreTextEffect20,
        msopreTextEffect21,
        msopreTextEffect22,
        msopreTextEffect23,
        msopreTextEffect24,
        msopreTextEffect25,
        msopreTextEffect26,
        msopreTextEffect27,
        msopreTextEffect28,
        msopreTextEffect29,
        msopreTextEffect30,
        msopreTextEffectMax, msopreTextEffectLast = msopreTextEffectMax - 1,

        msopre3DEffectMin, msopre3DEffectMinLessOne = msopre3DEffectMin - 1,
        msopre3DEffect1,
        msopre3DEffect2,
        msopre3DEffect3,
        msopre3DEffect4,
        msopre3DEffect5,
        msopre3DEffect6,
        msopre3DEffect7,
        msopre3DEffect8,
        msopre3DEffect9,
        msopre3DEffect10,
        msopre3DEffect11,
        msopre3DEffect12,
        msopre3DEffect13,
        msopre3DEffect14,
        msopre3DEffect15,
        msopre3DEffect16,
        msopre3DEffect17,
        msopre3DEffect18,
        msopre3DEffect19,
        msopre3DEffect20,
        msopre3DEffectMax, msopre3DEffectLast = msopre3DEffectMax - 1,

        msopre3DExtrusionMin, msopre3DExtrusionMinLessOne = msopre3DExtrusionMin - 1,
        msopre3DExtrusion1,
        msopre3DExtrusion2,
        msopre3DExtrusion3,
        msopre3DExtrusion4,
        msopre3DExtrusion5,
        msopre3DExtrusion6,
        msopre3DExtrusion7,
        msopre3DExtrusion8,
        msopre3DExtrusion9,
        msopre3DExtrusionMax, msopre3DExtrusionLast = msopre3DExtrusionMax - 1,

        msopre3DLightMin, msopre3DLightMinLessOne = msopre3DLightMin - 1,
        msopre3DLight1,
        msopre3DLight2,
        msopre3DLight3,
        msopre3DLight4,
        msopre3DLight5,                         // just a placeholder
        msopre3DLight6,
        msopre3DLight7,
        msopre3DLight8,
        msopre3DLight9,
        msopre3DLightMax, msopre3DLightLast = msopre3DLightMax - 1,

        msopre3DSurfaceMin, msopre3DSurfaceMinLessOne = msopre3DSurfaceMin - 1,
        msopre3DSurface1,                               // matte
        msopre3DSurface2,                               // plastic
        msopre3DSurface3,                               // metal
        msopre3DSurface4,                               // wire frame
        msopre3DSurfaceMax, msopre3DSurfaceLast = msopre3DSurfaceMax - 1,

        msopre3DLightHarshMin, msopre3DLightHarshMinLessOne = msopre3DLightHarshMin - 1,
        msopre3DLightHarsh1,                                            // FLAT
        msopre3DLightHarsh2,                                            // NORMAL
        msopre3DLightHarsh3,                                            // HARSH
        msopre3DLightHarshMax, msopre3DLightHarshLast = msopre3DLightHarshMax - 1,

        msopreArrowMin, msopreArrowMinLessOne = msopreArrowMin - 1,
        msopreArrow1,
        msopreArrow2,
        msopreArrow3,
        msopreArrow4,
        msopreArrow5,
        msopreArrow6,
        msopreArrow7,
        msopreArrow8,
        msopreArrow9,
        msopreArrow10,
        msopreArrow11,
        msopreArrowMax, msopreArrowLast = msopreArrowMax - 1,

        msopreLineDashMin, msopreLineDashMinLessOne = msopreLineDashMin - 1,
        msopreLineDash1,
        msopreLineDash2,
        msopreLineDash3,
        msopreLineDash4,
        msopreLineDash5,
        msopreLineDash6,
        msopreLineDash7,
        msopreLineDash8,
        msopreLineDashMax, msopreLineDashLast = msopreLineDashMax - 1,

        msopreArrowLSizeMin, msopreArrowLSizeMinLessOne = msopreArrowLSizeMin - 1,
        msopreArrowLSize1,
        msopreArrowLSize2,
        msopreArrowLSize3,
        msopreArrowLSize4,
        msopreArrowLSize5,
        msopreArrowLSize6,
        msopreArrowLSize7,
        msopreArrowLSize8,
        msopreArrowLSize9,
        msopreArrowLSizeMax, msopreArrowLSizeLast = msopreArrowLSizeMax - 1,

        msopreArrowRSizeMin, msopreArrowRSizeMinLessOne = msopreArrowRSizeMin - 1,
        msopreArrowRSize1,
        msopreArrowRSize2,
        msopreArrowRSize3,
        msopreArrowRSize4,
        msopreArrowRSize5,
        msopreArrowRSize6,
        msopreArrowRSize7,
        msopreArrowRSize8,
        msopreArrowRSize9,
        msopreArrowRSizeMax, msopreArrowRSizeLast = msopreArrowRSizeMax - 1,

        msoprePictureImageMin, msoprePictureImageMinLessOne = msoprePictureImageMin - 1,
        msoprePictureImage1,
        msoprePictureImage2,
        msoprePictureImage3,
        msoprePictureImage4,
        msoprePictureImageMax, msoprePictureImageLast = msoprePictureImageMax - 1,

        msopreShadePresetMin, msopreShadePresetMinLessOne = msopreShadePresetMin - 1,
        msopreShadePreset1,
        msopreShadePreset2,
        msopreShadePreset3,
        msopreShadePreset4,
        msopreShadePreset5,
        msopreShadePreset6,
        msopreShadePreset7,
        msopreShadePreset8,
        msopreShadePreset9,
        msopreShadePreset10,
        msopreShadePreset11,
        msopreShadePreset12,
        msopreShadePreset13,
        msopreShadePreset14,
        msopreShadePreset15,
        msopreShadePreset16,
        msopreShadePreset17,
        msopreShadePreset18,
        msopreShadePreset19,
        msopreShadePreset20,
        msopreShadePreset21,
        msopreShadePreset22,
        msopreShadePreset23,
        msopreShadePreset24,
        msopreShadePresetMax, msopreShadePresetLast = msopreShadePresetMax - 1,

        // The following Text Effect presets are for the text effect envelop swatch
        msopreTextEnvMin, msopreTextEnvMinLessOne = msopreTextEnvMin - 1,
        msopreTextEnv1,
        msopreTextEnv2,
        msopreTextEnv3,
        msopreTextEnv4,
        msopreTextEnv5,
        msopreTextEnv6,
        msopreTextEnv7,
        msopreTextEnv8,
        msopreTextEnv9,
        msopreTextEnv10,
        msopreTextEnv11,
        msopreTextEnv12,
        msopreTextEnv13,
        msopreTextEnv14,
        msopreTextEnv15,
        msopreTextEnv16,
        msopreTextEnv17,
        msopreTextEnv18,
        msopreTextEnv19,
        msopreTextEnv20,
        msopreTextEnv21,
        msopreTextEnv22,
        msopreTextEnv23,
        msopreTextEnv24,
        msopreTextEnv25,
        msopreTextEnv26,
        msopreTextEnv27,
        msopreTextEnv28,
        msopreTextEnv29,
        msopreTextEnv30,
        msopreTextEnv31,
        msopreTextEnv32,
        msopreTextEnv33,
        msopreTextEnv34,
        msopreTextEnv35,
        msopreTextEnv36,
        msopreTextEnv37,
        msopreTextEnv38,
        msopreTextEnv39,
        msopreTextEnv40,

        msopreTextEnvMax, msopreTextEnvLast = msopreTextEnvMax - 1,

        msopreLineStyleMin, msopreLineStyleMinLessOne = msopreLineStyleMin - 1,
        msopreLineStyle1,
        msopreLineStyle2,
        msopreLineStyle3,
        msopreLineStyle4,
        msopreLineStyleMax, msopreLineStyleLast = msopreLineStyleMax - 1,

        msopreLeftLineStyleMin, msopreLeftLineStyleMinLessOne = msopreLeftLineStyleMin - 1,
        msopreLeftLineStyle1,
        msopreLeftLineStyle2,
        msopreLeftLineStyle3,
        msopreLeftLineStyle4,
        msopreLeftLineStyleMax, msopreLeftLineStyleLast = msopreLeftLineStyleMax - 1,

        msopreRightLineStyleMin, msopreRightLineStyleMinLessOne = msopreRightLineStyleMin - 1,
        msopreRightLineStyle1,
        msopreRightLineStyle2,
        msopreRightLineStyle3,
        msopreRightLineStyle4,
        msopreRightLineStyleMax, msopreRightLineStyleLast = msopreRightLineStyleMax - 1,

        msopreTopLineStyleMin, msopreTopLineStyleMinLessOne = msopreTopLineStyleMin - 1,
        msopreTopLineStyle1,
        msopreTopLineStyle2,
        msopreTopLineStyle3,
        msopreTopLineStyle4,
        msopreTopLineStyleMax, msopreTopLineStyleLast = msopreTopLineStyleMax - 1,

        msopreBottomLineStyleMin, msopreBottomLineStyleMinLessOne = msopreBottomLineStyleMin - 1,
        msopreBottomLineStyle1,
        msopreBottomLineStyle2,
        msopreBottomLineStyle3,
        msopreBottomLineStyle4,
        msopreBottomLineStyleMax, msopreBottomLineStyleLast = msopreBottomLineStyleMax - 1,

        msopreColumnLineStyleMin, msopreColumnLineStyleMinLessOne = msopreColumnLineStyleMin - 1,
        msopreColumnLineStyle1,
        msopreColumnLineStyle2,
        msopreColumnLineStyle3,
        msopreColumnLineStyle4,
        msopreColumnLineStyleMax, msopreColumnLineStyleLast = msopreColumnLineStyleMax - 1,

        msopreLeftLineDashMin, msopreLeftLineDashMinLessOne = msopreLeftLineDashMin - 1,
        msopreLeftLineDash1,
        msopreLeftLineDash2,
        msopreLeftLineDash3,
        msopreLeftLineDash4,
        msopreLeftLineDash5,
        msopreLeftLineDash6,
        msopreLeftLineDash7,
        msopreLeftLineDash8,
        msopreLeftLineDashMax, msopreLeftLineDashLast = msopreLeftLineDashMax - 1,

        msopreRightLineDashMin, msopreRightLineDashMinLessOne = msopreRightLineDashMin - 1,
        msopreRightLineDash1,
        msopreRightLineDash2,
        msopreRightLineDash3,
        msopreRightLineDash4,
        msopreRightLineDash5,
        msopreRightLineDash6,
        msopreRightLineDash7,
        msopreRightLineDash8,
        msopreRightLineDashMax, msopreRightLineDashLast = msopreRightLineDashMax - 1,

        msopreTopLineDashMin, msopreTopLineDashMinLessOne = msopreTopLineDashMin - 1,
        msopreTopLineDash1,
        msopreTopLineDash2,
        msopreTopLineDash3,
        msopreTopLineDash4,
        msopreTopLineDash5,
        msopreTopLineDash6,
        msopreTopLineDash7,
        msopreTopLineDash8,
        msopreTopLineDashMax, msopreTopLineDashLast = msopreTopLineDashMax - 1,

        msopreBottomLineDashMin, msopreBottomLineDashMinLessOne = msopreBottomLineDashMin - 1,
        msopreBottomLineDash1,
        msopreBottomLineDash2,
        msopreBottomLineDash3,
        msopreBottomLineDash4,
        msopreBottomLineDash5,
        msopreBottomLineDash6,
        msopreBottomLineDash7,
        msopreBottomLineDash8,
        msopreBottomLineDashMax, msopreBottomLineDashLast = msopreBottomLineDashMax - 1,

        msopreColumnLineDashMin, msopreColumnLineDashMinLessOne = msopreColumnLineDashMin - 1,
        msopreColumnLineDash1,
        msopreColumnLineDash2,
        msopreColumnLineDash3,
        msopreColumnLineDash4,
        msopreColumnLineDash5,
        msopreColumnLineDash6,
        msopreColumnLineDash7,
        msopreColumnLineDash8,
        msopreColumnLineDashMax, msopreColumnLineDashLast = msopreColumnLineDashMax - 1,

        msopreMax, msopreLast = msopreMax -1,

        } MSOPRE;
/* MSOPRESET is bad hungarian, we actually call these variables "PRE"s
        everywhere.  So I changes the type to MSOPRE and added a temporary
        #define so that old MSOPRESET people still work. */
#define MSOPRESET MSOPRE


// MSOSPT -- Shape Type
/* NOTE: Shape types are saved in our file format, so ONLY ADD NEW TYPES.
   Do not rearrange this enum */
typedef enum
        {
        msosptMin                        = 0,
        msosptNotPrimitive               = msosptMin,
        msosptFirstAutoShape             = 1,
        msosptRectangle                  = msosptFirstAutoShape,
        msosptRoundRectangle             = 2,
        msosptEllipse                    = 3,
        msosptDiamond                    = 4,
        msosptIsocelesTriangle           = 5,
        msosptRightTriangle              = 6,
        msosptParallelogram              = 7,
        msosptTrapezoid                  = 8,
        msosptHexagon                    = 9,
        msosptOctagon                    = 10,
        msosptPlus                       = 11,
        msosptStar                       = 12,
        msosptArrow                      = 13,
        msosptThickArrow                 = 14,
        msosptHomePlate                  = 15,
        msosptCube                       = 16,
        msosptBalloon                    = 17,
        msosptSeal                       = 18,
        msosptArc                        = 19,
        msosptLine                       = 20,
        msosptPlaque                     = 21,
        msosptCan                        = 22,
        msosptDonut                      = 23,
        msosptTextSimple                 = 24,
        msosptTextOctagon                = 25,
        msosptTextHexagon                = 26,
        msosptTextCurve                  = 27,
        msosptTextWave                   = 28,
        msosptTextRing                   = 29,
        msosptTextOnCurve                = 30,
        msosptTextOnRing                 = 31,
        msosptStraightConnector1         = 32,
        msosptBentConnector2             = 33,
        msosptBentConnector3             = 34,
        msosptBentConnector4             = 35,
        msosptBentConnector5             = 36,
        msosptCurvedConnector2           = 37,
        msosptCurvedConnector3           = 38,
        msosptCurvedConnector4           = 39,
        msosptCurvedConnector5           = 40,
        msosptCallout1                   = 41,
        msosptCallout2                   = 42,
        msosptCallout3                   = 43,
        msosptAccentCallout1             = 44,
        msosptAccentCallout2             = 45,
        msosptAccentCallout3             = 46,
        msosptBorderCallout1             = 47,
        msosptBorderCallout2             = 48,
        msosptBorderCallout3             = 49,
        msosptAccentBorderCallout1       = 50,
        msosptAccentBorderCallout2       = 51,
        msosptAccentBorderCallout3       = 52,
        msosptRibbon                     = 53,
        msosptRibbon2                    = 54,
        msosptChevron                    = 55,
        msosptPentagon                   = 56,
        msosptNoSmoking                  = 57,
        msosptSeal8                      = 58,
        msosptSeal16                     = 59,
        msosptSeal32                     = 60,
        msosptWedgeRectCallout           = 61,
        msosptWedgeRRectCallout          = 62,
        msosptWedgeEllipseCallout        = 63,
        msosptWave                       = 64,
        msosptFoldedCorner               = 65,
        msosptLeftArrow                  = 66,
        msosptDownArrow                  = 67,
        msosptUpArrow                    = 68,
        msosptLeftRightArrow             = 69,
        msosptUpDownArrow                = 70,
        msosptIrregularSeal1             = 71,
        msosptIrregularSeal2             = 72,
        msosptLightningBolt              = 73,
        msosptHeart                      = 74,
        msosptPictureFrame               = 75,
        msosptQuadArrow                  = 76,
        msosptLeftArrowCallout           = 77,
        msosptRightArrowCallout          = 78,
        msosptUpArrowCallout             = 79,
        msosptDownArrowCallout           = 80,
        msosptLeftRightArrowCallout      = 81,
        msosptUpDownArrowCallout         = 82,
        msosptQuadArrowCallout           = 83,
        msosptBevel                      = 84,
        msosptLeftBracket                = 85,
        msosptRightBracket               = 86,
        msosptLeftBrace                  = 87,
        msosptRightBrace                 = 88,
        msosptLeftUpArrow                = 89,
        msosptBentUpArrow                = 90,
        msosptBentArrow                  = 91,
        msosptSeal24                     = 92,
        msosptStripedRightArrow          = 93,
        msosptNotchedRightArrow          = 94,
        msosptBlockArc                   = 95,
        msosptSmileyFace                 = 96,
        msosptVerticalScroll             = 97,
        msosptHorizontalScroll           = 98,
        msosptCircularArrow              = 99,
        msosptNotchedCircularArrow       = 100,
        msosptUturnArrow                 = 101,
        msosptCurvedRightArrow           = 102,
        msosptCurvedLeftArrow            = 103,
        msosptCurvedUpArrow              = 104,
        msosptCurvedDownArrow            = 105,
        msosptCloudCallout               = 106,
        msosptEllipseRibbon              = 107,
        msosptEllipseRibbon2             = 108,
        msosptFlowChartProcess           = 109,
        msosptFlowChartDecision          = 110,
        msosptFlowChartInputOutput       = 111,
        msosptFlowChartPredefinedProcess = 112,
        msosptFlowChartInternalStorage   = 113,
        msosptFlowChartDocument          = 114,
        msosptFlowChartMultidocument     = 115,
        msosptFlowChartTerminator        = 116,
        msosptFlowChartPreparation       = 117,
        msosptFlowChartManualInput       = 118,
        msosptFlowChartManualOperation   = 119,
        msosptFlowChartConnector         = 120,
        msosptFlowChartPunchedCard       = 121,
        msosptFlowChartPunchedTape       = 122,
        msosptFlowChartSummingJunction   = 123,
        msosptFlowChartOr                = 124,
        msosptFlowChartCollate           = 125,
        msosptFlowChartSort              = 126,
        msosptFlowChartExtract           = 127,
        msosptFlowChartMerge             = 128,
        msosptFlowChartOfflineStorage    = 129,
        msosptFlowChartOnlineStorage     = 130,
        msosptFlowChartMagneticTape      = 131,
        msosptFlowChartMagneticDisk      = 132,
        msosptFlowChartMagneticDrum      = 133,
        msosptFlowChartDisplay           = 134,
        msosptFlowChartDelay             = 135,
        msosptTextPlainText              = 136,
        msosptTextStop                   = 137,
        msosptTextTriangle               = 138,
        msosptTextTriangleInverted       = 139,
        msosptTextChevron                = 140,
        msosptTextChevronInverted        = 141,
        msosptTextRingInside             = 142,
        msosptTextRingOutside            = 143,
        msosptTextArchUpCurve            = 144,
        msosptTextArchDownCurve          = 145,
        msosptTextCircleCurve            = 146,
        msosptTextButtonCurve            = 147,
        msosptTextArchUpPour             = 148,
        msosptTextArchDownPour           = 149,
        msosptTextCirclePour             = 150,
        msosptTextButtonPour             = 151,
        msosptTextCurveUp                = 152,
        msosptTextCurveDown              = 153,
        msosptTextCascadeUp              = 154,
        msosptTextCascadeDown            = 155,
        msosptTextWave1                  = 156,
        msosptTextWave2                  = 157,
        msosptTextWave3                  = 158,
        msosptTextWave4                  = 159,
        msosptTextInflate                = 160,
        msosptTextDeflate                = 161,
        msosptTextInflateBottom          = 162,
        msosptTextDeflateBottom          = 163,
        msosptTextInflateTop             = 164,
        msosptTextDeflateTop             = 165,
        msosptTextDeflateInflate         = 166,
        msosptTextDeflateInflateDeflate  = 167,
        msosptTextFadeRight              = 168,
        msosptTextFadeLeft               = 169,
        msosptTextFadeUp                 = 170,
        msosptTextFadeDown               = 171,
        msosptTextSlantUp                = 172,
        msosptTextSlantDown              = 173,
        msosptTextCanUp                  = 174,
        msosptTextCanDown                = 175,
        msosptFlowChartAlternateProcess  = 176,
        msosptFlowChartOffpageConnector  = 177,
        msosptCallout90                  = 178,
        msosptAccentCallout90            = 179,
        msosptBorderCallout90            = 180,
        msosptAccentBorderCallout90      = 181,
        msosptLeftRightUpArrow           = 182,
        msosptSun                        = 183,
        msosptMoon                       = 184,
        msosptBracketPair                = 185,
        msosptBracePair                  = 186,
        msosptSeal4                      = 187,
        msosptDoubleWave                 = 188,
        msosptActionButtonBlank          = 189,
        msosptActionButtonHome           = 190,
        msosptActionButtonHelp           = 191,
        msosptActionButtonInformation    = 192,
        msosptActionButtonForwardNext    = 193,
        msosptActionButtonBackPrevious   = 194,
        msosptActionButtonEnd            = 195,
        msosptActionButtonBeginning      = 196,
        msosptActionButtonReturn         = 197,
        msosptActionButtonDocument       = 198,
        msosptActionButtonSound          = 199,
        msosptActionButtonMovie          = 200,
        msosptHostControl                = 201,
        msosptTextBox                    = 202,

        msosptMax, msosptLast = msosptMax - 1,
        msosptNil                        = 0x0FFF,
        } MSOSPT;


/* MSOSPC = Shape Category */
/* This is a MUCH fluffier notion than MSOSPT.  These are used,
        for example, when naming the Format Object dialog, menu item, and undo
        string.
        This table must match the table in ostrinit.pp(msoidsSpcoaXXX) and ostrman.pp(msoidsSpcuXXXX),
        so don't forget to update ostrinit.pp and ostrman.pp */
typedef enum
        {
        msospcMin = 0,
        msospcObject = msospcMin,
        msospcPicture,
        msospcControl,
        msospcWordArt,
        msospcTextBox,
        msospcAutoShape,
        msospcPlaceholder,
        msospcComment,
        msospcHorizRule,
        msospcCanvas,
        msospcOrgChart,
        msospcDiagram,
        msospcInk,
        msospcArc,
        msospcFreeform,
        msospcLine,
        msospcOval,
        msospcRectangle,
        msospcButton,
        msospcCheckBox,
        msospcDropDown,
        msospcEditBox,
        msospcGroupBox,
        msospcLabel,
        msospcListBox,
        msospcOptionButton,
        msospcScrollBar,
        msospcSpinner,
        msospcChart,
        msospcDialog,
        msospcGroup,
        msospcMax,
        msospcNil = 0xFF,
        } MSOSPC;

/****************************************************************************

// MSODGMT -- DiaGraM Type
/* NOTE: Diagram types are saved in our file format, so ONLY ADD NEW TYPES.
   Do not rearrange this enum */
typedef enum
        {
        msodgmtMin                       = 0,
        msodgmtCanvas                    = msodgmtMin,
        msodgmtFirstDiagramType          = 1,
        msodgmtOrgChart                  = msodgmtFirstDiagramType,
        msodgmtRadial                    = 2,
        msodgmtCycle                     = 3,
        msodgmtStacked                   = 4,
        msodgmtVenn                      = 5,
        msodgmtBullsEye                  = 6,
        msodgmtMax, msodgmtLast          = msodgmtMax - 1,
        msodgmtNil                       = 0x0FFF,
        } MSODGMT;


/* MSODGMC = DiaGraM Category */
/* This is a MUCH fluffier notion than MSODGMT.  These are used,
        for example, when naming the Format Object dialog, menu item, and undo
        string. */
typedef enum
        {
        msodgmcMin                 = 0,
        msodgmcCanvas              = msodgmcMin,
        msodgmcDiagram,
        msodgmcOrganizational,
        msodgmcStructural,
        msodgmcSchedule,
        msodgmcMax,
        msodgmcNil                 = 0xFF,
        } MSODGMC;


/* MSODGMLO = DiaGgraG LayOut */
typedef enum
        {
        msodgmloFirst = 0, msodgmloMin = msodgmloFirst, msodgmloMinLessOne = msodgmloMin - 1,

        // OrgChart layout
        msodgmloOrgChartMin,
        msodgmloOrgChartStd = 0,
        msodgmloOrgChartBothHanging,
        msodgmloOrgChartRightHanging,
        msodgmloOrgChartLeftHanging,
        msodgmloOrgChartMax, msodgmloOrgChartLast = msodgmloOrgChartMax - 1,

        // Cycle layout
        msodgmloCycleMin, msodgmloCycleMinLessOne = msodgmloCycleMin - 1,
        msodgmloCycleStd,
        msodgmloCycleMax, msodgmloCycleLast = msodgmloCycleMax - 1,

        // Radial layout
        msodgmloRadialMin, msodgmloRadialMinLessOne = msodgmloRadialMin - 1,
        msodgmloRadialStd,
        msodgmloRadialMax, msodgmloRadialLast = msodgmloRadialMax - 1,

        // Stacked layout
        msodgmloStackedMin, msodgmloStackedMinLessOne = msodgmloStackedMin - 1,
        msodgmloStackedStd,
        msodgmloStackedMax, msodgmloStackedLast = msodgmloStackedMax - 1,

        // Venn layout
        msodgmloVennMin, msodgmloVennMinLessOne = msodgmloVennMin - 1,
        msodgmloVennStd,
        msodgmloVennMax, msodgmloVennLast = msodgmloVennMax - 1,

        // BullsEye layout
        msodgmloBullsEyeMin, msodgmloBullsEyeMinLessOne = msodgmloBullsEyeMin - 1,
        msodgmloBullsEyeStd,
        msodgmloBullsEyeMax, msodgmloBullsEyeLast = msodgmloBullsEyeMax - 1,

        msodgmloMax, msodgmloLast = msodgmloMax -1,
        msodgmloNil = 0xFF,
        } MSODGMLO;


/* MSODGMLR = DiaGraM LineRout */
typedef enum
        {
        msodgmlrFirst = 0, msodgmlrMin = msodgmlrFirst, msodgmlrMinLessOne = msodgmlrMin - 1,

        // OrgChart Line Route
        msodgmlrOrgChart1 = 0,
        msodgmlrOrgChart2,
        msodgmlrOrgChart3,

        msodgmlrMax, msodgmlrLast = msodgmlrMax -1,
        msodgmlrNil = 0xFF,
        } MSODGMLR;


/* MSODGMPT = DiaGraM ProtoType */
typedef enum
        {

        /***NOTE: If you add a new prototype enum, you need to update vrgdgmptdesc ***/

        msodgmptFirst = 0, msodgmptMin = msodgmptFirst, msodgmptMinLessOne = msodgmptMin - 1,

        // OrgChart prototypes
        msodgmptOrgChartBlank = 0,
        msodgmptOrgChart1,
        msodgmptOrgChart2,

        // Cycle prototypes
        msodgmptCycleBlank,
        msodgmptCycle1,

        // Radial prototypes
        msodgmptRadialBlank,
        msodgmptRadial1,

        // Stacked prototypes
        msodgmptStackedBlank,
        msodgmptStacked1,

        // Venn prototypes
        msodgmptVennBlank,
        msodgmptVenn1,

        // BullsEye prototypes
        msodgmptBullsEyeBlank,
        msodgmptBullsEye1,

        msodgmptMax, msodgmptLast = msodgmptMax -1,
        msodgmptNil = 0xFF,
        } MSODGMPT;


/* MSODGMST = DiaGraM STyle */
typedef enum
        {
        /***WARNING: This is written out to the file format! ***/
        /***NOTE: If you add a new style enum, you need to update vrgdgmstdesc ***/

        msodgmstMin = 0, msodgmstFirst = msodgmstMin ,

        // OrgChart styles
        msodgmstOrgChartFirst = msodgmstFirst,
        msodgmstOrgChart2,
        msodgmstOrgChart3,
        msodgmstOrgChart4,
        msodgmstOrgChart5,
        msodgmstOrgChart6,
        msodgmstOrgChart7,
        msodgmstOrgChart8,
        msodgmstOrgChart9,
        msodgmstOrgChart10,
        msodgmstOrgChart11,
        msodgmstOrgChart12,
        msodgmstOrgChart13,
        msodgmstOrgChart14,
        msodgmstOrgChart15,
        msodgmstOrgChart16,
        msodgmstOrgChart17,
        msodgmstOrgChartMax,	
        msodgmstOrgChartLast = msodgmstOrgChartMax - 1,

        // Radial styles
        msodgmstRadialFirst = msodgmstFirst,
        msodgmstRadial2,
        msodgmstRadial3,
        msodgmstRadial4,
        msodgmstRadial5,
        msodgmstRadial6,
        msodgmstRadial7,
        msodgmstRadial8,
        msodgmstRadial9,
        msodgmstRadial10,
        msodgmstRadialMax,
        msodgmstRadialLast = msodgmstRadialMax - 1,

        // Cycle styles
        msodgmstCycleFirst = msodgmstFirst,
        msodgmstCycle2,
        msodgmstCycle3,
        msodgmstCycle4,
        msodgmstCycle5,
        msodgmstCycle6,
        msodgmstCycle7,
        msodgmstCycle8,
        msodgmstCycle9,
        msodgmstCycle10,
        msodgmstCycle2First,
        msodgmstCycle11 = msodgmstCycle2First,
        msodgmstCycle12,
        msodgmstCycle13,
        msodgmstCycle14,
        msodgmstCycle15,
        msodgmstCycle16,
        msodgmstCycle17,
        msodgmstCycle18,
        msodgmstCycle19,
        msodgmstCycle20,
        msodgmstCycle21,
        msodgmstCycleMax,
        msodgmstCycleLast = msodgmstCycleMax - 1,

        // Stacked styles
        msodgmstStackedFirst = msodgmstFirst,
        msodgmstStacked2,
        msodgmstStacked3,
        msodgmstStacked4,
        msodgmstStacked5,
        msodgmstStacked6,
        msodgmstStacked7,
        msodgmstStacked8,
        msodgmstStacked9,
        msodgmstStacked10,
        msodgmstStackedMax,
        msodgmstStackedLast = msodgmstStackedMax - 1,

        // Venn styles
        msodgmstVennFirst = msodgmstFirst,
        msodgmstVenn2,
        msodgmstVenn3,
        msodgmstVenn4,
        msodgmstVenn5,
        msodgmstVenn6,
        msodgmstVenn7,
        msodgmstVenn8,
        msodgmstVenn9,
        msodgmstVenn10,
        msodgmstVennMax,
        msodgmstVennLast = msodgmstVennMax - 1,

        // BullsEyeChart styles
        msodgmstBullsEyeFirst = msodgmstFirst,
        msodgmstBullsEye2,
        msodgmstBullsEye3,
        msodgmstBullsEye4,
        msodgmstBullsEye5,
        msodgmstBullsEye6,
        msodgmstBullsEye7,
        msodgmstBullsEye8,
        msodgmstBullsEye9,
        msodgmstBullsEye10,
        msodgmstBullsEyeMax,
        msodgmstBullsEyeLast = msodgmstBullsEyeMax - 1,

        msodgmstNil = 0xFFFF,
        } MSODGMST;


/* MSODGMPRE = DiaGraM property PREset */
typedef enum
        {
        msodgmpreNone = -1,
        msodgmpreFirst = 0, msodgmpreMin = msodgmpreFirst, msodgmpreMinLessOne = msodgmpreMin - 1,

        msodgmpreOrgChartMin, msodgmpreOrgChartMinLessOne = msodgmpreOrgChartMin - 1,
        msodgmpreOrgChart1 = 0,
        msodgmpreOrgChartMax, msodgmpreOrgChartLast = msodgmpreOrgChartMax - 1,

        msodgmpreMax, msodgmpreLast = msodgmpreMax -1,
        } MSODGMPRE;


/*
        Canvas and Diagram Extensibility
        Host apps decide which capabilities they support for objects that may have
        whatever features.

        WANT TO ADD A NEW BIT?
        1. Remember, only add bits for features that you want degraded behavior in
           previous Office versions or other host apps.
        2. Added a new MSODGMT?  You don't necessarily need to add a new feature
           capability bit.  However, you'll probably need a full build, apps too.
        3. Still here?  Okay, add the bit definition below, DGMT-specific.
        4. In the host apps that support it, set the new bit in
           MSODGGSI::rgGrfDgmtCaps.
        5. Figure out under what scenarios the bit should be set on a shape of that
           DGMT, and only set the bit for those cases.
        6. For cases where the degraded behavior should be as a static group
           instead of a canvas, clear the msofCapSupportsObj on msopidGrfCvsCaps.
        7. That should be it, enjoy!

        Note: For Office11, in order to refer to the specific release, You The
        Reader must add at least one new bit for each dgm type, and have it set by
        each host app's MSODGGSI!!
*/

#define msofCapSupportsObj    (1 << 0)  // used by all dgmt's, basic support
// don't add anymore right here, the one above is Special

// canvas-specific caps
#define msofCapCvsRotation    (1 << 1)  // supports canvas rotation
#define msofCapCvsNesting     (1 << 2)  // supports nested canvases

// add diagram-specific bits here
#if NEVER
#define msofCapOrgChartSample (1 << 1)  // Use me!!  See instructions above
#endif


/****************************************************************************

        MSODGCID = Drawing Command ID

        These are all the commands Office Drawing understands.  These are
        generally passed into methods of the DGUI (Drawing User Interface)
        object.

   *** WARNING *** tuann ***
   Office9 Developers -- PLEASE READ --
   New ids should be added only after msodgcidStartOffice9. This restriction
   is to prevent con*flict with ids that have been added by the Mac Office8
   team. I have reserved a block [msodgcidStartMacOffice8, msodgcidStartOffice9 -1]
   for Mac usage.

   Office Developers --

        Order really matters for these.  These are written into files, and many
        of them are indicies into various tables.  Do NOT ever change the value
        of an existing dgcid (you'll break old files). Also, if you modify this
        enum you should also update drvba.txt which is used for macro
        recording.

   *** WARNING *** neerajm ***
        Please pay attention to the reserved id ranges for different projects
        when you add a new dgcid.
                Upto 299 belongs to Office97
                300 - 309 belong to Mac Office 98
                310 - 319 belong to Win Office 2000
                320 - 400 belong to Win Office 10
                401 - 500 belong to Mac Office 2001 and beyond
                501 - 600 belong to Win Office 10 and beyond
        Special dgcid ranges:
                0x1000 - 0x3FFF belong to Win Office 97
                0x4000 - 0x4FFF reserved for Win Office 10
                0x5000 - 0x5FFF reserved for Mac Office

        Try to be good about the names you choose for commands.  If at all
        possible, be like an existing command.  Here's some examples...

        msodgcidToggleFooMode
        msodgcidEnterFooMode
        msodgcidExitFooMode

        Remember that developers in other apps with other command sets will be
        reading these names.  If you mean "TextEffect" then call it "TextEffect"
        (or even better "WordArt" -- don't just call it "Text").

****************************************************************************/

typedef enum
        {
        msodgcidRangeNormalMin                   = 0x0000,
        msodgcidRangeNormalMax                   = 0x0FFF,
        msodgcidRangeShapeMin                    = 0x1000,
        msodgcidRangeShapeMax                    = 0x1FFF,
        msodgcidRangeChangeShapeMin              = 0x2000,
        msodgcidRangeChangeShapeMax              = 0x2FFF,
        msodgcidRangeUndoMin                     = 0x3000,
        msodgcidRangeUndoMax                     = 0x30FF,
        msodgcidRangeUndoHostMin                 = 0x3100,
        msodgcidRangeUndoHostMax                 = 0x31FF,
        msodgcidRangeDiagramMin                  = 0x4000,
        msodgcidRangeDiagramMax                  = 0x4FFF,

        msodgcidNil                              = 0,
        msodgcidNYI                              = 1,

        msodgcidNormalMin                        = 2, msodgcidMin = msodgcidNormalMin, /* TODO peteren: Vape msodgcidMin */

        msodgcidCut                              = 2,
        msodgcidCopy                             = 3,
        msodgcidPaste                            = 4,
        msodgcidClear                            = 5,
        msodgcidSelectAll                        = 6,
        msodgcidDuplicate                        = 7,
        msodgcidRepeat                           = 8,
        msodgcidPickUpFormat                     = 9,
        msodgcidApplyFormat                      = 10,
        msodgcidApplyFormatToDefaults            = 11,
        msodgcidBringToFront                     = 12,
        msodgcidSendToBack                       = 13,
        msodgcidBringForward                     = 14,
        msodgcidSendBackward                     = 15,
        msodgcidBringInFrontOfDocument           = 16,
        msodgcidSendBehindDocument               = 17,
        msodgcidGroup                            = 18,
        msodgcidUngroup                          = 19,
        msodgcidRegroup                          = 20,
        msodgcidBeginFreeform                    = 21,
        msodgcidEndFreeform                      = 22,
        msodgcidAddFreeform                      = 23,
        msodgcidAddPolygonPt                     = 24,
        msodgcidInsertPolygonPt                  = 25,
        msodgcidDeletePolygonPt                  = 26,
        msodgcidChangePolygonPt                  = 27,
        msodgcidCopyPolygonPt                    = 28,
        msodgcidClosePolygon                     = 29,
        msodgcidOpenPolygon                      = 30,
        msodgcidPolygonCreate                    = 31,
        msodgcidPolygonReshape                   = 32,
        msodgcidAutoVertex                       = 33,
        msodgcidSmoothVertex                     = 34,
        msodgcidStraightVertex                   = 35,
        msodgcidCornerVertex                     = 36,
        msodgcidStraightSegment                  = 37,
        msodgcidCurvedSegment                    = 38,
        msodgcidAddWrapPolygon                   = 39,
        msodgcidWrapPolygonReshape               = 40,
        msodgcidRotateLeft90                     = 41,
        msodgcidRotateRight90                    = 42,
        msodgcidFlipHorizontal                   = 43,
        msodgcidFlipVertical                     = 44,
        msodgcidAlignLeft                        = 45,
        msodgcidAlignCenterHorizontal            = 46,
        msodgcidAlignRight                       = 47,
        msodgcidAlignTop                         = 48,
        msodgcidAlignCenterVertical              = 49,
        msodgcidAlignBottom                      = 50,
        msodgcidAlignPageRelative                = 51,
        msodgcidDistributeHorizontal             = 52,
        msodgcidDistributeVertical               = 53,
        msodgcidDistributePageRelative           = 54,
        msodgcidNudgeLeft                        = 55,
        msodgcidNudgeRight                       = 56,
        msodgcidNudgeUp                          = 57,
        msodgcidNudgeDown                        = 58,
        msodgcidNudgeLeftOne                     = 59,
        msodgcidNudgeRightOne                    = 60,
        msodgcidNudgeUpOne                       = 61,
        msodgcidNudgeDownOne                     = 62,
        msodgcidToggleReshapeMode                = 63,
        msodgcidToggleRotateMode                 = 64,
        msodgcidToggleCropMode                   = 65,
        msodgcidToggleWrapPolygonMode            = 66, msodgcidToggleWrapPolygon = msodgcidToggleWrapPolygonMode, // FUTURE peteren: Remove uses of old name
        msodgcidMoreFillColor                    = 67,
        msodgcidFillEffect                       = 68,
        msodgcidMoreLineColor                    = 69,
        msodgcidMoreLineWidth                    = 70,
        msodgcidMoreArrow                        = 71,
        msodgcidTextEffectRotateCharacters       = 72,
        msodgcidTextEffectStretchToFill          = 73,
        msodgcidTextEffectSameHeight             = 74,
        msodgcidTextEffectAlignLeft              = 75,
        msodgcidTextEffectAlignCenter            = 76,
        msodgcidTextEffectAlignRight             = 77,
        msodgcidTextEffectAlignLetterJustify     = 78, msodgcidTextEffectAlignJustify = msodgcidTextEffectAlignLetterJustify, msodgcidTextEffFullJustify = msodgcidTextEffectAlignJustify,
        msodgcidUnused79                         = 79,
        msodgcidTextEffectAlignWordJustify       = 80, msodgcidTextEffectAlignByWord = msodgcidTextEffectAlignWordJustify,
        msodgcidTextEffectAlignStretchJustify    = 81, msodgcidTextEffectAlignStretch = msodgcidTextEffectAlignStretchJustify, msodgcidTextEffStretchText = msodgcidTextEffectAlignStretch,
        msodgcidTextEffectSpacingVeryTight       = 82,
        msodgcidTextEffectSpacingTight           = 83,
        msodgcidTextEffectSpacingNormal          = 84,
        msodgcidTextEffectSpacingLoose           = 85,
        msodgcidTextEffectSpacingVeryLoose       = 86,
        msodgcidTextEffectKernPairs              = 87,
        msodgcidTextEffectEditText               = 88,
        msodgcidPictureMoreContrast              = 89,
        msodgcidPictureLessContrast              = 90,
        msodgcidPictureMoreBrightness            = 91,
        msodgcidPictureLessBrightness            = 92,
        msodgcidPictureReset                     = 93,
        msodgcidPictureImageAutomatic            = 94,
        msodgcidPictureImageGrayscale            = 95,
        msodgcidPictureImageBlackWhite           = 96,
        msodgcidPictureImageWatermark            = 97,
        msodgcidPictureInLine                    = 98,
        msodgcidUnused99                         = 99,  msodgcidTextWrapSquare    = 99,
        msodgcidUnused100                        = 100, msodgcidTextWrapTight     = 100,
        msodgcidUnused101                        = 101, msodgcidTextWrapNone      = 101,
        msodgcidUnused102                        = 102, msodgcidEditTextWrapCurve = 102,
        msodgcidMoreShadow                       = 103,
        msodgcidMoreShadowColor                  = 104,
        msodgcidNudgeShadowUp                    = 105,
        msodgcidNudgeShadowDown                  = 106,
        msodgcidNudgeShadowLeft                  = 107,
        msodgcidNudgeShadowRight                 = 108,
        msodgcidMore3D                           = 109,
        msodgcidMore3DColor                      = 110,
        msodgcid3DToggle                         = 111,
        msodgcid3DTiltForward                    = 112,
        msodgcid3DTiltBackward                   = 113,
        msodgcid3DTiltLeft                       = 114,
        msodgcid3DTiltRight                      = 115,
        msodgcid3DDepth0                         = 116,
        msodgcid3DDepth1                         = 117,
        msodgcid3DDepth2                         = 118,
        msodgcid3DDepth3                         = 119,
        msodgcid3DDepth4                         = 120,
        msodgcid3DDepthInfinite                  = 121,
        msodgcid3DPerspective                    = 122,
        msodgcid3DParallel                       = 123,
        msodgcid3DLightingFlat                   = 124,
        msodgcid3DLightingNormal                 = 125,
        msodgcid3DLightingHarsh                  = 126,
        msodgcid3DSurfaceMatte                   = 127,
        msodgcid3DSurfacePlastic                 = 128,
        msodgcid3DSurfaceMetal                   = 129, msodgcid3DSurfacePolishedMetalRemoveMe = 129,
        msodgcid3DSurfaceWireFrame               = 130,
        msodgcid3DSurfaceDullMetalRemoveMe       = 131, // TODO peteren: Remove this?
        msodgcidSelectShape                      = 132,
        msodgcidSelectShapeAdd                   = 133,
        msodgcidSelectShapeRemove                = 134,
        msodgcidToolPointer                      = 135, /* Don't put this on command bars, use msodgcidTogglePointerMode instead.  TODO peteren: Make this command unrunable. */
        msodgcidToolMarquee                      = 136,
        msodgcidToolFormatPainter                = 137,
        msodgcidToolRotate                       = 138,
        msodgcidToolCrop                         = 139,
        msodgcidToolLine                         = 140,
        msodgcidToolArrow                        = 141,
        msodgcidToolDoubleArrow                  = 142,
        msodgcidToolArc                          = 143,
        msodgcidToolPolygon                      = 144,
        msodgcidToolFilledPolygon                = 145,
        msodgcidToolCurve                        = 146,
        msodgcidToolFreeform                     = 147,
        msodgcidToolFilledFreeform               = 148,
        msodgcidToolFreehand                     = 149,
        msodgcidToolPicture                      = 150,
        msodgcidToolOleObject                    = 151,
        msodgcidToolText                         = 152,
        msodgcidToolCalloutWordRightAngle        = 153,
        msodgcidToolCalloutWordOneSegment        = 154,
        msodgcidToolCalloutWordTwoSegment        = 155,
        msodgcidToolCalloutWordThreeSegment      = 156,
        msodgcidToolStraightConnector            = 157,
        msodgcidToolAngledConnector              = 158,
        msodgcidToolCurvedConnector              = 159,
        msodgcidToolDoEdit                       = 160,
        msodgcidSwatchFillColorNone              = 161,
        msodgcidSwatchFillColorStandard          = 162,
        msodgcidSwatchFillColorMRU               = 163,
        msodgcidSwatchLineColorNone              = 164,
        msodgcidSwatchLineColorStandard          = 165,
        msodgcidSwatchLineColorMRU               = 166,
        msodgcidSwatchNoShadow                   = 167,
        msodgcidSwatchShadowColorStandard        = 168,
        msodgcidSwatchShadowColorMRU             = 169,
        msodgcidSwatchNo3D                       = 170,
        msodgcidSwatch3DColorAutomatic           = 171,
        msodgcidSwatch3DColorStandard            = 172,
        msodgcidSwatch3DColorMRU                 = 173,
        msodgcidSwatchShadowPresets              = 174,
        msodgcidSwatchLineWidthPresets           = 175,
        msodgcidSwatchLineDashPresets            = 176,
        msodgcidSwatchArrowPresets               = 177,
        msodgcidSwatch3DPresets                  = 178,
        msodgcidSwatchTextEffectEnvelopes        = 179, msodgcidSwatchTextEffectPresets = msodgcidSwatchTextEffectEnvelopes,
        msodgcidSwatch3DDirection                = 180,
        msodgcidSwatch3DLightDirection           = 181,
        msodgcidSwatchDlgPatternFgColorNone      = 182,
        msodgcidSwatchDlgPatternFgColorStandard  = 183,
        msodgcidSwatchDlgPatternFgColorMRU       = 184,
        msodgcidSwatchDlgPatternBgColorNone      = 185,
        msodgcidSwatchDlgPatternBgColorStandard  = 186,
        msodgcidSwatchDlgPatternBgColorMRU       = 187,
        msodgcidSwatchDlgGradientFgColorNone     = 188,
        msodgcidSwatchDlgGradientFgColorStandard = 189,
        msodgcidSwatchDlgGradientFgColorMRU      = 190,
        msodgcidSwatchDlgGradientBgColorNone     = 191,
        msodgcidSwatchDlgGradientBgColorStandard = 192,
        msodgcidSwatchDlgGradientBgColorMRU      = 193,
        msodgcidSwatchDlgLineColorNone           = 194,
        msodgcidSwatchDlgLineColorStandard       = 195,
        msodgcidSwatchDlgLineColorMRU            = 196,
        msodgcidSwatchDlgColorNone               = 197,
        msodgcidSwatchDlgColorStandard           = 198,
        msodgcidSwatchDlgColorMRU                = 199,
        msodgcidSwatchDlgLineStyle               = 200,
        msodgcidSwatchDlgLineWidth               = 201,
        msodgcidSwatchDlgDashStyle               = 202,
        msodgcidSwatchDlgArrowLStyle             = 203,
        msodgcidSwatchDlgArrowLSize              = 204,
        msodgcidSwatchDlgArrowRStyle             = 205,
        msodgcidSwatchDlgArrowRSize              = 206,
        msodgcidMenuDlgLineColor                 = 207,
        msodgcidMenuDlgLineStyle                 = 208,
        msodgcidMenuDlgLineWidth                 = 209,
        msodgcidMenuDlgLineDash                  = 210,
        msodgcidMenuDlgLineConnector             = 211,
        msodgcidMenuDlgArrowLStyle               = 212,
        msodgcidMenuDlgArrowLSize                = 213,
        msodgcidMenuDlgArrowRStyle               = 214,
        msodgcidMenuDlgArrowRSize                = 215,
        msodgcidMenuDlgFillShadeColor1           = 216,
        msodgcidMenuDlgFillShadeColor2           = 217,
        msodgcidMenuDlgFillPatternColor1         = 218,
        msodgcidMenuDlgFillPatternColor2         = 219,
        msodgcidEditTextEffectSpacingCustom      = 220,
        msodgcidEdit3DEffectDepthCustom          = 221,
        msodgcidSplitMenuFillColor               = 222,
        msodgcidSplitMenuLineColor               = 223,
        msodgcidSplitMenuShadowColor             = 224,
        msodgcidSplitMenu3DColor                 = 225,
        msodgcidRerouteConnections               = 226,
        msodgcidStraightStyle                    = 227,
        msodgcidAngledStyle                      = 228,
        msodgcidCurvedStyle                      = 229,
        msodgcidToggleFill                       = 230,
        msodgcidToggleLine                       = 231,
        msodgcidToggleShadow                     = 232,
        msodgcidFormatObject                     = 233,
        msodgcidEditObject                       = 234,
        msodgcidEditPicture                      = 235,
        msodgcidFormatShape                      = 236,
        msodgcidShadeFill                        = 237,
        msodgcidColorFill                        = 238,
        msodgcidTextEffectEdit                   = 239,
        msodgcidTextEffectInsert                 = 240,
        msodgcidTextEffectToolbarToggle          = 241,
        msodgcidDisconnect                       = 242,
        msodgcidSaveGraphicAsClipArt             = 243,
        msodgcidRememberSelection                = 244,
        msodgcidMoreColor                        = 245,
        msodgcidDlgMoreColor                     = 246,
        msodgcidDlgFillEffect                    = 247,
        msodgcidDlgLinePattern                   = 248,
        msodgcidDlgMoreLineColor                 = 249,
        msodgcidStraightButton                   = 250,
        msodgcidAngledButton                     = 251,
        msodgcidCurvedButton                     = 252,
        msodgcidLinePatternFill                  = 253,
        msodgcidSetFillStyle                     = 254, // What does this command do? Does it really set _only_ the fill style or the whole fill? Wouldn't msodgcidSetFill be better?
        msodgcidSetLineStyle                     = 255,
        msodgcidSetGeoTextStyle                  = 256, // What's a GeoTextStyle?  Does this pick a preset text effect?  Why not msodgcidSetTextEffectPreset?
        msodgcidSetColorStyle                    = 257, // What's a ColorStyle?  I think this should be msodgcidSetFillColor.
        msodgcidSetFmtObj                        = 258, // Another hard to understand command name.
        msodgcidDlgMoreGradientColor1            = 259,
        msodgcidDlgMoreGradientColor2            = 260,
        msodgcidDlgMorePatternColor1             = 261,
        msodgcidDlgMorePatternColor2             = 262,
        msodgcidGeoTextInsert                    = 263,
        msodgcidUnselectAll                      = 264,
        msodgcidDlgInsertPicture                 = 265,
        msodgcidActivateText                     = 266,
        msodgcidToggleShadowOpacity              = 267,
        msodgcidExitReshapeMode                  = 268,
        msodgcidToolVerticalText                 = 269,
        msodgcidExitRotateMode                   = 270,
        msodgcidTogglePictureToolbar             = 271,
        msodgcidSetDefaults                      = 272,
        msodgcidResetAdjustHandle                = 273,
        msodgcidToolStraightArrowConnector       = 274,
        msodgcidToolAngledArrowConnector         = 275,
        msodgcidToolCurvedArrowConnector         = 276,
        msodgcidToolStraightDblArrowConnector    = 277,
        msodgcidToolAngledDblArrowConnector      = 278,
        msodgcidToolCurvedDblArrowConnector      = 279,
        msodgcidToolSetTransparentColor          = 280,
        msodgcidTextEffectGallery                = 281,
        msodgcidShowAutoShapesAndDrawingToolbars = 282,
        msodgcidSetTransparentColor              = 283,
        msodgcidUnused284                        = 284, msodgcidSplitNodeRemoveMe = 284,
        msodgcidDeleteSegment                    = 285,
        msodgcidMarquee                          = 286, /* Used by the marquee tool to select the objects within a rectangle */
        msodgcidContextMenu                      = 287,
        msodgcidUpdateBlips                      = 288,
        msodgcidFormatTextGallery                = 289,
        msodgcidTogglePointerMode                = 290,
        msodgcidExitPointerMode                  = 291,
        msodgcidSwatchLineStylePresets           = 292,
        msodgcidToolMarqueeIfNoHit               = 293, /* pure tool, cannot be executed, used only for hits on DGVs. */
        msodgcidSetTextFrame                     = 294,
        msodgcidFillEffectOnlyTexture            = 295,
        msodgcidFillEffectWithoutTransparency    = 296,
        msodgcidMoreColorsWithoutTransparency    = 297,
        msodgcidDlgSetBackgroundFill             = 298,
        msodgcidActivateTextNoToggle             = 299, // like msodgcidActivateText but does not toggle; this is used internally when clicking on text shapes

        // Last valid command dgcid for Windows Office97.
        msodgcidNormalMaxWinOffice97             = 300,

        // Reserved id block for Mac Office 8 is from 300 thru 309 (inclusive)
        msodgcidStartMacOffice8                  = 300, msodgcidInsertMovie = 300,
        msodgcidMoviePlay                        = 301,
        msodgcidMovieShowController              = 302,
        msodgcidMovieLoop                        = 303,
        msodgcidMovieSetPosterFrame              = 304,
        msodgcidMacReserved305                   = 305,
        msodgcidMacReserved306                   = 306,
        msodgcidMacReserved307                   = 307,
        msodgcidMacReserved308                   = 308,
        msodgcidMacReserved309                   = 309,

        // Reserved id block for Win Office 9 ids is from 310 thru 319 (inclusive)
        msodgcidStartOffice9                     = 310, 
        msodgcidInsertScript                     = 310,
        msodgcidEditScript                       = 311,
        msodgcidRemoveAllScripts                 = 312,
        msodgcidRunCag                           = 313,
        msodgcidRunCagForPictures                = 314,
        msodgcidRunCagForMovies                  = 315,
        msodgcidRunCagForSounds                  = 316,
        msodgcidRunCagForShapes                  = 317,
        msodgcidInsertBlip                       = 318,       // specify pib in MSODGCB
        msodgcidMultiSelect                      = 319,       // Accessibility add-in: multi-selection

        // Reserved id block for Win office 10 is from 320 thru 400 (inclusive)
        msodgcidStartOffice10                    = 320, 
        msodgcidInsertDrawingCanvas              = 320,
        msodgcidInsertOrgChart                   = 321,
        msodgcidInsertRadialChart                = 322,
        msodgcidInsertCycleChart                 = 323,
        msodgcidInsertStackedChart               = 324,
        msodgcidInsertBullsEyeChart              = 325,
        msodgcidInsertVennDiagram                = 326,
        msodgcidOrgChartInsertSubordinate        = 327,
        msodgcidOrgChartInsertCoworker           = 328,
        msodgcidOrgChartInsertAssistant          = 329,
        msodgcidOrgChartDeleteNode               = 330,
        msodgcidOrgChartLayoutHorizontal1        = 331,
        msodgcidOrgChartLayoutHorizontal2        = 332,
        msodgcidOrgChartLayoutVertical1          = 333,
        msodgcidOrgChartLayoutVertical2          = 334,
        msodgcidDiagramStyle                     = 335,
        msodgcidDiagramApplyStyle                = 336,
        msodgcidConvertToVenn                    = 337,
        msodgcidConvertToRadial                  = 338,
        msodgcidConvertToCycle                   = 339,
        msodgcidConvertToBullsEye                = 340,
        msodgcidConvertToPyramid                 = 341,
        msodgcidMoveDiagramShapeUp               = 342,
        msodgcidMoveDiagramShapeDown             = 343,
        msodgcidInsertDiagramShape               = 344,
        msodgcidInsertDiagram                    = 345,
        msodgcidCreateDiagram                    = 346,

        msodgcidCanvasFit                        = 347,
        msodgcidCanvasResize                     = 348,
        msodgcidToggleCanvasToolbar              = 349,
        msodgcidToolResizeCanvas                 = 350,
        msodgcidCanvasExpand                     = 351,

        msodgcidMenuDlgFillTintBaseColor         = 352,
        msodgcidDlgMoreTintBaseColor             = 353,
        msodgcidDlgMoreGradientEffectsColor1     = 354,
        msodgcidDlgMoreGradientEffectsColor2     = 355,
        msodgcidDlgMorePatternEffectsColor1      = 356,
        msodgcidDlgMorePatternEffectsColor2      = 357,
        msodgcidShadowColorTint                  = 358,
        msodgcid3DColorTint                      = 359,

        msodgcidSwatchFillColorExtMRU            = 360,
        msodgcidSwatchLineColorExtMRU            = 361,
        msodgcidSwatchShadowColorExtMRU          = 362,
        msodgcidSwatch3DColorExtMRU              = 363,

        msodgcidSwatchDlgTintBaseColorStandard   = 364,
        msodgcidSwatchDlgPatternFgColorExtMRU    = 365,
        msodgcidSwatchDlgPatternBgColorExtMRU    = 366,
        msodgcidSwatchDlgGradientFgColorExtMRU   = 367,
        msodgcidSwatchDlgGradientBgColorExtMRU   = 368,
        msodgcidSwatchDlgTintBaseColorExtMRU     = 369,
        msodgcidSwatchDlgLineColorExtMRU         = 370,
        msodgcidSwatchDlgColorExtMRU             = 371,

        msodgcidSplitMenuFillColorExt            = 372,
        msodgcidSplitMenuLineColorExt            = 373,
        msodgcidSplitMenuShadowColorExt          = 374,
        msodgcidSplitMenu3DColorExt              = 375,
        msodgcidSetColorExtStyle                 = 376,

        msodgcidAlignCanvasRelative              = 377,

        msodgcidOrgChartSelectLevel              = 378,
        msodgcidOrgChartSelectBranch             = 379,
        msodgcidOrgChartSelectAllAssistants      = 380,
        msodgcidOrgChartSelectAllConnectors      = 381,
        msodgcidDiagramDeleteNode                = 382,
        msodgcidDiagramReverse                   = 383,
        msodgcidDiagramAutoLayout                = 384,
        msodgcidOrgChartAutoLayout               = 385,

        msodgcidMoreFillColorExt                 = 386,
        msodgcidMoreLineColorExt                 = 387,
        msodgcidMoreShadowColorExt               = 388,
        msodgcidMore3DColorExt                   = 389,

        msodgcidOptimizePict                     = 390,
        msodgcidOptimizePictDialog               = 391,

        msodgcidOrgChartLayoutAssistant          = 392,
        msodgcidConvertOrgChart                  = 393, // OBSOLETE
        msodgcidSelectShapes                     = 394,
        msodgcidPictureEditInPhotoDrawUNUSED     = 395, // OBSOLETE
        msodgcidPictureInsertFromPhotoDrawUNUSED = 396, // OBSOLETE
        msodgcidDiagramFit                       = 397,
        msodgcidDiagramResize                    = 398,
        msodgcidDiagramExpand                    = 399,
        msodgcidOrgChartFit                      = 400,

        // Reserved id block for Mac office 2001 is from 401 thru 500 (inclusive)
        msodgcidMacReserved401                   = 401,
        msodgcidMacReserved402                   = 402,
        msodgcidMacReserved403                   = 403,
        msodgcidMacReserved404                   = 404,
        msodgcidMacReserved405                   = 405,
        msodgcidMacReserved406                   = 406,
        msodgcidMacReserved407                   = 407,
        msodgcidMacReserved408                   = 408,
        msodgcidMacReserved409                   = 409,
        msodgcidMacReserved410                   = 410,
        msodgcidMacReserved411                   = 411,
        msodgcidMacReserved412                   = 412,
        msodgcidMacReserved413                   = 413,
        msodgcidMacReserved414                   = 414,
        msodgcidMacReserved415                   = 415,
        msodgcidMacReserved416                   = 416,
        msodgcidMacReserved417                   = 417,
        msodgcidMacReserved418                   = 418,
        msodgcidMacReserved419                   = 419,
        msodgcidMacReserved420                   = 420,
        msodgcidMacReserved421                   = 421,
        msodgcidMacReserved422                   = 422,
        msodgcidMacReserved423                   = 423,
        msodgcidMacReserved424                   = 424,
        msodgcidMacReserved425                   = 425,
        msodgcidMacReserved426                   = 426,
        msodgcidMacReserved427                   = 427,
        msodgcidMacReserved428                   = 428,
        msodgcidMacReserved429                   = 429,
        msodgcidMacReserved430                   = 430,
        msodgcidMacReserved431                   = 431,
        msodgcidMacReserved432                   = 432,
        msodgcidMacReserved433                   = 433,
        msodgcidMacReserved434                   = 434,
        msodgcidMacReserved435                   = 435,
        msodgcidMacReserved436                   = 436,
        msodgcidMacReserved437                   = 437,
        msodgcidMacReserved438                   = 438,
        msodgcidMacReserved439                   = 439,
        msodgcidMacReserved440                   = 440,
        msodgcidMacReserved441                   = 441,
        msodgcidMacReserved442                   = 442,
        msodgcidMacReserved443                   = 443,
        msodgcidMacReserved444                   = 444,
        msodgcidMacReserved445                   = 445,
        msodgcidMacReserved446                   = 446,
        msodgcidMacReserved447                   = 447,
        msodgcidMacReserved448                   = 448,
        msodgcidMacReserved449                   = 449,
        msodgcidMacReserved450                   = 450,
        msodgcidMacReserved451                   = 451,
        msodgcidMacReserved452                   = 452,
        msodgcidMacReserved453                   = 453,
        msodgcidMacReserved454                   = 454,
        msodgcidMacReserved455                   = 455,
        msodgcidMacReserved456                   = 456,
        msodgcidMacReserved457                   = 457,
        msodgcidMacReserved458                   = 458,
        msodgcidMacReserved459                   = 459,
        msodgcidMacReserved460                   = 460,
        msodgcidMacReserved461                   = 461,
        msodgcidMacReserved462                   = 462,
        msodgcidMacReserved463                   = 463,
        msodgcidMacReserved464                   = 464,
        msodgcidMacReserved465                   = 465,
        msodgcidMacReserved466                   = 466,
        msodgcidMacReserved467                   = 467,
        msodgcidMacReserved468                   = 468,
        msodgcidMacReserved469                   = 469,
        msodgcidMacReserved470                   = 470,
        msodgcidMacReserved471                   = 471,
        msodgcidMacReserved472                   = 472,
        msodgcidMacReserved473                   = 473,
        msodgcidMacReserved474                   = 474,
        msodgcidMacReserved475                   = 475,
        msodgcidMacReserved476                   = 476,
        msodgcidMacReserved477                   = 477,
        msodgcidMacReserved478                   = 478,
        msodgcidMacReserved479                   = 479,
        msodgcidMacReserved480                   = 480,
        msodgcidMacReserved481                   = 481,
        msodgcidMacReserved482                   = 482,
        msodgcidMacReserved483                   = 483,
        msodgcidMacReserved484                   = 484,
        msodgcidMacReserved485                   = 485,
        msodgcidMacReserved486                   = 486,
        msodgcidMacReserved487                   = 487,
        msodgcidMacReserved488                   = 488,
        msodgcidMacReserved489                   = 489,
        msodgcidMacReserved490                   = 490,
        msodgcidMacReserved491                   = 491,
        msodgcidMacReserved492                   = 492,
        msodgcidMacReserved493                   = 493,
        msodgcidMacReserved494                   = 494,
        msodgcidMacReserved495                   = 495,
        msodgcidMacReserved496                   = 496,
        msodgcidMacReserved497                   = 497,
        msodgcidMacReserved498                   = 498,
        msodgcidMacReserved499                   = 499,
        msodgcidMacReserved500                   = 500,

        // Reserved id block for Win Office 10 and beyond is from 501 thru 600 (inclusive)
        msodgcidOrgChartResize                   = 501,
        msodgcidOrgChartExpand                   = 502,
        msodgcidOrgChartStyle                    = 503,
        msodgciSplitMenuOrgChartInsertShape      = 504,
        msodgcidDiagramAutoFormat                = 505,
        msodgcidOrgChartAutoFormat               = 506,
        msodgcidCanvasScale                      = 507,
        msodgcidOrgChartScale                    = 508,
        msodgcidDiagramScale                     = 509,
        msodgcidConvertDrawOLE                   = 510, // OBSOLETE
        msodgcidAlignOrgChartRelative            = 511,
        msodgcidAlignDiagramRelative             = 512,

        // Office 11 command starts here
        msodgcidToggleInkMode                    = 513,
        msodgcidToggleInkEraseMode               = 514,
        msodgcidToolDoInkEdit                    = 515,
        msodgcidSplitMenuInkColor                = 516,
        msodgcidInkSessionEdit                   = 517,
        msodgcidSplitMenuInkAnntColor            = 518,
        msodgcidInkAnnotation                    = 519,
        msodgcidInkStyle1                        = 520,
        msodgcidInkStyle2                        = 521,
        msodgcidInkStyle3                        = 522,
        msodgcidInkStyle4                        = 523,
        msodgcidInkStyle5                        = 524,
        msodgcidInkStyle6                        = 525,
        msodgcidInkStyle7                        = 526,
        msodgcidInkStyle8                        = 527,
        msodgcidInkStyle9                        = 528,
        msodgcidInkAnnotationStyle1              = 529,
        msodgcidInkAnnotationStyle2              = 530,
        msodgcidInkAnnotationStyle3              = 531,
        msodgcidInkAnnotationStyle4              = 532,
        msodgcidInkAnnotationStyle5              = 533,
        msodgcidInkAnnotationStyle6              = 534,
        msodgcidInkAnnotationStyle7              = 535,
        msodgcidInkAnnotationStyle8              = 536,
        msodgcidInkAnnotationStyle9              = 537,
        msodgcidToggleInkToolbar                 = 538,
        msodgcidToggleInkAnnotationToolbar       = 539,
        msodgcidClearAllInkAnnotations           = 540,
        msodgcidCopyInkAsText                    = 541,
        msodgcidInkDrawing                       = 542,
        msodgcidInkSplitMenu                     = 543,
        msodgcidExitInkMode                      = 544,
        msodgcidInkEraser                        = 545,
        msodgcidInkAnnotationEraser              = 546,
        msodgcidExitInkAnnotationMode            = 547,
        msodgcidInkLabel1                        = 548,
        msodgcidInkLabel2                        = 549,
        msodgcidInkLabel3                        = 550,
        msodgcidOrgChartFitText                  = 551,

        msodgcidNormalMax, msodgcidMax/*old name*/ = msodgcidNormalMax, msodgcidNormalLast = msodgcidMax - 1, msodgcidLast/*old name*/ = msodgcidNormalLast,

        msodgcidShapeMin                        = 0x1000 + 1,

        msodgcidShapeRectangle                  = 0x1000 + 1,
        msodgcidShapeRoundRectangle             = 0x1000 + 2,
        msodgcidShapeEllipse                    = 0x1000 + 3,
        msodgcidShapeDiamond                    = 0x1000 + 4,
        msodgcidShapeIsocelesTriangle           = 0x1000 + 5,
        msodgcidShapeRightTriangle              = 0x1000 + 6,
        msodgcidShapeParallelogram              = 0x1000 + 7,
        msodgcidShapeTrapezoid                  = 0x1000 + 8,
        msodgcidShapeHexagon                    = 0x1000 + 9,
        msodgcidShapeOctagon                    = 0x1000 + 10,
        msodgcidShapePlus                       = 0x1000 + 11,
        msodgcidShapeStar                       = 0x1000 + 12,
        msodgcidShapeArrow                      = 0x1000 + 13,
        msodgcidShapeThickArrow                 = 0x1000 + 14,
        msodgcidShapeHomePlate                  = 0x1000 + 15,
        msodgcidShapeCube                       = 0x1000 + 16,
        msodgcidShapeBalloon                    = 0x1000 + 17,
        msodgcidShapeSeal                       = 0x1000 + 18,
        msodgcidShapeArc                        = 0x1000 + 19,
        msodgcidShapeLine                       = 0x1000 + 20,
        msodgcidShapePlaque                     = 0x1000 + 21,
        msodgcidShapeCan                        = 0x1000 + 22,
        msodgcidShapeDonut                      = 0x1000 + 23,
        msodgcidShapeTextSimple                 = 0x1000 + 24,
        msodgcidShapeTextOctagon                = 0x1000 + 25,
        msodgcidShapeTextHexagon                = 0x1000 + 26,
        msodgcidShapeTextCurve                  = 0x1000 + 27,
        msodgcidShapeTextWave                   = 0x1000 + 28,
        msodgcidShapeTextRing                   = 0x1000 + 29,
        msodgcidShapeTextOnCurve                = 0x1000 + 30,
        msodgcidShapeTextOnRing                 = 0x1000 + 31,
        msodgcidShapeStraightConnector1         = 0x1000 + 32,
        msodgcidShapeBentConnector2             = 0x1000 + 33,
        msodgcidShapeBentConnector3             = 0x1000 + 34,
        msodgcidShapeBentConnector4             = 0x1000 + 35,
        msodgcidShapeBentConnector5             = 0x1000 + 36,
        msodgcidShapeCurvedConnector2           = 0x1000 + 37,
        msodgcidShapeCurvedConnector3           = 0x1000 + 38,
        msodgcidShapeCurvedConnector4           = 0x1000 + 39,
        msodgcidShapeCurvedConnector5           = 0x1000 + 40,
        msodgcidShapeCallout1                   = 0x1000 + 41,
        msodgcidShapeCallout2                   = 0x1000 + 42,
        msodgcidShapeCallout3                   = 0x1000 + 43,
        msodgcidShapeAccentCallout1             = 0x1000 + 44,
        msodgcidShapeAccentCallout2             = 0x1000 + 45,
        msodgcidShapeAccentCallout3             = 0x1000 + 46,
        msodgcidShapeBorderCallout1             = 0x1000 + 47,
        msodgcidShapeBorderCallout2             = 0x1000 + 48,
        msodgcidShapeBorderCallout3             = 0x1000 + 49,
        msodgcidShapeAccentBorderCallout1       = 0x1000 + 50,
        msodgcidShapeAccentBorderCallout2       = 0x1000 + 51,
        msodgcidShapeAccentBorderCallout3       = 0x1000 + 52,
        msodgcidShapeRibbon                     = 0x1000 + 53,
        msodgcidShapeRibbon2                    = 0x1000 + 54,
        msodgcidShapeChevron                    = 0x1000 + 55,
        msodgcidShapePentagon                   = 0x1000 + 56,
        msodgcidShapeNoSmoking                  = 0x1000 + 57,
        msodgcidShapeSeal8                      = 0x1000 + 58,
        msodgcidShapeSeal16                     = 0x1000 + 59,
        msodgcidShapeSeal32                     = 0x1000 + 60,
        msodgcidShapeWedgeRectCallout           = 0x1000 + 61,
        msodgcidShapeWedgeRRectCallout          = 0x1000 + 62,
        msodgcidShapeWedgeEllipseCallout        = 0x1000 + 63,
        msodgcidShapeWave                       = 0x1000 + 64,
        msodgcidShapeFoldedCorner               = 0x1000 + 65,
        msodgcidShapeLeftArrow                  = 0x1000 + 66,
        msodgcidShapeDownArrow                  = 0x1000 + 67,
        msodgcidShapeUpArrow                    = 0x1000 + 68,
        msodgcidShapeLeftRightArrow             = 0x1000 + 69,
        msodgcidShapeUpDownArrow                = 0x1000 + 70,
        msodgcidShapeIrregularSeal1             = 0x1000 + 71,
        msodgcidShapeIrregularSeal2             = 0x1000 + 72,
        msodgcidShapeLightningBolt              = 0x1000 + 73,
        msodgcidShapeHeart                      = 0x1000 + 74,
        msodgcidShapePictureFrame               = 0x1000 + 75,
        msodgcidShapeQuadArrow                  = 0x1000 + 76,
        msodgcidShapeLeftArrowCallout           = 0x1000 + 77,
        msodgcidShapeRightArrowCallout          = 0x1000 + 78,
        msodgcidShapeUpArrowCallout             = 0x1000 + 79,
        msodgcidShapeDownArrowCallout           = 0x1000 + 80,
        msodgcidShapeLeftRightArrowCallout      = 0x1000 + 81,
        msodgcidShapeUpDownArrowCallout         = 0x1000 + 82,
        msodgcidShapeQuadArrowCallout           = 0x1000 + 83,
        msodgcidShapeBevel                      = 0x1000 + 84,
        msodgcidShapeLeftBracket                = 0x1000 + 85,
        msodgcidShapeRightBracket               = 0x1000 + 86,
        msodgcidShapeLeftBrace                  = 0x1000 + 87,
        msodgcidShapeRightBrace                 = 0x1000 + 88,
        msodgcidShapeLeftUpArrow                = 0x1000 + 89,
        msodgcidShapeBentUpArrow                = 0x1000 + 90,
        msodgcidShapeBentArrow                  = 0x1000 + 91,
        msodgcidShapeSeal24                     = 0x1000 + 92,
        msodgcidShapeStripedRightArrow          = 0x1000 + 93,
        msodgcidShapeNotchedRightArrow          = 0x1000 + 94,
        msodgcidShapeBlockArc                   = 0x1000 + 95,
        msodgcidShapeSmileyFace                 = 0x1000 + 96,
        msodgcidShapeVerticalScroll             = 0x1000 + 97,
        msodgcidShapeHorizontalScroll           = 0x1000 + 98,
        msodgcidShapeCircularArrow              = 0x1000 + 99,
        msodgcidShapeNotchedCircularArrow       = 0x1000 + 100,
        msodgcidShapeUturnArrow                 = 0x1000 + 101,
        msodgcidShapeCurvedRightArrow           = 0x1000 + 102,
        msodgcidShapeCurvedLeftArrow            = 0x1000 + 103,
        msodgcidShapeCurvedUpArrow              = 0x1000 + 104,
        msodgcidShapeCurvedDownArrow            = 0x1000 + 105,
        msodgcidShapeCloudCallout               = 0x1000 + 106,
        msodgcidShapeEllipseRibbon              = 0x1000 + 107,
        msodgcidShapeEllipseRibbon2             = 0x1000 + 108,
        msodgcidShapeFlowChartProcess           = 0x1000 + 109,
        msodgcidShapeFlowChartDecision          = 0x1000 + 110,
        msodgcidShapeFlowChartInputOutput       = 0x1000 + 111,
        msodgcidShapeFlowChartPredefinedProcess = 0x1000 + 112,
        msodgcidShapeFlowChartInternalStorage   = 0x1000 + 113,
        msodgcidShapeFlowChartDocument          = 0x1000 + 114,
        msodgcidShapeFlowChartMultidocument     = 0x1000 + 115,
        msodgcidShapeFlowChartTerminator        = 0x1000 + 116,
        msodgcidShapeFlowChartPreparation       = 0x1000 + 117,
        msodgcidShapeFlowChartManualInput       = 0x1000 + 118,
        msodgcidShapeFlowChartManualOperation   = 0x1000 + 119,
        msodgcidShapeFlowChartConnector         = 0x1000 + 120,
        msodgcidShapeFlowChartPunchedCard       = 0x1000 + 121,
        msodgcidShapeFlowChartPunchedTape       = 0x1000 + 122,
        msodgcidShapeFlowChartSummingJunction   = 0x1000 + 123,
        msodgcidShapeFlowChartOr                = 0x1000 + 124,
        msodgcidShapeFlowChartCollate           = 0x1000 + 125,
        msodgcidShapeFlowChartSort              = 0x1000 + 126,
        msodgcidShapeFlowChartExtract           = 0x1000 + 127,
        msodgcidShapeFlowChartMerge             = 0x1000 + 128,
        msodgcidShapeFlowChartOfflineStorage    = 0x1000 + 129,
        msodgcidShapeFlowChartOnlineStorage     = 0x1000 + 130,
        msodgcidShapeFlowChartMagneticTape      = 0x1000 + 131,
        msodgcidShapeFlowChartMagneticDisk      = 0x1000 + 132,
        msodgcidShapeFlowChartMagneticDrum      = 0x1000 + 133,
        msodgcidShapeFlowChartDisplay           = 0x1000 + 134,
        msodgcidShapeFlowChartDelay             = 0x1000 + 135,
        msodgcidShapeTextPlainText              = 0x1000 + 136,
        msodgcidShapeTextStop                   = 0x1000 + 137,
        msodgcidShapeTextTriangle               = 0x1000 + 138,
        msodgcidShapeTextTriangleInverted       = 0x1000 + 139,
        msodgcidShapeTextChevron                = 0x1000 + 140,
        msodgcidShapeTextChevronInverted        = 0x1000 + 141,
        msodgcidShapeTextRingInside             = 0x1000 + 142,
        msodgcidShapeTextRingOutside            = 0x1000 + 143,
        msodgcidShapeTextArchUpCurve            = 0x1000 + 144,
        msodgcidShapeTextArchDownCurve          = 0x1000 + 145,
        msodgcidShapeTextCircleCurve            = 0x1000 + 146,
        msodgcidShapeTextButtonCurve            = 0x1000 + 147,
        msodgcidShapeTextArchUpPour             = 0x1000 + 148,
        msodgcidShapeTextArchDownPour           = 0x1000 + 149,
        msodgcidShapeTextCirclePour             = 0x1000 + 150,
        msodgcidShapeTextButtonPour             = 0x1000 + 151,
        msodgcidShapeTextCurveUp                = 0x1000 + 152,
        msodgcidShapeTextCurveDown              = 0x1000 + 153,
        msodgcidShapeTextCascadeUp              = 0x1000 + 154,
        msodgcidShapeTextCascadeDown            = 0x1000 + 155,
        msodgcidShapeTextWave1                  = 0x1000 + 156,
        msodgcidShapeTextWave2                  = 0x1000 + 157,
        msodgcidShapeTextWave3                  = 0x1000 + 158,
        msodgcidShapeTextWave4                  = 0x1000 + 159,
        msodgcidShapeTextInflate                = 0x1000 + 160,
        msodgcidShapeTextDeflate                = 0x1000 + 161,
        msodgcidShapeTextInflateBottom          = 0x1000 + 162,
        msodgcidShapeTextDeflateBottom          = 0x1000 + 163,
        msodgcidShapeTextInflateTop             = 0x1000 + 164,
        msodgcidShapeTextDeflateTop             = 0x1000 + 165,
        msodgcidShapeTextDeflateInflate         = 0x1000 + 166,
        msodgcidShapeTextDeflateInflateDeflate  = 0x1000 + 167,
        msodgcidShapeTextFadeRight              = 0x1000 + 168,
        msodgcidShapeTextFadeLeft               = 0x1000 + 169,
        msodgcidShapeTextFadeUp                 = 0x1000 + 170,
        msodgcidShapeTextFadeDown               = 0x1000 + 171,
        msodgcidShapeTextSlantUp                = 0x1000 + 172,
        msodgcidShapeTextSlantDown              = 0x1000 + 173,
        msodgcidShapeTextCanUp                  = 0x1000 + 174,
        msodgcidShapeTextCanDown                = 0x1000 + 175,
        msodgcidShapeFlowChartAlternateProcess  = 0x1000 + 176,
        msodgcidShapeFlowChartOffpageConnector  = 0x1000 + 177,
        msodgcidShapeCallout90                  = 0x1000 + 178,
        msodgcidShapeAccentCallout90            = 0x1000 + 179,
        msodgcidShapeBorderCallout90            = 0x1000 + 180,
        msodgcidShapeAccentBorderCallout90      = 0x1000 + 181,
        msodgcidShapeLeftRightUpArrow           = 0x1000 + 182,
        msodgcidShapeSun                        = 0x1000 + 183,
        msodgcidShapeMoon                       = 0x1000 + 184,
        msodgcidShapeBracketPair                = 0x1000 + 185,
        msodgcidShapeBracePair                  = 0x1000 + 186,
        msodgcidShapeSeal4                      = 0x1000 + 187,
        msodgcidShapeDoubleWave                 = 0x1000 + 188,
        msodgcidShapeActionButtonBlank          = 0x1000 + 189,
        msodgcidShapeActionButtonHome           = 0x1000 + 190,
        msodgcidShapeActionButtonHelp           = 0x1000 + 191,
        msodgcidShapeActionButtonInformation    = 0x1000 + 192,
        msodgcidShapeActionButtonForwardNext    = 0x1000 + 193,
        msodgcidShapeActionButtonBackPrevious   = 0x1000 + 194,
        msodgcidShapeActionButtonEnd            = 0x1000 + 195,
        msodgcidShapeActionButtonBeginning      = 0x1000 + 196,
        msodgcidShapeActionButtonReturn         = 0x1000 + 197,
        msodgcidShapeActionButtonDocument       = 0x1000 + 198,
        msodgcidShapeActionButtonSound          = 0x1000 + 199,
        msodgcidShapeActionButtonMovie          = 0x1000 + 200,
        msodgcidShapeHostControl                = 0x1000 + 201,
        msodgcidShapeTextBox                    = 0x1000 + 202,

        // Add new shapes here.

        // Last valid shape value for Windows Office97.
        msodgcidShapeMaxWinOffice97             = 0x1000 + 203,

        msodgcidShapeMax = msodgcidShapeMaxWinOffice97,
        msodgcidShapeLast = msodgcidShapeMax - 1,

        msodgcidChangeShapeMin                        = 0x2000 + 1,

        msodgcidChangeShapeRectangle                  = 0x2000 + 1,
        msodgcidChangeShapeRoundRectangle             = 0x2000 + 2,
        msodgcidChangeShapeEllipse                    = 0x2000 + 3,
        msodgcidChangeShapeDiamond                    = 0x2000 + 4,
        msodgcidChangeShapeIsocelesTriangle           = 0x2000 + 5,
        msodgcidChangeShapeRightTriangle              = 0x2000 + 6,
        msodgcidChangeShapeParallelogram              = 0x2000 + 7,
        msodgcidChangeShapeTrapezoid                  = 0x2000 + 8,
        msodgcidChangeShapeHexagon                    = 0x2000 + 9,
        msodgcidChangeShapeOctagon                    = 0x2000 + 10,
        msodgcidChangeShapePlus                       = 0x2000 + 11,
        msodgcidChangeShapeStar                       = 0x2000 + 12,
        msodgcidChangeShapeArrow                      = 0x2000 + 13,
        msodgcidChangeShapeThickArrow                 = 0x2000 + 14,
        msodgcidChangeShapeHomePlate                  = 0x2000 + 15,
        msodgcidChangeShapeCube                       = 0x2000 + 16,
        msodgcidChangeShapeBalloon                    = 0x2000 + 17,
        msodgcidChangeShapeSeal                       = 0x2000 + 18,
        msodgcidChangeShapeArc                        = 0x2000 + 19,
        msodgcidChangeShapeLine                       = 0x2000 + 20,
        msodgcidChangeShapePlaque                     = 0x2000 + 21,
        msodgcidChangeShapeCan                        = 0x2000 + 22,
        msodgcidChangeShapeDonut                      = 0x2000 + 23,
        msodgcidChangeShapeTextSimple                 = 0x2000 + 24,
        msodgcidChangeShapeTextOctagon                = 0x2000 + 25,
        msodgcidChangeShapeTextHexagon                = 0x2000 + 26,
        msodgcidChangeShapeTextCurve                  = 0x2000 + 27,
        msodgcidChangeShapeTextWave                   = 0x2000 + 28,
        msodgcidChangeShapeTextRing                   = 0x2000 + 29,
        msodgcidChangeShapeTextOnCurve                = 0x2000 + 30,
        msodgcidChangeShapeTextOnRing                 = 0x2000 + 31,
        msodgcidChangeShapeStraightConnector1         = 0x2000 + 32,
        msodgcidChangeShapeBentConnector2             = 0x2000 + 33,
        msodgcidChangeShapeBentConnector3             = 0x2000 + 34,
        msodgcidChangeShapeBentConnector4             = 0x2000 + 35,
        msodgcidChangeShapeBentConnector5             = 0x2000 + 36,
        msodgcidChangeShapeCurvedConnector2           = 0x2000 + 37,
        msodgcidChangeShapeCurvedConnector3           = 0x2000 + 38,
        msodgcidChangeShapeCurvedConnector4           = 0x2000 + 39,
        msodgcidChangeShapeCurvedConnector5           = 0x2000 + 40,
        msodgcidChangeShapeCallout1                   = 0x2000 + 41,
        msodgcidChangeShapeCallout2                   = 0x2000 + 42,
        msodgcidChangeShapeCallout3                   = 0x2000 + 43,
        msodgcidChangeShapeAccentCallout1             = 0x2000 + 44,
        msodgcidChangeShapeAccentCallout2             = 0x2000 + 45,
        msodgcidChangeShapeAccentCallout3             = 0x2000 + 46,
        msodgcidChangeShapeBorderCallout1             = 0x2000 + 47,
        msodgcidChangeShapeBorderCallout2             = 0x2000 + 48,
        msodgcidChangeShapeBorderCallout3             = 0x2000 + 49,
        msodgcidChangeShapeAccentBorderCallout1       = 0x2000 + 50,
        msodgcidChangeShapeAccentBorderCallout2       = 0x2000 + 51,
        msodgcidChangeShapeAccentBorderCallout3       = 0x2000 + 52,
        msodgcidChangeShapeRibbon                     = 0x2000 + 53,
        msodgcidChangeShapeRibbon2                    = 0x2000 + 54,
        msodgcidChangeShapeChevron                    = 0x2000 + 55,
        msodgcidChangeShapePentagon                   = 0x2000 + 56,
        msodgcidChangeShapeNoSmoking                  = 0x2000 + 57,
        msodgcidChangeShapeSeal8                      = 0x2000 + 58,
        msodgcidChangeShapeSeal16                     = 0x2000 + 59,
        msodgcidChangeShapeSeal32                     = 0x2000 + 60,
        msodgcidChangeShapeWedgeRectCallout           = 0x2000 + 61,
        msodgcidChangeShapeWedgeRRectCallout          = 0x2000 + 62,
        msodgcidChangeShapeWedgeEllipseCallout        = 0x2000 + 63,
        msodgcidChangeShapeWave                       = 0x2000 + 64,
        msodgcidChangeShapeFoldedCorner               = 0x2000 + 65,
        msodgcidChangeShapeLeftArrow                  = 0x2000 + 66,
        msodgcidChangeShapeDownArrow                  = 0x2000 + 67,
        msodgcidChangeShapeUpArrow                    = 0x2000 + 68,
        msodgcidChangeShapeLeftRightArrow             = 0x2000 + 69,
        msodgcidChangeShapeUpDownArrow                = 0x2000 + 70,
        msodgcidChangeShapeIrregularSeal1             = 0x2000 + 71,
        msodgcidChangeShapeIrregularSeal2             = 0x2000 + 72,
        msodgcidChangeShapeLightningBolt              = 0x2000 + 73,
        msodgcidChangeShapeHeart                      = 0x2000 + 74,
        msodgcidChangeShapePictureFrame               = 0x2000 + 75,
        msodgcidChangeShapeQuadArrow                  = 0x2000 + 76,
        msodgcidChangeShapeLeftArrowCallout           = 0x2000 + 77,
        msodgcidChangeShapeRightArrowCallout          = 0x2000 + 78,
        msodgcidChangeShapeUpArrowCallout             = 0x2000 + 79,
        msodgcidChangeShapeDownArrowCallout           = 0x2000 + 80,
        msodgcidChangeShapeLeftRightArrowCallout      = 0x2000 + 81,
        msodgcidChangeShapeUpDownArrowCallout         = 0x2000 + 82,
        msodgcidChangeShapeQuadArrowCallout           = 0x2000 + 83,
        msodgcidChangeShapeBevel                      = 0x2000 + 84,
        msodgcidChangeShapeLeftBracket                = 0x2000 + 85,
        msodgcidChangeShapeRightBracket               = 0x2000 + 86,
        msodgcidChangeShapeLeftBrace                  = 0x2000 + 87,
        msodgcidChangeShapeRightBrace                 = 0x2000 + 88,
        msodgcidChangeShapeLeftUpArrow                = 0x2000 + 89,
        msodgcidChangeShapeBentUpArrow                = 0x2000 + 90,
        msodgcidChangeShapeBentArrow                  = 0x2000 + 91,
        msodgcidChangeShapeSeal24                     = 0x2000 + 92,
        msodgcidChangeShapeStripedRightArrow          = 0x2000 + 93,
        msodgcidChangeShapeNotchedRightArrow          = 0x2000 + 94,
        msodgcidChangeShapeBlockArc                   = 0x2000 + 95,
        msodgcidChangeShapeSmileyFace                 = 0x2000 + 96,
        msodgcidChangeShapeVerticalScroll             = 0x2000 + 97,
        msodgcidChangeShapeHorizontalScroll           = 0x2000 + 98,
        msodgcidChangeShapeCircularArrow              = 0x2000 + 99,
        msodgcidChangeShapeNotchedCircularArrow       = 0x2000 + 100,
        msodgcidChangeShapeUturnArrow                 = 0x2000 + 101,
        msodgcidChangeShapeCurvedRightArrow           = 0x2000 + 102,
        msodgcidChangeShapeCurvedLeftArrow            = 0x2000 + 103,
        msodgcidChangeShapeCurvedUpArrow              = 0x2000 + 104,
        msodgcidChangeShapeCurvedDownArrow            = 0x2000 + 105,
        msodgcidChangeShapeCloudCallout               = 0x2000 + 106,
        msodgcidChangeShapeEllipseRibbon              = 0x2000 + 107,
        msodgcidChangeShapeEllipseRibbon2             = 0x2000 + 108,
        msodgcidChangeShapeFlowChartProcess           = 0x2000 + 109,
        msodgcidChangeShapeFlowChartDecision          = 0x2000 + 110,
        msodgcidChangeShapeFlowChartInputOutput       = 0x2000 + 111,
        msodgcidChangeShapeFlowChartPredefinedProcess = 0x2000 + 112,
        msodgcidChangeShapeFlowChartInternalStorage   = 0x2000 + 113,
        msodgcidChangeShapeFlowChartDocument          = 0x2000 + 114,
        msodgcidChangeShapeFlowChartMultidocument     = 0x2000 + 115,
        msodgcidChangeShapeFlowChartTerminator        = 0x2000 + 116,
        msodgcidChangeShapeFlowChartPreparation       = 0x2000 + 117,
        msodgcidChangeShapeFlowChartManualInput       = 0x2000 + 118,
        msodgcidChangeShapeFlowChartManualOperation   = 0x2000 + 119,
        msodgcidChangeShapeFlowChartConnector         = 0x2000 + 120,
        msodgcidChangeShapeFlowChartPunchedCard       = 0x2000 + 121,
        msodgcidChangeShapeFlowChartPunchedTape       = 0x2000 + 122,
        msodgcidChangeShapeFlowChartSummingJunction   = 0x2000 + 123,
        msodgcidChangeShapeFlowChartOr                = 0x2000 + 124,
        msodgcidChangeShapeFlowChartCollate           = 0x2000 + 125,
        msodgcidChangeShapeFlowChartSort              = 0x2000 + 126,
        msodgcidChangeShapeFlowChartExtract           = 0x2000 + 127,
        msodgcidChangeShapeFlowChartMerge             = 0x2000 + 128,
        msodgcidChangeShapeFlowChartOfflineStorage    = 0x2000 + 129,
        msodgcidChangeShapeFlowChartOnlineStorage     = 0x2000 + 130,
        msodgcidChangeShapeFlowChartMagneticTape      = 0x2000 + 131,
        msodgcidChangeShapeFlowChartMagneticDisk      = 0x2000 + 132,
        msodgcidChangeShapeFlowChartMagneticDrum      = 0x2000 + 133,
        msodgcidChangeShapeFlowChartDisplay           = 0x2000 + 134,
        msodgcidChangeShapeFlowChartDelay             = 0x2000 + 135,
        msodgcidChangeShapeTextPlainText              = 0x2000 + 136,
        msodgcidChangeShapeTextStop                   = 0x2000 + 137,
        msodgcidChangeShapeTextTriangle               = 0x2000 + 138,
        msodgcidChangeShapeTextTriangleInverted       = 0x2000 + 139,
        msodgcidChangeShapeTextChevron                = 0x2000 + 140,
        msodgcidChangeShapeTextChevronInverted        = 0x2000 + 141,
        msodgcidChangeShapeTextRingInside             = 0x2000 + 142,
        msodgcidChangeShapeTextRingOutside            = 0x2000 + 143,
        msodgcidChangeShapeTextArchUpCurve            = 0x2000 + 144,
        msodgcidChangeShapeTextArchDownCurve          = 0x2000 + 145,
        msodgcidChangeShapeTextCircleCurve            = 0x2000 + 146,
        msodgcidChangeShapeTextButtonCurve            = 0x2000 + 147,
        msodgcidChangeShapeTextArchUpPour             = 0x2000 + 148,
        msodgcidChangeShapeTextArchDownPour           = 0x2000 + 149,
        msodgcidChangeShapeTextCirclePour             = 0x2000 + 150,
        msodgcidChangeShapeTextButtonPour             = 0x2000 + 151,
        msodgcidChangeShapeTextCurveUp                = 0x2000 + 152,
        msodgcidChangeShapeTextCurveDown              = 0x2000 + 153,
        msodgcidChangeShapeTextCascadeUp              = 0x2000 + 154,
        msodgcidChangeShapeTextCascadeDown            = 0x2000 + 155,
        msodgcidChangeShapeTextWave1                  = 0x2000 + 156,
        msodgcidChangeShapeTextWave2                  = 0x2000 + 157,
        msodgcidChangeShapeTextWave3                  = 0x2000 + 158,
        msodgcidChangeShapeTextWave4                  = 0x2000 + 159,
        msodgcidChangeShapeTextInflate                = 0x2000 + 160,
        msodgcidChangeShapeTextDeflate                = 0x2000 + 161,
        msodgcidChangeShapeTextInflateBottom          = 0x2000 + 162,
        msodgcidChangeShapeTextDeflateBottom          = 0x2000 + 163,
        msodgcidChangeShapeTextInflateTop             = 0x2000 + 164,
        msodgcidChangeShapeTextDeflateTop             = 0x2000 + 165,
        msodgcidChangeShapeTextDeflateInflate         = 0x2000 + 166,
        msodgcidChangeShapeTextDeflateInflateDeflate  = 0x2000 + 167,
        msodgcidChangeShapeTextFadeRight              = 0x2000 + 168,
        msodgcidChangeShapeTextFadeLeft               = 0x2000 + 169,
        msodgcidChangeShapeTextFadeUp                 = 0x2000 + 170,
        msodgcidChangeShapeTextFadeDown               = 0x2000 + 171,
        msodgcidChangeShapeTextSlantUp                = 0x2000 + 172,
        msodgcidChangeShapeTextSlantDown              = 0x2000 + 173,
        msodgcidChangeShapeTextCanUp                  = 0x2000 + 174,
        msodgcidChangeShapeTextCanDown                = 0x2000 + 175,
        msodgcidChangeShapeFlowChartAlternateProcess  = 0x2000 + 176,
        msodgcidChangeShapeFlowChartOffpageConnector  = 0x2000 + 177,
        msodgcidChangeShapeCallout90                  = 0x2000 + 178,
        msodgcidChangeShapeAccentCallout90            = 0x2000 + 179,
        msodgcidChangeShapeBorderCallout90            = 0x2000 + 180,
        msodgcidChangeShapeAccentBorderCallout90      = 0x2000 + 181,
        msodgcidChangeShapeLeftRightUpArrow           = 0x2000 + 182,
        msodgcidChangeShapeSun                        = 0x2000 + 183,
        msodgcidChangeShapeMoon                       = 0x2000 + 184,
        msodgcidChangeShapeBracketPair                = 0x2000 + 185,
        msodgcidChangeShapeBracePair                  = 0x2000 + 186,
        msodgcidChangeShapeSeal4                      = 0x2000 + 187,
        msodgcidChangeShapeDoubleWave                 = 0x2000 + 188,
        msodgcidChangeShapeActionButtonBlank          = 0x2000 + 189,
        msodgcidChangeShapeActionButtonHome           = 0x2000 + 190,
        msodgcidChangeShapeActionButtonHelp           = 0x2000 + 191,
        msodgcidChangeShapeActionButtonInformation    = 0x2000 + 192,
        msodgcidChangeShapeActionButtonForwardNext    = 0x2000 + 193,
        msodgcidChangeShapeActionButtonBackPrevious   = 0x2000 + 194,
        msodgcidChangeShapeActionButtonEnd            = 0x2000 + 195,
        msodgcidChangeShapeActionButtonBeginning      = 0x2000 + 196,
        msodgcidChangeShapeActionButtonReturn         = 0x2000 + 197,
        msodgcidChangeShapeActionButtonDocument       = 0x2000 + 198,
        msodgcidChangeShapeActionButtonSound          = 0x2000 + 199,
        msodgcidChangeShapeActionButtonMovie          = 0x2000 + 200,
        msodgcidChangeShapeHostControl                = 0x2000 + 201,
        msodgcidChangeShapeTextBox                    = 0x2000 + 202,

        // Last valid change shape value for Windows Office97.
        msodgcidChangeShapeMaxWinOffice97             = 0x2000 + 203,

        // Add new ChangeShape dgcids here.

        msodgcidChangeShapeMax = msodgcidChangeShapeMaxWinOffice97,
        msodgcidChangeShapeLast = msodgcidChangeShapeMax - 1,

        msodgcidUndoMin                        = 0x3000,

        msodgcidUndoUnknown                    = 0x3000 + 0,
        msodgcidUndoMove                       = 0x3000 + 1,
        msodgcidUndoSize                       = 0x3000 + 2,
        msodgcidUndoAdjust                     = 0x3000 + 3,
        msodgcidUndoPasteSpecial               = 0x3000 + 4,
        msodgcidUndoInsertControl              = 0x3000 + 5,
        msodgcidUndoRotate                     = 0x3000 + 6,
        msodgcidUndoFormatPicture              = 0x3000 + 7,
        msodgcidUndoFormatObject               = 0x3000 + 8,
        msodgcidUndoFormatControl              = 0x3000 + 9,
        msodgcidUndoFormatWordArt              = 0x3000 + 10,
        msodgcidUndoFormatTextBox              = 0x3000 + 11,
        msodgcidUndoFormatAutoShape            = 0x3000 + 12,
        msodgcidUndoFormatPlaceholder          = 0x3000 + 13,
        msodgcidUndoFormatComment              = 0x3000 + 14,
        msodgcidUndoCopy                       = 0x3000 + 15,
        msodgcidUndoReconcile                  = 0x3000 + 16,

        // Last valid undo dgcid for Windows Office97.
        msodgcidUndoMaxWinOffice97             = 0x3000 + 17,

        msodgcidUndoFormatHorizRule            = 0x3000 + 17,

        // Add new Undo dgcids here.
        msodgcidUndoFormatCanvas               = 0x3000 + 18,
        msodgcidUndoFormatOrgChart             = 0x3000 + 19,
        msodgcidUndoFormatDiagram              = 0x3000 + 20,
        msodgcidUndoFormatInk                  = 0x3000 + 21,

        msodgcidUndoMax, msodgcidUndoLast = msodgcidUndoMax - 1,

        msodgcidUndoHostMin                    = 0x3100,
        msodgcidUndoHostMax                    = 0x3200,
        msodgcidUndoHostLast                   = 0x31FF,

        msodgcidMaxNow,

        // Note: Reserved dgcid block for Mac office 2001 is from 0x5000 thru 0x5FFF (inclusive)

        msodgcidMaxForever                       = 0xFFFF,

        /* We promise that MSODGCIDs will always fit into 16 bits */
        } MSODGCID;

/* MsoFIsDgcidNormal returns TRUE iff a dgcid identifies a normal command.
        Note that msodgcidNil counts as "Normal" DGCID. */
__inline BOOL MsoFIsDgcidNormal(MSODGCID dgcid)
        {return ((dgcid & 0xF000) == 0x0000);}

/* MsoFIsDgcidInsertShape returns TRUE iff a dgcid identifies a command
        to insert some Shape or other. */
__inline BOOL MsoFIsDgcidInsertShape(MSODGCID dgcid)
        {return ((dgcid & 0xF000) == 0x1000);}

/* MsoFIsDgcidChangeShape returns TRUE iff a dgcid identifies a command
        to change the selected Shapes to some other Shape. */
__inline BOOL MsoFIsDgcidChangeShape(MSODGCID dgcid)
        {return ((dgcid & 0xF000) == 0x2000);}

/* MsoFIsDgcidUndo returns TRUE iff a dgcid identifies a command
        that exists only to name an undo record. */
__inline BOOL MsoFIsDgcidUndo(MSODGCID dgcid)
        {return ((dgcid & 0xFF00) == 0x3000);}

/* MsoFIsDgcidUndoHost returns TRUE iff a dgcid identifies a command
        that exists only to name an undo record, and which is named by the host. */
__inline BOOL MsoFIsDgcidUndoHost(MSODGCID dgcid)
        {return ((dgcid & 0xFF00) == 0x3100);}

/* MsoFIsDgcidInsertDiagram returns TRUE iff a dgcid identifies a command
        to insert some Diagram or other. */
__inline BOOL MsoFIsDgcidInsertDiagram(MSODGCID dgcid)
        {return (dgcid >= msodgcidInsertOrgChart &&
                dgcid <= msodgcidInsertVennDiagram);}

/* MsoSptFromDgcid returns the MSOSPT inserted by a DGCID.  This only
        works if dgcid is an "insert shape" or a "change shape" dgcid. */
__inline MSOSPT MsoSptFromDgcid(MSODGCID dgcid)
        {return (MSOSPT)(dgcid & 0x0FFF);}

/* MsoDgmtFromDgcid returns the MSODGMT inserted by a DGCID.  This only
        works if dgcid is an "insert diagram" dgcid. */
__inline MSODGMT MsoDgmtFromDgcid(MSODGCID dgcid)
        {return (MSODGMT)(dgcid & 0x0FFF);}

/* MsoFIsDgcidValidO97 returns TRUE if a MSODGCID existed in Office97. */
__inline BOOL MsoFIsDgcidValidO97(MSODGCID dgcid)
        {
        return ((msodgcidNormalMin <= dgcid && dgcid < msodgcidNormalMaxWinOffice97) ||
                (msodgcidShapeMin <= dgcid && dgcid < msodgcidShapeMaxWinOffice97) ||
                (msodgcidChangeShapeMin <= dgcid && dgcid < msodgcidChangeShapeMaxWinOffice97) ||
                (msodgcidUndoMin <= dgcid && dgcid < msodgcidUndoMaxWinOffice97));
        }

/* MsoFIsDgcidValid returns TRUE if a MSODGCID is a good one. */
__inline BOOL MsoFIsDgcidValid(MSODGCID dgcid)
        {
        return ((msodgcidNormalMin <= dgcid && dgcid < msodgcidNormalMax) ||
                (msodgcidShapeMin <= dgcid && dgcid < msodgcidShapeMax) ||
                (msodgcidChangeShapeMin <= dgcid && dgcid < msodgcidChangeShapeMax) ||
                (msodgcidUndoMin <= dgcid && dgcid < msodgcidUndoMax));
        }


#define msogrfdgciNil                   (0)
#define msofdgciSetRepeat               (1<<1)
#define msofdgciCanRepeat               (1<<2)
#define msofdgciAskEnabled              (1<<3)
#define msofdgciAskLatched              (1<<4)
#define msofdgciAskCreateCommand        (1<<5)
#define msofdgciAskFreeCommand          (1<<6)
#define msofdgciAskCloneCommand         (1<<7)
#define msofdgciNotVisibleToUser        (1<<8)
#define msofdgciPostsUndoRecords        (1<<9)
#define msofdgciPrefersTextButton       (1<<10)
#define msofdgciSwatchControl           (1<<11)
#define msofdgciEditControl             (1<<12)
#define msofdgciOnDialog                (1<<13)
#define msofdgciPrefersTextImageButton  (1<<14)
#define msofdgciSplitMenuControl        (1<<15)
#define msofdgciFeatureAnim             (1<<16)
#define msofdgciRecordSimple            (1<<17)
#define msofdgciNoRecord                (1<<18)
#define msofdgciUpdateAfterExecute      (1<<19)
#define msofdgciAskCreateRepeat         (1<<20)
#define msofdgciHideIfDisabled          (1<<21)

#define msofdgciUserInterface           (1<<22) // replace VisibleToUser?
        /* This command is a perfectly ordinary user-interface command.
                Host's should feel free to attach it to a command bar button. */

#define msofdgciAskTcid                 (1<<23)
#define msofdgciAskDgcidForUndo         (1<<24)
#define msofdgciShapeInsertionTool      (1<<25)

/* MSOFDGCD are flags describing when commands are disabled.  For
        simplicity ALL of these flags are of the form "disable if <foo>".
        Any command whos enabling is more complicated than can be expressed
        which such a mechanism should set msofdgciAskEnabled and then handle
        (in the FCommandFoo proc) dgcmAskEnabled. */
#define msogrfdgcdNil                 (0)
#define msofdgcdNoSelection           (1<<1)
#define msofdgcdNothingSelected       (1<<2)
#define msofdgcdNotMultipleSelection  (1<<3)
#define msofdgcdNotSingleSelection    (1<<4)
#define msofdgcdNoGroupsSelected      (1<<5)
#define msofdgcdNoDrawing             (1<<6)
#define msofdgcdNoView                (1<<7)
#define msofdgcdNoPicturesSelected    (1<<8)
#define msofdgcdNoPolygonsSelected    (1<<9)
#define msofdgcdInReshapeMode         (1<<10)
#define msofdgcdInRotateMode          (1<<12)
#define msofdgcdNotInReshapeMode      (1<<13)
#define msofdgcdNotInRotateMode       (1<<14)
#define msofdgcdInWrapPolygonMode     (1<<14)
#define msofdgcdNotInWrapPolygonMode  (1<<15)
#define msofdgcdNotAllOnSamePage      (1<<16)
#define msofdgcdNotAllInk             (1<<17)

/* In some cases these disable flags are cummulative.  For example,
        if you want your command disabled if you're not in rotate mode
        you almost certainly want it disabled if there's no selection
        to be in a mode.  So we define a few groups of flags that you should
        probably use instead of the single flags with the same name. */

#define msogrfdgcdNoSelection \
        (msofdgcdNoSelection | msofdgcdNoDrawing)
#define msogrfdgcdNothingSelected \
        (msofdgcdNothingSelected | msogrfdgcdNoSelection)
#define msogrfdgcdNotMultipleSelection \
        (msofdgcdNotMultipleSelection | msogrfdgcdNothingSelected)
#define msogrfdgcdNotSingleSelection \
        (msofdgcdNotSingleSelection | msogrfdgcdNothingSelected)
#define msogrfdgcdNoGroupsSelected \
        (msofdgcdNoGroupsSelected | msogrfdgcdNothingSelected)
#define msogrfdgcdNoPicturesSelected \
        (msofdgcdNoPicturesSelected | msogrfdgcdNothingSelected)
#define msogrfdgcdNoPolygonsSelected \
        (msofdgcdNoPolygonsSelected | msogrfdgcdNothingSelected)
#define msogrfdgcdNotInReshapeMode \
        (msofdgcdNotInReshapeMode | msogrfdgcdNoSelection)
#define msogrfdgcdNotInRotateMode \
        (msofdgcdNotInRotateMode | msogrfdgcdNoSelection)
#define msogrfdgcdNotInWrapPolygonMode \
        (msofdgcdNotInWrapPolygonMode | msogrfdgcdNoSelection)
#define msogrfdgcdNoView \
        (msofdgcdNoView | msofdgcdNoSelection | msofdgcdNoDrawing)

/* A few macros for hosts to use to ask common questions about commands. */

#define MsoFDgciMayDisable(pdgci) \
        (((pdgci)->grfdgcd != 0) || ((pdgci)->grfdgci & msofdgciAskEnabled))

#define MsoFDgciMayLatch(pdgci) \
        (((pdgci)->grfdgci & msofdgciAskLatched) != 0)

#define MsoFDgciPrefersTextButton(pdgci) \
        (((pdgci)->grfdgci & msofdgciPrefersTextButton) != 0)

#define MsoFDgciPrefersTextImageButton(pdgci) \
        (((pdgci)->grfdgci & msofdgciPrefersTextImageButton) != 0)

#define MsoFDgciIsSwatchControl(pdgci) \
        (((pdgci)->grfdgci & msofdgciSwatchControl) != 0)

#define MsoFDgciIsEditControl(pdgci) \
        (((pdgci)->grfdgci & msofdgciEditControl) != 0)

#define MsoFDgciIsCustomizable(pdgci) \
        (!MsoFDgciIsSwatchControl(pdgci) && !MsoFDgciIsEditControl(pdgci))

#define MsoFDgciIsSplitMenuControl(pdgci) \
        (((pdgci)->grfdgci & msofdgciSplitMenuControl) != 0)

#define MsoFDgciMayHide(pdgci) \
        (((pdgci)->grfdgci & msofdgciHideIfDisabled) != 0)

#define MsoFDgciMayChangeTcid(pdgci) \
        (((pdgci)->grfdgci & msofdgciAskTcid) != 0)

#define MsoFDgciIsShapeInsertionTool(pdgci) \
        (((pdgci)->grfdgci & msofdgciShapeInsertionTool) != 0)

/* MSODGE = Drawing Event.  These are passed to DGS/DGVS/DMS::OnEvent.

        Events come in four flavors: "Do", "Request", "Before", and "After".
        Their names should be of the form msodge[Do|Request|Before|After]Foo.
        The OnEvent methods all have the same form; they take a BOOL * called
        pfResult, an LPARAM, and a WPARAM.  The sink (that's the DGS, DGVS, or
        DMS, the implementer of OnEvent) may change *pfResult to affect the
        flow of control in the source (that's the DG, DGV, or DM, the caller
        of OnEvent) for "Do" and "Request" events.  For "Before" and "After"
        events, pf will be NULL.  The values in the LPARAM and WPARAM depend
        on the event.

        "Do" events are really just all those events for which *pfResult
        defaults to TRUE.  These are generally used to give the sink a chance
        to replace the default processing of an event. If the sink wishes to
        handle the event they should do so, and then set *pfResult (which in this
        case means "Should the source handle the event in the default way?") to
        FALSE before returning.  If they want the default behavior they can just
        return; *pfResult will have been initialized to TRUE ahead of time.

        "Request" events are really just all those events for which *pfResult
        defaults to FALSE.  These are often used to give the sink a chance to
        tell the source to not do something.  If an event is called
        msodgeRequestCancelFoo it means that if the sink just returns
        (leaving *pfResult == FALSE) we will do Foo.  If the sink doesn't
        want Foo to happen it should set *pfResult to TRUE.

        "Before" and "After" events notify the sink that something is about
        to happen or that something has just happened, but do not give the
        sink the opportunity to change anything.

        These events should all have comments by them explaining when
        and to whom (DGS, DGVS, or DMS) they are fired. */
/* TODO peteren: Some of our Request events are really requests
        to cancel things.  So for instance if the host sets *pfResult
        to TRUE in msodgeRequestSelChange it means to NOT change the
        selection.  I think the event names should have some negating
        word in them, like msodgeRequestCancelSelChange.  A request
        event doesn't imply cancelling, it's just an event who's default
        is FALSE. */

typedef enum
        {
        msodgeNil = 0,

        /* ----- Events passed to the DGS, the Drawing Site. */

        msodgeDgsFirst = 100,

        msodgeBeforeDgFree,
        /* Passed to DGS::OnEvent when the DG is about to go away.  We
                promise not to call any more DGS methods after this event.
                Among other things, this gives the DGS a chance to free anything
                it had hanging out of the dgsi.pvDgs. */
        msodgeDgsUseMeNext,

        msodgeRequestMarkShapeContext,
        /* Passed to the DGS when we're running a bring forward or a send
                backward command.  We'll have filled out dgeb.hsp.
                The host may mark (using DG::MarkShape, we'll
                have already called DG::FBeginMarkShapes) those Shapes in
                the same "context" as the specified Shape, and return TRUE
                to let us know they did so.  Or if you don't have "contexts"
                you can just ignore the event. */
        msodgeBeforeDeleteShape,
        msodgeAfterDeleteShape,
        /* We pass msodgeBeforeDeleteShape and msodgeAfterDeleteShape around
                moving the Shape to the deleted list. */
        msodgeBeforeUndeleteShape,
        msodgeAfterUndeleteShape,
                /* TODO peteren: Comment. */

        msodgeBeforeDelete,
        /* Passed to DGS::OnEvent before a shape is deleted.
                WPARAM is client data */
        msodgeBeforeBoundsChange,
        /* Passed to DGS::OnEvent before shape's bounds is changed.
                WPARAM is client data */
        msodgeAfterBoundsChange,
        /* Passed to DGS::OnEvent after shape's bounds has changed.
                WPARAM is client data */
        msodgeAfterPropertyChange,
        /* Passed to DGS::OnEvent after a Shape's property changes.
                This is obsolete; use msodgeAfterShapePropertyChange instead. */
        msodgeAfterAdd,
        /* Passed to DGS::OnEvent after shape is added to the DG.
                WPARAM is client data */
        msodgeAfterZOrderChange,
        /* Passed to DGS::OnEvent after shape's z-order is changed in the DG.
                WPARAM is client data */
        msodgeAfterShapeAdjust,
        /* This event is not fired.
                Passed to DGS::OnEvent after shape is adjusted in the DG.
                WPARAM is client data */

        msodgeBeforePurgeShape,
        msodgeAfterPurgeShape,
        msodgeBeforePurgeDeletedShape,
        msodgeAfterPurgeDeletedShape,
                /* These events are fired when purging shapes.  They fill
                        out dgeb.hps and dgeb.pvClient. */

        msodgeRequestCancelCreateShape,
                /* Passed to DGS::OnEvent after a brand new shape is created.
                        The shape has been entirely initialized and hooked up to its DG.  The
                        appropriate fields of the MSODGEB are filled out (hsp, pvClient, pvAnchor),
                        as well as additional hspModel field in dgexRequestCancelCreateShape.
                        If you return TRUE, we will treat the shape creation as though it
                        failed, and purge the shape. */

        msodgeBeforeClipboardClone,
                /* Passed before we clone shapes to the clipboard. */
        msodgeAfterClipboardClone,
                /* Passed after we clone shapes to the clipboard. */

        msodgeRequestChooseSpc,
                /* Lets the host override our choice of MSOSPC for a Shape. */

        msodgeRequestChooseShapeType,
                /* Lets the host override our choice of the shape type returned
                        from DispShape::get_Type for a Shape. */

        msodgeRequestCancelCloneShape,
                /* Called after a shape has been cloned and inserted into the new drawing. */

        msodgeBeforeChangeShapeType,
                /* Lets the host disable connection sites for a shape or do whatever
                        it wants to before a change shape. */

        msodgeAfterChangeShapeType,
                /* Lets the host disable connection sites for a shape or do whatever
                        it wants to after a change shape. */

        msodgeConstrainShape,
        /* Give the app a chance to constrain the new proposed size for the shape. */

        msodgeTextboxUndo,
                /* Passed after we perform a textbox undo operation. */

        msodgeBeforeShapePropertyChange,
        /* Passed to DGS::OnEvent before a Shape's property changes.
                dgeb.dgexBeforeshapePropertyChange will be filled out. */

        msodgeAfterShapePropertyChange,
        /* Passed to DGS::OnEvent after a Shape's property changes.
                dgeb.dgexAftershapePropertyChange will be filled out. */

        msodgeAfterDgBecomesInvalid,
                /* Passed to DGS::OnEvent whenever the DG is invalidated.  Note
                        that although DG::Invalidate is often called many times in
                        quick succession, you will only get this event the first time.
                        That's why it has NO arguments, you're not supposed to try
                        to look at the particular invalidation.  You're just supposed
                        to remember that the DG is invalid.  If you want to actually
                        deal with the invalidation, call DG::Validate
                        and you'll get an msodgeBeforeDgBecomesValid event. */
                /* TODO peteren: I suspect we should fire this event everytime
                        DG::Invalidate is called instead of just the first time.
                        Right now we fire this horrible event msodgeAfterZOrderChange
                        from DG::Invalidate everytime for PowerPoint; XL also has some
                        need for this. */

        msodgeBeforeDgBecomesValid,
                /* Passed to DGS::OnEvent from under DG::Validate.  Office will
                        have filled out dgexBeforeDgBecomesValid. */

        msodgeAfterShapeNameChange,
        /* Passed to DGS::OnEvent when a Shape Name (Automation) has changed */

        msodgeRequestFakeOleObject,
                /* Passed when we want to put a shape on the clipboard to give the host
                        an opportunity to create a fake OLE object that can represent the
                        the shape. Used by Excel for charts. */

        msodgeAfterCopy,
                /* I just copied some stuff. */

        msodgeBeforeUngroupShape,
                /* We're about to ungroup a group shape */

        msodgeUndoGroup,
                /* We're about to run an undo operation on a group. */

        msodgeRequestMatchShapeName,
                /* We're looking up a shape by name and we've found a shape that
                        matches the number part of the name, but it doesn't match the default
                        shape name that we know of and so we're about to return FALSE from
                        DG::FFindShapeFromName unless you tell us otherwise.  Passes the
                        hsp in question and the name being searched for (with the number
                        stripped off). */

        msodgeDoRelativeOffset,
                /* Ask the host to adjust the anchor to be a relative   anchor. */

        msodgeRequestHspCanPassThru,
                /* Request the host if the picture shape is a pass thru case. */

        msodgeRequestHspPibNameAlloc,
                /* Request the host to allocate the right string to set into
                        the msopidPibName.  used to keep Word field codes in sync with
                        the corresponding shape properties. */

        msodgeDoSetNextTxbx,
                /* Ask the host to set the next textbox. */

        msodgeDoReanchor,
        /* Passed to DGS::OnEvent for reanchoring a shape before
                another shape. This is needed in Word for anchoring of
                imagemap area shapes bafore the map user shape. */

        msodgeBeforeStyle,
                /* We are about to write the style attribute of a shape. */

        msodgeDoHRAdjustment,
                /* Ask the host to adjust the height for an HR. */

#ifdef NO96_162753
        msodgeFIconObject,
                /* Returns a BOOL value in fResult to tell whether the given
                        OLE object is represented as an Icon or not. */
#endif // NO96_162753

        msodgeBeforeConvertHToO,
                /* In Excel's R2L sheets, H coord's origin for x starts
                        at colMax. This event gives Excel a chane to modify x
                        so that OAFromH is called with logical H coordinates */

        msodgeAfterConvertOToH,

        msodgeBeforeShapeXML,
        /* Passed to DGS::OnEvent before writing shape XML. */

        msodgeAfterAsyncBlipDownload,
        /* Fired when a picture download completes.  Used by Word to invalidate
          inline shapes after download (DG::Invalidate does not cause a redraw
          on inline shapes.) */

        msodgeDoAutomaticBW,
        /* Fired when we want host to patch the automatic Fill/Line
                scheme indices. (e.g. XL creation of a bitmap surface) */

        msodgeOCXShape,
        /* Ask the host if this shape is OCX.  There's gotta be a better way */

        msodgeBeforeChangeSpid,
        /* Notify the host we're changing the shape's id.  Excel needs this */

        msodgeBeforePosPropXML,
        /* Fired before writing the msopidPosrelh/msopidPosrelv properties
                in XML.  Word needs this */

        msodgeBeforeBorderPropXML,
        /* Fired before writing the msopidBorder"X"Color properties
                in XML.  Word needs this */

        msodgeCreateGroupAnchorFromChild,
                /* Used by Word to create the anchor of a group shape based off of the
                   anchor info of one of its child shapes before the child anchor is
                   gone, during load of HTML/XML */

        msodgeNoUI,
                /* Asks the host if we want to do UI stuff, like bring up dialogs. */

        msodgeDoUpdateLinkedURLBlip,
                /* Ask the host if we want to update linked blips */

        msodgeIndentsText,
                /* During writing XML style, ask if host wants text-indent property. */

        msodgeLinkedWordart,
                /* Check for Excel's Linked Wordarts to block Wordart UI */

        msodgeDoRchForAlign,
        /* Ask the host for the bounding rectangle of the "aligning space" for
           the page which contains the selected objects. This is used by
           the "relative to page" align commands. Publisher needs this.
           MSO sets dgexDoRchForAlign.prchPage to default value.
           The host may set dgexDoRghForAlign.prchPage. */

        msodgeAfterShapeCreateFromUI,
        /* Alert the host that a shape has been created from the UI. Publisher
           needs this because it implements tables as msosptHostControl
                   objects, and the needed text edit activation is not fired on object
                   create because the text box tool was not used. Also note that
                   msodgeAfterAdd has too many callers to be useful here -- we'd end
                   up activating tables after undo, etc. */

        msodgeAfterNeedsSave,
                /* Publisher needs this for the auto save feature */

        msodgeAfterCreateDrawingCanvas,
                /* Notify the host after a drawing canvas object has been created. */

        msodgeDoClipboardClone,
                /* Passed to DGS:OnEvent to further allow hosts to customize copying of
                   shapes.  Not handling this event is the same as returning TRUE,
                   which is previous behavior; returning FALSE will have Escher use the
                   shapes already on the clipboard drawing as copy source */

        msodgeBeforeShapeVML,
                /* Fired after a shape's proto shapetype but before a shape's VML. */

        msodgeDoOffsetForDuplicate,
        /* Passed to DGS::OnEvent when shapes are about to be duplicated. This gives
                the host a chance to relatively adjust the coordinates of where the
                shapes will appear. If *pfResult is set to FALSE by the host then the
                duplicated shapes will adjust by the offset supplied by the host. */

        msodgeOffsetFromShapesOnPaste,
        /* Passed to DGS::OnEvent when shapes are about to be pasted. This gives
           the host a chance to specify the offset used to shift the shapes down
           if their coordinates collide with existing shapes to within a certain
           slop. If *pfResult is set to FALSE by the host then the pasted shapes
           will adjust by the offset supplied by the host. */

        msodgeRequestCancelPaste,
        /* Passed to DGS:OnEvent to further allow hosts to customize pasting of shapes.
           Not handling this event is the same as returning FALSE, which is previous behavior;
           returning TRUE will have Escher terminate the paste action. Occurs after the duplication
           of objects is complete but before the msodgeAfterPaste event. */

        msodgeResolveReadingOrder,
        /* We need to ask the host to resolve the text direction of some attached text.
           This is used by Excel, BIDI. */

        msodgeRequestCancelDeleteShape,
        /* Passed to DGS::OnEvent to allow the host to do any processing it needs to
                do that can fail. Not processing this event is the same as returning FALSE,
                which is previous behavior; returning TRUE will have Escher terminate the
                delete of the shape. Occurs before the msodgeBeforeDelete callback and
                only gets issued for the top-level shape, not its children. */

        msodgeRequestCanInsertInCanvas,
        /* Give the host a chance to disallow a shape from going into a canvas */

        msodgeAfterZOrderHasChanged,
        /* Passed to DGS::OnEvent. This event is used by Publisher since the other
                msodgeAfterZOrderChange event that gets fired is between the changing of
                the zorder. This event gets fired at the very end of the zorder.*/

        msodgeDoGetFrameButtons,
        /* This event is fired in GetHandleInfo. Publisher uses it to tell Escher
                if it has buttons associated with the object frame */

        msodgeDoScaleTextFonts,
        /* We are requesting the host app to change the text size in a textbox */

        msodgeDoDiagramAutoFit,
        /* We are requesting the host app to fit the shape to the text */

        msodgeSaveVisualElement,
        /* We are requesting the host app to save the visual element */

        msodgeLoadVisualElement,
        /* We are requesting the host app to load the visual element */

        msodgeResetTextRotation,
        /* This event is fired to tell, PowerPoint specifically, that the text
                rotation needs to be reset based on the current shape props (i.e.
                flipV, flipH ...) */

        msodgeDoUndo,
        /* Perform undo for top of stack */

        msodgeRequestPseudoInlinePaste,
        /* check for Word's insertion/paste layout option */

        msodgeAfterAddDiagramShape,
        /* just added a shape to a diagram needed by ppt */

        msodgeAfterRemoveDiagramShape,
        /* just removed a shape to a diagram needed by ppt */

        msodgeAfterAddShapeToGroup,
        msodgeAfterRemoveShapeFromGroup,
        /* Messages to allow host synchronization of canvas */

        msodgeBeforeDiagramReverse,
        msodgeAfterDiagramReverse,
        /* The diagram shape order, has been reversed*/

        msodgeAfterChangeDiagramType,
        /* The diagram type just changed, needed by ppt*/

        msodgeBeforeChangeDiagramType,
        /* The diagram type is about to be changed, needed by ppt*/

        msodgeAfterConvertOrgChart, // OBSOLETE
        /* an orgchart has been converted from the OLE server.  needed by ppt*/

        msodgeDoHspOfUndoRecord,
    /* sent from GetHspOfUndoRecord, needed by pub*/

        msodgeAddDelShapeGroupUndo,
        /* For dragging textboxes in and out of Word Canvas */

        msodgeRequestTextBounds,
        /* Called to get the text bounds of a Shape from the host. */

        msodgeAfterSolverApplyProps,
        /* Called after the SOLVER finishes applying properties to a shape
           since the event msodgeAfterPropertyChange isn't necessarily the last word. */

        msodgeDoSetDiagramDefSize,
        /* Allow the host to customize the default diagram size. */

        msodgeDoApplyTextStyleToShape,
        /* request the app to apply some text styles */

        msodgeGetPreferredTextStyles,
        /* get preferred font size, etc. */

        msodgeBeforeDiagramMoveShape,
        msodgeAfterDiagramMoveShape,
        /* The diagram shape has moved */

        msodgeResolveBestScaleForSlideShowSize,

        msodgeEnterCS,
        /* sent right after critical section of DG is entered */
        msodgeLeaveCS,
        /* sent right before critical section of DG is left */

        msodgeAfterCreateInkShape,
        /* sent after escher shape has been created, and the ink (wisp) properties */
        /* have been set */

        msodgeUseBackForLineColor,
        /* Ask site if backLineColor should be used in place of lineColor to determine
           the color in the "More Colors..." dialog. Used by Publisher for cases
           where border art have been applied to items in the selection. */

        msodgeRequestInkAnnotation,
        /* Ask site if hsp is an ink comment*/

        msodgeBeforeCreateDrawingCanvas,
        /* Notify the host before a drawing canvas is about to be created. */

        msodgeDoPseudoInline,
        /* Notify host app to set host specific pseudoinline shape properties. */

        msodgeBeforeClipboardMunge,
                /* Passed before we munge shapes to the clipboard. */
        msodgeAfterClipboardMunge,
                /* Passed after we munge shapes to the clipboard. */

        msodgeRequestAntiAliasInkGIF,
        /* Find out if we need to anti-alias GIF images */

        msodgeChildShapeVML,
        /* Asks client to write any client's tags for child shape */

        msodgeDoFinishCreateShapeFromModel,
        /* Asks app to make any adjustments to a shape that has just been created "in the image"
           of another model shape, with matching anchoring characteristics, etc.  Not necessarily
           a clone of the original shape, and may look different, but similar in behavior. */

        msodgeAfterTextBoxCreation,
        /* Notify the host that we have created one or more text boxes. 
           The text boxes might have been created as part of a diagram. */

         msodgeDoFindPageForInkShape,
        /* Find out what page the ink shape belongs to */

        msodgeBeforeWriteDgvgList,
        /* sent before writing to the list of DGVs in a DGVG, (useful for threading) */

        msodgeAfterWriteDgvgList,
        /* sent after writing to the list of DGVs in a DGVG, (useful for threading) */

        msodgeBeforeReadDgvgList,
        /* sent before reading from the list of DGVs in a DGVG, (useful for threading) */

        msodgeAfterReadDgvgList,
        /* sent after reading from the list of DGVs in a DGVG, (useful for threading) */

        /* !!! ADD NEW DGS EVENTS ABOVE THIS LINE  !!! */
        /* ||| Note: ADDING NEW EVENT MAY REQUIRE FULL BUILD FOR SHIP !!! */
        msodgeDgsMax,

        /* ----- Events passed to the DGGS, Drawing Group Site. */

        msodgeDggsFirst = 225,

        msodgeRequestAbortSave,
        /* Passed to DGGS::OnEvent periodically while we're saving.  If the host
                wants us to abort the save it should set *pfResult to TRUE. */
        msodgeRequestYieldSave,
        /* Passed to DGGS::OnEvent periodically while we're saving.  If the host
                wants us to yield it should set *pfResult to TRUE.  Yielding means we
                should clean up any state we have that's not on the stack and
                pass DGGS::OnEvent an msodgeRequestAbortAfterYieldSave event. */
        msodgeRequestAbortAfterYieldSave,
        /* Passed to DGGS::OnEvent after getting back TRUE from
                msodgeRequestYieldSave.  We promise to have cleaned up
                all of our state into the stack, so that the host can safely
                switch out of this thread and do whatever it wants.  When it returns
                if *pfResult is TRUE it means we should abort the save, FALSE
                means we should just continue. */
        msodgeBeforeSave,
        /* Passed to DGGS::OnEvent before doing an actual save in DGG::FSave. This
                allows the host to for example munge the svb for the clipboard case.
                LPARAM is the MSOSVB, WPARAM is the DG if any. */
        msodgeAfterSave,
        /* Passed to DGGS::OnEvent after doing an actual save in DGG::FSave.
                LPARAM is the MSOSVB, WPARAM is the DG if any. */
        msodgeBeforeLoad,
        /* Passed to DGGS::OnEvent before doing an actual load in DGG::FLoad. This
                allows the host to for example munge the ldb for the clipboard case.
                LPARAM is the MSOLDB, WPARAM is the DG if any. */
        msodgeAfterLoad,
        /* Passed to DGGS::OnEvent after doing an actual load in DGG::FLoad.
                LPARAM is the MSOLDB, WPARAM is the DG if any. */
        msodgeRequestAbortLoad,
        /* Passed to DGGS::OnEvent periodically while we're loading.  If the host
                wants us to abort the load it should set *pfResult to TRUE. */
        msodgeAfterDelayLoadFailure,
        /* Passed to DGGS::OnEvent on failure to load a picture from the delay stream.
                If the host wants to put up an alert, it is recommended that they wait until
                idle time. *pfResult is not checked. */
        msodgeRequestCancelWriteSMColor,
        /* Passed to DGGS::OnEvent during exporting split menu colors in XML. If
                the host        wants us to not write for the split menu it should set *pfResult
                to TRUE. */
        msodgeDoWritePidDefault,
        /* Passed to DGGS::OnEvent during exporting defaults properties in XML. If
                the host        wants us to not write for the property it should set *pfResult
                to TRUE. */
        msodgeDoInitInkSettings,
        /* Passed to DGGS::OnEvent to initialize ink style pens. If the host does not have its own settings,
           then use general pen settings. */
         msodgeInkPenColor,
        /* Passed to DGGS::OnEvent to get ink color for specific ink style pen. If the host does not have its own
           pen color then return false */

        msodgeDggsMax,

        /* ----- Events passed to the DGVS, Drawing View Site. */

        msodgeDgvsFirst = 300,

        msodgeDoMove,
        msodgeAfterMove,
        msodgeDoResize,
        msodgeAfterResize,
        msodgeDoRotate,
        msodgeAfterRotate,
        msodgeDoAdjustShape,
        msodgeAfterAdjustShape,
        msodgeDoCropPicture,
        msodgeAfterCropPicture,

        msodgeAfterDgvUsedTool,
        /* Passed to DGVS::OnEvent when the DGV has used a tool.  This gives
                the host a chance to reset the tool.  We'll pass in the tool we used
                (an MSODGCID) in wParam and in lParam an MSODGTU explaining what they
                did with the tool. */

        msodgeRequestCountShapesToView,
        /* Passed to DGVS::OnEvent when the DGV is rebuilding its SPVs.
                This gives the DGVS a chance to give us a guess on the number of
                Shapes that well be viewed.  We'll initially allocate our PLEX
                to have this many SPVs in it.  The wParam we pass you will be
                a pointer to an integer in which you can put your guess.  The
                default is to guess zero, and therefore grow the PLEX whenever
                we find a Shape in this View. */
        msodgeRequestChooseShapesToView,
        /* Passed to DGVS::OnEvent when the DGV is rebuilding its SPVs.
                This gives the DGVS a change to tell us which Shapes it wants
                viewed in this View.  The default is to view all the Shapes
                in the Drawing.  If the Host wants to filter the shapes they
                can call Pidg->MarkShape(hsp) on a bunch of shapes and cancel
                the default behavior.  Note that if you leave fOmitOffscreenShapes
                TRUE we'll still not view shapes outside the view.  For now if
                you mark a group we'll include all the children of the group
                (as this is the behavior Word is after).  It's an error for
                you to mark a Shape that's not top-level (that is, that's the
                child of a group) without marking all its parents as well. */
        msodgeRequestCancelViewShape,
        /* Passed to DGVS::OnEvent when the DGV is rebuilding its SPVs.
                This gives the DGVS a chance to (by returning TRUE) cause
                us to NOT view a Shape we've decided to view.  We'll fill out
                dgveb.hspCancelView and dgveb.pvClientCancelview for the site
                to examine. */

        msodgeConstrainSizeOfShape,
        /* Passed to DGVS::OnEvent when we set up a direct manipulation to allow
                the host to constrain the size of the dragged shape. */

        msodgeBeforeDrag, msodgeBeginDrag = msodgeBeforeDrag,
        /* Passed to DGVS::OnEvent when we begin a drag.
                TODO peteren: get rid of msodgeBeginDrag */

        msodgeAfterDrag, msodgeEndDrag = msodgeAfterDrag,
        /* Passed to DGVS::OnEvent when we end a drag.
                TODO peteren: get rid of afterdrag. */

        msodgeShapeExtentInDrag,
        /* Passed to DGVS::OnEvent during a drag to show the bounding rectangle of
                the shapes being dragged. */

        msodgeDoDrawShape,
        /* Passed to DGVS::OnEvent just before we draw a shape.  We will have
                filled out dgexDrawShape.  If you set fResult to FALSE we'll not
                draw the Shape, trusting you to have done it all. */

        msodgeAfterDrawShape,
        /* Passed to DGVS::OnEvent just after we draw a shape.  We will have
                filled out dgexDrawShape. */

        msodgeDoClickShape,
        /* Passed to DGVS::OnEvent, WPARAM is client data, LPARAM is hsp */

        msodgeAfterChangeShapeOnView,
        /* Passed to DGVS::OnEvent when an SPV appeares in a DGV, vanishes from
                a DGV, or moves in a DGV such that it's HSPV has changed.  Fills
                out hspvOld and hspvNew in the DGVEB (just hspvOld on a vanish,
                just hspvNew on an "appear" and both on a "move"). */

        msodgeAfterInsertShape,
        /* Passed to DGVS::OnEvent when a shape has just been inserted. This gives
                the host a chance to edit the default shape properties. The lParam
                is the HSP being inserted. */

        msodgeRequestCancelButtonClick, msodgeRequestCancelButtonDrag = msodgeRequestCancelButtonClick/*old name*/,
        /* Gives the host a chance to cancel a button drag before it starts
                and change it into a move drag instead. */

        msodgeAfterButtonClick, msodgeAfterButtonShapeClick = msodgeAfterButtonClick,
        /* Passed to DGVS::OnEvent when a click happens on a shape which
                has the button property set to TRUE. */

        msodgeAfterDgvFree,
        /* Passed to DGVS::OnEvent when the DGV is about to go away.  This
                is called "After" instead of "Before" because we've started
                to dismantle the DGV and so some methods on it may not work.
                You should use use this event to free stuff you might have
                hanging off the MSODGVSI, say in the pvDgvs field. */

        msodgeDragShapeStart,
        msodgeDragShapeFinish,
        msodgeDragShapePaint,
        msodgeDragShapeDraw,
        msodgeDragShapeErase,

        msodgeDragShapesStart,
        msodgeDragShapesFinish,

        msodgeDoDblClickShape,
        /* Passed to DGVS::OnEvent, when only one shape is selected and
                msopidFOnDblClickNotify property is set for the shape.
                WPARAM is client data, LPARAM is hsp  */

        msodgeAfterAttachedObjectDeleted,
        /* Passed to DGVS::OnEvent, WPARAM is client data, LPARAM is hsp.
                If a shape is being deleted whose msopidFDeletedAttachedObject
                property is set then ONLY the attached object is deleted and
                this notification is sent. */

        msodgeDoHitTestShape,
        /* Passed to DGVS::OnEvent when blah. */

        msodgeBeforeRegisterShapeRcvRemoveMe, // TODO peteren: Remove this!

        msodgeRequestCancelInsertShape,
        /* Passed to DGVS::OnEvent when a shape has just been inserted. This gives
                the host a chance to cancel the insertion of a drag inserted shape
                and edit the default shape properties. The lParam is the HSP being
                inserted. */

        msodgeDoScrollOrDragDrop,
        /* Passed to the host when the mouse leaves their window and they don't
                have an inset scroll region.  Currently, only Excel has this behavior. */

        msodgeAfterHitTestShape, // Was "msodgeRequestHitTestShape"
        /* Passed to DGVS after Office has hit-tested a Shape.  The host
                can look at dgexAfterHitTestShape and change the drgh and fHit
                values therein. */

        msodgeRequestPositionSelection,
                /* Called to get the host a chance to mess with the MSOSVI
                        we'll use to position the selection dots, fuzz, etc.
                        of a Shape. */

        msodgeRequestPositionText,
                /* Called to get the host a chance to mess with where the
                        text goes on a Shape. */

        msodgeRequestUnionShapeCoverage,
        /* Passed to DGVS when Office is messing around with the regions of
                Shapes.  The host should look at hspCoverage, hspvCoverage,
                and psviCoverage, and also cvrkCoverage (to tell which kind of
                coverage region we're interested in).  If it does anything it
                should set fResult to TRUE.  It should union it's region into
                hrgnCoverage.  If something goes wrong it should set fErrorCoverage.
                If it wants Office to NOT continute to add it's own coverage it
                should set fCompletedCoverage to TRUE.  Note that hrgnCoverage
                may be msohrgnNil when we call you, you should use MsoFEnsureRgn
                if you have something to add. */

        msodgeInsertDefaultSizedShape,
        /* Passed to DGVS when we insert a default sized shape. */

        msodgeAfterGetShapeBounds,
        /* Passed to the DGVS after Office has determined the bounding rectangle
                of a Shape it's thinking of viewing.  We'll have filled out
                dgveb.hspGetBounds to identifying the Shape, dgveb.pvClientGetBounds
                with the client data of the Shape, dgveb.psviGetBounds to let
                the DGVS know where the Shape is, and dgveb.prcvGetBounds with what
                Office thinks the bounding rectangle should be.  The host can
                modify *prcvGetBounds if it so desires. */

        msodgeBeforeDgvValidate,
        msodgeAfterDgvValidate,
        /* These two events bracket the work of DGV::Validate.  These
                let the host set up state into which to receive various
                notification events that happen during DGV::Validate. */
        /* TODO peteren: Some Tuesday order these in with other Validate events. */

        msodgeAfterDgvBecomesInvalid, msodgeAfterDgvInvalidate = msodgeAfterDgvBecomesInvalid, /* TODO peteren: Remove old msodgeAfterDgvInvalidate */
        msodgeAfterDgvBecomesValid,
        /* Passed to DGVS::OnEvent when the DGV's being invalid changes. */

        msodgeAfterDgvCopeWithDgInval,
        /* Passed to DGVS::OnEvent after a DGV has "coped" with the invalidation
                accumulated in a DG.  This is in many ways the hosts's best chance
                to do its own invalidation.  Note that if in coping with the DG
                invalidation the DGV has itself become invalid it's quite possible
                that you'll also be receiving an "msodgeAfterDgvBecomeInvalid" event. */

        msodgeAfterDgvChooseGrfspvs,
        /* Passed to DGVS::OnEvent after DGV::Validate has decided on
                the grfspvs to assign to a particular SPV.  The host can
                change dgexAfterDgvChooseGrfspvs if they want. */

        msodgeQueryDragCopyShape,
                /* Asks if I can drag-copy a shape. */

        msodgeAfterIsShapeOpaque,
                /* Passed after Office decides whether or not a shape is
                        opaque in a particular view.  Host gets to change it. */

        msodgeCalcRchPage,
                /* Asks for a bounding rectangle for dragging. */

        msodgeRequestSelectionForDrag,
                /* Asks for a different selection other than the UI selection
                        for direct manipulation. */

        msodgeAfterPaste,
                /* I just pasted some stuff. */

        msodgeBeforeRegisterShapeDEs,
                /* We're in DGV::Validate and about to register DEs for
                        a shape and (if it has any) its text.  This lets the host
                        mess with the argumest the DGV is going to pass to the DM.
                        In particular this is your chance to or in msofdmrdeBadClip.
                        .dgexBeforeRegisterShapeDEs will be filled out. */

        msodgeRequestCancelDrop,
                /* Passed to DGVS::OnEvent() to give the host a chance to use the
                        shapes being dropped (and are currently in the clipboard
                        drawing) to some use other than the regular drop. */

        msodgeDoAdjustDrchForPaste,
        /* Passed to DGVS::OnEvent when shapes are about to be pasted. This gives
           the host a chance to relatively adjust the coordinates of where the
           shapes will appear. */

        msodgeAfterRequestSpvRedraw,
        /* Passed to DGVS::OnEvent when a particular SPV has been asked to
                redraw.  Anyone using a DM can ignore this, but Word needs it.
                We'll fill out dgveb.dgexAfterRequestSpvRedraw. */

        msodgeDoAddTextboxColorsBMS,
        /* Passed to DGVS::OnEvent just before we add textbox colors into
                IMsoBitmapSurface.  We will have filled out
                dgexAddTextboxColorsBMS. If you set fResult to FALSE we'll not
                add the textbox colors, trusting you to have done it all. */

        msodgeDoDrawTextboxBMS,
        /* Passed to DGVS::OnEvent just before we draw a textbox into
                IMsoBitmapSurface.  We will have filled out dgexDrawTextboxBMS.
                If you set fResult to FALSE we'll not
                draw the textbox, trusting you to have done it all. */

        msodgeDoScrollToShape,
        /* Passed to DGVS::OnEvent to scroll a shape to view */

        msodgeRequestNoOriginW,
        /* Passed to DGVS:OnEvent. Return TRUE if view origin should not
        be accounted for conversion V<->W. */

        msodgeRequestDrag,
        /* Passed to DGVS:OnEvent. Calls back before starting an OLE drag.
         the dgexRequestDrag will contain the IDataObject and IDropSource pointer that is
         used by default. You can the call the DoDragDrop() yourself instead of Escher
         doing it for you. This is used when you need to add native/custom formats
         */

        msodgeDoTitleShape,
        /* Passed to get the title shape of the current slide, used in Power Point
                only.  PowerPoint returns the HSP of the title shape or NULL if there
                is no title shape.  This is passed to DGVS when the background object is
                being drawn. */

        msodgeRequestPosition,
        /* Passed to DGVS::OnEvent to get the position of a RECT in
                host specific way.*/

        msodgeCanDisplayPlaceholderText,
        /* Used by PPT to find out if on Master, then do not display text */

        msodgeRequestCancelResize,
        /* This event fires before a resizing operation occurs to allow the host to cancel
           the action. */

        msodgeRequestInputRcv,
        /* ask host to specify the displayable area for msohsp's in the DGV's HWND
           Used to limit inking input to the specified area of the hwnd
           Sure would be nice to not need this...
        */

        msodgeRequestRcviFromZoom,
        /* Return the "true" RCVI from XL given their constant-pixel cell sizing. */

        msodgeRequestBlockResizeInkCanvas,
        /* Return TRUE if the ink canvas should not allow resize */


        /* ----- to DGVS that should really go to DGSLS if we had one */

        msodgeDgvsFirstShouldBeDgsls/* = 380*/,

        msodgeDoSelectShape, // see dgexDoSelectShape
        msodgeRequestSelectShape,
        msodgeDoUnselectShape, // see dgexDoUnselectShape
        msodgeDoChangeDgslMode, // see dgexDoChangeDgslMode
        msodgeAfterChangeDgsl, // see dgexAfterChangeDgsl
        msodgeDoPaste,
        /* Passed to DGVS::OnEvent to allow hosts to customize pasting
                of shapes. pcsd is a pointer to 'MSOCSD' and pidgsl is a
                pointer to IMsoDrawingSelection in MSODGVEB. */

        msodgeRequestRightClickDragMenuOwnership,
        /*
        Passed to the client to allow them to tell MSO not to display the right click menu.
        If the client takes ownership (returns TRUE in fResult), msodgeDoDisplayRightClickDragMenu
        will be generated and sent
        to the client so that the menu can be displayed and the actions carried out.
        The reason for making these into two separate events is so that DGVDRG::ExecuteShape
        can run correctly, erasing any ghosts created during the drag before the host displays
        and executes any copy or move command.
        */
        msodgeDoDisplayRightClickDragMenu,
        /*
        Passed to client so that they can display the context menu after a right-click
        drag button up. Publisher needs this for format painting, copying and for
        inline shapes as well as to get some other legacy behavior currently not supported
        in the DGVDRG::ExecuteShape method. Leave fResult TRUE to allow MSO to display
        its existing context menu; set it to false to do nothing -- this will simply
        unghost the dragged shapes and return.
        */

        msodgeDoInitDeleteMarkedShapesPaste,
        /* Passed to DGVS:OnEvent to further allow hosts to customize pasting of shapes.
           Not handling this event is the same as returning TRUE, which is previous behavior;
           returning FALSE will leave items on the clipboard drawing after pasting is complete */

        msodgeRequestCancelDrag,
        /* Give the host a chance to not allow a drag on a shape.  The difference
           between this event and msodgeDoMove is that we won't try to bring up the
           context menu for right-clicks. */

        msodgeBeforeShapeFill,
        /* This event is fired before rendering the fill of a shape. It gives the
                host a chance to return the background drawing used by the view.
                This event is used by PPT only. */


        msodgeAfterGetTopLevelGroupShapeBounds,
        /* This event is fired after the final bounds of a top level shape from a
           group has been calculated. This is used in Frontpage. */

        msodgeBeforeSnappingShape,
        /*      This event is fired before calculating the snappng points of a shap. It
                is mainly used for Publisher. When a rotated shape is snapping to the grid,
                Publisher requires the rotated anchor rect to be used, rather than Escher's
                way of using the original anchor rect of the shape.
        */

        msodgeRequestShapeDragRch,
        /* Allows host to specify the rch for a shape to be dragged.  Otherwise, we
           retrieve the rch from the drawing site shape anchor. */

        msodgeBeforeColorScheme,
        /* This event is fired before obtaining the color scheme of a drawing.
                It gives the host a chance to return the drawing whose color scheme is
                used by the view. This event is used by PPT only. */

        msodgeRequestCancelShapeRotateKnob,
        /* Allows the host to cancel display of the rotate knob on a particular shape.
           This is used by PPT to avoid rotate knobs on placeholder shapes.     */

        msodgeQueryGetData,
        /* Asks if I can QueryGetData on shape, used to block dragdrop of motion paths */

        msodgeDoClearUndoStack,
        /* ask the app to clear the undo stack. */

        msodgeRequestOptimizeInkRender,
        /* FUTURE: remove this unused event here and in %XL%\shr\msodg.c (Triage NO O11#322611) */
        msodgeAfterInkCursorDown,
        /* Asks app to do any after InkCursorDown work*/
        msodgeBeforeInkShapize,
        /* Asks app to clean up drawing view before shapizing, used for Word */
        msodgeRequestInkSelectionUpdate,
        /* Asks app to redraw the ink shape after a selection change, only for
           Word to use a particular invalidation function to force redraw */
        msodgeDoInkEditingView,
        /* Asks app if view is still able to edit ink (only if MSODGVSI::fInkingView is true) */
        msodgeExclueInkAnnotationForErase,
        /* Asks app if it needs to exclude ink annotation for erasing */
        msodgeRequestClipInk,
        /* Asks app if we should clip ink at the drawing view boundaries */
        msodgeRequestDgvIsVisible,
        /* Asks app if the drawing view is visible, used to ask Excel if a
           sheet is the topmost sheet in its window */

        msodgeDgvsMax,


        /* ----- Events passed to the DMS, the Display Manager Site. */

        msodgeDmsFirst = 450,

        /* TODO peteren: These shouldn't actually by DGEs, since DMs are supposed
                to be seperate from Drawing. */

        msodgeAfterDmInvalidate,
        /* Passed to DMS::OnEvent when the DM has been invalidated somehow
                such that DM::Update needs to be called. */

        msodgeAfterDmValidate,
        /* Passed to DMS::OnEvent when the DM has decided it no longer needs
                updating.  This doesn't necessarily happen at the end of a call
                to DM::Update, it may happen right near the beginning. */

        /* Note: adding event may require full build for ship */
        msodgeDmsMax,


        /* ----- Events passed to the DGUIS, the Drawing User Interface Site. */

        msodgeDguisFirst = 500,

        msodgeAfterDguiSetRepeat,
        msodgeRequestNumOfTextSelected,
        msodgeRequestGetTextSelection,

        msodgeAfterPickUpFormat,
        msodgeAfterApplyFormat,
        msodgeAfterApplyFormatToDefaults,

        msodgeDoEnterPointerMode, // set fResult to FALSE if we shouldn't enter pointer mode
        msodgeDoExitPointerMode, // set fResult to FALSE if we shouldn't exit pointer mode

        msodgeAfterIsDgcEnabled,
                /* DGUIS may set dgexAfterIsDgcEnabled.fEnabled to FALSE if they
                        have some reason to disable this particular MSODGC at the moment. */
        msodgeAfterIsDgcidEnabled,
                /* DGUIS may set dgexAfterIsDgcidEnabled.fEnabled to FALSE if they
                        have some reason to disable this particular DGCID at the moment. */

        msodgeBeforeLongOperation,
                /* Fire this before a long operation like Picture Disassembly so that
                   the client can put up a wait cursor */

        msodgeBeforeInsertShape,
        /* REVIEW: Currently we fire this before insert a wordart object, so word can switch to page view
                        other app ignores this event. No clear whether to fire this event when inserting other
                        shapes in word */

        msodgeBeforeRepeatInsertShape,
                /* Fire this event before we are about repeat inserting a shape, if the shape is a hostcontrol, this
                   give excel a chance to set the corresponding vpCli */

        msodgeAfterRepeatInsertShape,
                /* Fire this after we are done with repeat inserting a shape, this asks excel to reset vpCli */

        msodgeAfterDguiRunRepeat,

        msodgeDoChangeModes,
                /* Fired whenever the DGUI changes modal states.  Office will
                        have filled out dgex.grfdguimOld and dgex.grfdguimNew.  Host
                        may change grfdguimNew as they wish or set fResult = FALSE to
                        cancel all changes.  This replaces msodgeDoEnterPointerMode
                        and msodgeDoExitPointer Mode. */

        msodgeDoFormatObjectDialog,
                /* Fire this before bringing up the format object dialog so the host
                        gets a chance to bring up its own if it wishes.*/

        msodgeDoHostRecording,
                /* Fire this event for allowing hosts to do VBA recording. */

        msodgeDoGetSelectionForShape,
                /* Fire this to allow the host to choose a selection to select the given
                        shape into. Used by DispShape::Select, defaulting to current one. */

        msodgeBeforeRecording,
                /* Fire this event for allowing hosts to modify some data used for
                macro recording before VBA recording is done by Escher. Also
                see dgexBeforeRecording. */

        msodgeDoGiveDataObject,
                /* Give the host a chance to give a data object that is not
                        necessarily on the clipboard. */

        msodgeDoInsertCagFile,
                /* Insert a file catalog by the Clip Art Gallary based on the passed in information in DGEB.
                        This  event if fired when the users presses insert in Clip Art Gallery. Hosts should
                        process this event like an InsertPictureCommand. */

        msodgeBeforePickUpFormat,
        msodgeBeforeApplyFormat,

        msodgeDoDrawLineStyleMenu,
                /* Needed by Publisher. Give the host a chance to draw the currently selected menu
                item in the line style menu in the format object dialog. Needed by Publisher to
                display the string "Border Art" when the Border Art swatch is selected */

        msodgeAfterDiagramCreated,
                /* Tell the host that we finish creating a Diagram. */

        msodgeDoInsertTextEffectDlg,
                /*Fire this before bringing up the insert text effect gallery dialog so the host
                        gets a chance to refuse bringing it up. */

        msodgeDoExitFillEffectsDlg,
                /* Fire this before the Fill Effects dialog is about to come down due to user clicking
                        OK. Parameters in dgexDoExitFillEffectsDlg include the extended color information for color1 and color2
                        in the gradient tab. Gives the host a chance to check the values of these colors
                        (depending on current printing mode) and also a chance to change the values of them
                        if needed. Return FALSE in fResult if you don't want the dialog to exit else return TRUE
                        (and maybe update color values) to exit the dialog. Needed by Publisher*/

        msodgeDoFilterDefaultTextureFillColor,
                /* Fire this event when the Fill Effects dialog is coming down and the user selected
                        a texture fill. Needed by Publisher so that they can change the default color
                        applied for a texture to an appropriate spot color.*/

        msodgeDoResolveSpotColor,
                /* This event gets fired from the More Colors dialog to give the host a chance
                        to resolve spot colors. */

        msodgeBeforeInkingStart, // REVIEW sabrinan (michkim): Is this still used? The comment doesn't seem to fit the name either.
        /* ask the app to clear the undo stack. */

        msodgeAfterInkingFinished, // REVIEW sabrinan (michkim): Is this still used? The comment doesn't seem to fit the name either.
        /* ask the app to clear the undo stack. */

        msodgeBeforeInsertInkCanvas,
        /* ask Word to change selection cursor */

        msodgeEnsureInkAnnotationVisible,
        /* ask the app to ensure ink annotations are visible */

        msodgeDoExitInkingModesFromAllDrawingGroups,
        /* ask the app to exit ink mode from all their presentations/books/documents. */

        /* Note: adding event may require full build for ship */
        msodgeDguisMax,


        /* ----- Events passed to the DGSLS, the Drawing Selection Site. */

        msodgeDgslFirst = 600,

        msodgeAskDgslAnchoredOnOnePage,
                /* Passed to DGSLS::OnEvent when Office needs to know if all the
                        Shapes references by a particular DGSL are on the same "page".
                        Office doesn't really understand "pages"; it just knows that they
                        exist, that two Shapes can be on the same page or not, and that
                        certain commands (like Group) only work on sets of objects that
                        are all on the same page */
                /* TODO peteren: Are we going to use this? */

        msodgeDgslMax,

        } MSODGE;

typedef int MSODGEX; // Hey! Not a real type!
        /* MSODGEX is not really a type; it's the hungarian we use to name
                all the gross fields inside MSODG*EB structures.  I just stuck
                the type in here in case someone tagged on it. */

/* TODO davebu: Use the MSODGVEB instead of this little structure. */

typedef struct _MSODGEBDRAG
        {
        int cShapes;
        RECT rchBounds;
        } MSODGEBDRAG;


#define msogrfescNil     0x00000000
#define msofescSelBefore 0x00000001
/* Above is set if at least one shape selected before selection ops done. */
#define msofescSelAfter  0x00000002
/* Above is set if at least one shape selected after selection ops done. */


/* TODO peteren: move this somewhere more general */
/* MM is Mouse Message.  These are designed to correspond
        exactly with the flags in the wParam value passed by Windows with
        mouse messages. */
typedef WPARAM MSOMM; // Mouse Message
#define msommNil ((MSOMM)0)
#define msogrfmmNil        0x00000000
/* The low 16 bits of an MSOMM are reserved to line up with MK_ values
        that are passed in the WPARAM of mouse messages. */
#define msofmmLButton      0x00000001 // MK_LBUTTON
#define msofmmRButton      0x00000002 // MK_RBUTTON
#define msofmmShift        0x00000004 // MK_SHIFT
#define msofmmControl      0x00000008 // MK_CONTROL
#define msofmmMButton      0x00000010 // MK_MBUTTON
/* TODO peteren: Should these next two, which are only ever used on
        the MAC, but which there's no good reason not to have #defined
        on Windows (certainly we'd rather not have collisions) be under
        #if MAC? */
#define msofmmFGSwitch     0x00002000 // MK_FGSWITCH (Mac Only)
#define msofmmOption       0x00004000 // MK_OPTION (Mac Only)
#define msofmmCommand      0x00008000 // MK_COMMAND (Mac Only)
/* We put our own flags in the high 16 bits of the MSOMM. */
#define msofmmAlt          0x00010000 // Win Only (See msofmmOption)
#define msofmmDoubleClick  0x00020000

#define msogrfmmInWParamWin \
        (msofmmLButton | msofmmRButton | msofmmMButton | msofmmShift | \
        msofmmControl)
#define msogrfmmInWParamMac \
        (msofmmLButton | msofmmRButton | msofmmShift | msofmmControl | msofmmOption | \
        msofmmCommand | msofmmFGSwitch)
#define msogrfmmInWParam msogrfmmInWParamWin

MSOAPI_(MSOMM) MsoMmBuild(UINT wm, WPARAM wParam, BOOL fShift,
        BOOL fControl, BOOL fAlt, BOOL fDoubleClick);

MSOAPI_(BOOL) MsoFMmImpliesContextMenu(MSOMM mm);

// RUT == Rule type
typedef enum
        {
        msorutConnector = -1,
        msorutAlign              = -2,
        msorutArc                = -3,
        msorutCallout    = -4,
        msorutNone               = 0,
        msorutAny                = 0,
        msorutFirstHost = 1             // Host rules have positive types
        } MSORUT;

// RUC == Rule Change
typedef enum
        {
        msorucAddHSP,                           // HSP addition
        msorucDelHSP,                           // HSP subtraction
        msorucParam                                     // parameters changed
        } MSORUC;

// CX == Connection Shape (Values for GetProxy())
typedef enum
        {
        msocxspConnector = 0,   // HSP of the connector
        msocxspStart = 1,                       // HSP of the attached start shape
        msocxspEnd = 2                          // HSP of the attached end shape
        } MSOCXSP;

// CXK the kind of connection sites
typedef enum
        {
        msocxkNone = 0,                 // No connection sites
        msocxkSegments = 1,             // Connection sites at the segments
        msocxkCustom = 2,                       // Sites defined by msopidPConnectionSites
        msocxkRect = 3,                 // Use the connection sites for a rectangle
        } MSOCXK;

// DGSLK = DrawinG SeLection Kind.
typedef enum
        {
        msodgslkNormal,      // Normal Selection Mode.
        msodgslkRotate,      // Rotate selection mode
        msodgslkReshape,         // Reshape Selection Mode.
        msodgslkConnectorDoesNotSeemUsed, // Connector Selection Mode.
        msodgslkWrapPolygon, // Display and edit of wrap polygons.
        msodgslkTextEdit,    // Text Edit Mode.
        msodgslkScale,       // Scale Selection Mode.
        msodgslkCrop,        // Cropping Selection Mode.
        } MSODGSLK;

// FDGVI = Drawing View Invalidation flag
#define msogrfdgviNil            0
#define msofdgviFull             (1<<0)
        /* Invalidates everything. */
#define msofdgviShapesCheck      (1<<1)
        /* Means we need to check all the Shapes to see which have been
                invalidated (by edits to the Drawing) and fix any that need it. */
#define msofdgviSelectionCheck   (1<<2)
        /* Check the selectedness (if so and how) of Shapes. */
#define msofdgviSelectionFull    (1<<3)
        /* Redo all of the selectedness of Shapes */
// USE ME NEXT                   (1<<4)
#define msofdgviSlide            (1<<5)
        /* TODO peteren: Comment? */
#define msofdgviOneShape         (1<<6)
        /* Means we should invalidate the shape specified in the hspv argument. */
#define msofdgviEffects          (1<<7)
        /* Rebuild all the GEL effects */
#define msofdgviOneShapeZOrder   (1<<8)
        /* Means we should invalidate the shape specified in the hspv argument
                for a Z-Order change. */
#define msofdgviOneShapeEffect   (1<<9)
        /* Rebuild the GEL effect of the given shape. */

/* MSOSPVSK = Shape in View Selection Kind */
/* These are an enumeration for the different kinds of visible
        selection "adornments" there are for Shapes. */
#define msospvskInvisible           0
#define msospvskActiveBorder        1
#define msospvskActiveBorderNoDots  2
#define msospvskNormalBorder        3
#define msospvskNormalBorderAdjust  4
#define msospvskNormal              5
#define msospvskNormalAdjust        6
#define msospvskRotate              7
#define msospvskReshape             8
#define msospvskWrapPolygon         9
#define msospvskCrop                10
#define msospvskCanvas              11
#define msospvskLocked              12
#define msospvskNormalBorderNoDots  13
#define msospvskMax                 14
#define msospvskTypeScale           15
/* Hey!  msospvsk's are stored in 4 bits inside msogrfspvs's */

/* GRFSPVS = Shape in View Selection flags */
/* These flags describes the selection adornments drawn on a particular
        SPV (that is, on a Shape in a DGV).  This ULONG (along with the
        Shape itself) is supposed to as nearly as possible COMPLETELY
        describe the shape's selected state. */

#define msogrfspvsNil         0
#define msogrfspvsNotSelected 0
#define msogrfspvsMaskSpvsk   0x0000000F
        /* In the low 4 bits of a grfspvs we store an spvsk. */
#define msofspvsSelected      (1<<4)
#define msofspvsHasButtons    (1<<5)
#define msofspvsShowTextLabel (1<<6)
#define msofspvsDisableCrop   (1<<7)

/* PeterEn thinks we should get rid of msofspvsSelected and replace it
   with msofspvsSelectedDeep and msofspvsSelectedRoot that mean precisely
   that this shape is selected (deep or root) in this DGV's selection,
   and then write code at the end of DGV::Validate to verify that the
   bits actually do line up with the selection.  If you find yourself
   tempted to reference msofspvsSelected, you might consider making this
   change. */
/* Hey! We store grfspvs's in 8 bits internally */
/* Begin old names for spvsk's */
#define msogrfspvsTypeInvisible          msospvskInvisible
#define msogrfspvsTypeActiveBorder       msospvskActiveBorder
#define msogrfspvsTypeActiveBorderNoDots msospvskActiveBorderNoDots
#define msogrfspvsTypeNormalBorder       msospvskNormalBorder
#define msogrfspvsTypeNormalBorderAdjust msospvskNormalBorderAdjust
#define msogrfspvsTypeNormal             msospvskNormal
#define msogrfspvsTypeNormalAdjust       msospvskNormalAdjust
#define msogrfspvsTypeRotate             msospvskRotate
#define msogrfspvsTypeReshape            msospvskReshape
#define msogrfspvsTypeWrapPolygon        msospvskWrapPolygon
#define msogrfspvsTypeCrop               msospvskCrop
#define msogrfspvsTypeCanvas             msospvskCanvas
#define msogrfspvsTypeLocked             msospvskLocked
#define msogrfspvsTypeNormalBorderNoDots msospvskNormalBorderNoDots
#define msogrfspvsTypeScale              msospvskTypeScale
#define msogrfspvsTypeMask               0x0000000F
/* End old names for spvsk's */


// DGCXM = Drawing Context Menus
typedef enum
        {
        msodgcxmGeneralShape,
        msodgcxmCurveShape,
        msodgcxmPictureShape,
        msodgcxmOLEObjectShape,
        msodgcxmConnectorShape,
        msodgcxmTextEffectShape,
        msodgcxmRotateMode,
        msodgcxmEditCurveSegment,
        msodgcxmEditCurveVertex,
        msodgcxmOLEControlShape,
        msodgcxmOrgChart,
        msodgcxmDiagram,
        msodgcxmCanvasShape,
        msodgcxmInlineCanvas,
        msodgcxmInvalid,
        msodgcmMax,
        } MSODGCXM;

// DGRST = DrawinG Related Shape Type
typedef enum
        {
        msodgrstNextRoot,
        msodgrstNextAll,
        msodgrstParent,
        msodgrstRoot,
        msodgrstRootInCanvas,
        } MSODGRST;


/******************************************************************************
        Shape Properties

        For simplicity, every shape in a drawing supports the same number of
        properties, although they may have no meaning for some shapes.

        Each property is 32 bits in length. Although somewhat restrictive, 32
        bits allows for the common variants of property types. Supported
        property types include BOOL, long, unsigned long, string, array,
        enumeration, color and pointer to shape. The MSOPTYPE enumeration
        lists all the possible types of shape properties.

        Each shape property is identified by unique office property id
        (MSOPID). The FSetProp and FetchProp methods are used to access
        properties by property id.      The property table in currently in
        MSODRP.H lists all the properties defined for shapes.

        A set of related properties are grouped into property set. Examples of
        property sets include shape fill properties and connections
        properties. "C" structures are defined each property set. The
        PFetchPropSet and SetPropSet methods are used to access property sets.
        It is expected that it will often be more convenient to access
        properties in sets rather than individually. Since property sets can be
        cached, it will often be faster to Fetch sets rather than individual
        properties. MSODRP.H list all the properties set

********************************************************************** RICKH **/

// MSOANCHOR
typedef enum
        {
        msoanchorTop,
        msoanchorMiddle,
        msoanchorBottom,
        msoanchorTopCentered,
        msoanchorMiddleCentered,
        msoanchorBottomCentered,
        msoanchorTopBaseline,
        msoanchorBottomBaseline,
        msoanchorTopCenteredBaseline,
        msoanchorBottomCenteredBaseline
        } MSOANCHOR;

// MSOCDIR
typedef enum
        {
        msocdir0,               // Right
        msocdir90,              // Down
        msocdir180,             // Left
        msocdir270              // Up
        } MSOCDIR;
#define msoMaxConnections 64

// MSOCXSTYLE
typedef enum
        {
        msocxstyleNinch = -1,
        msocxstyleStraight = 0,
        msocxstyleBent,
        msocxstyleCurved,
        msocxstyleNone
        } MSOCXSTYLE;

// MSOTXFL -- text flow
typedef enum
        {
        msotxflHorzN,           //Horizontal non-@
        msotxflTtoBA,           //Top to Bottom @-font
        msotxflBtoT,            //Bottom to Top non-@
        msotxflTtoBN,           //Top to Bottom non-@
        msotxflHorzA,           //Horizontal @-font
        msotxflVertN,           //Vertical, non-@
        msotxflMax,     msotxflLast = msotxflMax - 1,
        } MSOTXFL;

// MSOTXDIR is hungarian for text direction (needed for Bi-Di support)
typedef enum
        {
        msotxdirLTR,                    // left-to-right text direction
        msotxdirRTL,                    // right-to-left text direction
        msotxdirContext,                // context text direction
        } MSOTXDIR;

// MSOPRH -- position relative horizontal
typedef enum
        {
        msoprhMargin = 0, msoprhMin = msoprhMargin,
        msoprhPage,
        msoprhText,
        msoprhChar,
        msoprhMax,
        } MSOPRH;

#define MsoFValidPrh(prh) (FBetween((prh), msoprhMin, msoprhMax - 1))

// MSOPV -- position vertical
typedef enum
        {
        msopvAbs = 0, msopvMin = msopvAbs,
        msopvTop,
        msopvCenter,
        msopvBottom,
        msopvInside,
        msopvOutside,
        msopvMax,
        } MSOPV;

// MSOPRV -- position relative vertical
typedef enum
        {
        msoprvMargin = 0, msoprvMin = msoprvMargin,
        msoprvPage,
        msoprvText,
        msoprvLine,
        msoprvMax,
        } MSOPRV;

#define MsoFValidPrv(prv) (FBetween((prv), msoprvMin, msoprvMax - 1))


// MSOSPCOT -- Callout Type
#define msospcotNinch                   (0)
#define msospcotMin                     (1)
#define msospcotRightAngle              (1)
#define msospcotOneSegment              (2)
#define msospcotTwoSegment              (3)
#define msospcotThreeSegment            (4)
#define msospcotMax                     (5)

// MSOSPCOA -- Callout Angle
#define msospcoaMin                     (-1)
#define msospcoaAny                     (0)
#define msospcoa30                      (1)
#define msospcoa45                      (2)
#define msospcoa60                      (3)
#define msospcoa90                      (4)
#define msospcoa0                       (5) /* This one's not stored in any SPs. */
#define msospcoaMax                     (5)

// MSOSPCOD -- Callout Drop
#define msospcodMin                     (0)
#define msospcodTop                     (0)
#define msospcodCenter                  (1)
#define msospcodBottom                  (2)
#define msospcodSpecified               (3)
#define msospcodMax                     (4)

/*Used in the built in office text effect preset to store the strings in the
  intl dll.  */
#define MSOPRESETSZFLAG                 (0x80000000)
#define MSOPRESETBLIPTAGFLAG            (0X40000000)
#define MSOPRESETSTR(ids)               (MSOPRESETSZFLAG|ids)
#define MSOPRESETBLIP(tag)              (MSOPRESETBLIPTAGFLAG|tag)

/*----------------------------------------------------------------------------
        Shape path representation

        A shape consists of a number of vertices which are interpreted according
        to either the pSegmentInfo property in the geometry property set or, if
        the segment info array is absent or empty, the shapePath property.  The
        shapePath property is an enumeration whose sole purpose is to reduce
        storage requirements.  In all cases it is possible to construct a
        segment info array which is equivalent to the shapePath.

        shapePath indicates a closed or open figure whose sides consist of just
        one form of line segment - straight line, curve, elliptical quadrant or
        arc.

        See msogel.h for discussion of the segment info array format - consult the
        definition of MSOPATHINFO.  In general do NOT use the msoshapeComplex
        type - it is always sufficient just to set the pSegmentInfo array and
        leave shapeType as it was, and so the property set maximum is set to
    exclude this value.
------------------------------------------------------------------- JohnBo -*/
// MSOSHAPEPATH
typedef enum
        {
        msoshapeLines,        // A line of straight segments
        msoshapeLinesClosed,  // A closed polygonal object
        msoshapeCurves,       // A line of Bezier curve segments
        msoshapeCurvesClosed, // A closed shape with curved edges
        msoshapeComplex,      // pSegmentInfo must be non-empty
        msoshapeMin = msoshapeLines,
        msoshapeMax = msoshapeComplex
   } MSOSHAPEPATH;

typedef enum
{
  msowrapSquare,
  msowrapByPoints,
  msowrapNone,
  msowrapTopBottom,
  msowrapThrough,
} MSOWRAPMODE;

// MSOBWMODE is hungarian for Black and White rendering Modes
typedef enum
        {
        msobwColor,          // only used for predefined shades
        msobwAutomatic,      // depends on object type
        msobwGrayScale,      // shades of gray only
        msobwLightGrayScale, // shades of light gray only
        msobwInverseGray,    // dark gray mapped to light gray, etc.
        msobwGrayOutline,    // pure gray and white
        msobwBlackTextLine,  // black text and lines, all else grayscale
        msobwHighContrast,   // pure black and white mode (no grays)
        msobwBlack,          // solid black
        msobwWhite,          // solid white
        msobwDontShow,       // object not drawn
        msobwNumModes        // number of Black and white modes
        } MSOBWMODE;


/*----------------------------------------------------------------------------
        Utility APIS.
------------------------------------------------------------------- JohnBo -*/
/* Given a text color cr, the grfdraw from an MSODC (see the definition below)
        and an MSOBWMODE from a shape return the appropriate color which should be
        used to display the text.  This resolves the MSOBWMODE information in the
        "standard" way and deals with all the relevant recolor control flags in
        grfdraw.  Note that the color must not be a scheme color or palette index -
        it must be an RGB value. */
MSOAPI_(COLORREF) MsoCrFromTextColor(COLORREF cr, COLORREF crDrawColor,
        ULONG grfdraw, int ccrScheme, const COLORREF *pcrScheme, MSOBWMODE bwMode, BOOL fEmbossText);

/*      MsoFIsClsidPicture
        Returns TRUE if the passed in CLSID is one that we consider to be a
        picture (e.g. a WMF, paintbrush image, DIB). */
MSOAPI_(BOOL) MsoFIsClsidPicture(REFCLSID rclsid);

/*      Returns TRUE if the passed in CLSID is one that we consider to be an
        OLE orgchart object. */
MSOAPI_(BOOL) MsoFIsClsidOrgChart(REFCLSID rclsid);

/*      Returns TRUE if the passed in CLSID is one that we consider to be a
        MS drawing object. */
MSOAPI_(BOOL) MsoFIsClsidMSDraw(REFCLSID rclsid);

/*      Returns TRUE if regkey DoNotAllowOrgChartConversion is 0. */
MSOAPI_(BOOL) MsoFAllowOrgChartConversion();

/*      Returns TRUE if regkey DoNotAllowDrawOLEConversion is 0. */
MSOAPI_(BOOL) MsoFAllowDrawOLEConversion();

/*      MsoFIsClsidOptimizable
        Returns TRUE if the passed in CLSID is an OLE object that can have its
        OLE data stripped for optimization */
MSOAPI_(BOOL) MsoFIsClsidOptimizable(REFCLSID rclsid);

#define msocopMax 64            // The maximum number of properties a property set can have
#define msocOPSSIZE msocopMax


// MSOPTYPE is hungarian for property type
//
// All properties must be one of the following types
//
typedef enum
{
        msoptypeBool,           // BOOL
        msoptypeLong,           // LONG
        msoptypeULong,          // ULONG
        msoptypeSZ,             // CHAR*
        msoptypeEnum,           // Any enumeration
        msoptypeColor,          // COLORREF
        msoptypePPlex,          // IMsoArray*
        msoptypePBLIP,          // IMsoBlip*
        msoptypePRule,          // IMsoRule*
        msoptypeMSOHSP,         // MSOHSP
        msoptypePIHlink,        // IHLink*
        msoptypeWZ,             // WCHAR*
        msoptypeUnknown,        // Unknown property (from a future version)
        msoptypeGrfBlip,        // MSOBLIPFLAGS
        msoptypeMovie,          // QuickTime movie
        msoptypePMSOSP,         // MSOSP*.  Like MSOHSP but never copied
        msoptypePInkData,       // IMsoInkData*
        msoptypeMax,            // Always last
} MSOPTYPE;


// Ninch - no input, no change values for each property type.
//
//      Set with a ninch value has no effect. Fetch on a mixed selection
// returns a ninch value.
//
#define msofNinch                       ((BOOL)-1)
#define msolNinch                       ((LONG)MINLONG+1)
#define msouNinch                       ((ULONG)(LONG)MAXDWORD)
#define msoszNinch                      ((char*)(LONG_PTR)(LONG)MAXDWORD)
#define msoEnumNinch                    ((ULONG)(LONG)MAXDWORD)
#define msopiaNinch                     ((IMsoArray*)(LONG_PTR)(LONG)MAXDWORD)
#define msopiruNinch                    ((IMsoRule*)(LONG_PTR)(LONG)MAXDWORD)
#define msopibNinch                     ((IMsoBlip*)(LONG_PTR)(LONG)MAXDWORD)
#define msohspNinch                     ((MSOHSP)(LONG)MAXDWORD)
#define msopihlNinch                    ((IHlink*)(LONG_PTR)(LONG)MAXDWORD)
#define msowzNinch                      ((WCHAR*)(LONG_PTR)(LONG)MAXDWORD)
#define msoiidNinch                     ((IMsoInkData*)(LONG_PTR)(LONG)MAXDWORD)


// Reset values for properties
//
//      Set with a reset value removes the property from a property list.
// Equivalent to MsoResetProp call.
//
#define msofReset                       ((BOOL)-2)
#define msolReset                       ((LONG)MINLONG+2)
#define msouReset                       ((ULONG)(LONG)MAXDWORD-1)
#define msoszReset                      ((char*)(LONG_PTR)(LONG)MAXDWORD-2)
#define msoEnumReset                    ((ULONG)(LONG)MAXDWORD-1)
#define msocrReset                      ((COLORREF)0x80000002)
#define msopiaReset                     ((IMsoArray*)(LONG_PTR)(LONG)MAXDWORD-2)
#define msopiruReset                    ((IMsoRule*)(LONG_PTR)(LONG)MAXDWORD-2)
#define msopibReset                     ((IMsoBlip*)(LONG_PTR)(LONG)MAXDWORD-2)
#define msohspReset                     ((MSOHSP)(LONG_PTR)(LONG)MAXDWORD-2)
#define msopihlReset                    ((IHlink*)(LONG_PTR)(LONG)MAXDWORD-2)
#define msowzReset                      ((WCHAR*)(LONG_PTR)(LONG)MAXDWORD-2)
#define msoiidReset                     ((IMsoInkData*)(LONG_PTR)(LONG)MAXDWORD-2)


// MSOPINFO     is hungarian for Property INFO
//
// See MsoFLoadPropName for the name of the property
// Be sure to update vopinfo* in drpinfof.cpp if you
// change the fileds in this structure.
typedef struct
{
        MSOPTYPE                        optype;                         //
        LONG                            lDefault;                       // The default property value
        LONG                            min;                            // The minimum allowable property value
        LONG                            max;                            // The maximum allowable property value
        struct
                {
                ULONG                   fMaster:1;                      // Could be a master property
                ULONG                   fDefault:1;                     // Tracked in default properties
                ULONG                   fStyle:1;                       // Tracked in style properties
                ULONG                   fRtf:1;                         // Property used in RTF
                ULONG                   fComplex:1;                     // Property needs be freed by MsoFreeProp
                ULONG                   fLocal:1;                       // Property requires special fetching code
                ULONG                   fBuiltIn:1;                     // Property is found in shape definitions
                ULONG                   fCache:1;                       // Property could be cached in the shape
                ULONG                   fWrap:1;                        // Property will cause wrap polygons to recalc.
                ULONG                   fAfter97:1;                     // Property added since Office 97 ship
                ULONG                   fSaved:1;                       // Property persisted in binary file
                ULONG                   fVisual:1;                      // Property affects redering (NOT position)
                };
} MSOPINFO;

typedef void* MOVIE;

// MSOPID, MSOPSID, MSOPS... are defined in the following generated header
//
#ifndef MSODRP_H
#include <msodrp.h>
#endif

#define msopsidNil ((MSOPSID)0xFFFF)

// Old names for
typedef MSOPSXfrm MSOPSTransform;
#define msopsidTransform msopsidXfrm


// MSOPSINFO is hungarian for Property Set INFO
//
// A property set consist of range non-bools followed
// by a range of bools. The OPIDs of the non-bools are contiguous
// and the OPIDs of the bools are contiguous.
//
typedef struct
{
        ULONG                           cbSize;                         // The size of the property set
        MSOPID                          opidFirst;                      // The first non-bool property
        ULONG                           copid;                          // The number of non-bool props
        MSOPID                          opidFirstBool;                  // The first bool prop
        ULONG                           copidBool;                      // The number of bool props
        ULONG                           copidBoolIn97;                  // The number of bool props we had in O97
        struct
                {
                ULONG                   fGroup:1;                       // The property set can be applied to a group
                ULONG                   fMaster:1;                      // Could follow the master
                ULONG                   fLocal:1;                       // Could requires special handling
                ULONG                   fBuiltIn:1;                     // Could be built in
                ULONG                   fCache:1;                       // Could be cached in the shape
                };
} MSOPSINFO;


/* MSOPIDEN is an OPID generator. The OPID enum is not contiguous.
        This enumerator generates the appropriate valid OPIDs. It doesn't
        apply to specific shape. */
typedef struct
{
        union
                {
                ULONG grf;
                struct
                        {
                        ULONG                           fPropSet : 1;   // Enum the PIDs in a property set
                        ULONG                           fDefault : 1;   // Enum defaults properties only
                        ULONG                           fStyle : 1;     // Enum style properties
                        ULONG                           fMaster : 1;    // Enum master properties
                        ULONG                           fAll : 1;       // Enum all properties
                        ULONG                           res : 27;       // Reserved
                        };
                };
        MSOPSID                         opsid;                  // Current propset
        const MSOPSINFO*                popsinfo;               // Current prop set info
        MSOPID                          opid;                   // Current opid
        const MSOPINFO*                 popinfo;                // Current propert info;
        int                             iop;                    // The current index into the property set range
        int                             ispp;                   // Just a counter
} MSOPIDEN;


/* MSOPENUM is OP value generator. It applies to specific shape. */
typedef struct
{
        MSOPINFO pinfo;
        MSOPID pid;
        unsigned long op;
        void *ppv1;  // These fields are used only internally by Office.
        void *ppv2;  // Don't touch.
        void *ppv3;
        void *ppv4;
} MSOPENUM;


/* Return the Office Property Set ID given the property identified by 'opid'. */
__inline MSOPSID MsopsidGet(MSOPID opid) { return (MSOPSID)((long)opid / msocOPSSIZE); }

/* Return a pointer to the Office Property Information given the Property
   identified by 'opid'. */
MSOAPI_(const MSOPINFO*) MsoPopinfoGet(MSOPID opid);

/* Return a pointer to the Office Property Set Information given the Property
   Set identified by 'opsid'. */
MSOAPI_(const MSOPSINFO*) MsoPopsinfoGet(MSOPSID opsid);

/* Return TRUE iff the Property Value pointed to by 'pop' is the NINCH value
   for the Property identified by 'opid', otherwise return FALSE. */
MSOAPI_(BOOL) MsoFIsNinch(MSOPID opid, void* pop);

/* Return TRUE iff the Property Value pointed to by 'pop' is the default value
   for the Property identified by 'opid', otherwise return FALSE. */
MSOAPIXX_(BOOL) MsoFIsDefault(MSOPID opid, void* pop);

/* Return TRUE iff the Property Value pointed to by 'pop' is the reset value
   for the Property identified by 'opid', otherwise return FALSE. */
MSOAPIX_(BOOL) MsoFIsReset(MSOPID opid, void* pop);

/* Set the passed in structure to Ninch values for the specified property */
MSOAPI_(void) MsoGetPropNinch(MSOPID opid, void* pop);

/* Set the passed in structure to reset values for the specified property */
MSOAPI_(void) MsoGetPropReset(MSOPID opid, void* pop);

/* Set the passed in structure to the default values for the specified property */
MSOAPI_(void)   MsoGetPropDefault(MSOPID opid, void* pop);

/* Set the passed in structure to Ninch values for the specified property set */
MSOAPI_(void) MsoGetPropSetNinch(MSOPSID opsid, void* pops);

/* Set the passed in structure to Default values for the specified property set */
MSOAPI_(void) MsoGetPropSetDefault(MSOPSID opsid, void* pops);

/* Return TRUE iff the Property identified by 'opid' is a valid Property ID
   (ie the Property exists), otherwise return FALSE. */
MSOAPI_(BOOL) MsoFIsValidOpid(MSOPID opid);

/* Return TRUE iff the value is valid for the property . */
MSOAPI_(BOOL) MsoFIsValidValue(MSOPID opid, const void* pop);

/* Return TRUE iff the Property Set identified by 'opsid' is a valid Property
   set ID (ie the Property Set exists), otherwise return FALSE. */
MSOAPIX_(BOOL) MsoFIsValidKnownOpsid(MSOPSID opsid);

/* Deep copy the property including any objects it may point to.  */
MSOAPI_(BOOL) MsoFCopyProp(MSOPID opid, const void* popFrom, void* pop);

/* Deep copy the property set including the objects pointed to by
   properties in the property set. */
MSOAPI_(BOOL) MsoFCopyPropSet(MSOPSID opsid, const void* popsFrom, void* pops);

/* Is a property equal to another */
MSOAPI_(BOOL) MsoFIsPropEqual(MSOPID opid, const void* popA, const void* popB);

/* Is a property set equal to another */
MSOAPI_(BOOL) MsoFIsPropSetEqual(MSOPSID opsid, const void* popsA, const void* popsB);

/* Free the contents of a property. Helper function that will
   free any objects pointed to by the property.  */
MSOAPI_(void) MsoFreeProp(MSOPID opid, const void* pop);

/* Free the contents of a property set. Helper function that will
   free any objects pointed to by the property set.  */
MSOAPI_(void) MsoFreePropSet(MSOPSID opsid, const void* pops);

#if DEBUG
/* Write a BE record for a property */
MSOAPIXX_(BOOL) MsoFWriteBeProp(MSOPID opid, const void* pops, HMSOINST hmsoinst,
        LPARAM lParam);

/* Write a BE record for a property set */
MSOAPI_(BOOL) MsoFWriteBePropSet(MSOPSID opsid, const void* pops, HMSOINST hmsoinst,
        LPARAM lParam);
#endif

/* Get a name for a given MSOPID. Specify, if you want the debug name or
   the RTF name. The which is a buffer of length cch.  Returns the
        actual length of the string.  The returned string is guaranteed to be
        null-terminated. */
typedef enum {msolpnDebug,      msolpnRTF} MSOLPN;
MSOAPI_(BOOL) MsoFLoadPropName(MSOPID opid, MSOLPN lpn, char* szName, int cchMax);

/* Enum all of MSOPIDs of the a specified prop set. */
MSOAPI_(void) MsoInitPropSetOPIDEN(MSOPIDEN *popiden, MSOPSID opsid);

/* Enum all of MSOPIDs defined. */
MSOAPIX_(void) MsoInitOPIDEN(MSOPIDEN *popiden);

/* Generate the opids */
MSOAPI_(BOOL) MsoFEnumOPIDEN(MSOPIDEN *popiden);

/*      Saves the property set 'opsid' from the buffer 'pops' into 'pistm'.
        NOTE: The caller retains the ownership of the properties in 'pops'. */
MSOAPI_(BOOL) MsoFSavePropSet(IStream *pistm, MSOPSID opsid, const void *pops);

/*      Loads the property set 'opsid' from 'pistm' into the buffer 'pops'.
        NOTE: The caller would own the properties in 'pops' if this call is
        successful. */
MSOAPI_(BOOL) MsoFLoadPropSet(IStream *pistm, MSOPSID opsid, void *pops);

/* Return true if the prop opid will be fetched from the built-in OPT when
   calling FetchProp on the hsp. */
MSOAPI_(BOOL) MsoFPropFromBuiltIn(MSOHSP hsp, MSOPID opid);

/* Return true iff the prop 'opid' will be fetched from the OPT of the shape
        when calling FetchProp on the 'hsp' else return false. */
MSOAPI_(BOOL) MsoFPropFromPropertyTable(MSOHSP hsp, MSOPID opid);

/* Write out a BE reocrd for a MSOHUR */
#if DEBUG
MSOAPI_(BOOL) MsoFWriteBeMSOHUR(HMSOINST hinst, MSOHUR hur, LPARAM lParam,
        BOOL fSaveObj, interface IMsoDrawing *pidg);
#endif

/****************************************************************************
        Shape Properties, Peter's MSOSPP version

        This is really just alternate packaging for the Shape properties.
        There are a few problems with the old ones, all of which resulted from
        the incremental changes they were forced to undergo during development.

        1.  The names were chosen before public Office naming conventions were
            made solid.
        2.  BLIP properties broke the notion that all properties fit into 32 bits.
            It's terribly convienient and actually no memory loss to treat
                 properties as coming in 8-byte chunks (OPTEs, really).  So now we
                 export these.
        3.  BLIP properties also break the notion that Shape properties are
            autonomous.  A Shape property may be a BID that assumes a particular
                 Drawing Group (user document).

        I'm not sure these will actually work.  But I thought I'd throw them in
        and write some new routines that needed writing in terms of them and see
        what happens.
****************************************************************** PeterEn **/

/* MSOSPPID -- Shape Property ID */
typedef MSOPID MSOSPPID; // Shape Property ID
        /* TODO peteren: We call these MSOPIDs everywhere else.  Change? */
        /* TODO peteren: Ensure property IDs never get above FFFF? */
#define msosppidMax ((MSOSPPID)0xFFFF)
#define msosppidNil ((MSOSPPID)0xFFFF)
#define msopidNil ((MSOPID)0xFFFF)
/* MSOSPPI -- Shape Property Info */
/* TODO peteren: Rename MSOPINFO to this */
typedef MSOPINFO MSOSPPI;

MSOAPIX_(const MSOSPPI *) MsoPsppiFromSppid(MSOSPPID sppid);

/* MSOSPPSI -- Shape Property Set Info */
/* TODO peteren: Rename MSOPSINFO to this */
typedef MSOPSINFO MSOSPPSI;

MSOAPIX_(const MSOSPPSI *) MsoPsppsiFromSppid(MSOSPPID sppid);

/* MSOSPP -- Shape Property */
/* Shape Properties are stored, fetched, and set as lists of MSOSPPs.
        This 8-byte structure contains the MSOSPPID for the property in question,
        a few interesting bits, and the actual data of the property.
        The interesting bits are:
         fComplex  Means that this property has stuff hanging off of it that
                   might need to be freed, marked, released, etc.
         fNinch    Means that this property is "No Input No Change".  For example,
                   if you fetch a property from a selection of multiple Shapes,
                   and it turns out the Shapes have DIFFERENT values for the
                   property, you'll get back an MSOSPP with fNinch == TRUE.
         fReset   Means that you should reset the SPPID and ignore lValue.
                                  Only valid for apply.
         fIncr     Means that the value in 'lValue' is the increment value that
                                  value that should be added to the current property value.
         lValue   The "value" of the property is stored in the following fields
*/
typedef struct MSOSPP
        {
        union
                {
                ULONG grf;
                struct
                        {
                        ULONG sppid : 16; // What property is this
                        ULONG fComplex : 1; // Needs freeing, marking, etc.
                        ULONG fNinch : 1; // Is this property "No Input No Change"
                        ULONG fReset : 1;       // Reset the property to default values
                        ULONG fIncr : 1; // Increment the current property value by 'lValue'
                        ULONG fAlreadyFetched : 1;      //Set by FetchSpp after the property  is
                                                                                                // fetched for the first shape in the selection.
                        ULONG : 11;
                        };
                };
        LONG lValue; // main 32-bit portion of property value
        } MSOSPP;
#define cbMSOSPP (sizeof(MSOSPP))

/* MSOFASPP */
/* These flags are passed to a bunch of methods that apply MSOSPPs. */
#define msogrfasppNil                   (0)
#define msofasppShapes                  (1<<0)
#define msofasppDefaults                (1<<1)
#define msofasppGroups                  (1<<2)
#define msofasppChildren                (1<<4)
#define msofasppResetAll                (1<<5)  // Reset all deltas before apply
#define msofasppEnsureUndo              (1<<6)
        /* If an undo record hasn't yet been begun, begin and end one around
                changing the properties. */
#define msofasppIgnoreCanHaveProperty (1<<7)
        /* Don't check whether a shape can have a given property before
                applying it. UI-level code shouldn't use this flag; in general,
                use it only if you know what you're doing. */
#define msofasppAnchorBotRight (1<<8)
   /* Use bottom-right as anchor for resizing. FALSE: use TopLeft (default) */
#define msofasppAnchorCenter  (1<<9)
   /* Use center as anchor for resizing. Mutually exclusive with
      msofasppAnchorTopLeft */
#define msofasppJustSetProperty (1<<10)
        /* Do nothing except directly call FSetProp.  This means that you
                can't expect FCanHaveProperty to be checked, an AfterShapePropertyChange
                event to be passed, or invalidation to be done.  Additionally, you
                can't use fake properties or anything else that cannot be applied
                through IMsoDrawing::FSetProp.

                Passing this flag can speed up property application quite a bit.
                It can also cause bugs.  Use with care. */
#define msofasppPropagate   (1 << 11)
#define msofasppFromDiagram (1 << 12)

/* MSOFFSPP */
/* These flags are passed to a bunch of methods that fetch MSOSPPs. */
#define msogrffsppNil           (0)
#define msoffsppShapes          (1<<1)
#define msoffsppDefaults        (1<<2) // Not yet supported!!!!!
#define msoffsppGroups          (1<<3)
#define msoffsppChildren        (1<<4)
#define msoffsppAlreadyFetched  (1<<5)
#define msoffsppDeltasOnly                (1<<6) // Return NINCH for non-deltas
#define msoffsppFetchResetValue (1<<7) // Return reset value of a property
#define msoffsppFetchFShapeCanHaveProperty (1<<8) //fetch property value iff this property
                                                                                                                                        //can be set for the shape
#define msoffsppTable           (1<<9)  // Shape belongs to PPT native table


/* Most commands want to set properties on all selected Shapes that aren't
        Groups and on the defaults.  They can use msogrfasppUsual. */
#define msogrfasppUsual \
        (msofasppDefaults | msofasppShapes | msofasppChildren)

/* MsoInitSpp initalizes an SPP.  MsoMakeSpp is often more useful. */
__inline void MsoInitSpp(MSOSPP *pspp, MSOSPPID sppid)
        {
        pspp->grf = 0;
        AssertInline(!pspp->fComplex && !pspp->fNinch);
        pspp->sppid = sppid;
        }

MSOAPIX_(void) MsoInitRgspp(MSOSPP *rgspp, int *pcspp, MSOPSID opsid);
MSOAPIX_(void) MsoInitDefaultsRgspp(MSOSPP *rgspp, int *pcspp);

/* MsoMakeSpp is the normal way to fill out a new SPP. */
__inline void MsoMakeSpp(MSOSPP *pspp, MSOSPPID sppid, LONG lValue)
        {
        pspp->grf = 0;
        AssertInline(!pspp->fComplex && !pspp->fNinch);
        pspp->sppid = sppid;
        pspp->lValue = lValue;
        }

/* MsoFreeSppCore is called by MsoFreeSpp to do real work if needed. */
MSOAPIX_(void) MsoFreeSppCore(MSOSPP *pspp,
        interface IMsoDrawingGroup *pidgg);

/* MsoFreeSpp frees any allocations hanging off an MSOSPP.  It's OK
        to call this multiple times on the same MSOSPP; it's neatly do nothing
        the second time. */
__inline void MsoFreeSpp(MSOSPP *pspp, interface IMsoDrawingGroup *pidgg)
        {
        if (pspp->fComplex)
                MsoFreeSppCore(pspp, pidgg);
        }

/* MsoFreeRgspp frees a whole array of MSOSPPs. */
MSOAPI_(void) MsoFreeRgspp(MSOSPP *rgspp, int cspp,
        interface IMsoDrawingGroup *pidgg);

/* MsoClearSpp clears out any pointers, etc. in an SPP without actually
        freeing them (useful if owenership of such pointers has been passed
        on to someone else). */
__inline void MsoClearSpp(MSOSPP *pspp)
        {
        pspp->fComplex = FALSE;
        }

/* MsoClearRgspp clears out a whole array of MSOSPPs. */
MSOAPIX_(void) MsoClearRgspp(MSOSPP *rgspp, int cspp);

/* MsoFCloneSppCore does real work under MsoFCloneSpp. */
MSOAPIX_(BOOL) MsoFCloneSppCore(MSOSPP *pspp,
        interface IMsoDrawingGroup *pidgg);

/* MsoFCloneSpp, well, clones an MSOSPP. */
__inline BOOL MsoFCloneSpp(const MSOSPP *psppSrc, MSOSPP *psppDest,
                interface IMsoDrawingGroup *pidgg)
        {
        *psppDest = *psppSrc;
        return (psppDest->fComplex ? MsoFCloneSppCore(psppDest, pidgg) : TRUE);
        }

/* MsoFCloneRgspp clones a whole range of MSOSPPs. */
MSOAPIX_(BOOL) MsoFCloneRgspp(MSOSPP *rgsppSrc, int cspp,
        MSOSPP *rgsppDest, interface IMsoDrawingGroup *pidgg);

/* MsoMoveSpp moves the contents of one MSOSPP into another. */
__inline void MsoMoveSpp(MSOSPP *psppSrc, MSOSPP *psppDest)
        {
        *psppDest = *psppSrc;
        psppSrc->fComplex = FALSE;
        psppSrc->sppid = msosppidNil; // People better not look here.
        }

/* MsoMoveRgspp moves a whole array of MSOSPPs. */
MSOAPIX_(void) MsoMoveRgspp(MSOSPP *rgsppSrc, int cspp, MSOSPP *rgsppDest);

/* MsoFIsSppNinched tells you if an MSOSPP is ninched or not. */
__inline BOOL MsoFIsSppNinched(const MSOSPP *pspp)
        {
        return (pspp->sppid == msosppidNil || pspp->fNinch ||
                MsoFIsNinch((MSOPID)(pspp->sppid), (void *)&pspp->lValue));
        }
/* TODO peteren: Remove MsoFSppNinched in favor of MsoFIsSppNinched */
__inline BOOL MsoFSppNinched(const MSOSPP *pspp)
        { return MsoFIsSppNinched(pspp);        }


/* NOTE!!: KEEP the tables vrgFakeGroupPropInfo[] and vrgFakeNoGroupPropInfo[]
        in drpinfof.cpp in sync with this MSOSPPIDFAKE enum */
#define msosppidFakeMin 0x8000
#define msosppidFakeNoGroupMin (msosppidFakeMin + msocOPSSIZE)
#define msopsidFakeGroup (msosppidFakeMin/msocOPSSIZE)
#define msopsidFakeNoGroup (msosppidFakeNoGroupMin/msocOPSSIZE)
// the properties marked as "set only" can only be set, those marked "fetch only"
// can only be fetched, the others can be set and fetched.
typedef enum
        {
        // Group properties
        msosppidApplyPreset = msosppidFakeMin, //set only
        msosppidHeight,
        msosppidWidth,
        msosppidGrowHeightByPercent, //set only - Relative to Original size is FALSE
        msosppidGrowWidthByPercent,  //set only - Relative to Original size is FALSE
        msosppidLeftCropping,
        msosppidRightCropping,
        msosppidTopCropping,
        msosppidBottomCropping,
        msosppidHorizontalPosition,
        msosppidVerticalPosition,
        msosppidRotation,
        msosppidContrastChangeByPercent,//set only
        msosppidGtextRotation,          //set only
        msosppid3DTiltVertical,         //set only
        msosppid3DTiltHorizontal,       //set only
        msosppidGrowHeightRelativeToOrigByPercent, //set only - Relative to Original size is TRUE
        msosppidGrowWidthRelativeToOrigByPercent,        //set only - Relative to Original size is TRUE
        msosppidOriginalHeight, //fetch only
        msosppidOriginalWidth,  //fetch only
        msosppidPictureHeightScale,//fetch only
        msosppidPictureWidthScale, //fetch only
        msosppidFInsetPen,
        msosppidFInsetPenOK, //fetch only
        msosppidFakeGroupMax, msosppidFakeGroupLast = msosppidFakeGroupMax - 1,


        // NoGroup properties
        msosppidPibDownload = msosppidFakeNoGroupMin,
                                                                                   // fetch only - blip with synchronous download
        msosppidFillBlipDownload,     // fetch only - blip with synchronous download
        msosppidLineFillBlipDownload, // fetch only - blip with synchronous download
        msosppidFakeMax, msosppidFakeLast = msosppidFakeMax - 1
        } MSOSPPIDFAKE;


/*****************************************************************************
        Save and Load

        Escher's file format is tree-structured, with data stored only at the
        leaves.  Effectively, there are two kinds of file blocks: containers and
        data blocks.  (Containers may be nested.)

        An MSOFBH is saved at the beginning of each file block.  It contains
        the type of the block, the version number, the instance, and the length.
        The block types are enumerated as MSOFBT below.  Containers always have
        version msofbvContainer; data blocks may use the version however they see
        fit.  The instance may be used for anything.  The length value does NOT
        count the MSOFBH itself; that is, someone wishing to skip over a certain
        file block would read the MSOFBH and then seek ahead cbLength bytes.

        Here's a map of Escher's file structure:

        When IMsoDrawingGroup::FSave is called with fBeforeDrawings:
                Drawing Group Container (msofbtDggContainer)
                        |
                        |- Drawing Group data (msofbtDgg)
                        |- Blip Store Container (msofbtBstoreContainer)
                        |               |
                        |               |- Blip Store Entry (msofbtBSE)
                        |               |- ...
                        |- Default properties

        When it's called with fOneDrawing:
                Drawing Container (msofbtDgContainer)
                        |
                        |- Drawing data (msofbtDg)
                        |- Shape (msofbtSpgrContainer or msofbtSpContainer)
                        |               |
                        |               |- Group Shape data, if the shape is a group shape (msofbtSpgr)
                        |               |- Shape data (msofbtSp)
                        |               |- Properties (msofbtOPT)
                        |               |- Textbox (msofbtClientTextbox)
                        |               |- Anchor (msofbtAnchor or msofbtClientAnchor)
                        |               |- Client data (msofbtClientData)
                        |- Other shapes and group shapes...
                        |- Solver (msofbtSolverContainer)
                                        |
                                        |- Rule (msofbtConnectorRule or msofbtAlignRule)
                                        |- ...

        (The flag fAllDrawings just saves each drawing with a
        client-supplied drawing label before it.)

        This format is somewhat flexible: many of the above items are left out if
        they don't exist (e.g. textbox, solver).  Also, the order of items in
        a container doesn't matter.

        MSOSVB and MSOLDB are parameter blocks passed around to save and load
        functions, respectively.  This includes client callbacks for saving
        and loading textboxes, anchors, and client data.  Be aware that the same
        procedures are called for clipboard copy/paste operations; the flag
        fClipboard will be set in that case.

        Saving:
                If you are passed an MSOSVB when saving, you should first write an
                FBH using MsoFWriteFbh().  (It's okay to do nothing if you aren't
                saving anything, but don't expect to be called in the corresponding
                load function if you do this.)  The length value need not be accurate if
                we're only doing a counting pass (i.e. fSaveImmediate is not
                passed in).

                Then, you should only actually save if svb.fSaveImmediate
                is true. Save to the stream pistmImmediate.  In all cases, you
                should add the number of bytes that were saved (or that would be saved)
                to dfoImmediate.  (Note: MsoFWriteFbh() does its own dfo updating.)

                The pistmDelay, etc. members of MSOSVB are for saving big things
                to a delay stream that can be lazily loaded.  Currently, only the
                blip store uses these.

        Loading:
                The FBH that you wrote out should be passed in via the MSOLDB.
                The version number you saved will be in the FBH's ver field.  You're
                free to do whatever you want with it.

                Load from pistmImmediate, and update dfoImmediate with the number
                of bytes that you loaded.  (The number of bytes loaded should match
                the cbLength in the FBH.)

                If ldb.fClipboard is TRUE, then clsid will contain the CLSID of the
                app that put the information on the clipboard.


****************************************************************** MMorgan **/

// File Blocks

/* MSOFBT - File Block Type */
/* Office developers -- When you add a new FBT (to the END, of course),
        be sure to add a new entry to the version number table in drfile.cpp. */
typedef enum
        {
        msofbtMin=0xF000,
        msofbtDggContainer=msofbtMin,
        msofbtBstoreContainer,
        msofbtDgContainer,
        msofbtSpgrContainer,
        msofbtSpContainer,
        msofbtSolverContainer,
        msofbtDgg,
        msofbtBSE,
        msofbtDg,
        msofbtSpgr,
        msofbtSp,
        msofbtOPT,
        msofbtTextbox,
        msofbtClientTextbox,
        msofbtAnchor,
        msofbtChildAnchor,
        msofbtClientAnchor,
        msofbtClientData,
        msofbtConnectorRule,
        msofbtAlignRule,
        msofbtArcRule,
        msofbtClientRule,
        msofbtCLSID,
        msofbtCalloutRule,
        msofbtBlipFirst,
        msofbtBlipFirstClient = msofbtBlipFirst+msoblipFirstClient,
        msofbtBlipLast        = msofbtBlipFirst+msoblipLastClient,
        msofbtRegroupItems,
        msofbtSelection,
        msofbtColorMRU,
        msofbtUndoContainer,
        msofbtUndo,
        msofbtDeletedPspl,
        msofbtSplitMenuColors,
        msofbtOleObject,
        msofbtColorScheme,
        msofbtSecondaryOPT,
        msofbtTertiaryOPT,
        msofbtTimeNodeContainer,
        msofbtTimeConditionList,
        msofbtTimeConditionContainer,
        msofbtTimeModifierList,
        msofbtTimeNode,
        msofbtTimeCondition,
        msofbtTimeModifier,
        msofbtTimeBehaviorContainer,
        msofbtTimeAnimateBehaviorContainer,
        msofbtTimeColorBehaviorContainer,
        msofbtTimeEffectBehaviorContainer,
        msofbtTimeMotionBehaviorContainer,
        msofbtTimeRotationBehaviorContainer,
        msofbtTimeScaleBehaviorContainer,
        msofbtTimeSetBehaviorContainer,
        msofbtTimeCommandBehaviorContainer,
        msofbtTimeBehavior,
        msofbtTimeAnimateBehavior,
        msofbtTimeColorBehavior,
        msofbtTimeEffectBehavior,
        msofbtTimeMotionBehavior,
        msofbtTimeRotationBehavior,
        msofbtTimeScaleBehavior,
        msofbtTimeSetBehavior,
        msofbtTimeCommandBehavior,
        msofbtClientVisualElement,
        msofbtTimePropertyList,
        msofbtTimeVariantList,
        msofbtTimeAnimationValueList,
        msofbtTimeIterateData,
        msofbtTimeSequenceData,
        msofbtTimeVariant,
        msofbtTimeAnimationValue,
        msofbtExtTimeNodeContainer,
        msofbtSlaveContainer,
        msofbtMax=0xFFFF
        } MSOFBT;

#define MSOFBTBLIPFROMBT(bt)  MSOFBT(msofbtBlipFirst+(bt))
#define MSOBTFROMFBTBLIP(fbt) MSOBLIPTYPE((fbt)-msofbtBlipFirst)

typedef BYTE    MSOFBV; // File Block Version - 4 bits (see MSOFBH definition)
#define msofbvMin       ((MSOFBV)0)
#define msofbvMax ((MSOFBV)0xF)
#define msofbvContainer msofbvMax // Containers don't really have a version

typedef USHORT MSOFBI; // File Block Instance - 12 bits
#define msofbiMin       ((MSOFBI)0)
#define msofbiMax ((MSOFBI)0xFFF)
#define msocbElemCompressMark   0xFFF0 // indicates a compressed array

// Creates an FBH from the given arguments and writes it out
// in keeping with the MSOSVB conventions.
MSOAPI_(BOOL) MsoFWriteFbh(struct MSOSVB *psvb, MSOFBV ver,
                                                                        MSOFBI inst, MSOFBT fbt, ULONG cbLength);

// Arguments to save/load functions

/* MSOSVB - SaVe Block
        structure passed around among save functions (contains items that
        would otherwise be arguments to the save functions) */
/* HEY!!! Office developers!  There's an internal structure (actually
        a class) called the SVB that mirrors this structure field by field.
        If you change one you need to change the other! */
typedef struct MSOSVB
        {
        union
                {
                struct
                        {
                        ULONG fBeforeDrawings : 1;
                        ULONG fAllDrawings : 1;
                        ULONG fOneDrawing : 1;
                        ULONG fAfterDrawings : 1;
                        ULONG fSaveImmediate : 1;
                        ULONG fSaveDelay : 1;
                        ULONG fUpdateDelayReferences : 1;
                        ULONG fUseDelay : 1;
                        ULONG fIncrementalDelay : 1;
                        ULONG fClipboard : 1;
                        ULONG fNonNative : 1;
                        ULONG fExportLinks : 1;
                        ULONG fSaveDeletedShapes : 1;
                        ULONG fCloneImmediate : 1;
                        ULONG fClientKeepsUndo : 1;     // Does not purge undo before save
                        ULONG : 17;
                        };
                ULONG grf;
                };

        MSOBLIPFORMATFLAGS grfBlipFormat;

        IStream *pistmImmediate;
        MSODFO   dfoImmediate;

        IStream *pistmDelay;
        MSOFO            foDelay;
        MSODFO   dfoDelay;

        // Link management fields - used only when fExportLinks is set
        struct _DOCSUMINFO *lpdsiobj;
        DWORD           dwApp;

        /* The next two fields are placeholders for fields used internally
                by Office. */
        void *pv1;
        void *pv2;

        #if DEBUG
        union
                {
                struct
                        {
                        ULONG fDontSaveIdcls : 1;
                        ULONG fSaveExtraFileBlocks : 1;
                        ULONG grfDebugUnused : 30;
                        };
                ULONG grfDebug;
                };
        LONG lVerifyInitSvb; // Used to make sure people use InitSaveBlock
        #endif
        } MSOSVB;
#define cbMSOSVB (sizeof(MSOSVB))

/* MSOLDB - LoaD Block
        structure passed around among load functions (contains items that
        would otherwise be arguments to the load functions) */
/* HEY!!! Office developers!  There's an internal structure (actually
        a class) called the LDB that mirrors this structure field by field.
        If you change one you need to change the other! */
typedef struct MSOLDB
        {
        union
                {
                struct
                        {
                        ULONG fBeforeDrawings : 1;
                        ULONG fAllDrawings : 1;
                        ULONG fOneDrawing : 1;
                        ULONG fAfterDrawings : 1;
                        ULONG fUseDelay : 1;
                        ULONG fClipboard : 1;
                        ULONG fLoadAllDelay : 1;
                        ULONG fCloneImmediate : 1;
                        ULONG fFileProbablyLastSavedInO9OrEarlier : 1;
                        ULONG : 23;
                        };
                ULONG grf;
                };

        IStream *pistmImmediate;
        MSODFO dfoImmediate;
        IStream *pistmDelay;
        MSODFO dfoDelay;
        MSOFBH fbh;
        CLSID            clsid;

        /* The remaining fields are placeholders for fields used internally
                by Office. */
        void *pv1;
        void *pv2;

        #if DEBUG
        LONG lVerifyInitLdb; // Used to make sure people use InitLoadBlock
        #endif
        } MSOLDB;
#define cbMSOLDB (sizeof(MSOLDB))


/****************************************************************************
        MsoDrEventBit (MsoDEB)

        Site callback event trigger map.

        This is for optimizing out event firing for the host apps, by implementing
        the FWantsEvent infrastructure.  Generally, Excel and Word use different
        events, and PowerPoint rarely uses any (except the drawing view).

        We do a 'lazy' implementation of our internal checks for FWantsEvent.  We
        assume that the host wants all our events (as in '97 and 2k), but the host
        has the option to unset from the event map any events for which it is not
        interested, in the default switch/case of OnEvent.

        Optionally, the host app can call MsoDrEventBitClear during initializing
        of the site info block (ie. MSODGSI) for events they know are unwanted.
        (ie. after calling MsoInitDgsi)

        Pre-processor macros can be skanky, but I wanted to keep this chunk of
        code simple to call, compatible with our C clients, and with as little
        overhead as reasonably possible.

        For the debug build, we have iEventEnd for error-checking of the event
        declarations, and we heavily pad the map so that adding new events won't
        cause the overall size to change and force a full rebuild.

        Note: making changes to the below APIs means you'll probably need to do
        a full build!
*****************************************************************************/

// the bit mapping type
#define MsoDEBType(name) _ ## name ## _mapEvents
// the bit mapping name
#define MsoDEB(name) name ## _mapEvents
// pointer to bit mapping name
#define MsoDEBP(name) p ## name ## _mapEvents

// declares a bit set type
// the debug padding is so that adding new events won't require full builds
// for debug, at least up to 1024 events (all bets are off for ship)
#if DEBUG
#define MsoDEBDeclare(name, start, end)\
typedef struct {\
        unsigned short iEventStart;\
        unsigned short iEventEnd;\
        union {\
                BYTE rgEvents[(((end)-(start))/8)+1];\
                BYTE padding[128]; };\
        } MsoDEBType(name)
#else
#define MsoDEBDeclare(name, start, end)\
typedef struct {\
        unsigned short iEventStart;\
        BYTE rgEvents[(((end)-(start))/8)+1];\
        } MsoDEBType(name)
#endif

// defines a bit set
#define MsoDEBDefine(name)\
        MsoDEBType(name) MsoDEB(name)

// defines a pointer to a bit set
#define MsoDEBPDefine(name)\
        MsoDEBType(name) *MsoDEBP(name)

// initialize a bit set
#define MsoDEBInit(pcontainer, name, start, end, f)\
        {\
        Assert((end) >= (start));\
        Assert((LONG)(end) < (LONG)((unsigned short)~0));\
        AssertSz((((end)-(start))/8)+1 < 128, "MsoDEB: # events > debug padding");\
        (pcontainer)->MsoDEB(name).iEventStart = (start);\
        Debug((pcontainer)->MsoDEB(name).iEventEnd = (end);)\
        memset((pcontainer)->MsoDEB(name).rgEvents, ((f) ? ~0 : 0),\
                (((end)-(start))/8)+1);\
        }

// initialize bit-set-pointer to an existing bit set
#define MsoDEBPInit(pcontainer, name, pcontainerfrom)\
        (pcontainer)->MsoDEBP(name) = &(pcontainerfrom)->MsoDEB(name)

// validate bit set
#define MsoDEBAssert(pcontainer, name, event)\
        {\
        Assert((event) >= (pcontainer)->MsoDEB(name).iEventStart);\
        Assert((event) <= (pcontainer)->MsoDEB(name).iEventEnd);\
        }

// validate bit-set-pointer
#define MsoDEBPAssert(pcontainer, name, event)\
        {\
        Assert((pcontainer)->MsoDEBP(name));\
        Assert((event) >= (pcontainer)->MsoDEBP(name)->iEventStart);\
        Assert((event) <= (pcontainer)->MsoDEBP(name)->iEventEnd);\
        }

// clear bit in bit set
#define MsoDEBClear(pcontainer, name, event)\
        {\
        MsoDEBAssert(pcontainer, name, event);\
        (pcontainer)->MsoDEB(name).rgEvents[\
                ((event) - (pcontainer)->MsoDEB(name).iEventStart)/8] &=\
                 ~(1 << (((event) - (pcontainer)->MsoDEB(name).iEventStart)%8));\
        }

// clear bit in bit-set-pointer
#define MsoDEBPClear(pcontainer, name, event)\
        {\
        MsoDEBPAssert(pcontainer, name, event);\
        (pcontainer)->MsoDEBP(name)->rgEvents\
                [((event) - (pcontainer)->MsoDEBP(name)->iEventStart)/8] &=\
                 ~(1 << (((event) - (pcontainer)->MsoDEBP(name)->iEventStart)%8));\
        }

// check bit in bit set
#define MsoDEBF(pcontainer, name, event)\
        (F((pcontainer)->MsoDEB(name).rgEvents\
                [((event) - (pcontainer)->MsoDEB(name).iEventStart)/8] &\
                (1 << (((event) - (pcontainer)->MsoDEB(name).iEventStart)%8))))


/****************************************************************************
        Drawing User Interface (and related stuff)

        This Drawing User Interface is the top-level object of the Office Drawing
        Component.  Basically we expect each host to create exactly one of these.
        This is where "global" UI interactions (mode-setting, etc.) happen.
*****************************************************************************/

/* MSODGCI -- Drawing Command Info. */
/* This is the CONSTANT description of a particular DGC.  We keep a big
        tables of these.  These should be read-only to almost everyone in the
        universe, don't go setting values in here. */
/* Hey! Office developers; take care to keep this in ssync with the DGCI. */
typedef struct MSODGCI
        {
        ULONG grfdgci; // Random info about command
        ULONG grfdgcd; // When is this command disabled?
        int tcid;
        void *pfndgc; // Lines up with DGCI.pfndgc
        int idsName; // Command Name (unique, non-localized, no spaces)
        int idsUndo; // Undo String (spaces, not unique)
        int idsHelp; // Help String (I suspect should be not an ids but a help id of some sort)
        ULONG lArgument;
#if DEBUG
        MSODGCID dgcidCheck;
                /* We store DGCIs in a table indexed by MSODGCID.  We verify that
                        the table lines up with the definitions by including the
                        MSODGCIDs in the Debug version. */
#endif
        } MSODGCI;
#define cbMSODGCI (sizeof(MSODGCI))

/* This structure is passed in as a parameter to several polygon reshape
        commands. */
typedef struct _MSOCURVEPT
        {
        union
                {
                MSODGCID dgcidTool;
                void *pvDrgcrv; // Used for state-based polygon creation.
                };
        MSOHSP hsp;
        int ipt;
        MSOPATHESCAPE pesc;
        union
                {
                ULONG grf;
                struct
                        {
                        ULONG fHasHsp    : 1; // True if hsp is valid.
                        ULONG fHasDrgcrv : 1; // True if DRGCRV is valid.
                        ULONG fHasDgcid  : 1; // True if dgcid is valid.
                        ULONG fFinish    : 1; // True if finish creation with this point.
                        ULONG cpt        : 2; // Fill this out with the number of points in rgpth.
                        ULONG fIsVbaIpt  : 1; // Is the ipt in VBA indices?
                        ULONG grfUnused  : 25;
                        };
                };
        POINT rgpth[1];
                /* We need 3 points when we are creating a corner curve, for all
                        other cases we just need a single point. */
        } MSOCURVEPT;

typedef struct MSOCOLOREXT
        {
#if 0
        union
                {
                ULONG grf;
                struct
                        {
                        ULONG fRGB      : 1;// color specified in sRGB color space
                        ULONG fCMYK     : 1;// color specified in CMYK color space
                        ULONG fInk      : 1;// color is specified as an ink color
                        ULONG fColorMod : 1;// color modification value is specified
                        ULONG fCMY      : 1;// CMY color channels are specified
                        ULONG fK        : 1;// black color in CMYK color space is specified
                        ULONG fClrRGB   : 1;// clrRGB has valid color value
                        ULONG : 25;
                        };
                };
#endif
        COLORREF         clr;    // sRGB value if color is a simple sRGB else it is a flattened color
        COLORREF         clrRGB; // MSOCOLORNONE if color is not specified in sRGB
        LONG             clrCMY; // MSOCOLORNONE if color is not specified in CMYK
        LONG      clrK;   // MSOCOLORNONE if color is not specified in CMYK
        WCHAR    *pwzInk; // NULL if not an ink color
        COLORREF  clrMod; // MSOCOLORMODUNDEFINED if color mod is not specified
        } MSOCOLOREXT;
#define cbMSOCOLOREXT (sizeof(MSOCOLOREXT))

/* MsoInitColorExt initalizes a MSOCOLOREXT. */
__inline void MsoInitColorExt(MSOCOLOREXT *pclr)
        {
        //pclr->grf    = 0L;
        pclr->clr    = MSOCOLORNONE;
        pclr->clrRGB = MSOCOLORNONE;
        pclr->clrCMY = MSOCOLORNONE;
        pclr->clrK   = MSOCOLORNONE;
        pclr->pwzInk = NULL;
        pclr->clrMod = MSOCOLORMODUNDEFINED;
        }

/* MsoFreeColorExt frees any allocations hanging off a MSOCOLOREXT. */
__inline void MsoFreeColorExt(MSOCOLOREXT *pclr)
        {
        if (pclr->pwzInk && msowzNinch != pclr->pwzInk)
                {
                MsoFreePv(pclr->pwzInk);
                pclr->pwzInk = NULL;
                }
        }

/* MsoInitColorExt initalizes a MSOCOLOREXT. */
__inline void MsoInitColorExtRgb(MSOCOLOREXT *pclr, COLORREF rgb)
        {
        MsoInitColorExt(pclr);
        pclr->clr = rgb;
        }

/* Given all of the color properties that go into a MSOCOLOREXT, set them in
        the MSOCOLOREXT. Note that this API assumes an uninitalized MSOCOLOREXT is
        passed in and so makes no attempts to free allocated mem references that
        it may contain. Also, it doesn't clone any of the properties that are
        passed in. The caller should make sure that it is doing so when needed. */
MSOAPI_(void) MsoBuildMsoColorExt(__no_count MSOCOLOREXT *pclr, COLORREF clr,
        COLORREF clrRGB, COLORREF clrCMY, COLORREF clrK, __no_count WCHAR *pwzInk,
        COLORREF clrMod);

/* MSODGC = Drawing Command */
/* Hey!! Office Devlopers, take heed.  The MSODGC structure is
        basically copy-pasted into the DGC structure.  If you change
        one you'll have to change both. */
typedef struct MSODGC
        {
        MSODGCID dgcid; // Identifer for the command we're running
        MSODGCI dgci; // Lots of useful data about this command

        interface IMsoDrawingGroup *pidgg;
        interface IMsoDrawing *pidg;
        interface IMsoDrawingSelection *pidgsl;
        interface IMsoDrawingView *pidgv;
        union
                {
                struct
                        {
                        ULONG fSetRepeat : 1; // Change the repeat state in the DGUI
                        ULONG fRecordVBA : 1; // Record VBA under FExecuteCommand
                        ULONG fExecute : 1; // Actually run under FExecuteCommand
                        ULONG fSetDefaults : 1; // Set the new shape defaults when executing
                        ULONG fNoDelay : 1; // If TRUE we need the command run immediately
                        ULONG fUpdate : 1; // Command should do an update if needed
                        ULONG fNoWarning : 1; // Command should do put up any warning dialog
                        ULONG fForceCreate : 1; // Force new shape creation and skip
                                 // the shape insertion mode
                        ULONG fNoUpdateCaret : 1; // Disable the the caret update
                        ULONG fExtendedColors : 1;
                                /* If true, extended color properties are used in the UI. Namely,
                                        the Tints tab of the Fill Effects dialog will be visible and
                                        CMYK and CMS options will appear in the More Colors dialog. */
                        ULONG fSpotMode : 1;
                                /* If true, Fill Effects dialog has special behavior to disallow
                                        effects that are 'illegal' in spot mode. */
                        ULONG fProcessMode : 1;
                                /* If true, More Colors dialog has special behavior to disable
                                        RGB and HSL controls. */
                        ULONG fDrillDownSelection : 1;
                                /* If true, Object Model knows that the command that will be executed
                                        will cause the selection to be in a drill down mode and hence can
                                        choose to not do any macro recording. When the selection changes
                                        from drill down to top level, this flag is set to false. */
                        ULONG fBlockAutoCanvas : 1; // don't auto-create a gray-state canvas
                        ULONG fOnlySolidFill : 1; // Publisher: only solid fill tabs are displayed
                        ULONG fNoColorTransparency : 1; // Publisher: hide transparency controls in more color dialog
                        ULONG grfUnused : 16;
                        };
                ULONG grf;
                };
        union
                {
                MSOGV gvClient; // Default is 0
                void *pvClientContext; // old name for gvClient, remove
                };
        MSOGV gvUndoContext;
                /* We expect the host's implementation of FRequestExecuteCommand
                        to set a value here as it sets up whatever notion of undo the
                        host has.  Then it can look at it in all subsquent calls to
                        DG::FPostUndoRecord. */

        /* Arguments for commands. */
                // TODO peteren: Figure out how much unioning should go on here.
        union
                {
                struct
                        {
                        void *pvParm1;
                        void *pvParm2;
                        void *pvParm3;
                        void *pvParm4;
                        void *pvParm5;
                        void *pvParm6;
                        void *pvParm7;
                        void *pvParmExtra[10];
                        };
                MSOHSP hsp;
                MSOCR cr;
                MSOPX *ppxspp;
                MSOPX *ppxpt;
                MSOCURVEPT *pcrvpt;
                CLIPFORMAT cf;
                ULONG lSwatchResult;
                RECT rchTextFrame;
                MSOMM mm; //Keyboard status for msodgcidNudgexxxx
                struct   //msodgcidInsertBlip
                        {
                        interface IMsoBlip *pib;        // if NULL, then wzName must be the URL
                        RECT               rch;         // Position of the blip
                        MSOBLIPFLAGS       bf;
                        WCHAR              *wzName;     // Maybe NULL, if pib is specified
                        } dgcInsertBlip;
						struct   //msodgcidOptimizePict
							{
							int resampleType; // what type of resampling to do
							BOOL fSelectedShapes : 1; // operate only on selected shapes
							BOOL fCompress : 1; // do lossless compressions
							BOOL fCrop : 1; // delete cropped areas of images

							/* a host-supplied array of drawing groups on which the cmd
								should operate.  if empty, the command operates on the 
								DGG implied by the context (MSODGC::pidgg/pidgsl).  
								Hosts wishing to have the command operate on multiple
								DGGs should re-initialize this array EVERY TIME in 
								IMsoDrawinguserInterfaceSite::FRequestExecuteCommand.
								Thus FCommandOptimizePict can trust these fields even
								though the original DGGs in question may have gone away
							*/
							int cEntries;
							struct 
								{
								interface IMsoDrawingGroup *pdgg;
								/* allow 2 shapes with the same blip in this DGG to get
									2 different blips after compression.  Desirable when
									(as in Word inline layer) a single blip referenced
									multiple times gets multiple copies saved out. */
								BOOL fAllowBlipSplit : 1; 
								/* skip unmarked shapes when analyzing the DGG for 
									shapes to optimize and for repeated blip use.  Set
									for cases (Like the undo doc in Word) where we want 
									to operate on some of the sha[es in a DG but not 
									all. */
								BOOL fOnlyMarkedShapes : 1;
								} rgEntries[6];
							} dgcOptimizePict;
                MSOCOLOREXT clrExt;
                POINT ptvContextMenu; // msodgcidContextMenu
                };

#if DEBUG
        LONG lVerifyInitDgc;
#endif // DEBUG
        } MSODGC;
#define cbMSODGC (sizeof(MSODGC))

/* MsoFDgcPostsUndoRecords returns TRUE if a command's going to post
        any undo records. */
#define MsoFDgcPostsUndoRecords(pmsodgc) \
        (((pmsodgc)->dgci.grfdgci & msofdgciPostsUndoRecords) != 0)


// MSODGUIHRSID -- Host Recorded VBA String(s) ID
/* These are sent with msodgeDoHostRecording event to the host, they
        means different things depending on the dguihrsid passed. */
typedef enum
        {
        msodguihrsidActivateText,  // Host records placement of the insertion point
                                                                                // in the currently selected textbox.
        msodguihrsidUnselect,      // Host records dropping of the current selection
        msodguihrsidDuplicate,     // Host records duplicate command
        } MSODGUIHRSID;


// MSODGUIBRID -- Before Recording ID
/* These are sent with msodgeBeforeRecording event to the host, they
        means different things depending on the MSODGUIBRID passed. */
typedef enum
        {
        msodguibridLabelTextBox,  // Before recording a label or textbox
        } MSODGUIBRID;

MsoDEBDeclare(dgui, msodgeDguisFirst, msodgeDguisMax);

/* MSODGUIEB -- Drawing User Interface Event Block */
typedef struct MSODGUIEB
        {
        /* These first fields are filled out for EVERY
                Drawing User Interface Event. */
        MSODGE dge;
        BOOL fResult;
        interface IMsoDrawingUserInterface *pidgui;

        MsoDEBPDefine(dgui);

        union
                {
                struct
                        {
                        MSOCOLOREXT *pclrExt1;  //extended clr of color1 on gradient tab
                        MSOCOLOREXT *pclrExt2;  //extended clr of color2 on gradient tab
                        LONG nShadeTintLevel; // Level of tint or shade of the color
                        BOOL fTwoColor; // Two color gradient?
                        } dgexDoExitFillEffectsDlg;

                struct
                        // fields for msodgeRequestNumOfTextSelected and msodgeGetTextSelection.
                        {
                        BOOL fWordLimit; // TRUE if we set cWord below.
                        BOOL fDelSel; // TRUE if we want the host to delete the text selection.
                        BOOL fIsCreateDiagram;
                        int cWord; // number of word to read
                        int cch; // max cch for msodgeGetTextSelection;
                                 // cch selected for msodgeRequestNumOfTextSelected.
                        WCHAR *wzText; // for msodgeGetTextSelection.
                        } dgexRequestTextSelection;
                struct
                        {
                        MSODGC *pdgc;
                        BOOL fEnabled;
                        } dgexAfterIsDgcEnabled;
                struct
                        {
                        MSODGCID dgcid;
                        BOOL fEnabled;
                        } dgexAfterIsDgcidEnabled;
                struct
                        {
                        MSOHSP hsp;
                        void* pvClient;
                        } dgexRepeatInsertShape;
                struct
                        {
                        ULONG grfdguimOld;
                        ULONG grfdguimNew;
                        } dgexDoChangeModes;
                struct
                        {
                        HDC     hdc;
                        RECT rect;
                        } dgexDoDrawLineStyleMenu;
                struct
                        {
                        BOOL fReturn;
                        } dgexDoFormatObjectDialog;
                struct
                        {
                        MSODGC *pdgc;
                        MSODGUIHRSID dguihrsid;
                        } dgexDoHostRecording;
                struct
                        {
                        MSOHSP hsp;
                        interface IMsoDrawing *pidg;
                        interface IMsoDrawingSelection *pidgsl;
                        } dgexDoGetSelectionForShape;

                struct
                        {
                        interface IMsoDrawing *pidg;
                        BOOL      fProceed;
                        } dgexBeforePickUpFormat;
                struct
                        {
                        interface IMsoDrawing *pidg;
                        BOOL      fProceed;
                        } dgexBeforeApplyFormat;

                struct
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        } dgexAfterPickUpFormat;
                struct
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        } dgexAfterApplyFormat;
                struct
                        {
                        MSODGUIBRID dguibrid;
                        union
                                {
                                struct
                                        {
                                        MSOHSP hsp;
                                        MSOTXFL *ptxflow;
                                        RECT *prch;
                                        } dgexTB;       // for textbox and label recording
                                };
                        } dgexBeforeRecording;
                struct
                        {
                        IDataObject *pido;
                        } dgexDoGiveDataObject;
                struct
                        {
                        MSODGCID dgcid;                         // From the command block
                        interface IMoniker* pimk;               // The name of the file to insert
                        void *pvParm1;                          // From pvParm1 of the command block
                        BOOL fSuccess;                          // Set fFalse to indicate failure
                        }dgexDoInsertCagFile;
                struct
                        {
                        MSOCOLOREXT *pclrExt;
                        }dgexDoFilterDefaultTextureFillColor;
                struct
                        {
                        MSOCOLOREXT *pclrExt;
                        }dgexDoResolveSpotColor;
                struct
                        {
                        BOOL fSupportHighContrastInk;
                        BOOL fEraseMode;
                        } dgexBeforeInkingStart;
                };
        } MSODGUIEB;
#define cbMSODGUIEB (sizeof(MSODGUIEB))


// MSODGRVS -- Record VBA String
/* These are passed to DGUIS::FGetVBAString, along with a wParam that
        means different things depending on the dgrvs passed. */
typedef enum
        {
        msodgrvsSelection,   // send Selection object prefix, wParam is ignored
        msodgrvsDrawing,     // send Drawing object prefix, wParam is ignored
        msodgrvsSchemeColor, // send SchemeColor string for scheme color index in 'wParam'
        } MSODGRVS;

// MSOCDID -- Clipboard Drawing ID
/* Passed to FGetClipboardDrawing and ReleaseClipboardDrawing. */
typedef enum
        {
        msocdidNormal,
        msocdidDragDrop,
        } MSOCDID;

/* PopulateRgtcr options (msopopoxxx). */
#define msopopoDefault                          0x0000  // have Office do the default thing (MRU, More Colors, Fill Effects added when deemed necessary)
#define msopopoMoreColors                       0x0001  // more colors are supported, so add MRU and More Colors button
#define msopopoNoDefaultControls                0x0002  // tell Office not to add any controls

/* IMsoDrawingUserInterfaceSite (idgcmps).
        This is the interface you the host needs to speak to create a
        Drawing User Interface. */

interface IMsoTFC;

#undef  INTERFACE
#define INTERFACE IMsoDrawingUserInterfaceSite
DECLARE_INTERFACE_(IMsoDrawingUserInterfaceSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- Standard Office Drawing methods

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Office Drawing will call DGUIS::OnEvent whenever a "global" drawing
                event (MSODGE) occurs.  See the definition of MSODGE for a more
                complete explanation of events.  We'll not call this function
                if you set msodgcmpsi.fWantsEvents to FALSE. */
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDguis, MSODGUIEB *pdguieb) PURE;

        // ----- Methods for Office Drawing to get current context

        /* These methods are called while we're running commands.  Note that
                we pass a pointer to the command we're running.  This is primarily
                to allow the host to look at the pvClientContext field.

                If you have already created one of these objects and we ask you for it,
                you should return it.  However, if you haven't created the object yet
                and fEnsure is FALSE, just return FALSE (don't bother creating one). */

        MSOMETHOD_(BOOL, FGetCurrentDrawingGroup) (THIS_ void *pvDguis,
                interface IMsoDrawingGroup **ppidgg, MSODGC *pdgcExecuting, BOOL fEnsure) PURE;

        MSOMETHOD_(BOOL, FGetCurrentDrawing) (THIS_ void *pvDguis,
                interface IMsoDrawing **ppidg, MSODGC *pdgcExecuting, BOOL fEnsure) PURE;

        MSOMETHOD_(BOOL, FGetCurrentView) (THIS_ void *pvDguis,
                interface IMsoDrawingView **ppidgv, MSODGC *pdgcExecuting, BOOL fEnsure) PURE;

        /* FGetCurrentSelection should return a drawing selection
                whenever possible, even if it's currently empty.  Commands like
                SelectAll need a Selection to Select All into. */
        MSOMETHOD_(BOOL, FGetCurrentSelection) (THIS_ void *pvDguis,
                interface IMsoDrawingSelection **ppidgsl, MSODGC *pdgcExecuting,
                BOOL fEnsure) PURE;

        // ----- Methods used while executing Commands

        /* Office Drawing (the DGUI, actually) will call FRequestExecuteCommand
                when for some reason (Direct Manipulation, say?) Office would
                like a command run.  The host can either call DGUI::FExecuteCommand
                (perhaps setting up undo first and taking down after) or they can
                (PowerPoint does this) store the pointer to the command and execute
                it later.  If they do the latter they should set *pfDelayExecution
                to TRUE (we initialize it to FALSE).  Return TRUE if everything
                succeeds, FALSE otherwise.  If you just run the command you don't
                have to worry about freeing it; we'll take care of that.  If you
                set *pfDelayExecution to TRUE then you'll have to free the command
                yourself after executing it. */
        MSOMETHOD_(BOOL, FRequestExecuteCommand) (THIS_ void *pvDguis,
                MSODGC *pdgc, BOOL *pfDelayExecution) PURE;

        // ----- Other methods

        /* Office Drawing will call RequestToolbarUpdate when we see something
                change that we know will change the enabled/disabled/etc. state of
                one of our commands. */
        MSOMETHOD_(void, RequestToolbarUpdate) (THIS_ void *pvDguis) PURE;

        /* If when running a command we find we want to set some kind of "tool"
                mode we'll call FSetCurrentTool with the dgcid of that command.
                We'll pass msodgcidNil when we're finished with the tool and think
                it should go away. */
        MSOMETHOD_(BOOL, FSetCurrentTool) (THIS_ void *pvDguis,
                MSODGCID dgcidTool) PURE;

        /* When Office calls DgcidGetCurrentTool, the host should return the
                last dgcidTool passed to FSetCurrentTool or msodgcidNil if
                there isn't an Office Drawing tool selection. */
        MSOMETHOD_(MSODGCID, DgcidGetCurrentTool) (THIS_ void *pvDguis) PURE;

        /*
        Returns the IMsoDrawing you would like me to use to store the clipboard.
        */
        MSOMETHOD_(BOOL, FGetClipboardDrawing) (THIS_ void *pvDguis,
                interface IMsoDrawing **ppidg, MSOCDID cdid) PURE;

        /*
        Returns the IMsoDrawing you would like me to use to store the clipboard.
        */
        MSOMETHOD_(BOOL, FGetScrapDrawing) (THIS_ void *pvDguis,
                interface IMsoDrawing **ppidg) PURE;

        /*
        Releases the clipboard IMsoDrawing.
        */
        MSOMETHOD_(void, ReleaseClipboardDrawing) (THIS_ void *pvDguis,
                interface IMsoDrawing *pidg, MSOCDID cdid) PURE;

        /*
        Asks the host to set the clipboard for us.  If the host cannot do so, it is
        responsible for calling Release on the IDataObject, in order to free it.
        */
        MSOMETHOD_(void, RequestSetClipboard) (THIS_ void *pvDguis,
                IDataObject *pidoClipboard) PURE;

        /*
        Gives a chance for the host to clear up its data structures to
        indicate that the Office clipboard has been deleted by someother copy.
        */
        MSOMETHOD_(void, ReleaseClipboard) (THIS_ void *pvDguis,
                IDataObject *pidoClipboard) PURE;

        /*
        Is this data the current OLE clipboard?
        */
        MSOMETHOD_(BOOL, FIsCurrentClipboard) (THIS_ void *pvDguis,
                IDataObject *pidoClipboard) PURE;


        /*
        Ask the host to fill in the rgtcr with swatch controls, and return the number of
        controls created in pctcr. The host can also specify popo, one of msopopoxxx above.
        */
        MSOMETHOD_(BOOL, FPopulateRgtcr) (THIS_ void *pvDguis, MSOTCR *rgtcr,
                interface IMsoControlSite *pics, int *pctcr, MSODGCID dgcid,
                int *ppopo) PURE;

        /*
        Asks the host to save whatever data is appropriate for the given
        drawing control user.  (This data will be read back by the host in its
        implementation of IMsoToolbarSet::PicuLoadUserData.)

        The ppv will be whatever the host set the ppv to when creating the control.
        If the control is ever cloned, the new control will use the
        pidguiClone specified in the MSODGUISI as its DGUI; if that's NULL, it will
        use this DGUI.
        */
        MSOMETHOD_(BOOL, FSaveDrawingControlUser) (THIS_ void *pvDguis,
                MSODGCID dgcid, IMsoControl *pic, void **ppv, IStream *pistm) PURE;

        /* ----- Methods to help with recording VBA ----- */
        /* For several reasons, Office Drawing VBA is NOT recorded
                through IMsoSimpleRecorder.  But we tried to make it easy for
                the host to share implementation. */

        /* Office will pass wz's to FRecordVBALine whenever the host
                calls FExecuteCommand on a command with the "fRecordVBA" bit set.
                The string returned by 'wzLine' is owned by Office. The buffer
                is allocated on the stack. Host should not delete this buffer. */
        MSOMETHOD_(BOOL, FRecordVBALine) (THIS_ void *pvDguis,
                WCHAR *wzLine, MSODGC *pdgcRecording) PURE;

        /* Office will call FGetVBAString  to get various strings from the
                host while constructing recorded VBA.  In dgrvs we'll pass an
                identifier for the particular string we want.  In wParam we'll pass
                some random data dependent on the dgrvs we passed.  The host must
                copy the required string into the buffer pointed to by pwch,
                and return in *pcwch the number of characters they copied.
                We'll pass the size of the buffer in cwchMax; we promise that this
                will always be >= 255, but if the string you want to give us is longer
                than 255 characters you'll have to check this length to be sure not
                to overwrite the buffer, returning FALSE if you don't fit.
                By the way, we'll set *pcwch to 0 before we call you so if your
                version of the requested string is the empty string you can just
                return TRUE. */
        MSOMETHOD_(BOOL, FGetVBAString) (THIS_ void *pvDguis,
                MSODGRVS dgrvs, WPARAM wParam, WCHAR *pwch, int *pcwch, int cwchMax,
                MSODGC *pdgcRecording) PURE;

};

/* MSODGUISI (Drawing User Interface Site Information)
        You pass one of these to MsoFCreateDrawingUserInterface.  It contains
        a pointer to your DGUIS as well as the pvDgcmps that allows you to
        overload DGUISs.  It also contains any reasonably constant information
        that describes your DGUISI. */
/* Hey! Office developers, if you add a field in here, add code
        in MsoInitDguisi to initialize it and AssertValidDguisi to
        verify it. */
/* After each flag is a little comment of the form dWxP.  These letters
        indicate the default value of the flag and whether the flag is on
        or off in Word, Excel, and PowerPoint.  Capital letter are "on"
        and lower-case are "off". */
typedef struct _MSODGUISI
        {
        IMsoDrawingUserInterfaceSite *pidguis;
        void *pvDguis;

        HMSOINST hmsoinst;
        interface IMsoToolbarSet *pitbs;
        union
                {
                interface IMsoTFC *pitfc;
                        /* Bogus field left here only to preserve backward compatibility.
                                FUTURE mmorgan: remove this */
                interface IMsoDrawingUserInterface *pidguiClone;
                        /* DGUI to clone toolbar control users into when they get
                                an FCloneUserClient. */
                };

        int ut;           // unit (msoutFoo) used by the host
        WCHAR wchDecimal; // Decimal symbol used by the host

        SIZE dpthDefaultTextboxSize;

        int iPixelsPerInch;     // If fUsePixelsInDialogs is true, use this resolution for conversion

        union
                {
                struct
                        {
                        ULONG fWantsEvents : 1; // D
                                /* Means that this DGUIS wants its OnEvent method
                                        called whenever a DGUI Event happens. */
                        ULONG fRecordVBA : 1; // d
                                /* TRUE when commands run in this DGUI should be emitting VBA. */
                        ULONG fSetDefaults : 1; // d
                                /* TRUE when commands set defaults */
                        ULONG fPointerMode : 1; // dwx
                                /* This DGUI is in Pointer Mode */
                        ULONG fPointerModeMeansOnlyShapes : 1; // DWXp
                                /* We'll not hit text, objects, etc. when in pointer mode. */
                        ULONG fExitTextEditBeforeChangeSelection : 1; // DWXp
                                /* The hosts text edit mode depends on the selection, so if we're
                                        in text edit mode and want to select some other shape we have
                                        to first close text edit and then select the new shape. */
                        ULONG fDontRenderLinkFormats : 1;
                                /* Should we provide cfLink* formats when we copy an OLE
                                        object. */
                        ULONG fUsePidguiClone : 1; // dwxP
                                /* If true, we treat the pidgui/pitfc union above as a pidgui.
                                        This is a HACK placed here to preserve binary compatibility
                                        so we don't break apps that still set the pitfc field.
                                        FUTURE mmorgan: remove this along with the pitfc field */
                        ULONG fUsePixelsInDialogs : 1; // dW
                                /* If true, format some controls in some dialogs with utPixel,
                                        instead of pdguisi->ut. Pixel-formattability of control is
                                        TMC-based and can be found by FIsPixelAwareTmc. */
                        ULONG fHypCntrlClickFollow : 1; // dW
                                /* If true, then a control click follows a hyperlink on a drawing.  Otherwise,
                                   use single click to follow the hyperlink (default) */
                        ULONG fAutoCreateCanvas : 1;
                                /* For Word's canvas auto-create feature */
                        ULONG fIgnoreTextSelInPointerMode : 1;
                                /* Do not use text selection to create diagram when in pointer mode. */
                        ULONG grfUnused : 20;
                        };
                ULONG grf;
                };

        MsoDEBDefine(dgui);

#if DEBUG
        /* Hey! Leave this field at the end! */
        LONG lVerifyInitDguisi; // Used to make sure people use MsoInitDgcmpsi
#endif // DEBUG
        } MSODGUISI;

#define cbMSODGUISI (sizeof(MSODGUISI))

/* You must call MsoInitDguisi on any new DGUISI you're filling out
        before passing it into MsoFCreateDrawingUserInterface. */
MSOAPI_(void) MsoInitDguisi(MSODGUISI *pdguisi);

/* MSODMBDWBD -- Debug Message argument Block for msodmDguiWriteBeForDgc. */
#if DEBUG
typedef struct MSODMBDWBD
{
        MSODGC *pdgc; // The command to memory-mark.
        LPARAM lParam; // The "lParam" passed around during memory marking
} MSODMBDWBD;
#endif // DEBUG

// MSOFDGUIGCS -- Flags passed to DGUI::FGetCommandString
#define msofdguigcsUndo     (1<<0)
#define msofdguigcsRedo     (1<<1)
#define msofdguigcsRepeat   (1<<2)
#define msofdguigcsToolTip  (1<<3) // Used only for split buttons
#define msofdguigcsMode     (1<<4)
#define msofdguigcsName     (1<<5)
        /* Returns the "Name" of the command, which has no spaces in it. */
#define msofdguigcsHelp     (1<<6)
#define msofdguigcsDrawPrefix (1<<7)
        /* Ignored unless "Name" is specified, in which case "Draw" is prepended
                to the command name. */

#define msogrfdguigcsType \
        (msofdguigcsUndo | msofdguigcsRedo | msofdguigcsRepeat | \
        msofdguigcsToolTip | msofdguigcsMode | msofdguigcsHelp | msofdguigcsName)
        /* These are all the "type" options to DGUI::FGetCommandString;
                you must pass in exactly one of these. */
        // TODO peteren(davebu) -- why exactly are these a bitfield?

/* TODO peteren: Get rid of these old non-flag constants and then these
        temporary renamings. */
#define msodguigcsUndo msofdguigcsUndo
#define msodguigcsRedo msofdguigcsRedo
#define msodguigcsRepeat msofdguigcsRepeat
#define msodguigcsToolTip msofdguigcsToolTip

/* MSOFDGUIND -- Flags to pass to DGUI::FGetNextDgcid */
#define msogrfdguindNil  0
#define msofdguindUI     (1<<0)
        /* Filters for commands that can be in the UI (on toolbars, etc.) */

// MSOTTID -- Tab Template Identifiers
typedef enum
        {
        msottidNil,
        msottidFillNLine,
        msottidSize,
        msottidPosition,
        msottidPicture,
        msottidTextbox,
        msottidWeb,
        msottidFormatObject, //means we are passing in a struct with all the tt's
        } MSOTTID;

/* MSOFDGUIM -- Drawing User Interface Mode */
/* These are passed to DGUI::GrfdguimSetModes */
#define msogrffdguimNil  (0)
#define msofdguimPointer (1<<0)
#define msogrfdguimAll (msofdguimPointer)

/* IMsoDrawingUserInterface (idgui).
        TODO peteren: Comment? */
#undef  INTERFACE
#define INTERFACE IMsoDrawingUserInterface
DECLARE_INTERFACE_(IMsoDrawingUserInterface, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingUserInterface methods

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Call Free to make a DGUI go away.  This will NOT free any drawings,
                or other stuff allocated under a DGUI. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* PdguisiBeginUpdate returns a pointer to our MSODGUISI.  This will
                always succeed. The caller may examine and (unless they passed in
                fReadOnly == TRUE) change the fields in the MSODGUISI.  The caller must
                in any case call EndDguisiUpdate when they're finished. */
        MSOMETHOD_(MSODGUISI *, PdguisiBeginUpdate) (THIS_ BOOL fReadOnly) PURE;

        /* After you call PdguisiBeginUpdate you have to call EndDguisiUpdate,
                passing back in the pointer you got from PdguisiBeginUpdate and
                TRUE for fChanged if you changed any fields in it. */
        MSOMETHOD_(void, EndDguisiUpdate) (THIS_ struct _MSODGUISI *pdguisi,
                BOOL fChanged) PURE;

        // ----- Command methods

        MSOMETHOD_(BOOL, FCreateCommand) (THIS_ MSODGC **ppdgc, MSODGCID dgcid) PURE;

        MSOMETHOD_(void, FreeCommand) (THIS_ MSODGC *pdgc) PURE;

        /* Call FCloneCommand to make a copy of a command. */
        MSOMETHOD_(BOOL, FCloneCommand) (THIS_ MSODGC *pdgc,
                MSODGC **ppdgcClone) PURE;

        /* Call GetCommandInfo to get the MSODGCI  corresponding to a
                particular Office Drawing Command.  You can look in here
                for such helpful stuff as the recommended tcid for this dgcid,
                whether or not this Commmand can latch, when it's disabled,
                whether it's implemented or not, etc. */
        MSOMETHOD_(void, GetCommandInfo) (THIS_ MSODGCI *pdgci, MSODGCID dgcid) PURE;

        /* Call DgcidFromTcid to get the Office Drawing Command that corresponds
                to a particular MSOTCID.  This method is horribly slow.  Do anything
                you can to avoid calling it!  If you can't we could possibly
                make it faster. */
        MSOMETHOD_(MSODGCID, DgcidFromTcid) (THIS_ int tcid) PURE;

        /* GetCommandTcid returns the MSOTCID for a command.
                You can specify the command _either_ by passing in a pointer to
                an MSODGC (and dgcid == msodgcidNil) or by passing in pdgc == NULL
                and the MSODGCID you care about.
                Note that there's an MSOTCID field in the MSODGCI structure
                you can get from GetCommandInfo, but that only works for commands that
                always have exactly the same MSOTCID.  This method works for commands
                that change tcid based on state.
                This method actually returns TWO tcid's, one for the general case and
                another for when the command is on its "home" menu.  For example,
                "Format Picture" is usually on the "Format" menu where it's string
                should be just "Picture".  You can pass NULL for ptcidHome if you
                dont' care about it. */
        MSOMETHOD_(void, GetCommandTcid) (THIS_ MSODGC *pdgc, MSODGCID dgcid,
                int *ptcid, int *ptcidHome) PURE;

        /* If you've been running a DGC and posting undo records, you
                probably want to store an MSODGCID with the undo record
                to use to get strings to put on the Undo menu.  However, storing
                the DGCID in the DGC you're running occasionally gives you
                overly general strings.  So call this method instead. */
        MSOMETHOD_(void, GetCommandDgcidForUndo) (THIS_ MSODGC *pdgc,
                MSODGCID *pdgcid) PURE;

        /* FIsCommandEnabled returns TRUE if the given command
                can currently be run.  If this returns FALSE you probably want
                to gray out menu items / toolbar buttons for this command. */
        MSOMETHOD_(BOOL, FIsCommandEnabled) (THIS_ MSODGC *pdgc) PURE;
        /* Quickie version of FIsCommandEnabled if you don't need to
                pass in arguments. */
        MSOMETHOD_(BOOL, FIsDgcidEnabled) (THIS_ MSODGCID dgcid) PURE;

        MSOMETHOD_(MSOSPT, ShapeTypeFromDgcid) (THIS_ MSODGCID dgcid) PURE;

        MSOMETHOD_(BOOL, FIsCommandLatched) (THIS_ MSODGC *pdgc) PURE;
        /* Quickie version of FIsCommandLatched if you don't need to
                pass in arguments. */
        MSOMETHOD_(BOOL, FIsDgcidLatched) (THIS_ MSODGCID dgcid) PURE;

        MSOMETHOD_(void, BeginQueryCommandEnabling) (THIS) PURE;

        MSOMETHOD_(void, EndQueryCommandEnabling) (THIS) PURE;

        /* Call FExecuteCommand to actually run a command. */
        MSOMETHOD_(BOOL, FExecuteCommand) (THIS_ MSODGC *pdgc) PURE;
        /* Quickie version of FExecuteCommand if you don't need to
                pass in arguments. */
        MSOMETHOD_(BOOL, FExecuteDgcid) (THIS_ MSODGCID dgcid) PURE;

        // ----- Methods for making complicated command bar controls

        /* FCreateDrawingControlUser will create a control user of the type
                specified by dgcid.  We return a pointer to the new control user
                but no ppv.  Our control users actually assert if you pass them a
                non-NULL ppv. To dispose of one of these control users,
                just call its NotifyDelete method. */
        MSOMETHOD_(BOOL, FCreateDrawingControlUser) (THIS_
                interface IMsoControlUser **ppicu, MSODGCID dgcid) PURE;

        /* FCreateDrawingControlUser will create a menu user of the type
                specified by dgcid.  We return a pointer to the new menu user*/
        MSOMETHOD_(BOOL, FCreateDrawingMenuUser) (THIS_
                interface IMsoControlUser **ppicu, MSODGCID dgcid,
                IMsoSwatchUser *pisu, IMsoSwatchUser* pisu2 ) PURE;

        /* FQuerySwatchSelection lets app-implemented color swatches
                ask us which color should be the current color.  If no color
                is selected, FALSE is returned; otherwise, the cr field of
                the MSODGC contains the current color. */
        MSOMETHOD_(BOOL, FQuerySwatchSelection) (THIS_ MSODGC *pdgc) PURE;

        // ----- Other methods

        /* Call FIsModalToolbarVisible() to find out of a modal toolbar
                that depends on Drawing state should be up.  Right now this
                can handle msotbidTextEffects and msotbidPictureTools. */
        MSOMETHOD_(BOOL, FIsModalToolbarVisible) (THIS_ int tbid,
                void *pvContext) PURE;

        /* FRecordVBA returns TRUE if this DGUI is set to record VBA.
                Note that all this really means is that Command created by
                this DGUI will have their fRecordVBA bit set by default. */
        MSOMETHOD_(BOOL, FRecordVBA) (THIS) PURE;

        /* Call SetRecordVBA to change whether or not the DGUI is setting
                the fRecordVBA bit in the Commands it creates. */
        MSOMETHOD_(void, SetRecordVBA) (THIS_ BOOL fRecordVBA) PURE;

        /* Cal FCreateDlgControls to create the controls to go into the GCC in
                an sdm dialog */
        MSOMETHOD_(BOOL, FCreateDlgControls) (THIS_
                interface IMsoControlContainer *picc, unsigned int tmc,
                unsigned int wBtn, BOOL *pfStop, int *pdx, int *pdy) PURE;

        /* Call FGetCommandString to get a string from a command.  Pass us
                a pointer to a buffer of WCHARs (wz) and the buffer's length (cchMax).
                We fill out the buffer and return TRUE on success.  If you pass a
                non-NULL pcch we'll return in *pcch the length of the string (not
                including the terminating zero, so that the largest value you can get
                back in *pcch is cchMax-1).  Please pass in NULL for pdgc, we're hoping
                to remove that argument.  Pass in the MSODGCID identifying the command
                you're interested in.  In grfdguigcs you pass flags about what kind of
                string you want (undo, redo, repeat, tooltip, etc). */
        MSOMETHOD_(BOOL, FGetCommandString) (THIS_ MSODGC *pdgc, MSODGCID dgcid,
                ULONG grfdguigcs, WCHAR *wz, int cchMax, int *pcch) PURE;

        /* SetPib sets up the given blip to be inserted by the next picture
                insertion. The DGUI will take over ownership of the blip and the
                UNICODE string and will mark and free them as appropriate.  The blip
                may be passed in or specified as a file name or URL - the MSOBLIPFLAGS
                say which and also control whether the blip is saved with the document
                (note: it may be saved as a result of being in some other shape).  If
                the fUpdate flag is passed the blip is a "placeholder" and the name
                (which must be passed and be a URL name) will be resolved and a new
                blip created.  If the fSynchronous flag is true a new thread will not
                be spawned to do any slow resolution or loading.  Arguments are
                interpreted in exactly the same way as those passed to the IMsoDrawing
                FCreateBlipShape method. */
        MSOMETHOD_(void, SetPib) (THIS_ IMsoBlip *pib, const WCHAR *pwzFile,
                const WCHAR *pwzDescription, MSOBLIPFLAGS blipFlags, BOOL fUpdate,
                BOOL fSynchronous) PURE;

        /* Call FAddTtToArgs after a dialog with office tabs has come down
                and the app has created a command and wishes to place the
                spp's to be applied to the selection in the command arguments. */
        MSOMETHOD_(BOOL, FApplyTt) (THIS_ interface IMsoDrawingSelection *pidgsl,
                MSOTTID ttid, void *ptt, void *pttMask) PURE;

        /* Fill in tab templates for office tabs looking at the selection,
                called before the host brings up a dialog with office tabs. */
        MSOMETHOD_(BOOL, FGetTtFromSelection) (THIS_
                interface IMsoDrawingSelection *pidgsl, MSOTTID ttid, void *ptt,
                void *pttMask) PURE;

        /* Use FGetNextDgcid to enumerate the Office Drawing Commands.  Set
                a local MSODGCID to msodgcidNil to start.  You can specify what
                kind of commands you want to see in grfdguind. */
        MSOMETHOD_(BOOL, FGetNextDgcid) (THIS_ MSODGCID *pdgcid,
                ULONG grfdguind) PURE;

        /* Use GrfdguimSetModes to change any UI modes (of which there's
                currently exactly one, "Pointer Mode").  You pass in three
                grfdguim's, one of modes to turn on, one of modes to turn off,
                and one of modes to toggle.  We return the modes we're in
                after any changes. */
        MSOMETHOD_(ULONG, GrfdguimSetModes) (THIS_ ULONG grfdguimOn,
                ULONG grfdguimOff, ULONG grfdguimToggle) PURE;

        /* Set the fill blip used by the shape formatting dialog.  Word
        needs this to support old macros. */
        MSOMETHOD_(void, SetFormatShapeBlip) (THIS_ IMsoBlip *pib) PURE;

        /* Retrieve the fill blip used by the shape formatting dialog.
                Word needs this to support old macros. */
        MSOMETHOD_(interface IMsoBlip *, PibFormatShape) (THIS) PURE;

        /* FQuerySwatchSelectionExt like FQuerySwatchSelection, but supports
        extended colors. Should only be used for extended color dgcids. */
        MSOMETHOD_(BOOL, FQuerySwatchSelectionExt) (THIS_ MSODGC *pdgc) PURE;

        MSOMETHOD_(BOOL, FInCanvasAutoCreateMode) (THIS) PURE;

        /* NotifyCancelCurrentTool is used by Word to end canvas auto-create. */
        MSOMETHOD_(void, NotifyCancelCurrentTool) (THIS) PURE;

        MSOMETHOD_(BOOL, FInkAnnotationMode) (THIS) PURE;
        MSOMETHOD_(BOOL, FMadeInkOrEraseEdits) (THIS_ BOOL fCheckSurface) PURE;
};

/* MsoFCreateDrawingUserInterface makes a new Drawing User Interface,
        attached via the specified MSODGUISI to a host-supplied
        IMsoDrawingUserInterfaceSite. Most hosts will probably call this
        exactly once per run. */
MSOAPI_(BOOL) MsoFCreateDrawingUserInterface(
        IMsoDrawingUserInterface **ppidgui, MSODGUISI *pdguisi);


/****************************************************************************
        Drawing

        A Drawing is a set of Shapes ... TODO peteren: Better comment.
        This section has Drawing Sites too.
*****************************************************************************/

/* DGDCR -- Drawing default colors. An optional parameter to allow clients
        change the defaults colors for shapes. Although this effect can be achieved
        by setting defaults properties, passing this structure is more
        efficient. */
typedef struct
        {
        COLORREF crFill;
        COLORREF crFillBack;
        COLORREF crLine;
        COLORREF crLineBack;
        COLORREF crShadow;
        COLORREF crShadowHighlight;
        } MSODGDCR;


/* DGSI = DrawinG Site Info.  Every Drawing is connected to a Drawing
        Site object, which provides necessary call-back functions for that
        Drawing.  However, a Drawing Site is also defined by some relatively
        constant pieces of data.  Instead of having tons of Get functions
        in the Drawing Site or tons of Set functions in the Drawing itself
        we've defined a single structure to contain all of this data, the
        MSODGSI.  You have to pass a pointer to one of these to any methods
        that create a Drawing.  Probably the most important thing it contains
        is the pointer to the site object.

        The trick is how to ensure that this structure is kept up to date.
        There's two different possibilities, which I'll call "push" and "pull".
        A "push" is when the host realizes that one of the values it passed to
        its Drawing is wrong and should be changed.  A "pull" is when the
        Drawing calls back to the host to get the most recent value of
        something.  In general the things in the DGSI are things that change
        very rarely, and then at moments known to the host, so we'll choose a
        primarily "push" model. The host (implementor of Drawing Site) can call
        a method of the Drawing (DG::PdgsiBeginUpdate) to get a pointer to the
        DGSI, then change the data, and then call another method
        (DG::EndDgsiUpdate) to let the DG know it's finished.  These scheme
        gives us great opportunities for debug-checks on the changes the host
        makes.

        Anyway, there still may be cases where it's hard for the host to know
        when one of these values has changed, and yet its important for the
        Drawing to have the updated value before some important event takes
        place.  This the case where a pull might have been useful.  To handle
        these cases we'll have the Drawing fire events; the DGS can then
        on some events call back into the Drawing to fix the DGSI.

        If all this seems a bit overblown, please realize that it allows to DG
        to run without calling fetch functions, it allows the DGS to update all
        the data describing it with only a pair of functions, it allows for
        either immediate or delayed updating of data, and it comes with only
        two additional DG methods and (kind of) one DGS method (the event
        method).  We use exactly the same mechanism with Drawing Views and
        Drawing View Sites. */

MsoDEBDeclare(dg, msodgeDgsFirst, msodgeDgsMax);

typedef struct _MSODGSI
        {
        interface IMsoDrawingSite *pidgs;
                /* DGS interface pointer.  This is a push field; you can change it
                        by calling DGI think you can change this whenever you
                        want. */
        void *pvDgs;
                /* client data pased to all DGS methods.  I think you can also change
                        this whenever you want. */
        int cxyhInch;
                /* The meaning of H units in this Drawing.  You can't ever
                        change this field; too much stuff depends on it. */
        RECT rchBounds;
                /* The size of the Drawing. */
                /* FUTURE peteren: Eventually DGs won't have a "size"; instead
                        individual "pages" of a DG will have a "size".  But this
                        "page" concept (for Word) doesn't really exist yet. */
        int dxyhAnchorSlop;
                /* For clients that return a different a value for rchAnchor
                        than we gave you (shame on you), tell us maximum amount of difference
                        we will see. */
        int dxyhConnectorMargin;
                /* For clients that use connectors, the margin for connector routing. */
        int dxhDefaultRect;
                /* The default width of rectangle shape. The default sizes for other
                   shapes are sized proportionally. MsoInitDgsi sets this
                        value to 1 inch. */
        MSODGDCR *pdgdcr;
                /* The initial default colors for new shapes. May be null. Changing
                   the defaults of the DGG will override these defaults. Should not change
                        during the life of the drawing. */
        int cMaxSchemeColors;
        /* For clients that use color schemes. Specify the maximum of scheme colors
         known at init time. Escher Object Model needs it to do range-checking.
         MsoInitDgsi sets it to 0 */
        LONG dxaMax;
        /* Maximum width of a shape in emus.  The reason why these are in emus
                rather than H units is so that we can initialize these in MsoInitDgsi
                prior to knowing what cxyhInch is*/
        LONG dyaMax;
        /* Maximum height of a shape in emus */
        MSOSPID spidMax;
           /* TODO KeChengH: Get rid of this. */
        union
                {
                struct
                        {
                        ULONG fAnchorIsRch : 1;
                                /* TRUE iff the pvAnchors this Drawing Site returns
                                        are really just prch's.  This will let us avoid making
                                        stupid calls to methods in the DGS.  This will be TRUE,
                                        I believe, in PowerPoint and in the Server. */
                        ULONG fWantsEvents : 1;
                                /* Means that this DGS wants its OnDrawingEvent method
                                        called whenever a Drawing Event happens. */
                        ULONG fBackgroundShape : 1;
                                /* Drawing has a background shape which is created during
                                        creation of the drawing. */
                        ULONG fConnectors : 1;
                                /* Drawing supports connectors. Host will tell about all
                                   anchor changes. */
                        ULONG fBoundShapes : 1;
                                /* Constrain shapes to within the rchBounds. Followed by connectors
                                   and other rules. */
                        ULONG fNoAtFonts : 1;
                                /* This host does not support "at" fonts, so when setting
                                        msotxfl properties in far east versions we should rotations
                                        that don't rely on "at" fonts. */
                        ULONG fCannotContainLabels : 1;
                                /* Used to indicate that you would prefer to record Shapes.AddTextbox
                                        instead of Shapes.AddLabel. */
                        ULONG fSuspendInvalidation : 1;
                                /* To be used when the host does not want the drawing to
                                        record invalidation. */
                        ULONG fCancelCloneShapeMaimsTempSpace : 1;
                                /* Bogus flag needed until we get interface changes again. */
                        ULONG fCanHaveTextShadow : 1;
                        ULONG fCanHavePerspectiveTextShadow : 1;
                        ULONG fCanHaveDoubleTextShadow : 1;
                                /* Used by hosts that can render shadows for text, so we know
                                        when to enable our shadow UI.  (All rendering of text shadows
                                        is done by the host.) */
                        ULONG fNoTextRotation : 1;
                                /* If this flag is true, we disable rotation for shapes created
                                        with the textbox tool.  Shapes that have had text attached to
                                        them are not affected. */
                        ULONG fEnableAutoMargins : 1;
                                /* If this flag is true, the msopidFAutoMargins flag on shape
                                        is followed. */
                        ULONG fGetClsidImplemented : 1;
                                /* Used to show that we can call GetClsidFromHsp on this host */
                        ULONG fDebugPrecisionLossOkay : 1;
                                /* Quasi-bogus flag, allows XL Macro sheet to lose anchor precision */
                        ULONG fGroupTangle : 1;
                                /* This flag is set only in Pub to support the groupTangle prop.
                                        The groupTangle prop's behavior is enabled when this flag is set.
                                        All other apps treat the grouptangled shape as a group shape*/
                        ULONG fCreateInsetPenShapes : 1;
                                /* Set by host app to create shapes with insetpen. */
                        ULONG fCanvasConnectersOnly : 1;
                                /* Indicates that connectors don't work at the top-level */
                        ULONG fSupportsPseudoInline : 1;
                                /* Word's pseudo-inline layer. */
                        ULONG fIgnoreMarginsForHitTest : 1;
                                /* MSOSP::FGetTextRcvi won't include the margins when calculating the
                                   text frame RCVI if this flag is set. Currently used only by Publisher,
                                   this flag allows Publisher to retain control of mouse event handling
                                   during hit testing while still using the msopidD*Text properties for
                                   margins storage. */
                        ULONG fNoClearClipDataObject : 1;
                                /*  CDO:Stupify() checks this flag before blowing away
                                        the contents of its associated drawing.  Currently used only by
                                        Publisher to be able to re-use clipboard drawings without having Draw
                                        reset their content. */
                        ULONG fNoLineDrawDash : 1;
                                /* If set forces a dashed line to draw when shapes don't
                                        actually have a line. Currently only used by Pub. One
                                        could argue that this should be in the MSODGVSI, but
                                        there are problems with that because the DGV isn't always
                                        available when we need to check this flag. */
                        ULONG fNoFillHitTest : 1;
                                /* Set if app wants diagram styles to use color schemes. */
                        ULONG fUseColorSchemeForDiagrams : 1;
                                /* If set forces non-filled shapes to hit test as if they
                                        were filled. Currently only used by Pub. This really
                                        should be in the MSODGVSI too (like fNoLineDrawDash), but
                                        isn't for the same reason. */
                        ULONG fOverrideCanvasBehavior : 1;
                                /* If set forces shapes that are dragged over diagrams or drag/dropped
                                        into diagram bounds to be pasted floating on the drawing. This is set
                                        to TRUE only in PPT. */
                        ULONG fSupportCF_HTMLForShapeSelection : 1;
                                /*If set, the app supports postinf CF_HTML format for shape selection.
                                        This is currently supported only in Word. */
                        ULONG fOkToApplyProperty : 1;
                                /* If set, the app says it's OK for property application to occur
                                        at this time, despite other debug checks. */
                        ULONG fDoNotConvertOrgChart : 1;
                                /* If set, the app says do not convert an orgchart OLE object.*/
                        ULONG fPublisherWebView: 1;
                                /* If set, turns off features which aren't supported in HTML.
                                   Used only by Pub.  Turns off inset pen and extra line styles (left,
                                   top, right, bottom, and column). */
                        ULONG fNoLineDrawDashForPic : 1;
                                /* similar to fNoLineDrawDash, but this one is for picture shapes.
                                        If set forces a dashed line to draw when pictures don't
                                        actually have a line. Currently only used by Pub.  */
                        ULONG grfUnused : 2;
                        };
                ULONG grf;
                };
        POINT ptBlipSize;
        MsoDEBDefine(dg);

#if DEBUG
        LONG lVerifyInitDgsi; // Used to make sure people use MsoInitDgsi
#endif // DEBUG
        } MSODGSI;
#define cbMSODGSI (sizeof(MSODGSI))

/* You should call MsoInitDgsi on any new DGSI you're filling out
        before passing it into a method that creates a Drawing.  This way
        we can add new fields to the structure without adding new code in
        all the hosts; just a recompile will be sufficent. */
MSOAPI_(void) MsoInitDgsi(MSODGSI *pdgsi);

/* MSOFDGIs (Drawing Invalidation Flag) are passed as arguments to
        IMsoDrawing::Invalidate.  Many of these flags are buffered up
        in the DG or even in individual SPs (until DG::EveryoneCopeWithDgInval).
        Be very careful about adding new bits in here. */
#define msogrfdgiNil                   0
/* First come bits are buffered in the DG and in each SP. */
#define msofdgiOneShapeAppear          (1<<0)
        /* Shape appeared (was inserted, undeleted, etc.) */
#define msofdgiOneShapeVanish          (1<<1)
        /* Shape vanished (was deleted, insertion undone, etc.) */
#define msofdgiOneShapeCoverMore       (1<<2)
        /* Shape draws in pixels it didn't before. */
#define msofdgiOneShapeRevealMore      (1<<3)
        /* Shape doesn't cover pixels it didn't before. */
#define msofdgiOneShapeZOrder          (1<<4)
        /* Shape moved in the Z Order. */
#define msofdgiOneShapeRedraw          (1<<5)
        /* Shape needs to be redrawn */
#define msofdgiOneShapeText            (1<<6)
        /* Shape's text needs to be redrawn. */
#define msofdgiOneShapeLayout         (1<<7)
   /* Shape's associated Diagram needs to be re-layout. */
#define msofdgiOneShapeUnused8         (1<<8)
#define msofdgiOneShapeUnused9         (1<<9)
        /* Because adding one of these bits is such a pain (it messes up the
                DG structure and the SP structure and forces you to write code in
                AssertOfficeDrawing) I added a few bits for future expansion. */

/* Next come bits that are buffered in the DG but not in each SP. */
#define msofdgiFull                    (1<<10)
#define msofdgiUnused11                (1<<11)
#define msofdgiSlide                   (1<<12)
        /* Lots of shapes may have changed positions; check them all next time
                through DGV::Validate. */

/* Finally come bits that are just arguments to DG::Invalidate; they're
        not stored anywhere.  Add these from the top down */
#define msofdgiOneShapeWrap            (1<<31)

#define msofdgiOneShapeInsert msofdgiOneShapeAppear
#define msofdgiOneShapeDelete msofdgiOneShapeVanish
        /* TODO peteren: Remove "insert" and "delete", replacing them
                with their new names, "appear" and "vanish". */

#define msogrfdgiOneShapeChange \
        (msofdgiOneShapeCoverMore | msofdgiOneShapeRevealMore)
        /* Most of the code that changes a Shape will probably use
                msogrfdgiOneShapeChange, since it doesn't really know if
                more or less stuff is showing. */
#define msogrfdgiOneShapeChangeWithZ \
        (msofdgiOneShapeCoverMore | msofdgiOneShapeRevealMore | \
        msofdgiOneShapeZOrder)
        /* A little more invalidating that OneShapeChange, this one also
                marks the Shape as having moved in the Z-order. */
#define msogrfdgiShape \
        (msofdgiOneShapeAppear | msofdgiOneShapeVanish | \
        msofdgiOneShapeCoverMore | msofdgiOneShapeRevealMore | \
        msofdgiOneShapeZOrder | msofdgiOneShapeRedraw | \
        msofdgiOneShapeText | msofdgiOneShapeLayout | \
        msofdgiOneShapeUnused8 | msofdgiOneShapeUnused9)
        /* msogrfdgiShape should contain all of the invalidation flags
                that can be applied to a single Shape.  These bits are stored in SPs. */
#define msogrfdgiOneShape msogrfdgiShape // TODO peteren: Remove this old name for msogrfdgiOneShape
#define msogrfdgiDrawing \
        (msogrfdgiShape | msofdgiFull | msofdgiUnused11 | msofdgiSlide)
        /* msogrfdgiDrawing contains all the invalidation bits that are stored
                in the DG (this includes those stored in SPs). */
#define msogrfdgiArgument \
        (msofdgiOneShapeWrap)
        /* msogrfdgiArgument contains all the invalidation bits that are just
                arguments to DG::Invalidate. */
#define msogrfdgiAll \
        (msogrfdgiDrawing | msogrfdgiArgument)
        /* msogrfdgiAll is the union of all of the msofdgi values. */
#define msogrfdgiViewsCare \
        (msofdgiOneShapeAppear | msofdgiOneShapeVanish | \
        msofdgiOneShapeCoverMore | msofdgiOneShapeRevealMore | \
        msofdgiOneShapeZOrder | msofdgiOneShapeRedraw | \
        msofdgiOneShapeText | msofdgiFull | msofdgiSlide | msofdgiOneShapeLayout)
        /* msogrfdgiViewCares is the set of all invalidations about which DGVs
                care.  This shouldn't really be public, but it's simplest to #define
                it next to these other definitions. */
#define msogrfdgiSelectionsCare \
        (msofdgiOneShapeAppear | msofdgiOneShapeVanish | \
        msofdgiOneShapeCoverMore | msofdgiOneShapeRevealMore | \
        msofdgiOneShapeZOrder | msofdgiOneShapeRedraw | \
        msofdgiOneShapeText | msofdgiFull | msofdgiOneShapeLayout)
        /* msogrfdgiSelectionsCare is the set of all invalidations about which
                DGSLs care.  This shouldn't really be public either. */
#define msogrfdgiDgsCare msogrfdgiDrawing
        /* Set of all invalidations that the drawing sites care about. */

/* MSOFDGENs are flags that describe an enumeration of Shapes.
        These are filled out in an MSODGEN before passing it to
        DG::BeginEnumerateShapes or DGSL::BeginEnumerateSelectedShapes. */
#define msogrfdgenNil                     0x00000000
#define msofdgenWantGroupOnWayIn          0x00000001
#define msofdgenWantChildren              0x00000002
#define msofdgenWantGroupOnWayOut         0x00000004
#define msofdgenOnlyWantMarked            0x00000008
#define msofdgenClearMarked               0x00000010
#define msofdgenWantAllChildrenOfMarked   0x00000020
#define msofdgenWantNoChildrenOfUnmarked  0x00000040
#define msofdgenSpecifyRootShape          0x00000080
#define msofdgenSpecifyStartShape         0x00000100
        /* TODO peteren: Clean up this hack! */
#define msofdgenUseFSelectedToo           0x00000200
#define msofdgenIncludeRootShape          0x00000400
#define msofdgenIncludeBackgroundShape    0x00000800

        /* If you pass in msofdgenSpecifyRootShape we'll look in
                MSODGEN.hspRoot for the hsp to a Group Shape among who's
                children you wish to enumerate.  Ordinarily we enumerate
                among all the Shapes in the whole Drawing. */

#define msogrfdgenShallow (msofdgenWantGroupOnWayIn)
#define msogrfdgenDeep (msofdgenWantGroupOnWayIn | msofdgenWantChildren)

#define msogrfdgenAll (msogrfdgenDeep)
        /* This should also enumerate deleted and background shapes.  Really.
                If we figure out any new wierd sorts of SPs this will return them
                too.  Callers should probably try to not make many assumptions
                about the state of the SPs this returns. */

#define msogrfdgenEnumerateShapes \
        (msofdgenWantGroupOnWayIn | msofdgenWantChildren | \
        msofdgenWantGroupOnWayOut | msofdgenOnlyWantMarked | \
        msofdgenClearMarked | msofdgenWantAllChildrenOfMarked | \
        msofdgenWantNoChildrenOfUnmarked | msofdgenSpecifyRootShape)
        /* These are all the values you can legally pass to
                DG::BeginEnumerateShapes. */

/* MSOFDGENOs are passed out of DG::FEnumerateShapes to describe the
        state of the enumeration. */
#define msogrfdgenoNil                    0x00000000
#define msofdgenoAtGroup                  0x00000001
#define msofdgenoAtGroupOnWayIn           0x00000002
#define msofdgenoAtGroupOnWayOut          0x00000004
#define msofdgenoAtGroupAboutToSkip       0x00000008
#define msofdgenoDone                     0x00000010
#define msofdgenoDeleted                  0x00000020
        /* TODO peteren: Comment */
#define msofdgenoPatriarch                0x00000040
        /* TODO peteren: Comment */
#define msofdgenoBackgroundShape          0x00000080

typedef struct MSODGEN
        {
        /* You need to set grfdgen before you call BeginEnumerateShapes,
                and then NOT CHANGE IT during the ensuing enumeration! */
        ULONG grfdgen;
        union
                {
                MSOHSP hspRoot; // For msofdgenSpecifyRootShape
                MSOHSP hspStart; // For msofdgenSpecifyStartShape
                };

        /* You can read hsp and grfdgeno during the enumeration, but
                DON'T CHANGE them. */
        MSOHSP hsp;
        ULONG grfdgeno;
        MSOHSP hspTop;

        /* Internally Office stores a bunch of other fields in here, but
                you don't get to see them.  They all fit in this array of bytes. */
        BYTE rgb[28];

#if DEBUG
        LONG l1; // lVerifyDgen2;
#endif // DEBUG
        } MSODGEN;
#define cbMSODGEN (sizeof(MSODGEN))

#define msogrfchpNil            0x00000000
#define msofchpUndoable         0x00000001
#define msofchpCreateGroup      0x00000002
#define msofchpCreatePoly       0x00000004
#define msofchpCreateFirstChild 0x00000008
#define msofchpNoChildFixups    0x00000010
#define msofchpHtmlImport       0x00000020
#define msofchpHtmlZOrderFixups 0x00000040

/* Parameter Block to FCreatePolyLineShape.

        Specifies the MsoPath and releated information need to
        create a shape. The path should be drawn in the (0,0),(dxg,dyg)
        coordinate system.
*/
typedef struct MSOCPLB
        {
        MSOSHAPEPATH spp;                       // tells if you have segement information
        BOOL                     fFilled;       // Is the polyline filled
        LONG                     dxg;                   // (dxg,dyg) maps to the bottom right of the shape
        LONG                     dyg;                   //
        int                      cPts;          // Number of points
        POINT                   *rgpt;          // the points of the shape
        MSOPATHINFO *ppi;                       // Path segment information
        int                      cSegments; // Number of path segments
        } MSOCPLB;

/* The MSORCVI structure (that's Rectangle, V and I coordinates) is used
        to _accurately_ position something in a view.  One might think that
        a simple RECT (rcv) would suffice, and indeed the MSORCVI contains
        an rcv.  If an object contains no internal coordinate information (say
        it's just a rectangle) then all it needs to render itself is that rcv.
        However, if an object contains its own internal coordinates (say it's
        a group or a polygon or something like that) that are higher resolution
        that the target device, then we need to know how to convert those
        coordinates to device coordinates, and we'd also like to know the "basis"
        for the rcv in a higher-than-device-resolution coordinate space so that
        we leave scaling ourselves down to device coordinates until the very end
        (thereby avoiding some round-off cases). */
typedef struct _MSORCVI
        {
        RECT rcv; // rectangle in V's (pixels)
        RECT rci; // same rectangle in I's (possibly higher-res than pixels)
        int cxiUnit; // how many I's are there in an arbitrary unit (X-axis)
        int cxvUnit; // how many V's are there in the same unit (X-axis)
        int cyiUnit; // how many I's are there in an arbitrary unit (Y-axis)
        int cyvUnit; // how many V's are there in the same unit (Y-axis)
        int cxiInch; // how many I's are there in an inch (X-axis)
        int cyiInch; // how many I's are there in an inch (Y-axis)
        } MSORCVI;
#define cbMSORCVI (sizeof(MSORCVI))

/* MsoFEqRcvi returns TRUE iff two MSORCVIs are identical. */
__inline BOOL MsoFEqRcvi(const MSORCVI *prcvi1, const MSORCVI *prcvi2)
        {
        return (MsoMemcmp(prcvi1, prcvi2, cbMSORCVI) == 0);
        }

/* Set the scale on an MSODC given an MSORCVI, ONLY the c... values are used -
        the rcv/rci are not required to set the scale, this routine should be
        called at least once to ensure that an MSODC is set up with valid scale
        information for the view, internally the display manager initializes any DC
        which it gets.  Failing to call this routine on a client generated MSODC
        will cause performance problems in rendering.  The routine sets the lzoom
        value, typically to 1 (to maintain maximum accuracy). */
MSOAPI_(void) MsoSetDcScale(MSODC* pdc, const MSORCVI* prcvi);


/* MSOSVI -- Shape View Info */
/* Structure full of information about a Shape's existance in a
        particular view (DGV).  We don't actually store one of these for each
        Shape in a View; we store enough to be able to generate one when
        needed.  The MSORCVI is stored relative to the closest axis (X or Y) to
        the X axis of the shape - if the shape is rotated >=45 and <135 degrees
        the X axis is closest to the Y axis of the drawing and the MSORCVI
        width gives the Y dimension of the shape, and the height the X
        dimension. */
typedef struct _MSOSVI
        {
        void **ppv;             // Per view informations
        union
                {
                MSORCVI rcviAnchor;
                        /* rcviAnchor contains the view-relative upright (that is, snapped to
                                the nearest multiple of 90 degrees) anchor rectangle of the
                                object. */
                MSORCVI rcvi; // Old name TODO remove me!
                };
        union
                {
                MSOANGLE angle; // Angle of rotation (0 <= x < 360)
                MSOANGLE        langle; // Old name TODO remove me!
                };
        union
                {
                struct
                        {
                        BOOL fFlipH : 1; // Flipped horizontally (x := -x)
                        BOOL fFlipV : 1; // Flipped vertically (y := -y)
                        ULONG axis : 2;  // Axis corresponding to langle
                        BOOL fInited : 1;  // fAxisFlip is set (remove this)
                        BOOL fHackHack161574: 1; // we are trying to hittest text as button
                        ULONG grfUnused : 2;  // Must be zero
                        MSOANGLE angle90 : 24; // Angle of rotation modulo (-45 <= x <= 45)
                        };
                struct
                        {
                        ULONG : 8;
                        MSOANGLE langle90 : 24; // Old name
                        };
                };
                void *poptView;  // Per-view property overrides used when drawing.
        } MSOSVI;
#define cbMSOSVI (sizeof(MSOSVI))

/* MSOSVIT == Shape View Info with Text */
/* Extended version of an MSOSVI that also has information about the
        location of the shape's text. */
typedef struct _MSOSVIT
        {
        MSOSVI svi;
        MSORCVI rcviTextFrame;
        } MSOSVIT;
#define cbMSOSVIT (sizeof(MSOSVIT))

/* Intialize the angle part of an MSOSVI - this code just sets the angle
        plus the axis of rotation and flip information. */
MSOAPI_(void) MsoInitSviAngle(MSOSVI *psvi, MSOANGLE langle,
        BOOL fFlipH, BOOL fFlipV);

/* Given an MSOSVI return the height or width of the shape in I units -
        the height or width is derived from the appropriate dimension of the
        rci, according to the rotation of the shape. */
MSOAPIX_(LONG) MsoDxiFromPsvi(const MSOSVI *psvi);  // Width
MSOAPIX_(LONG) MsoDyiFromPsvi(const MSOSVI *psvi);  // Height

/* Given an MSOSVI return the height or width of the shape in V units -
        the height or width is derived from the appropriate dimension of the
        rcv, according to the rotation of the shape. */
MSOAPIX_(LONG) MsoDxvFromPsvi(const MSOSVI *psvi);  // Width
MSOAPIX_(LONG) MsoDyvFromPsvi(const MSOSVI *psvi);  // Height

/* These two APIs give an angle in the range 0..<360 corresponding to the
        rotation in an MSOSVI.  Flipping the shape changes the reported rotation;
        if a shape is flipped the X axis and Y axis rotation will differ by
        180 degrees, otherwise they will be the same, to obtain an angle in the
        range -180 <= angle < 180 subtract MsoLAngle(180). */
MSOAPIX_(MSOANGLE) MsoXRotation(const MSOSVI *psvi);
MSOAPIX_(MSOANGLE) MsoYRotation(const MSOSVI *psvi);

/*      Given the current angle and the relative anchor position, this
        routine rotates the anchor by +/- 90 degrees to align it to the
        new angle. */
MSOAPI_(void) MsoRcFromAngle(RECT *prc, MSOANGLE langleOld, MSOANGLE langleNew);

/* Rotates the rectangle around pptCenter by an angle reduced to +/-45deg.
        Therefore, prc should actually be adjusted using MsoRcFromAngle(langle). */
MSOAPI_(void) MsoRotateRc(RECT *prc, MSOANGLE langle, const POINT *pptCenter);

/* MSOSPQ -- Shape Query
        These are an enumeration of "Shape Queries", this is, TRUE/FALSE
        questions you can ask of a Shape.  These are to be passed to
        IMsoDrawing::FQueryShape. */
typedef enum
        {
        msospqIsMarked,
        msospqIsMarkedSelected,
        msospqIsDeleted,
        msospqIsGroup,
        msospqIsChild,
        msospqIsConnector,
        msospqHasAttachedText,
        msospqHasAttachedOLEObject,
        msospqIsBackground,
        msospqCanHaveAttachedText,
        msospqCanConvertToPolygon,
        msospqNeedRasterMode,
        msospqSuggestRasterMode,
        msospqIsPurged,
        msospqHasFill,
        msospqHasLine,
        msospqHasTextbox,
        msospqIsPicture,          //blip is NOT guaranteed to be loaded
        msospqLinksTextbox,       //works during html load
        msospqFHidden,
        msospqIsWordArt,
        msospqIsCanvas,
        msospqIsBackgroundOfCanvas,
        msospqIsInDiagram,
        msospqIsInOrgChart,
        msospqIsDiagram,
        msospqIsLine,             //Lines, Arrows and double Arrows
        msospqSupportsInsetPen,   //shape supports insetpen rendering
        msospqIsConnectedConnector,
        msospqHasInk,             // shape contains ink (can be regular or annotation shape)
        msospqIsInk,              // shape contains ink (not an annotation)
        msospqIsInkAnnotation,    // ink annotation shape
        msospqIsAllInkAnnotationsInGroup, // shape is a group and all its children are ink annotation shapes except background shape
        msospqIsBaseInkShape,
         } MSOSPQ;



/* MSODGAR -- Drawing Access Request */
/* These are used by the IMsoDrawing thread synchronization routines. */
typedef struct MSODGAR
        {
        union
                {
                struct
                        {
                        ULONG fRead : 1;
                        ULONG fWrite : 1;
                        ULONG fMarked : 1;
                        ULONG fSelected : 1;
                        ULONG fTempSpace : 1;
                        ULONG  : 27; // Use me next
                        };
                ULONG grf;
                };
        /* TODO peteren: More fields? */
        } MSODGAR;
#define cbMSODGAR (sizeof(MSODGAR))


/* MSOFELF - Flags for FEnumFonts */
#define msofelfAddTextEffects           (1 << 0)
#define msofelfAddPictures              (1 << 1)
#define msofelfAddAll                   (msofelfAddTextEffects | msofelfAddPictures)

/* MSOPFNELF - Callback function for FEnumFonts. Return TRUE to continue. */
typedef BOOL (CALLBACK* MSOPFNELF) (const LOGFONTW *plf, void *pvParm);


/* MSOGRFDGSN - Shape Name manipulation */
/* FUTURE peteren: Those that are a single flag shouldn't have "gr" in their name */
#define msogrfdgsnCheckCollision        (1 << 0)   // Check for name collision
#define msogrfdgsnUserDefined           (1 << 1)   // Only Lookup User-Defined name
#define msogrfdgsnDoNotify              (1 << 2)   // Notify client via msodgeAfterShapeNameChange
#define msogrfdgsnLocalizedNames        (1 << 3)   // Lookup auto-generated names from localized locale
#define msogrfdgsnUnlocalizedNames      (1 << 4)   // Lookup auto-generated names from Non-Localized locale
#if NO96_165378
#define msofdgsnValidateCh              (1 << 5)   // (PUT) check to make sure only valid characters in name before applying
#endif
#define msofdgsnHTML                    (1 << 6)   // Lookup HTML generated names
#define msogrfdgsnDefaults          msogrfdgsnLocalizedNames | msogrfdgsnUnlocalizedNames
// Default setting for Automation
#define msogrfdgsnOADefaults        msogrfdgsnCheckCollision | msogrfdgsnDoNotify | msogrfdgsnUnlocalizedNames

/* Flag for FGetShapeIDString */
#define msospidsSpType          0    // String is formed for shapetype (VML)
#define msospidsSpid            1    // String is formed from spid
#define msospidsObjName         2    // String is formed from msopidWzName and is unique
#define msospidsSpidDupObjName  3    // String is formed from spid because same msopidWzName is found

/* Flag for EndHTMLImport */
#define msogrfehiNil            0               // No action for spid stuff
#define msogrfehiAssignSpid     (1 << 0)        // Re-assign spid for shapes with temp spid
#define msogrfehiCopySpid       (1 << 1)        // Copy spid to cloned shape
#define msogrfehiCloned         (1 << 2)        // The DG is a cloned one

/* MSOFDGQ - DG queries for GrfdgqQuery */
#define msofdgqAll                      0xFFFFFFFF
#define msofdgqHasInkShape              (1<< 0) // contains at least 1 ink shape that's not an annotation
#define msofdgqHasInkAnnotation         (1<< 1) // contains at least 1 ink annotation shape
#define msofdgqHasInk                   (msofdgqHasInkShape | msofdgqHasInkAnnotation) // contains at least 1 shape with ink data

/* MSOSPI -- Shape Position Info */
/* This should replace the XFRM structure */
typedef struct _MSOSPI
        {
        MSORECT rch; // Flattened coordinates of the anchor rect
        MSOANGLE angle; // Angle of rotation (16.16, 0 <= x < 360)
        struct
                {
                BOOL fFlipH : 1; // Flipped horizontally (x := -x)
                BOOL fFlipV : 1; // Flipped vertically (y := -y)
                ULONG/*MSOAXIS*/ axis : 2;  // Axis corresponding to angle
                ULONG grfUnused : 4;
                MSOANGLE angle90 : 24; // Angle of rotation modulo (-45 <= x <= 45)
                };
        } MSOSPI;
#define cbMSOSPI (sizeof(MSOSPI))

/* MSOSPICE -- Shape Position Info Collection Element */
typedef struct _MSOSPICE
        {
        MSOHSP hsp;       // hsp of the shape for which spi contains data
        int ispiceParent; // index of parent of this shape; index within containing MSOSPIC, or ispiceNil if root shape
        MSOSPI spi;
        } MSOSPICE;

#define ispiceNil -1

/* MSOSPIC -- Shape Position Info Collection */
/* Dynamic array of shape position info plus group hierarchy information (ispiceParent)
   returned by FGetSpicOfSp */
typedef struct _MSOSPIC
   {
   int cspice;
   MSOSPICE rgspice[1];
   } MSOSPIC;

/* MSOPSP -- Proto ShaPe used for shape XML */
typedef struct MSOPSP
        {
        MSOSPT spt;    //Shape Type of the shape
        MSOSPID spid;  //shape id of the master shape, is set to msoidNil for
                                                //shapes which are not master shapes
        union
                {
                ULONG grf;
                struct
                        {
                        ULONG  fMaster : 1;     // set to TRUE only for a master shape
                        ULONG  fUnused : 31;    // Unused
                        };
                };
        } MSOPSP;
#define cbMSOPSP (sizeof(MSOPSP))

/* MSOXMLW - XML Write */
typedef struct MSOXMLW
        {
        union
                {
                struct
                        {
                        ULONG fWriteShapeID : 1;
                        ULONG fWriteStyleAttribute : 1;
                        ULONG fWriteAllBgFill : 1;     // Write out all Background Fill Type in XML
                        ULONG fDoAttachedText : 1;
                        ULONG fRelativePositions : 1;
                        ULONG fOmitXmlTag : 1;         // suppress the writing of the <xml> ... </xml> per shape
                        ULONG fRelyOnVML : 1;
                        ULONG fInlineShapes : 1;       // host supports inline shapes
                        ULONG fOmitZIndex : 1;         // suppress the writing of the shape's z-index style attribute
                        ULONG fAskForMasterShape : 1;  // ask host if the shape is a master shape
                        ULONG fWriteWrapBlock: 1;      // ask host if the shape needs to output wrapblock element
                        ULONG fHasParaAlign: 1;        // Word's non-left-aligned paragraphs
                        ULONG fWriteTopLeft : 1;       // Write top/left instead of margin-top/margin-right
                        ULONG fHasClientAttr : 1;      // host wants DGS::FWriteHspClientAttrXML to be called
                        ULONG fOmitIHlinkAttrs : 1;    // do not write msopidPihlShape and msopidWzTooltip properties
                        ULONG fSupportsPosProps : 1;   // set by hosts that support msopidPosrelh/msopidPosrelv props
                        ULONG fSupportsBorderProps : 1;// set by hosts that support msopidBorder"X"Color props
                        ULONG fRtl:1;                  // Use for RTL sheet orientation
                        ULONG fWantsDirectionLTR:1;    // Write direction:ltr on shapes
                        ULONG fIndentsText:1;          // Write direction:ltr on shapes
                        ULONG fDoFpRelativePos:1;          // Force relative positionning, Used by FP.
                        ULONG grfUnused: 11;
                        };
                ULONG grf;
                };

        interface IMsoHTMLExport *pihe;
        interface IMsoDrawing *pidgOuter;
        MSOPX *ppxMasterPsp; //list of PROTO shapes of master drawing
        MSOPX *ppxpsp; //list of PROTO shapes of drawing
        MSOPX *ppxHspZIndex; // z-indices of shapes in drawing
        void *pvClient; //client data

        // context info passed to FCreateSupportingFileStream calls
        void *pvCurrentExportContext;

        #if DEBUG
        union
                {
                struct
                        {
                        ULONG grfDebugUnused : 32;
                        };
                ULONG grfDebug;
                };
        LONG lVerifyInitXmlw; // Used to make sure people use InitXMLWrite
        #endif
        } MSOXMLW;
#define cbMSOXMLW (sizeof(MSOXMLW))

#if NO96_165378
/* MsoFIsWchValidForShapeName returns TRUE if a character may be in
        a shape name. */
MSOAPI_(BOOL) MsoFIsWchValidForShapeName(WCHAR wch);
#endif

/* IMsoDrawing (idg).  This is the actual interface spoken by DGs.
        TODO peteren: This should really get a longer comment too. */
#if !MKCSM
#undef  INTERFACE
#define INTERFACE IMsoDrawing
DECLARE_INTERFACE_(IMsoDrawing, ISimpleUnknown)
        {
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawing methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Call Free to make an entire Drawing go away.  This includes all
                the DGVGs, DGVs, and DGSLs that might viewing or selecting in
                this Drawing. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* PdgsiBeginUpdate returns a pointer to an MSODGSI.  This will
                always succeed (never return NULL).  The caller may examine and
                (unless they passed in fReadOnly == TRUE) change the fields in
                the MSODGSI.  The caller must in any case call EndDgsiUpdate
                when they're finished.  See comments by the declaration of
                MSODGSI. */
        MSOMETHOD_(MSODGSI *, PdgsiBeginUpdate) (THIS_ BOOL fReadOnly) PURE;

        /* After you call PdgsiBeginUpdate you have to call EndDgsiUpdate,
                passing back in the pointer you got from PdgsiBeginUpdate and
                TRUE for fChanged if you changed any fields in it.  See comment
                by PdgsiBeginUpdate. */
        MSOMETHOD_(void, EndDgsiUpdate) (THIS_ struct _MSODGSI *pdgsi,
                BOOL fChanged) PURE;


        // ----- IMsoDrawing methods for thread synchronization

        /* To access an IMsoDrawing in a thread-safe manner you should
                first allocate an MSODGAR on your stack and call InitAccessRequest
                on it.  Then you set any fields you want to in it and
                pass it to FMakeAccessRequest.  If this returns TRUE you have
                the access you requested.  If it returns FALSE you don't.
                If it returns TRUE you must call ReleaseAccessRequest when you're
                finished.  If it returns FALSE you can just abandon the MSODGAR
                or you can pass it into FMakeAccessRequest again. */
        MSOMETHOD_(void, InitAccessRequest) (THIS_ MSODGAR *pdgar) PURE;
        MSOMETHOD_(BOOL, FMakeAccessRequest) (THIS_ MSODGAR *pdgar) PURE;
        MSOMETHOD_(void, ReleaseAccessRequest) (THIS_ MSODGAR *pdgar) PURE;


        // ----- IMsoDrawing methods (general methods for processing Shapes)

        /* FQueryShape is a general-purpose method for asking a shape
                yes/no questions.  Kinda the Mr. Eight Ball of the Office
                Drawing world.  You pass in the MSOHSP of the shape you
                want to query and an MSOSPQ specifying the question. */
        MSOMETHOD_(BOOL, FQueryShape) (THIS_ MSOHSP hsp, MSOSPQ spq) PURE;

        /* Call FBeginMarkShapes to let the DG know you're going to begin
                an operation on it that involves marking shapes.  This can fail
                because only one person at a time can mark shapes. */
        MSOMETHOD_(BOOL, FBeginMarkShapes) (THIS) PURE;

        /* Call MarkShape to, well, mark the shape specified by hsp.  This just
                means setting a bit in the shape. You must have called
                FBeginMarkShapes before you call this.
                TODO peteren: If needed add the "fMarkAncestors" option. */
        MSOMETHOD_(void, MarkShape) (THIS_ MSOHSP hsp /*, BOOL fMarkAncestors*/) PURE;

        /* Call EndMarkShapes to finish an operation involving marking shapes.
                You should pass in TRUE for fClearMarks if you haven't already
                cleared all the shapes you marked via some other means.  If you have
                your own means you should use it, as clearing these bits means
                a pass through the DG. */
        MSOMETHOD_(void, EndMarkShapes) (THIS_ BOOL fClearMarks) PURE;

        /* FIsShapeMarkedSelected returns TRUE iff the specified Shape is
                marked as selected. See IMsoDrawingSelection::FMarkSelectedShapes. */
        /* TODO peteren: Remove this in favor of FQueryShape(msospqMarkedSelected) */
        MSOMETHOD_(BOOL, FIsShapeMarkedSelected) (THIS_ MSOHSP hsp) PURE;


        // ----- IMsoDrawing methods (enumerating Shapes)

        /* Use BeginEnumerateShapes to initiate an enumeration of the
                Shapes in a Drawing (or the Shapes in a Group Shape).  You
                pass us a pointer to an MSODGEN (presumably on your stack) in which
                the grfdgen (and optionally hspRoot) fields are filled out. */
        MSOMETHOD_(void, BeginEnumerateShapes) (THIS_ MSODGEN *pdgen) PURE;

        /* Call FEnumerateShapes to continue an enumeration begun by a call
                to BeginEnumerateShapes.  We return TRUE if there were any more
                Shapes over which to enumerate and FALSE if you're done.  If we
                return TRUE you can examine pdgen->hsp and pdgen->grfdgeno to learn
                the state of the enumeration. */
        MSOMETHOD_(BOOL, FEnumerateShapes) (THIS_ MSODGEN *pdgen) PURE;

        /* This reasonably special-purpose method may be called while
                enumerating through the Shapes in a Drawing or Group in the case where
                you've just arrived at a Group for the first time.  In other words,
                pdgen->grfdgeno & msofdgenoAtGroupOnWayIn will be TRUE.  Ordinarily
                when next you called FEnumerateShapes we'd dive into the Group
                and return you it's first child.  But if for some reason you don't
                like this Group and don't want to see ANY of its children you
                can call this method.  Thereafter when you call FEnumerateShapes
                we'll move you to the Groups next sibling instead. */
        MSOMETHOD_(void, SkipChildrenWhileEnumerateShapes) (THIS_ MSODGEN *pdgen)
                PURE;

        // ----- IMsoDrawing methods (editing methods)

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FBeginUndoRecord) (THIS) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FEndUndoRecord) (THIS_ BOOL fPost,
                interface IMsoDrawingSelection *pidgsl) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FMakingUndoRecord) (THIS) PURE;

        /* Call FPurge to purge a particular Drawing from a Drawing Group.
                This methods does NOT free DGVs, DGSLs, and other stuff that might
                be referencing this Drawing; it's up to the caller to have gotten
                rid of such stuff.  The meaning of the return value is as yet
                undefined.  This method does not actually free the DG in question,
                as it might contain shapes that are referenced from elsewhere in
                the DGG.  You have to call DGG::FreePurgedStuff or DGG::Free
                to actually force the purged stuff to be deallocated.  Once you've
                purged something there's no way to get it back. */
        MSOMETHOD_(BOOL, FPurge) (THIS_ void *pvContext) PURE;

        /* Call FMove to get move a Drawing into a new Drawing Group.  We
                return TRUE if we succeed.  This is not undoable. */
        MSOMETHOD_(BOOL, FMove) (THIS_ interface IMsoDrawingGroup *pidggDest,
                void *pvContext) PURE;

        /* Call FClone to make a copy of a Drawing either in the same
                drawing group (either by passing in the correct pidggClone or by
                passing NULL for it) or in a different Drawing Group.  You have to
                pass in a pointer to a DGSI to use for the clone.  We pass you
                back a pointer to the new Drawing.  If you pass in TRUE for
                fCloneShapes we clone all the Shapes as well, otherwise the
                cloned Drawing will be initially empty. */
        MSOMETHOD_(BOOL, FClone) (THIS_ interface IMsoDrawingGroup *pidggClone,
                MSODGSI *pdgsiClone, IMsoDrawing **ppidgClone, BOOL fCloneShapes,
                void *pvContext) PURE;

        /* Call FDeleteShape to remove a shape from a DG.  If you pass in
                fUndo == TRUE we will post undo records to the DG's DGS that allow
                this Shape to come back.  Otherwise it will be permanently gone.
                We return TRUE on success (if !fUndo we always succeed). */
        MSOMETHOD_(BOOL, FDeleteShape) (THIS_ MSOHSP hsp, BOOL fUndo,
                void *pvContext) PURE;

        /* Call FCloneShape to make a duplicate of a particular
                Shape, either in the same DG (pass in the correct pidgClone
                or NULL) or a different DG.  The new Shape will appear at the
                top of the Z order in the new DG.  We return the hsp for the
                new Shape in *phshpClone.  We we return TRUE on success.
                If fUndo is TRUE we'll post undo records into pidgClone's DGS
                that will delete the cloned shape if executed. */
        MSOMETHOD_(BOOL, FCloneShape) (THIS_ MSOHSP hspOriginal,
                interface IMsoDrawing *pidgClone, MSOHSP *phspClone, BOOL fUndo,
                void *pvContext) PURE;

        /* Same as FCloneShape except setting flag fDiffApps=true means that
                        the cloning is from one App to another*/

        MSOMETHOD_(BOOL, FCloneShapeEx) (THIS_ MSOHSP hspOriginal,
                interface IMsoDrawing *pidgClone, MSOHSP *phspClone, BOOL fUndo,
                void *pvContext, BOOL fDiffApps) PURE;

        /* Call FDeleteMarkedShapes to remove all of the marked Shapes in the
                DG.  If (fUndo) we'll post undo records that will bring them back if
                executed.  If (!fUndo) they'll be gone permanently.  Note that they
                may not actually be freed from memory until you call
                DGG:FFreePurgedStuff.  If you pass in fClearMarks we'll clear
                the marks in the original Shapes as we do this.  If we fail, then
                we left the shapes half-deleted.  Sorry. */
        MSOMETHOD_(BOOL, FDeleteMarkedShapes) (THIS_ BOOL fUndo,
                BOOL fClearMarks, void *pvContext) PURE;

        /* Call FCloneMarkedShapes to make copies in a second DG (at the top
                of that DG's Z-order) of the marked shapes.  Pass in TRUE for
                fClearMarks if you want us to clear the Shapes' "marked" bits as
                we do this.  pidgClone can be NULL if you want to clone the
                Shapes within the same DG.  If (fUndo) we'll post undo records
                to pidgClone's DGS that will delete the cloned shapes if executed. */
        MSOMETHOD_(BOOL, FCloneMarkedShapes) (THIS_ interface IMsoDrawing *pidgClone,
                BOOL fUndo, BOOL fClearMarks, void *pvContext) PURE;

        /* copy properties NOT saved in binary from hspFrom to hspTo.  You should
                never call this method--it exists only so that Word's
                FCloneExactDrawingInfo can clone the shape properties that are
                not saved in binary */
        MSOMETHOD_(BOOL, FCopyUnsavedProps)(THIS_ MSOHSP hspFrom, MSOHSP hspTo, BOOL fCopyResaveID) PURE;

        /* Call FApplyPropertiesToShape to set some properties on a Shape.
                If you call this while the Drawing is building an Undo Record
                we'll add data to that undo record, otherwise we'll apply the
                properties non-undoably.

                hsp
                rgspp
                cspp
                grfaspp

                TODO peteren: Finish comment. */
        MSOMETHOD_(BOOL, FApplyPropertiesToShape) (THIS_ MSOHSP hsp,
                MSOSPP *rgspp, int cspp, ULONG grfaspp) PURE;

		/*
		FApplyPropsToShapes will apply the given properties to every shape
		in the shape iter.  The shape iter may be a set of shapes from a
		selection or some other set.

		This is better than FApplyPropertiesToShape because it doesn't guess
		at which shapes should have what props set while enumerating through
		groups and children.  Here we assume that every shape in the iter wants
		every prop set.  If you want to differentiate groups and children, then
		create two iterators and call this twice with different props.

		We don't own the pIter, but the rgspp will be consumed.
		*/ 
        MSOMETHOD_(BOOL, FApplyPropsToShapes) (THIS_
			interface IMsoShapeIter *pIter, MSOSPP *rgspp, int cspp) PURE;

        /* Call FFetchPropertiesFromShape to get the values of some properties
                from a Shape.
                TODO peteren: Better comment. */
        MSOMETHOD_(BOOL, FFetchPropertiesFromShape) (THIS_ MSOHSP hsp,
                MSOSPP *rgspp, int cspp, ULONG grffspp) PURE;


        // ----- IMsoDrawing methods (MSODGIDs, MSOSPIDs, etc.)

        /* DgidGetDrawingId returns this Drawing's ID (unique within its Drawing
                Group). */
        MSOMETHOD_(MSODGID, DgidGetDrawingId) (THIS) PURE;

        /* SpidGetShapeId returns the DGG-unique identifier (MSOSPID)
                of a particular Shape.  This is the new (better type, better
                name) version of LIdFromHsp. */
        MSOMETHOD_(MSOSPID, SpidGetShapeId) (THIS_ MSOHSP hsp) PURE;

        /* FFindShape looks for the Shape in this DG with the specified MSOSPID,
                returning TRUE and filling out *phsp if it finds one and FALSE
                otherwise.  This is O(1) if we have a hash-table constructed
                (at load time) and O(n) thereafter.  We may try to improve this.
                This replaces HspFromLId. */
        MSOMETHOD_(BOOL, FFindShape) (THIS_ MSOHSP *phsp, MSOSPID spid) PURE;

        /* SpidfFromSpid converts a real MSOSPID (which is a unique
                identifier for a Shape relative to Drawing Group) into a more
                user-friendly (they actually go 1, 2, 3, etc.) identifier which is
                unique only within this Drawing.  We use these, for example,
                in VBA. */
        MSOMETHOD_(MSOSPIDF, SpidfFromSpid) (THIS_ MSOSPID spid) PURE;

        /* SpidFromSpidf converts a friendly MSOSPIDF (unique within the
                Drawing) into a real MSOSPID (unique within the Drawing Group). */
        MSOMETHOD_(MSOSPID, SpidFromSpidf) (THIS_ MSOSPIDF spidf) PURE;


        // ----- IMsoDrawing methods (other)

        /* TODO peteren: Rename this method PidggGetDrawingGroup. */
        MSOMETHOD_(interface IMsoDrawingGroup *, PidggGet) (THIS) PURE;

        /*
        Is the DG dirty?
        */
        MSOMETHOD_(BOOL, FNeedsSave) (THIS) PURE;

        /*
        A way for the host to explicitly set the dirtiness of the DG.
        */
        MSOMETHOD_(void, SetFNeedsSave) (THIS_ BOOL) PURE;

        /*
        Delete all SPS's in the purgable SPS list in the DG.
        */
        MSOMETHOD_(int, NCollectGarbage) (THIS) PURE;

        /* FExecuteUndoRecord runs an undo record we passed previously to the
                DGS.  Set fRedo to TRUE if you're actually redoing this HUR.  If
                you pass fRedo == FALSE (undo) and pidgsl == NULL we promise to
                succeed (so that you can use this for error recovery). */
        MSOMETHOD_(BOOL, FExecuteUndoRecord) (THIS_ MSOHUR hur, BOOL fRedo,
                struct IMsoDrawingSelection *pidgsl) PURE;

        /*
        The given undo record we gave to the host has either fallen off the
        end of the stack or is dropped due to new actions done.  Set fPurgeShapes
        to TRUE if the undo record couldn't ever come back and you want any
        shapes it references to be purged.
        */
        MSOMETHOD_(void, FreeUndoRecord) (THIS_ MSOHUR hur, BOOL fPurgeShapes) PURE;

        /*
        Called by the host to save the given undo record to a file.
        */
        MSOMETHOD_(BOOL, FSaveUndoRecord) (THIS_ MSOSVB *psvb, MSOHUR hur) PURE;

        /*
        Called by the host to allocate and load an undo record from a file.
        */
        MSOMETHOD_(BOOL, FLoadUndoRecord) (THIS_ MSOLDB *pldb, MSOHUR *phur) PURE;

        /*
        FCloneUndoRecord copies the undo record for subsequent reconciliation.
        */
        MSOMETHOD_(BOOL, FCloneUndoRecord) (THIS_ MSOHUR hur,
                MSOHUR *phurClone) PURE;

        /*
        Called by a shape
        */
        MSOMETHOD_(BOOL, FRegisterShapeUndoRecord) (THIS_ MSOHSUR, MSOHSP hsp) PURE;

        /*
        Called by a rule to register a rule undo record
        */
        MSOMETHOD_(BOOL, FRegisterRuleUndoRecord)
                (THIS_ MSOHRUR, interface IMsoRule* piru) PURE;

        MSOMETHOD_(void, GetHspOfUndoRecord) (MSOHUR hur) PURE;

        /*
        Create a primitive non-inheriting shape of type spt, at coordinates rch
        (or using pvAnchor directly instead of making an anchor if prch is NULL)
        in the drawing. Returns TRUE on successful insert after which phsp points
        to the new HSP. Undo records will not be generated if fUndoable is FALSE.
        This function does not ever deal with any selection. Note that ppxhspGroup
        is only for the host's convenience, and the host is responsible to free it.
        If we fail, we will NOT free pvAnchor.
        */
        MSOMETHOD_(BOOL, FCreateHsp) (THIS_ MSOHSP *phsp, MSOSPT spt, RECT *prch,
                void *pvAnchor, MSOHSP hspPrepared, MSOHSP hspPrev, void **ppxhspGroup,
                int grfchp) PURE;

        /*
        Create a primitive non-inheriting shape of type spt, at coordinates rch
        (or using pvAnchor directly instead of making an anchor if prch is NULL)
        in at the top of the drawing. Returns TRUE on successful. phsp points
        to the new HSP. Properties are a fixed the default.
        */
        MSOMETHOD_(BOOL, FCreateBuiltInShape) (THIS_ MSOHSP *phsp, // [out] result
                RECT *prch,
           void *pvAnchor,
                MSOSPT spt,
                BOOL fApplyDefaults, BOOL fUndo) PURE;

        /*
        Calls DG::FBeginPolyLineShape to create a polyline shape (msoINonPrmitive)
        at coordinates rch (or using pvAnchor directly instead of making an anchor
        if prch is NULL) in at the top of the drawing. Returns TRUE on success.
        Phsp points to the new HSP. Properties are a fixed default.
        */
        MSOMETHOD_(BOOL, FCreatePolyLineShape)(THIS_ MSOHSP *phsp,
                MSOSPT spt,
                RECT *prch,
                void *pvAnchor,
                BOOL fApplyDefaults,
                const MSOCPLB *pcpl, BOOL fUndo) PURE;

        /*
        Creates a shape (msosptNotPrimitive) at coordinates rch (or using pvAnchor directly
        instead of making an anchor if prch is NULL) in at the top of the drawing. Returns
        TRUE on success. Phsp points to the new HSP. Properties are a fixed default.  If
        the function successfully creates a shape, it will set the ink data property.
        */
        MSOMETHOD_(BOOL, FCreateInkShape)(THIS_ MSOHSP *phsp,
                RECT *prch, void *pvAnchor, BOOL fApplyDefaults, const MSOCPLB *pcpl,
                interface IInkDisp *piInkDispPlain, interface IInkStrokes* piInkStrokes,
                MSOGV gvPage, MSOHSP hspModel, BOOL fUndo, BOOL fInkAnnotation) PURE;

        /*
        Create a blip (ie. a rectangle shape with a picture) shape. Ownership
        of the blip is transfered if successful.

        Returns TRUE on successful insert after which phsp points
        to the new HSP. Properties are a fixed default.  The blip may be passed
        in or specified as a file name or URL - the MSOBLIPFLAGS say which and
        also control whether the blip is saved with the document (note: it may
        be saved as a result of being in some other shape).  If the fUpdate flag
        is passed the blip is a "placeholder" and the name (which must be passed
        and be a URL name) will be resolved and a new blip created.  If the
        fSynchronous flag is true a new thread will not be spawned to do any
        slow resolution or loading.

        wzName:  file name: Win32: a UNICODE path name (may be in UNC form)
                                URL name:  Win32: UNICODE URL
                                file name: MacOS: a packed WLM format FSSpec
                                URL name:  MacOS: UNICODE URL
        wzDescription: UNICODE text describing the picture, equivalent to the
                        ALT="foo" text in a HTML IMG tag (should normally be provided by
                        whoever provided the picture!)
        */
        MSOMETHOD_(BOOL, FCreateBlipShape)(THIS_ MSOHSP *phsp,
                RECT *prch,
                void *pvAnchor,
                BOOL fApplyDefaults,
                IMsoBlip* pib,
                const WCHAR *wzName,
                const WCHAR *wzDescription,
                MSOBLIPFLAGS blipFlags,
                BOOL fUpdate,
                BOOL fSynchronous, BOOL fUndo) PURE;

        /*
        Update all the blips in the document which can be updated - if the blip
        flags indicate that the name is that of a file or URL the blip is
        updated by resolving the URL and reloading the file (if the blip occurs
        multiple times in the document it will only be updated once).  The flag
        indicates that the method should be synchronous - normally the update
        should be done in the background.  FALSE is returned on some internally
        detected error - if a blip cannot be resolved TRUE is still returned.
        If the MSOHSP parameter is non-nil then only that shape is updated.
        */
        MSOMETHOD_(BOOL, FUpdateBlips) (THIS_
                MSOHSP hsp,
                BOOL fSynchronous) PURE;

        /*
        Special update functions for the envelope and stationery cases.
        Both synchronously load every blip; the envelope one also sets the
        No Send flag, while the stationery one clears the link-to-file flag.
        */
        MSOMETHOD_(BOOL, FUpdateBlipsForEnvelope) (THIS) PURE;
        MSOMETHOD_(BOOL, FUpdateBlipsForStationery) (THIS) PURE;


        /*      CShapes returns the number of Shapes in this DG.  If fAll is TRUE
                it returns the total number of Shapes in the DG, including those
                inside groups.  If fFall is FALSE it returns the total number of
                top-level shapes in the DG. */
        /* TODO peteren: I don't like the "fAll" argument.  In particular
                it's unclear whether or not it counts groups themselves.  We should
                have two arguments, fGroups, and fChildren. */
        MSOMETHOD_(int, CShapes) (THIS_ BOOL fAll) PURE;

        /* IMsoDrawing::Invalidate is called when a Drawing has changed in
                some way, to let Views, Selections, etc. know that they have some
                updating to do.  This method will usually be called from other
                Office Drawing code, but there are cases when the host will want
                to do so (changes in pvClientData that should cause redraw?).
                This method is really fast, especially when called several times
                in quick successions.  See comments by the #define's for various
                msofdgi values to see what they actually do. */
        MSOMETHOD_(void, Invalidate) (THIS_ ULONG grfdgi, MSOHSP hsp) PURE;

        /*
        This is a top level shape enumerator usable by the host. Given a generic
        MSOHSP, we return a next one. Passing in a pointer to NULL means start at
        the beginning. A return value of FALSE means the end of the list.
        */
        MSOMETHOD_(BOOL, FGetRelatedShape) (THIS_ MSOHSP *phsp,
                MSODGRST dgrst) PURE;

        /* FGetTheFirstShape returns TRUE and the MSOHSP of the first Shape in
                this Drawing if there is one and FALSE if there isn't.  If you
                pass in a non-NULL ppvClient we'll also return the pvClient of
                the first Shape. */
        MSOMETHOD_(BOOL, FGetFirstShape) (THIS_ MSOHSP *phsp,
                void **ppvClient) PURE;

        /*
        Get the last shape in the drawing, returning FALSE if there aren't any.
        */
        MSOMETHOD_(BOOL, FGetLastShape) (THIS_ MSOHSP *phsp,
                void **ppvClient) PURE;



        /*
        A way for a host (like Word) to force change the pvAnchor of a shape.
        */
        MSOMETHOD_(void, SetPvAnchorOfHsp) (THIS_ void *, MSOHSP) PURE;

        /*
        Returns a pointer to the anchor pointer of the given HSP.
        */
        MSOMETHOD_(void, GetPvAnchorOfHsp) (THIS_ void **, MSOHSP) PURE;

        /*
        A way for the host to set the client data of an HSP to a given value.
        Note that this will set the fNeedsSave flag of the DG as it should.
        */
        MSOMETHOD_(void, SetClientDataOfHsp) (THIS_ void *, MSOHSP) PURE;

        /*
        Returns a pointer to the client data pointer of the given HSP.
        */
        MSOMETHOD_(void, GetClientDataOfHsp) (THIS_ void **, MSOHSP) PURE;

        // Don't use FLocatePrchOfHsp in new code. It should be replaced by FGetSpiOfHsp.
        MSOMETHOD_(BOOL, FLocatePrchOfHsp) (THIS_ RECT *prch, MSOHSP hsp) PURE;

        /*
        Fill out an SPI describing the position of a shape.  Flattens coordinates
        and rotation for shapes in groups.  Works without using marking bits.
        This should REPLACE FLocatePrchOfHsp.
        */
        MSOMETHOD_(BOOL, FGetSpiOfHsp) (THIS_ MSOSPI *pspi, MSOHSP hsp) PURE;

        /*
        Fill out a SPIC describing the position of a group's child shapes.
        If hsp == msohspNil, fill it out for all shapes and their children.
        Flattens coordinates and rotation for child shapes.  Works without using
        marking bits.  Call ReleaseSpic to dispose of the SPIC
        */
        MSOMETHOD_(BOOL, FGetSpicOfHsp) (THIS_ MSOSPIC **ppspic, MSOHSP hsp) PURE;

        /*
        Release a SPIC returned by FGetSpicOfSp
        */
        MSOMETHOD_(void, ReleaseSpic) (THIS_ MSOSPIC *pspic) PURE;

        /*
        Returns the bounding rectangle of the shape's effect which is calculated
        without a view being present (default view parameters). If
        hsp == msohspNil it gives the bounding rectangle of the effects of
        all the shapes in the drawing.
        */
        MSOMETHOD_(BOOL, FGetEffectBounds) (THIS_ MSODC *pdc, POINT *ppthOrigin,
                RECT *prchDraw, RECT *prchEffect, MSOHSP hsp) PURE;

        /*      Set the properties of a shape. WARNING: No undo record is made when calling
        this method and no invalidation is done. Intended to be used in
        conversion, where failure can   be handled cleanly.  You probably want
        to be calling FApplyPropertiesToShape instead! */
        MSOMETHOD_(BOOL, FSetProp)(THIS_ MSOHSP hsp, MSOPID opid, void* pop) PURE;

        /*
        Set the properties of a shape. WARNING: No undo record is made when calling
        this method. Intended to be used in conversion, where failure can
        be handled cleanly.
        */
        MSOMETHOD_(BOOL, FResetProp)(THIS_ MSOHSP hsp, MSOPID opid) PURE;

        /*
        Set the property sets of a shape. WARNING: No undo record is made when calling
        this method. Intended to be used in conversion, where failure can
        be handled cleanly.
        */
        MSOMETHOD_(BOOL, FSetPropSet)(THIS_ MSOHSP hsp, MSOPSID opsid, void* pops) PURE;

        /* Call FetchProp to retrieve a property (identified by opid) of a Shape
                (identified by hsp) into the OP pointed to by pop. */
        MSOMETHOD_(void, FetchProp) (THIS_ MSOHSP hsp, MSOPID opid, void* pop) PURE;

        /* Call FecthPropPreset if you want to retrieve a property that is in the form
        of a preset. preFirst and preLast specify where to start and end. pidIgnore will
        not be considered (if found) in the comparison. This was needed by Publisher to retrieve
        the line dash preset value for a shape
        */
        MSOMETHOD_(void, FetchPropPreset) (THIS_ MSOHSP hsp, MSOPRESET preFirst, MSOPRESET preLast,
        MSOPID pidIgnore, void* pop) PURE;
        /*
        Return the property set identified in the shape indentified.
        */
        MSOMETHOD_(void, FetchPropSet)(THIS_ MSOHSP hsp, MSOPSID opsid, void* pops) PURE;

        /*
        The following two methods are used to examine the properties of a shape.
        We will try to avoid giving you useless information for properties that are
        set to the default so as to minimize the number of non-trivial properties.
        */
        MSOMETHOD_(void, InitEnumProps)(THIS_ MSOHSP hsp, MSOPENUM *ppenum) PURE;
        MSOMETHOD_(BOOL, FEnumProps)(THIS_ MSOHSP hsp, MSOPENUM *ppenum) PURE;
        MSOMETHOD_(BOOL, FPropInfoFromSz)(THIS_ char *sz, MSOPENUM *ppenum) PURE;

        MSOMETHOD_(BOOL, FGetTxid) (THIS_ MSOHSP hsp, MSOTXID *ptxid) PURE;

        /* FSetTxid sets the text id of a shape. The fNotUndoable flag is very
                scary and should always be set to FALSE unless you really know what you
                are doing. */
        MSOMETHOD_(BOOL, FSetTxid) (THIS_ MSOHSP hsp, MSOTXID txid,
                BOOL fNotUndoable) PURE;
        MSOMETHOD_(BOOL, FGetNextTextHsp) (THIS_ MSOHSP hsp, MSOHSP *phspNext) PURE;
        MSOMETHOD_(BOOL, FSetNextTextHsp) (THIS_ MSOHSP hsp, MSOHSP hspNext,
                BOOL fNotUndoable) PURE;

        /* Get the text frame for the shape.
                The text frame takes into account the design of the shape,
                the text margin and the current adjustments. The text frame
                doesn't change size when the shape is rotated. */
        MSOMETHOD_(BOOL, FGetTextFrameRch) (THIS_ MSOHSP hsp, RECT *prchText) PURE;

        /* Set the anchor of the shape to fit passed in text frame.
                The text frame takes into account the design of the shape,
                the text margin and the current adjustments. */
        MSOMETHOD_(BOOL, FSetTextFrameRch) (THIS_ MSOHSP hsp, RECT *prchText) PURE;

        /* Same as FSetTextFrameRch but doesn't expand parent bounds. */
        MSOMETHOD_(BOOL, FSetTextFrameRch2) (THIS_ MSOHSP hsp, RECT *prchText) PURE;

        /* FCreateViewGroup creates a new IMsoDrawingViewGroup object
                to view this drawing and returns a pointer to it.  Since this interface
                inherits from ISimpleUnknown instead of IUnknown there's no ref-counting
                to worry about. */
        MSOMETHOD_(BOOL, FCreateViewGroup)
                (THIS_ interface IMsoDrawingViewGroup **ppidgvg) PURE;

        /* FCreateSelection creates a new IMsoDrawingSelection object to
                select Shapes within this drawing.  The new DGSL will have been
                ref-counted for you.  You may if you want pass in pointer to
                an MSODGSLSI for us to copy into the new DGSL. */
        MSOMETHOD_(BOOL, FCreateSelection) (THIS_
                interface IMsoDrawingSelection **ppidgsl, struct _MSODGSLSI *pdgslsi)
                PURE;

        /*
        Create and/or return an AddRef'ed IDispatch object for the Shapes
        collection automation object corresponding to this drawing.
        */
        MSOMETHOD_(BOOL, FGetDispShapes) (THIS_ IMsoDrawingUserInterface *pidgui, IDispatch **) PURE;

        /*
        Create and/or return an AddRef'ed IDispatch object for the Shape
        site at index iShape in this drawing. Return NULL if iShape is invalid.
        */
        MSOMETHOD_(IDispatch *, PidispFromHsp) (THIS_ MSOHSP, IMsoDrawingUserInterface *pidgui) PURE;

        /*
        Given an IDispatch pointer to a shape site, return the hsp
        corresponding to it.
        */
        MSOMETHOD_(MSOHSP, HspFromPidisp) (THIS_ IDispatch *) PURE;

        /* Call FGetShapeName to get the unique friendly name (e.g. "Rectangle 1",
                "Oval 2") associated to 'hsp'. The user-defined name stored in
                msopidWzName is returned if it's defined. Otherwise,
                the name is autogenerated based on the shape SPC and its internal
                SPID value. if succeeded, the name is returned in 'wz'. 'grfdgsn'
                is for future options. For now, pass in msogrfdgsnDefault.

                Pass in wz a pointer to a buffer of WCHARs and in cchMax the length
                of the buffer.  If you pass in a non-NULL pcch we'll return in *pcch
                the number of characters in the name, not counting the terminating 0.
        */
        MSOMETHOD_(BOOL, FGetShapeName) (THIS_ MSOHSP hsp, ULONG grfdgsn,
                WCHAR *wz, int cchMax, int *pcch) PURE;

        /* Call FPutShapeName to set the 'hsp' user-defined name to 'wz'.
        Pass msogrfdgsnCheckCollision in 'grfdgsn' to enforce name uniqueness
        (slow). For XL loading old files, use msogrfdgsnDefault that
        will skip the collision check and is thus very fast
        (only valid during loading time).
        */
        MSOMETHOD_(BOOL, FPutShapeName) (THIS_ MSOHSP hsp, ULONG grfdgsn,
                const WCHAR *wz) PURE;


        /* Find the shape 'sp' that matches the input name 'wzName'.
        Pass msogrfdgsnUserDefined in 'grfdgsn' to lookup only the User-defined names.
        msogrfdgsnDefaults means to lookup user-defined names first and then match the
        auto-generated ones.
        */
        MSOMETHOD_(BOOL, FFindShapeFromName) (THIS_ const WCHAR *wz, ULONG grfdgsn,
                MSOHSP* phsp) PURE;

        MSOMETHOD_(BOOL, FGetAltText) (THIS_ MSOHSP hsp, ULONG grfdgsn,
                WCHAR *wz, int cchMax, int *pcch) PURE;

        MSOMETHOD_(BOOL, FPutAltText) (THIS_ MSOHSP hsp, ULONG grfdgsn,
                const WCHAR *wz) PURE;

        MSOMETHOD_(BOOL, FCropCanvas) (THIS_ MSOHSP hsp, const RECT *prchNew) PURE;

        MSOMETHOD_(void, GetBaseFlipsForHsp) (THIS_ MSOHSP hsp, BOOL *pfFlipH, BOOL *pfFlipV) PURE;

        MSOMETHOD_(BOOL, FDiagramAutoFit) (THIS_ MSOHSP hsp) PURE;
        // Begin OLE API section:

        /*
        The given OLE shape's ojbect has been activated or deactivated by the
        host, so its picture should have hatch marks drawn over it appropriately.
        */
        MSOMETHOD_(void, SetFActiveHsp) (THIS_ MSOHSP hsp, BOOL fActive,
                BOOL fNotUndoable) PURE;

        /*
        Create and insert a new OLE shape, whose host id will be the given OID.
        Any host initiated object creation in the drawing layer should call this.
        */
        MSOMETHOD_(MSOHSP, HspCreateShapeFromOid) (THIS_ MSOOID oid, RECT *prch,
                void *pvAnchor, interface IMsoDrawingSelection *pidgsl, BOOL fOCX) PURE;

        /*
        Change the presentation picture of the given OLE shape to be the passed
        value. The host should call this on getting an update from the server.
        */
        MSOMETHOD_(void, SetPictOfHsp) (THIS_ MSOHSP hsp, HANDLE hPict, LONG cf,
                BOOL fNotUndoable) PURE;

        /*
        Change the host's persistent id of the given OLE shape to be the passed
        OID. The host should call this on doing a Convert to a new class.
        */
        MSOMETHOD_(BOOL, FSetNewOidOfHsp) (THIS_ MSOHSP hsp, MSOOID oid,
                BOOL fNotUndoable) PURE;

        /*
        Given an HSP, return the host id OID corresponding to it, or msooidNil if
        it isn't an OLE shape. This will always run on O(1) time.
        */
        MSOMETHOD_(MSOOID, OidFromHsp) (THIS_ MSOHSP hsp) PURE;

        /*
        Given a host id OID, return the hsp corresponding to it. This loops over
        all the DG's shapes and runs on O(n) time.
        */
        MSOMETHOD_(MSOHSP, HspFromOid) (THIS_ MSOOID oid) PURE;

        /* Image Optimization helper functions
        */
        MSOMETHOD_(MSOHSP, FConvertOleObjectFromHspToBlip) (THIS_ MSOHSP hsp) PURE;

        MSOMETHOD_(BOOL, FIsOleObjectOptimizable) (THIS_ MSOHSP hsp) PURE;

        MSOMETHOD_(BOOL, FIsOleObjectConvertibleToBitmap) (THIS_ MSOHSP hsp) PURE;

        MSOMETHOD_(BOOL, FIsOleObjectAOrgChart) (THIS_ MSOHSP hsp) PURE;

        MSOMETHOD_(BOOL, FIsOleObjectMSDraw) (THIS_ MSOHSP hsp) PURE;

        //------ Connector Section

        /*
        Add a connector to the front of the drawing at the specified location.
        The connector is not attached to any shape. See also FCreateBuiltInShape.
        */
        MSOMETHOD_(BOOL, FCreateConnector) (THIS_ MSOHSP *phsp, RECT *prch,
                void *pvAnchor, MSOCXSTYLE cxstyle) PURE;

        /*
        Is the hsp a connector shape?
        */
        MSOMETHOD_(BOOL, FIsConnector) (THIS_ MSOHSP hsp) PURE;

        /*
        Attach the start/end of the connector to the specified shape at
        the specified site. If hsp is NULL, then deattach the end. An Undo
        record is recorded.
        */
        MSOMETHOD_(BOOL, FAttach) (THIS_ MSOHSP hspCon, MSOCXSP cxsp, MSOHSP hsp, int site) PURE;

        /*
        Is the connector attached? If so return the shape (phsp) and site (psite).
        MSOCXSP specifies whether you are asking about the start or the end of the
        connector.
        */
        MSOMETHOD_(BOOL, FIsAttached) (THIS_ MSOHSP hspCon, MSOCXSP cxsp, MSOHSP *phsp, int *psite) PURE;



        //------ Rule Section

        /*
        Add a rule to the list of rules managed by the drawing. The rule is
        is tested for consistency with other rules. If necessary, the drawing
        is changed to comply with the new rule. As part of the adding
        process is a rule is queried for all the shapes it governs. If 'fOwn'
        is set, the drawing owns the rule and is responsible for calling
        Free on it.
        */
        MSOMETHOD_(BOOL, FAddRule) (THIS_ interface IMsoRule* piru) PURE;

        /*
        Remove a rule from the list of rules managed by the drawing.
        */
        MSOMETHOD_(void, RemoveRule) (THIS_ interface IMsoRule* piruRemove) PURE;

        /*
        Called to notify that a particular rule has changed, either by
        adding or deleting HSPs that it governs or by changing it paramters.    Must be called
        after a HSP was added to the rule, or before an HSP is deleted from a rule
        After the calling this the rule is marked for testing by the drawings solver.
        */
        MSOMETHOD_(BOOL, FOnRuleChange)
                (THIS_ interface IMsoRule* piruChanged, MSORUC ruc, int iProxy) PURE;

        /*
        The number of rules
        */
        MSOMETHOD_(int, CRules) (THIS) PURE;

        /*
        Get the i'th (0-based) rule.
        */
        MSOMETHOD_(void, GetRule) (THIS_ int i, interface IMsoRule** piru) PURE;

        /*
        A way for the host to get a RULE from a specific RUID.
        */
        MSOMETHOD_(interface IMsoRule*, PiruFromRuid) (THIS_ MSORUID ruid) PURE;

        /*
        Rules should use these interfaces to fetch and set properties
        */

        /*
        Do we have change pending in particular property. A range of
        properties can be specified.
         */
        MSOMETHOD_(BOOL, FHasPropChanges) (THIS_ MSOHSP hsp,
                MSOPID opidStart, MSOPID opidEnd) PURE;

        /*
        Remove the property from the pending changes table for this shape
        */
        MSOMETHOD_(void, RemovePropChange) (THIS_ MSOHSP hsp, MSOPID opid) PURE;

        /*
        Remove the property set from the pending changes table for this shape
        */
        MSOMETHOD_(void, RemovePropSetChange) (THIS_ MSOHSP hsp, MSOPSID opsid) PURE;

        /*
        Fetch a prop taking into account the pending changes
        */
        MSOMETHOD_(void, FetchPropWithChanges) (THIS_ MSOHSP hsp,
                MSOPID opid, void* pop) PURE;

        /*
        Fetch a prop set taking into account the pending changes
        */
        MSOMETHOD_(void, FetchPropSetWithChanges) (THIS_ MSOHSP hsp,
                MSOPSID opsid, void* rgop) PURE;


        // ----- Conversion Section

        /* We allow the client to convert a shape into either a METAFILE or
           a polygon. The client needs to choose the best representation depending
                on its policy.  */

        /* Convert the shape to the single polygon that is most representive of the shape.
        All shapes have a polygon representation. At best it is exactly the shape, at
        worst it could be the bounding box. The current adjustment, rotation, position
        of the shape is taken into account.

                hsp            -- Shape [in]
                dg             -- data group to allocate the array [in]
                fHUnits        -- if TRUE, points are output in H units,
                                                                  otherwise in G units. [in]
                prgpt          -- Array of points in G (default) or H units [out]
                pc             -- Count [out]
                pfPolyline     -- Return is an open polyline [out]

        A polygon is allocated in 'dg'. See msospqCanConvertToPolygon to check
        if shape can be converted well to a polygon.

        IMPLEMENTAION: The exact conversion routine is still be refined. Typically,
        the first path is flattened. */
        MSOMETHOD_(BOOL, FConvertShapeToPolygon) (THIS_ MSOHSP hsp, int dg,
                BOOL fHUnits, _Out_ POINT **prgpt, _Out_ ULONG *pc, _Out_ BOOL *pfPolyline) PURE;

        // ----- Misc Section

        /* FDrawShape draws the Shape specified by hsp into an MSODC (*pdc) at
                the position specified in an MSOSVI (*psvi), WITHOUT any blip
                recoloring */
        MSOMETHOD_(BOOL, FDrawShape) (THIS_ MSOHSP hsp, MSODC *pdc,     MSOSVI *psvi) PURE;

        /* This method needs to be called in to draw all the shapes in this
                DG into a metafile. Typically, the host from inside its implementation
                of IMsoDrawingSite creates a DGVS that this DG can use to generate
                a DC, fills out a MSODGVSI and then calls this method, so that
                Office can use a uniform procedure to draw the shapes into that
                DC and generate a metafile. */
        MSOMETHOD_(HANDLE, HDrawShapesInMetafile) (THIS_ BOOL fEmf) PURE;

        MSOMETHOD_(void, BeginUsingTempShapeSpace) (THIS_ BOOL fClearSpace) PURE;
        MSOMETHOD_(void, EndUsingTempShapeSpace) (THIS) PURE;
        MSOMETHOD_(void, SetTempShapeSpace) (THIS_ MSOHSP hsp, DWORD dw) PURE;
        MSOMETHOD_(void, GetTempShapeSpace) (THIS_ MSOHSP hsp, DWORD *pdw) PURE;

        /*
        I'm going to try to sort the top-level shapes in this Drawing by z-order.
        Before you call this API, you must have called BeginUsingTempShapeSpace
        and given each top-level shape a unique temp value between 0 and n-1,
        inclusive.  I return FALSE if you call me without fulfilling my entry
        conditions, or if I can't allocate the memory to do this correctly.

        At the moment, I don't record undo because I don't feel like it.  This
        might be incorrect.
        */
        MSOMETHOD_(BOOL, FSortShapesByTempSpace) (THIS_ MSOHSP hsp) PURE;

        /* This method adds a default wrap polygon to the given shape and
                stores it in the shape's property list. If fUndo is TRUE, we
                assume that there is a undo operation already started, and
                will add undo records. */
        MSOMETHOD_(BOOL, FAddWrapPolygon) (THIS_ MSOHSP hsp, BOOL fUndo) PURE;


        /* Gets the transformed points of the wrap polygon of the given shape. */
        MSOMETHOD_(BOOL, FGetWrapPolygon) (THIS_ MSOHSP hsp, POINT *rgpth,
                int *pcpt) PURE;

        /* fetches msopidPVerticies, creates a new array (client's responsibility to free),
                and resolves any guided coordinates in place */
        MSOMETHOD_(BOOL, FGetResolvedVerticesHsp)(THIS_ MSOHSP hsp, MSOSVI *psvi,
                IMsoArray **ppirg) PURE;

        /* Special save functions for Word's inline shapes. Many restrictions apply.
                Use only if you know what you are doing. */
        MSOMETHOD_(BOOL, FLoadShape) (THIS_ MSOLDB *pldb, MSOHSP *phsp, void *pvAnchor) PURE;
        MSOMETHOD_(BOOL, FSaveShape) (THIS_ MSOSVB *psvb, MSOHSP hsp) PURE;
        MSOMETHOD_(void, AddPictureToCanvas) (THIS_ interface IMsoDrawingSelection *pidgsl, MSOHSP hspPict) PURE;

        // ----- IMsoDrawing methods TO GET RID OF !!!

        /* FIsShapeGroup returns TRUE iff the specified Shape is a Group. */
        /* TODO peteren: Remove this in favor of FQueryShape(msospqGroup) */
        MSOMETHOD_(BOOL, FIsShapeGroup) (THIS_ MSOHSP hsp) PURE;

        /* FIsShapeChild returns TRUE iff the specified Shape is the child of a Group. */
        /* TODO peteren: Resort all the DG methods. */
        MSOMETHOD_(BOOL, FIsShapeChild) (THIS_ MSOHSP hsp) PURE;

        /* FIsShapeChild returns TRUE iff the specified Shape has been deleted. */
        /* TODO peteren: Resort all the DG methods. */
        MSOMETHOD_(BOOL, FIsShapeDeleted) (THIS_ MSOHSP hsp) PURE;

        /* FIsShapeMarked returns TRUE iff the specified Shape is marked.
                Brilliant, no? */
        /* TODO peteren: Remove this in favor of FQueryShape(msospqMarked) */
        MSOMETHOD_(BOOL, FIsShapeMarked) (THIS_ MSOHSP hsp) PURE;

        /*
        Returns true if the given shape is on the Drawing's deleted list.
        */
        /* TODO peteren: Replace with FQueryShape(msospqDeleted) */
        MSOMETHOD_(BOOL, FDeletedHsp) (THIS_ MSOHSP) PURE;

        /*      Update means hunt down all the Drawing Views viewing this Drawing and
                tell them (okay, actually their Display Managers) to Update. This
                doesn't do any invalidation first; it's assumed that methods that do
                edits to the Drawing will do the necessary invalidation. These methods
                don't do updates, though, so that the client will have some flexibility
                in deciding when the updates happen.  They might choose, for instance,
                to update a "current" view immediately and others at idle.  Or they can
                just use this to update all of them. */
        /* TODO peteren: Make this method not public? */
        MSOMETHOD_(VOID, Update) (THIS) PURE;

        /* TODO peteren: Comment? */
        MSOMETHOD_(BOOL, FBeginReadAccess) (THIS) PURE;

        /* TODO peteren: Comment? */
        MSOMETHOD_(void, EndReadAccess) (THIS) PURE;

        /* TODO peteren: Comment? */
        MSOMETHOD_(BOOL, FBeginWriteAccess) (THIS) PURE;

        /* TODO peteren: Comment? */
        MSOMETHOD_(void, EndWriteAccess) (THIS) PURE;

        /* Call this before you call FCloneUndoRecord */
        MSOMETHOD_(BOOL, FBeginReconcile) (THIS_ IMsoDrawing *pidgDest) PURE;

        /* This function is called to do all the high-level drawing manipulation
                that needs to occur before undo records from this drawing can be
                run against another.  This should be called between FBeginReconcile
                and EndReconcile, and after you have unwound all the undo records
                back to their original state. */
        MSOMETHOD_(BOOL, FReconcile) (THIS) PURE;

        /* Call this after you are done making calls to FCloneUndoRecord to
                notify Office to clean up after itself. */
        MSOMETHOD_(void, EndReconcile) (THIS) PURE;

        // ----- IMsoDrawing methods stuck at the end to avoid full builds
        MSOMETHOD_(MSOHSP, HspGetBackgroundShape) (THIS) PURE;
        MSOMETHOD_(MSOHSP, HspGetRootParent) (THIS_ MSOHSP hsp) PURE;
        MSOMETHOD_(MSOHSP, HspGetParent) (THIS_ MSOHSP hsp) PURE;

        //--- Font Enumerations

        /*      Enumerate the fonts in the drawing. These fonts include the
                those used for text effects and those found in pictures. The callback
                is called with the fonts Warning: demand loads the pictures
                in the drawing. */
        MSOMETHOD_(BOOL, FEnumFonts) (THIS_ MSOPFNELF pfnelf, void *pvParam, ULONG grfelf) PURE;
        //--- Arc

        /* Set the parameters of an arc shape. An ellipse formed by the specified
           rectangle defines the curve of the arc. The arc extends from the point
                where it intersects the radial from the center of the rectangle to
                the start point. The arc ends where it intersects the radial from the
                center of the rectangle to the end point.       Modeled after the WIN32
                Arc function. */
        MSOMETHOD_(BOOL, FSetArc) (THIS_ MSOHSP hsp,
                RECT *prchEllipse,
                POINT *ppthStart,
                POINT *ppthEnd
                ) PURE;

        /* Get the parameters of the arc shape. See FSetArc */
        MSOMETHOD_(void, GetArc) (THIS_ MSOHSP hsp,
                RECT *prchEllipse,
                POINT *ppthStart,
                POINT *ppthEnd
                ) PURE;

        /* Get the Drawing Invalidation Flags for the given shape
                set hsp==NULL for the flags for the drawing  */
        MSOMETHOD_(BOOL, FGetShapeInvalidation) (THIS_ MSOHSP hsp, ULONG *pgrfdgi) PURE;

        /* Get the single invalid shape, returns FALSE if there are multiple */
        MSOMETHOD_(MSOHSP, HspGetSingleInvalidShape) (THIS) PURE;

        /* Call Validate to have the DG have all its dependents
                cope with the invalidation stored up in the DG, including
                firing the msodgeBeforeDgBecomesValid event to the DGS.
                The DG will then clear all the invalidation it has stored. */
        MSOMETHOD_(void, Validate) (THIS) PURE;
#if NEVER
        /* Preferred, handles non-top-level shapes better than FApplyPrchToHsp */
        MSOMETHOD_(BOOL, FApplyPrchToHsp) (THIS_ MSOHSP hsp, RECT *prch, BOOL fUndo) PURE;
#endif
        /* Sets the anchor rectangle of the shape to the given rch.
                Works on child shapes! */
        MSOMETHOD_(BOOL, FSetPrchOfHsp) (THIS_ RECT *prch, MSOHSP hsp, BOOL fUndo) PURE;

        /* An alternate version of FCreateHsp (see above) that takes a spidf.
                Only a separate function to preserve binary compatibility.
                FUTURE mmorgan: merge this with FCreateHsp once interface changes
                are allowed again. */
        MSOMETHOD_(BOOL, FCreateHsp2) (THIS_ MSOHSP *phsp, MSOSPT spt, RECT *prch,
                void *pvAnchor, MSOHSP hspPrepared, MSOHSP hspPrev, void **ppxhspGroup,
                int grfchp, MSOSPIDF spidf) PURE;

        /* Load an OLE shape from data collected form HTML
                TODO barnwl (hailiu): IMsoDrawingImport may be a better place for this method*/
        MSOMETHOD_(BOOL, FLoadOleShapeFromHtml)(THIS_
                LPVOID pvClient,
                REFCLSID rclsid,
                LPUNKNOWN punkData,
                LPRECT prch,
                MSOOID *poid,
                MSOHSP *phsp,
                BOOL fUndo) PURE;

        /* Create shapes from data in Clipboard file format */
        MSOMETHOD_(BOOL, FCreateShapesFromCF)(THIS_
                IMsoDrawing *pidgScrap,
                IMsoBlip *pib,
                interface IMsoDrawingSelection *pidgsl,
                RECT *prch) PURE;

        /* Draws shapes in a supplied IMsoBitmapSurface (T-MARKW)  This just
                calls the drawing site method. */
        MSOMETHOD_(BOOL, HDrawShapesInBitmapSurface) (THIS_
                interface IMsoBitmapSurface *pbms) PURE;

        /* HTML Export methods */
        /* tPassthroughHint is a tristate flag: msotYes: unmodified pasthru
                msotNo: not unmodified passthru
                msotMaybe: don't know, FWriteShapeXML should guess */
        MSOMETHOD_(BOOL, FWriteShapeXML)(THIS_ MSOXMLW *pxmlw, MSOHSP hsp,
                DWORD hetnUser, int tPassthroughHint) PURE;
        MSOMETHOD_(BOOL, FGetShapeIDString)(THIS_ MSOHSP hsp,   WCHAR *wzName,
                int cch, MSOXMLW *pxmlw, int* piKind) PURE;
        MSOMETHOD_(BOOL, FGetSpidString)(THIS_ MSOHSP hsp, WCHAR *wzName, int cch)
                PURE;
        MSOMETHOD_(BOOL, FExportShapeTypesVML)(THIS_ MSOXMLW *pxmlw,
                DWORD hetnUser) PURE;

        /* HTML Import methods */
        MSOMETHOD_(BOOL, FCreateHspFromImgHtml)(THIS_ MSOHISD *phisd,
                POINT *pptTopLeft, MSOHSP *phsp, BOOL fCreateHsp) PURE;
        MSOMETHOD_(BOOL, FCreateHspFromAreaHtml)(THIS_ MSOHISD *phisd,
                MSOHSP *phsp, RECT *prcBounds) PURE;
        MSOMETHOD_(BOOL, FCleanupAreaHspsHtml)(THIS_ MSOHISD *phisd, BOOL fUseSameDg) PURE;
        MSOMETHOD_(BOOL, SetShapeHtmlId)(THIS_ MSOHSP hsp, WCHAR *wzId, int cchId, BOOL fCFHTML) PURE;
        MSOMETHOD_(void, SetHTMLImport) (THIS) PURE;
        MSOMETHOD_(void, EndHTMLImport) (THIS_ BOOL fCfhtml, interface IMsoReducedHTMLImport *pirhi, ULONG msogrfehi,
                IMsoDrawing *pidgClone) PURE;

        MSOMETHOD_(BOOL, FWriteBackgroundXML)(THIS_ MSOXMLW *pxmlw, DWORD hetnUser) PURE;
        MSOMETHOD_(BOOL, FLoadBackgroundFromHtml)(THIS_ MSOHISD *phisd, MSOETK *petk) PURE;
        MSOMETHOD_(BOOL, FApplyHTMLBackground)(THIS_ MSOCVS *pcvs,
                interface IMsoReducedHTMLImport *pirhi) PURE;

        /* Returns the EMU size of the blip in this drawing - takes account
                of the DGG target DPI if necessary. */
        MSOMETHOD_(void, PtaFromBlip)(THIS_ POINT *ppta, IMsoBlip *pib) CONST_METHOD PURE;

        /* Return the DPI of this drawing, x and y DPI values may be different,
                so a POINT is filled in. */
        MSOMETHOD_(void, DrawingDPI)(THIS_ POINT *pptDpi) CONST_METHOD PURE;

        /* Returns an IMsoGE for the HSP. Very similar to DGV method, but doesn't require
           a DGV (user has to fill in the svi himself, though).  Client is responsible
           for calling pige->Free() when he is done with it. */
        MSOMETHOD_(IMsoGE*, PigeFromHsp) (THIS_ MSOHSP hsp, MSOSVI *psvi) PURE;

        /* Our next-generation shape iterator, better than FEnumX.
           We support msofShapeIterSafe, msofShapeIterWantGroups, and
           msofShapeIterWantChildren.
           The hspGroup is optional; we will iterate its children (or contents
           of a drawing canvas), otherwise the whole drawing. */
        MSOMETHOD_(BOOL, FCreateShapeIter) (THIS_ ULONG msogrfShapeIterFlags,
                MSOHSP hspGroup, interface IMsoShapeIter** ppISI) PURE;
        MSOMETHOD_(BOOL, FPrepareDiagramsForRTFOut) (THIS) PURE;
        MSOMETHOD_(BOOL, FFixupDiagrams) (THIS) PURE;
        
        /* Explicitly set diagram bounds.  Currently used to do font/connector
        scaling for org charts. */
        MSOMETHOD_(void, ScaleDiagramContents) (THIS_ MSOHSP hspDiagram,
                RECT *rchOld, RECT *rchNew) PURE;

        /* Allows re-ordering of top level shapes.  Will fail if either
        hspToMove or hspPrev is not top-level.  hspToMove is re-oredered in the
        z-list to be directly after hspPrev.  If hspPrev is NULL then hspToMove
        is set as the first shape.*/
        MSOMETHOD_(BOOL, FReorderShape) (THIS_ MSOHSP hspPrev, MSOHSP hspToMove) PURE;

        /* This method updates the flattened colors in the shape 'hsp'. If hsp
                is equal to msohspNil then all shapes in the DG are updated.*/
        MSOMETHOD_(BOOL, FUpdateColors) (THIS_ MSOHSP hsp) PURE;

        /* GrfdgqQuery can be used to ask yes/no questions of the DG.
                msogrfdgq's are an #defined questions we support. */
        MSOMETHOD_(ULONG, GrfdgqQuery) (THIS_ ULONG grfdgq) PURE;
        };
#endif // !MKCSM

MSOAPI_(BOOL) MsoFValidHsp(IMsoDrawing *pidg, MSOHSP hsp);
MSOAPI_(BOOL) MsoFValidDg(IMsoDrawing *pidg);
MSOAPI_(BOOL) MsoFGetRelatedShape(MSOHSP *phsp, MSODGRST dgrst,
        void **ppvClient);
MSOAPI_(void) MsoGetPvAnchorOfHsp(void **ppvClient, MSOHSP hsp);
MSOAPI_(void) MsoGetClientDataOfHsp(void **ppvClient, MSOHSP hsp);
MSOAPI_(IMsoDrawing *) MsoPidgFromHsp(MSOHSP hsp);
MSOAPI_(BOOL) MsoFSetPropertyOfHsp(MSOHSP hsp, MSOPID opid,
                                                                                                const void* pop, BOOL fUndo);
MSOAPI_(BOOL) MsoFGetPropertyOfHsp(MSOHSP hsp, MSOPID opid, void* pop);
MSOAPI_(MSOHSP) MsoHspNextRoot(MSOHSP hsp, void **ppvClient);
MSOAPI_(MSOHSP) MsoHspNextAll(MSOHSP hsp, void **ppvClient);
MSOAPI_(BOOL) MsoFSetPiViewObject(MSOHSP hsp, IViewObject *piViewObject);
MSOAPI_(void) MsoInitXMLWrite(MSOXMLW *pxmlw,
        interface IMsoHTMLExport *pihe, MSOPX *ppxMasterPsp, MSOPX *ppxpsp,
        void *pvClient);

/* determines whether a particular <IMG shapes> attribute should be
        kept or ignored.  rgpidg contains an array of drawings where the
        shapes referenced in wzShapes might be found. */
MSOAPI_(BOOL) MsoFWantImgShapesAttribute(const MSOHISD *phisd,
        const WCHAR *wzShapes, IMsoDrawing **rgpidg, int cpidg);

/*      Saves the fill property set from the buffer 'pops' in XML.
        NOTE: The caller retains the ownership of the properties in 'pops', the
        caller must pass in the array of scheme colors used. */
MSOAPI_(BOOL) MsoFWriteFillPropSetXML(MSOXMLW *pxmlw, const void *pops,
        DWORD hetnUser, int ccrScheme, const COLORREF *prgcrScheme, interface IMsoDrawing* pidg);

/*      Saves the line property set from the buffer 'pops' in XML.
        NOTE: The caller retains the ownership of the properties in 'pops', the
        caller must pass in the array of scheme colors used. */
MSOAPI_(BOOL) MsoFWriteLinePropSetXML(MSOXMLW *pxmlw, const void *pops,
        DWORD hetnUser, int ccrScheme, const COLORREF *prgcrScheme, interface IMsoDrawing* pidg);

/* Initialize the scheme to sys color mapping array - call once for each
        scheme color which needs to be written as a system color. */
MSOAPI_(void) MsoInitSysColorOfScheme(int ischeme, COLORREF crSysColor);

/*      FGetAltText gets the Alternative Text of the shape.  If the Alt. text
        has not be defined, it will generated base on the shape types.
*/
MSOAPI_(BOOL) MsoFGetAltText(MSOHSP hsp, WCHAR *pwz, int *pcch, BOOL fNoDefault);
MSOAPIX_(BOOL) MsoFGetAltTextFromPdgsl(interface IMsoDrawingSelection *pidgsl,
        WCHAR *pwz, int *pcch, BOOL fNoDefault);

/* Adjust the transform of an existing ink surface to match a current drawing view. */
MSOAPI_(void) MsoSetInkSurfaceXformFromDgv(interface IMsoInkSurface *piSurface, interface IMsoDrawingView* pidgv);
/* Adjust the transform of an existing ink renderer to match a current drawing view. */
MSOAPI_(void) MsoSetInkRendererXformFromDgv(interface IInkRenderer *piRenderer, HDC hdc, interface IMsoDrawingView* pidgv);

/* Extract the ink strokes from the specified IInkDisp and turn them into shapes. */
//MSOAPI_(void) MsoCreateShapesFromInkDisp(IInkDisp *piInkDisp, IMsoDrawing *pidg, BOOL fSingleShape);
MSOAPI_(void) MsoCreateShapesFromInkDisp(interface IMsoInkSurface *piSurface, interface IMsoDrawing *pidg,
                                         interface IMsoDrawingView *pidgv, interface IInkStrokes *piInkUndo, 
                                         BOOL fUndo, BOOL fWantCallbackEvent);
// TODO SteveBre: Temporary signature -- should use params above.

/* MsoBlipFromInk Creates an IMsoBlip from strokes in IInkDisp. */
MSOAPI_(BOOL) MsoFCreateBlipFromInk(interface IInkDisp* piid, interface IMsoBlip **ppib);

MSOAPI_(BOOL) MsoFDeleteInkAnnotations(interface IMsoDrawing* pidg, int idsgdu);

/*
Optimization of images -- helper functions
*/
#define msogrfOptNil                    0x00000000
#define msogrfOptResample               0x00000001
#define msogrfOptCrop                   0x00000002
#define msogrfOptCompress               0x00000004
/* If this flag is set, one blip that appears in multiple shapes will be 
	resampled separately for each shape */
#define msogrfOptAllowBlipSplit         0x00000008
#define msogrfOptOnlyMarkedShapes       0x00000010
#define msogrfOptOnlySelected           0x00000020
#define msogrfOptUndo                   0x00000040


/*****************************************************************************
        TODO comment
*****************************************************************************/
typedef enum
        {
        msoblipDefault  = 0UL + (4UL<<2),

        /* Selection */
        msoblipNotUsed            =(0UL), // Default Evaluation
        msoblipNotSelected        =(1UL),
        msoblipPartiallySelected  =(2UL),
        msoblipSelected           =(3UL),
        msoblipSelectedMask       =(3UL), // Mask to extract selection info for blip

        /* Rotation */
        msoblipRotatedAxisX       =(0UL<<2), // Note: these values are MSOAXIS<<2
        msoblipRotatedAxisY       =(1UL<<2),
        msoblipRotatedAxisNegX    =(2UL<<2),
        msoblipRotatedAxisNegY    =(3UL<<2),
        msoblipRotatedNotUsed     =(4UL<<2), // Default Evaluation
        msoblipMultipleTransforms =(5UL<<2), // Requires blip split to implement rotation
        msoblipRotationMask       =(7UL<<2), // Mask to extract rotation info

        /* Flipping [only valid if Rotation is 1-4] */
        msoblipNotFlippedHoriz    =(0UL<<5), // Default
        msoblipFlippedHoriz       =(1UL<<5),
        msoblipFlippedHorizMask   =(1UL<<5), // Mask to extract flipping info

        msoblipTransformMask      =(15UL<<2) // Mask flipping + rotation
        } MSOOBEINFO;

MSOAPI_(BOOL) MsoPibOptimizeFromHsp(interface IMsoDrawing *pidg, MSOHSP hsp,
                                                                                                          ULONG grfOptValues);
MSOAPI_(BOOL) MsoFCanOptimize(MSOHSP hsp);
MSOAPI_(BOOL) MsoFCanOptimizeLossLessly(MSOHSP hsp);

MSOAPIX_(BOOL) MsoFReducible(const IMsoBlip* pibIn);
MSOAPIX_(BOOL) MsoFLossLessPalettize(const IMsoBlip* pibIn);
MSOAPIX_(BOOL) MsoFLessThanNColors(const IMsoBlip* pibIn);
MSOAPIX_(BOOL) MsoFIsBig(MSOHSP hsp);

MSOAPI_(BOOL) MsoFTryToOptimizeFromHsp(IMsoDrawing *pidg, MSOHSP hsp,
        ULONG grfOptValues);

/* A limit on the number of characters in Alt Text */
#define MsoCchMaxAltText 4096



/****************************************************************************
        Drawing Site

        The host must provide an IMsoDrawingSite interface pointer for each
        DG it creates.  Actually that's not entirely true; whenever the host
        passes an IMsoDrawingSite * into a DG they may also pass a void * of
        client data (pvDgs); this will be remembered and passed back as the
        first argument of every DGS method.  This allows the host to implement
        the DGSs of lots of DGs with a single actual object.
*****************************************************************************/

/* MSODGEB -- Drawing Event Block */
typedef struct MSODGEB
        {
        /* These first fields are filled out for EVERY Drawing Event. */
        MSODGE dge;
        BOOL fResult;
        interface IMsoDrawing *pidg;
        MsoDEBPDefine(dg);

        /* All the rest of these fields are only filled out by some Events.
                See the comments by the definitions of the MSODGEs to see which
                Events fill out which fields.  Nice new events use
                their own named struct in the union, dgexEventNameHere. */

        /* TODO peteren: These fields are hideous.  Some events even
                fill out BOTH hsp and pvClient AND wParam and lParam. */

        WPARAM wParam;
        LPARAM lParam;

        MSOHSP hsp;
        void *pvClient;
        void *pvAnchor;

        /* Proper new events each use their own dgexFoo */
        union
                {
                struct
                        {
                        MSOHSP hspModel; // Model from which new shape copies anchor, z-order, and other attributes.
                        } dgexRequestCancelCreateShape;
                struct
                        {
                        interface IMsoDrawing *pidgOriginal;
                        void *pvDgsOriginal;
                        MSOHSP hspOriginal;
                        void *pvClientOriginal;

                        interface IMsoDrawing *pidgClone;
                        void *pvDgsClone;
                        MSOHSP hspClone;
                        void *pvClientClone;
                        } dgexRequestCancelCloneShape;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        MSOSPC spc;
                        } dgexRequestChooseSpc;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        int spt;
                        } dgexRequestChooseShapeType;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        } dgexBeforeChangeShapeType;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        } dgexAfterChangeShapeType;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        MSODRGPRC *pdrgprc;
                        ULONG drghn; // TODO (ilanb): this is actually an MSODRGH
                        } dgexConstrainShape;
                struct
                        {
                        MSOHSP hsp;
                        BOOL fAddText;
                        MSOTXID txid;
                        BOOL fOid;
                        } dgexTextboxUndo;
		struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        BOOL fUndo;
                        MSOPSID opsid;                  // Modified propset id. If more than one
                                                                                // propset is changed it is set to msopsidNil.
                        MSOSPPID sppid;    // set when fake sppid is used for a property change
                                                                                // opsid is set to msopsidNil 
                        } dgexBeforeShapePropertyChange;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        BOOL fUndo;
                        MSOPSID opsid;                  // Modified propset id. If more than one
                                                                                // propset is changed it is set to msopsidNil.
                        MSOSPPID sppid;    // set when fake sppid is used for a property change
                                                                                // opsid is set to msopsidNil
                        } dgexAfterShapePropertyChange;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        IStorage *pistg;
                        } dgexRequestFakeOleObject;
                struct
                        {
                        HENHMETAFILE hemf;
                        HANDLE hmfp;
                        BOOL fNeedLnkSrc;
                        MSOHSP hsp;
                        IStream *pstmLink;
                        } dgexAfterCopy;
                struct
                        {
                        MSOHSP hspGroup;
                        } dgexBeforeUngroupShape;
                struct
                        {
                        MSOHSP hspSingleInvalid;
                        } dgexBeforeDgBecomesValid;
                struct
                        {
                        BOOL fGroup;
                        MSOHSP hsp;
                        } dgexUndoGroup;
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        WCHAR *wzName;
                        } dgexRequestMatchShapeName;
                struct
                        {
                        RECT* prch;
                        } dgexDoRelativeOffset;
                struct  // for msodgeDoSetNextTxbx
                        {
                        MSOHSP hspSrc;
                        MSOHSP hspNext;
                        } dgexDoSetNextTxbx;
                struct
                        {
                        MSOHSP hspAnchor;
                        void *pvDgsAnchor; //client data of hspAnchor drawing
                        } dgexDoReanchor;
                struct
                        {
                        MSOHSP hspChild;
                        RECT rch;
                        void *pvAnchor;
                        } dgexCreateGroupAnchorFromChild;
                struct
                        {
                        BOOL *fInlineShape;//initialized to false before the
                        } dgexBeforeStyle; //event is fired
#ifdef NO96_162753
                struct
                        {
                        MSOHSP hsp;
                        MSOOID oid;
                        } dgexFIconObject;
#endif // NO96_162753
#if BIDI
                struct
                        {
                        LONG *pxh;
                        RECT *prch;
                        BOOL fRightToLeftSheet; //XL needs that for WordArt placement (XL10 bug 200424)
                        } dgexBeforeConvertHToO;
                struct
                        {
                        LONG *pxh;
                        RECT *prch;
                        } dgexAfterConvertOToH;
#endif //BIDI
                struct
                        {
                        BOOL fMaster;
                        } dgexBeforeShapeXML;
                struct
                        {
                        MSOHSP hsp;
                        BOOL   fSuccess;
                        BOOL   fAppearanceChange;
                        const WCHAR *wzFile;
                        IMsoBlip *pib;
                        } dgexAfterAsyncBlipDownload;
                struct
                        {
                        MSOHSP hsp;
                        void   *pvClient;
                        BOOL   fInlineOle;
                        } dgexOCXShape;
                struct
                        {
                        MSOHSP hsp;
                        void   *pvClient;
                        BOOL fSpidF; // TRUE if the following spid is actually spidF
                        MSOSPID spidOld;
                        MSOSPID spidNew;
                        } dgexBeforeChangeSpid;
                struct
                        {
                        WCHAR *wzPibName;
                        } dgexRequestHspPibNameAlloc;
                struct
                        {
                        RECT* prch;
                        } dgexDoHRAdjustment;
                struct
                        {
                        BOOL   fAtParaStart;
                        } dgexBeforePosPropXML;
                struct
                        {
                        COLORREF rgclr[4];
                        } dgexBeforeBorderPropXML;
                struct
                        {
                        MSOHSP hsp;
                        void   *pvClient;
                        } dgexIndentsText;
                struct // for msodgeDoRchForAlign
                        {
                        RECT *prchPage;
                        interface IMsoDrawingSelection  *pidgsl;
                        } dgexDoRchForAlign;
                struct // for msodgeAfterShapeCreateFromUI
                        {
                        interface IMsoDrawingSelection  *pidgsl;
                        } dgexAfterShapeCreateFromUI;
                struct // for msodgeAfterCreateDrawingCanvas
                        {
                        BOOL fResize;
                        RECT rchResize;
                        BOOL fSnapGrid;
                        BOOL fGrow;
                        } dgexAfterCreateDrawingCanvas;
                struct
                        {
                        int fAfterSaveDg; //-1 is the ninch value
                        int fAfterSaveDgg; //-1 is the ninch value
                        } dgexAfterNeedsSave;
                struct // for msodgeDoClipboardClone
                        {
                        interface IMsoDrawing *pidgSrc;
                        interface IMsoDrawing *pidgDest;
                        MSOCDID cdid;
                        } dgexDoClipboardClone;
                struct  // for msodgeDoOffsetForDuplicate
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        int       dxh;  // x offset in h coords
                        int       dyh;  // y offset in h coords
                        } dgexDoOffsetForDuplicate;
                struct  // for msodgeOffsetFromShapesOnPaste
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        int       dxh;  // x offset in h coords
                        int       dyh;  // y offset in h coords
                        } dgexOffsetFromShapesOnPaste;
                struct // for msodgeRequestCancelPaste
                        {
                        interface IMsoDrawing *pidgSrc;
                        interface IMsoDrawing *pidgDest;
                        MSOCDID cdid;
                    } dgexRequestCancelPaste;
                struct // for msodgeResolveReadingOrder
                        {
                        MSOHSP hsp;
                        MSOTXDIR txidResolved;
                    } dgexResolveReadingOrder;
                struct // for msodgeRequestCancelDeleteShape
                        {
                        MSOHSP hsp;
                        void   *pvClient;
                        BOOL   fGroupOnly;
                        BOOL   fUndo;
                        } dgexRequestCancelDeleteShape;

                struct // for msodgeDoGetFrameButtons
                        {
                        MSOHSP hsp;
                        void   *pvClient;
                        BOOL   fJumpBack;
                        BOOL   fJumpForth;
                        BOOL   fOverflow;
                        BOOL   fGroup;
                        BOOL   fUngroup;
                        BOOL   fSmartObject;
                        BOOL   fGroupTL;
                        BOOL   fWebComponent;
                        } dgexDoGetFrameButtons;
                struct // for msodgeDoScaleTextFonts
                        {
                        MSOTXID txid;
                        UINT uiScale; // 100% would mean no change
                        UINT defaultTextSize;
                        BOOL fAbsolute; // if this is set, then uiScale is absolute font size.
                        BOOL fUndo;
                        BOOL fForDiagram;
                        } dgexDoScaleTextFonts;
                struct // for msodgeSaveVisualElement
                        {
                        MSOSVB* pmsosvb;
                        void* pVisual;
                        } dgexSaveVisualElement;
                struct // for msodgeLoadVisualElement
                        {
                        MSOLDB* pmsoldb;
                        void** ppVisual;
                        } dgexLoadVisualElement;
                struct // for msodgeAfterAddShapeToGroup and AfterRemoveShapeFromGroup
                        { // and msodgeAddDelShapeGroupUndo
                        MSOHSP hspGroup;
                        BOOL fFireUndo;
                        BOOL fDrag;
                        BOOL fGoingIn;
                        BOOL fCreate;
                        } dgexAddRemoveToGroup;
                struct // for msodgeDoUndo
                        {
                        BOOL fPause;
                        BOOL fResume;
                        } dgexDoUndo;
                struct // for msodgeAfterConvertOrgChart; OBSOLETE
                        {
                        MSOHSP hspOrig;
                        MSOHSP hspNew;
                        } dgexConvertOrgChart;
                struct //msodgeDoHspOfUndoRecord
                        {
                        MSOHSP hsp;
                        } dgexDoHspOfUndoRecord;
                struct
                        {
                        MSOHSP hsp;
                        MSOTXID txid;
                        RECT *prchTextFrame;
                        } dgexRequestTextBounds;
                struct // for msodgeADoSetDiagramDefSize
                        {
                        MSODGMT dgmt;
                        RECT rcvSize;
                        } dgexDoSetDiagramDefSize;
                struct
                        {
                        MSOHSP hsp;
                        MSODGMT dgmt;
                        MSOTXALIGN alignment;
                        } dgexApplyTextStyleToShape;
                struct
                        {
                        UINT preferredTextSize;
                        } dgexGetPreferredTextStyles;
                struct
                        {
                        POINT ptScreenExtent;            // Full screen slideshow size chosen (in)
                        POINT ptPicturePixelDimensions;  // Pixel dimensions of the picture     (in)

                        POINT ptResolvedPictureExtentsH; // picture extents in H coords (out)
                        } dgexResolveBestScaleForSlideShowSize;
                struct
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        BOOL fUseBackLineColor;
                        } dgexUseBackForLineColor;
                struct
                        {
                        MSOXMLW *pxmlw;
                        BOOL fBegin; // begin tag for child shape
                        } dgexChildShapeVML;
                struct
                        {
                        MSOHSP hspOrig;
                        MSOHSP hspNew;
                        } dgexDoFinishCreateShapeFromModel;
                struct
                        {
                        MSOHSP hspInk;
                        int gvPage;
                        } dgexDoFindPageForInkShape;

                };
        } MSODGEB;
#define cbMSODGEB (sizeof(MSODGEB))


typedef enum
        {
        msodrwskShape,
        msodrwskShapeRange,
        msodrwskGroupShapes,
        msodrwskShapeNodes,
        msodrwskShapeNode,
        msodrwskCalloutFormat,
        msodrwskConnectorFormat,
        msodrwskFillFormat,
        msodrwskLineFormat,
        msodrwskShadowFormat,
        msodrwskTextEffectFormat,
        msodrwskThreeDFormat,
        msodrwskPictureFormat,
        msodrwskColor,
        msodrwskFreeformBuilder,
        msodrwskDiagramNode,
        msodrwskDiagramNodes,
        msodrwskDiagramNodeChildren,
        msodrwskDiagram,
        msodrwskCanvasShapes,
        } MSODRWSK;

/* MSOISD = Init Shape Data */
typedef struct MSOISD
        {
        interface IMsoDrawing *pidg;
        MSOHSP hsp;
        void *pvContext;

        BOOL fClient;
        void *pvClient;
        BOOL fAnchor;
        void *pvAnchor;
        BOOL fAttachedText;
        MSOTXID txid;
        BOOL fOleObject;
        MSOOID oid;
        } MSOISD;

/* MSOPSD = Purge Shape Data */
typedef struct MSOPSD
        {
        interface IMsoDrawing *pidg;
        MSOHSP hsp;
        void *pvContext;

        BOOL fClient;
        void *pvClient;
        BOOL fAnchor;
        MSODGSAK dgsak;
        void *pvAnchor;
        BOOL fAttachedText;
        MSOTXID txid;
        BOOL fOleObject;
        MSOOID oid;
        } MSOPSD;

/* MSOCSD = Clone Shape Data */
typedef struct MSOCSD
        {
        interface IMsoDrawing *pidgOriginal;
        MSOHSP hspOriginal;
        interface IMsoDrawing *pidgClone;
        MSOHSP hspClone;
        interface IMsoDrawingSite *pidgsClone;
        void *pvDgsClone;
        BOOL fPurgeOriginal;
        void *pvContext;

        BOOL fClient;
        void *pvClientOriginal;
        BOOL fAnchor;
        void *pvAnchorOriginal;
        BOOL fAttachedText;
        MSOTXID txidOriginal;
        BOOL fOleObject;
        MSOOID oidOriginal;

        void *pvClientClone;
        void *pvAnchorClone;
        MSOTXID txidClone;
        MSOOID oidClone;
        } MSOCSD;

/* MSOCSD = Move Shape Data */
typedef struct MSOMSD
        {
        interface IMsoDrawing *pidgOriginal;
        MSOHSP hspOriginal;
        interface IMsoDrawing *pidgClone;
        MSOHSP hspClone;
        interface IMsoDrawingSite *pidgsClone;
        void *pvDgsClone;
        BOOL fPurgeOriginal;
        void *pvContext;

        BOOL fClient;
        void *pvClientOriginal;
        BOOL fAnchor;
        void *pvAnchorOriginal;
        BOOL fAttachedText;
        MSOTXID txidOriginal;
        BOOL fOleObject;
        MSOOID oidOriginal;

        void *pvClientClone;
        void *pvAnchorClone;
        MSOTXID txidClone;
        MSOOID oidClone;
        } MSOMSD;

/* IMsoDrawingSite (idgs).  This is the actual interface a Drawing Site
        must speak. */
#undef  INTERFACE
#define INTERFACE IMsoDrawingSite
DECLARE_INTERFACE_(IMsoDrawingSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingSite methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* The Drawing will call DGS::OnEvent whenever a drawing
                event (MSODGE) occurs.  See the definition of MSODGE for a more
                complete explanation of events.  This function will only be
                called at all if you set msodgsi.fWantsEvents to TRUE. */
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDgs, MSODGEB *pdgeb) PURE;

        // ----- IMsoDrawingSite methods (low-level editing)

        /* This method is called when we're initializing a Shape. We will pass
                you a MSOISD of all the host data that we think that the shape
                contains, with fClient, fAnchor, fAttachedText and fOleObject
                reflecting which fields we believe to be initialized. */
        MSOMETHOD_(BOOL, FInitShapeData) (THIS_ void *pvDgs, MSOISD *pisd) PURE;

        /* This method is called when we're purging a Shape.  By purging I mean
                that the Shape is really truly going away, as opposed to "deleting"
                which simply means the Shape is being moved out of the main drawing and
                into the "deleted" list (from whence undo can get it back if
                necessary). The return value of these methods is not yet well defined,
                so we currently assert that you return TRUE.  We will pass you a MSOPSD
                of all the host data that we think that the shape contains, with
                fClient, fAnchor, fAttachedText and fOleObject reflecting which fields
                we believe to be purged. */
        MSOMETHOD_(BOOL, FPurgeShapeData) (THIS_ void *pvDgs, MSOPSD *ppsd) PURE;

        /* This method is called when we're cloning a Shape from one DG into
                another one.  These should return FALSE if they fail. We will pass you
                a MSOPSD of all the host data that we think that the shape contains.
                It will always have a pvClientData, but the pvAnchor, MSOTXID and
                MSOOID are optional.  Consequently we have fAnchor, fAttachedText and
                fOleObject, respectively, that says whether or not they are included. */
        MSOMETHOD_(BOOL, FCloneShapeData) (THIS_ void *pvDgs, MSOCSD *pcsd) PURE;

        /* This method is called when we're moving a Shape from one DG into
                another one.  These should return FALSE if they fail. We will pass you
                a MSOPSD of all the host data that we think that the shape contains.
                It will always have a pvClientData, but the pvAnchor, MSOTXID and
                MSOOID are optional.  Consequently we have fAnchor, fAttachedText and
                fOleObject, respectively, that says whether or not they are included. */
        MSOMETHOD_(BOOL, FMoveShapeData) (THIS_ void *pvDgs, MSOMSD *pmsd) PURE;

        // ----- IMsoDrawingSite methods (anchors and coordinates)

        /* FCreateAnchor should look at the rch we pass in and store
                in *ppvAnchor enough data (probably a pointer to a RECT or something
                like it that you allocate) to return that rch later if we call
                FLocateAnchor.  It doesn't have to be a rectangle, though, it
                can be some kind of reference to your own document.  PowerPoint
                allocates a rectangle and gives us a pointer to it.  Excel allocates
                an ORC (SPRC?) and gives us a pointer to that.  Word allocates
                a new entry in the PLCSPA and gives us an _index_ to that.
                You can probably ignore gvPage, unless you're Word.  For Word it does
                not suffice to be given a rectangle, we also need to know which
                page the Shape goes on.  So someway or other we let the host pass in
                a uvPage so that we can pass it back to them.  At least initially
                (and quite possibly forever) there are cases when we don't have
                a uvPage to pass back.  And since there is no well-defined value
                for uvNil, we also need an fPage BOOL that's TRUE iff gvPage is valid. */
        /* TODO peteren: There was something here about not calling this method
                if fAnchorIsRch is TRUE?  Lose this. */
        MSOMETHOD_(BOOL, FCreateAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSOHSP hsp, RECT *prch, BOOL fPage, MSOGV gvPage) PURE;

        /* FLocateAnchor should convert whatever data the host hung off
                pvAnchor into an RCH, return TRUE on success. */
        /* TODO peteren: What's up with fAnchorIsRch? */
        /* TODO peteren: Do we need pvPage here? */
        MSOMETHOD_(BOOL, FLocateAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSOHSP hsp, BOOL fPage, MSOGV gvPage, RECT *prch) PURE;

        /* When saving a Shape we'll FSaveAnchor to save the data hanging
                off its pvAnchor.  This should return TRUE iff it succeeds. If
                fAnchorIsRch is TRUE we may just save the RCH as flat data
                without calling you. */
        MSOMETHOD_(BOOL, FSaveAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSODGSAK dgsak, MSOHSP hsp, MSOFBI fbi, MSOSVB *psvb) PURE;

        /* FLoadAnchor should load a shape saved by FSaveAnchor, returning
                TRUE iff it succeeds.  If fAnchorIsRch is TRUE we may just load
                the RCH ourselves (although only if we saved it ourselves). */
        MSOMETHOD_(BOOL, FLoadAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSODGSAK dgsak, MSOHSP hsp, MSOLDB *pldb) PURE;

        /* The FChangeAnchor method is to move a Shape that already has an
                anchor (from FCreateAnchor or FLoadAnchor) to a new rch.  This gives
                the host a chance to be efficient, although if you're lazy you can
                probably just call your own FreeAnchor method and the FCreateAnchor.
                See comment by FCreateAnchor for explanation of pvPage.  The
                ppvAnchorUndo parameter allows the host to store its own data in
                the undo stack; most hosts will just ignore this.  The grfsac indicates
                what we're doing, i.e. an initial edit in which case the host should
                initialize ppvAnchorUndo, an undo or redo action in which case the
                host can potentially use and/or change the value there, or a special
                non-undoable action in which case ppvAnchorUndo is to be ignored. */
        MSOMETHOD_(BOOL, FChangeAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSOHSP hsp, RECT *prch, void **ppvAnchorUndo, ULONG grfsac,
                BOOL fPage, MSOGV gvPage) PURE;

        /* We'll call FDeleteAnchor when the anchor of a shape is being deleted;
                that is, relegated to the undo or redo stack.  We call this when the
                Shape itself is being deleted (moved to the undo stack) and in cases
                like Group where we're just moving the anchor itself to the undo stack.
                It's likely that only Word will have to do anything in here. */
        MSOMETHOD_(BOOL, FDeleteAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSOHSP hsp) PURE;

        /* We'll call FUndeleteAnchor when we're undoing the deletion of
                an anchor; that is, moving it from the undo/redo stack back into
                a real shape. It's likely that only Word will have to do any work here. */
        MSOMETHOD_(BOOL, FUndeleteAnchor) (THIS_ void *pvDgs, void **ppvAnchor,
                MSOHSP hsp) PURE;

        // ----- IMsoDrawingSite methods (client data)

        /*      The host is given a pointer to a 32 bit value unique to the given shape
                which it may fill out as it pleases. Return FALSE if memory
                failure, etc.  We call this method whenever a Shape is created.  Note,
                however, that we may call this method BEFORE the Shape is entirely
                finished (for example, if we're inserting a Shape we'll call this
                before we know how big the Shape is going to be*/
        MSOMETHOD_(BOOL, FCreateClientData) (THIS_ void *pvDgs, void **ppvClient,
                MSOHSP hsp) PURE;

        /*
        The shape is being written to the given stream, so the host needs to save
        what corresponds to their client data value, if anything.
        */
        MSOMETHOD_(BOOL, FSaveClientData) (THIS_ void *pvDgs, MSOSVB *psvb,
                void **ppvClient, MSOHSP hsp) PURE;

        /*
        The shape is being read from the given stream, so the host needs to load
        and set their client data value given a pointer to it.
        */
        MSOMETHOD_(BOOL, FLoadClientData) (THIS_ void *pvDgs, MSOLDB *pldb,
                void **ppvClient, MSOHSP hsp) PURE;


        // ----- IMsoDrawingSite methods (undo)

        /*      FBeginUndo is called by Office to indicate that a series of calls to
                FPostUndoRecord will commence, together indicating an undoable
                command. */
        MSOMETHOD_(BOOL, FBeginUndo) (THIS_ void *pvDgs) PURE;

        /*      FEndUndo is called by Office to indicate that the current command is
                complete. If fCommit is FALSE the command failed and we expect but do not
                require the host to "roll-back" (that is, immediately call
                FExecuteUndoRecord) any undo records that were passed since the last
                FBeginUndo.  In szCommand is a string you can use to lable your undo
                command.  TODO peteren: Perhaps this should be an identifier instead? */
        MSOMETHOD_(BOOL, FEndUndo) (THIS_ void *pvDgs, BOOL fCommit,
                WCHAR *wzCommand) PURE;

        /* When Office Drawing makes changes to a Drawing it packages up
                an undo record (referenced via an MSOHUR) and gives it to the
                Drawing Site to hold onto.  In pdgcExecuting you'll find a pointer
                to the Command that's being run. */
        MSOMETHOD_(BOOL, FPostUndoRecord) (THIS_ void *pvDgs, MSOHUR hur,
                MSODGC *pdgcExecuting) PURE;

        // ----- IMsoDrawingSite methods (textboxes)


        /* FCreateTextbox creates a new textbox and returns a TXID to
                identify it. */
        MSOMETHOD_(BOOL, FCreateTextbox) (THIS_ void *pvDgs, MSOTXID *ptxid,
                MSOHSP hsp) PURE;

        // Return the text stored for this textbox
        // TODO davebu -- do we need this method?
        MSOMETHOD_(BOOL, FGetText) (THIS_ void *pvDgs, MSOTXID txid, MSOHSP hsp,
                WCHAR **pwz) PURE;

        // Return the text stored for this textbox
        // TODO davebu -- do we need this method?
        MSOMETHOD_(void, ReleaseText) (THIS_ void *pvDgs, MSOTXID txid,
                MSOHSP hsp, WCHAR *wz) PURE;

        // Return the RTF stored for this textbox
        MSOMETHOD_(BOOL, FGetRTF) (THIS_ void *pvDgs, MSOTXID txid, MSOHSP hsp,
                char **psz) PURE;

        // Return the RTF stored for this textbox
        // TODO davebu -- do we need this method?
        MSOMETHOD_(void, ReleaseRTF) (THIS_ void *pvDgs, MSOTXID txid, MSOHSP hsp,
                char *sz) PURE;

        // Set the RTF stored for this textbox, used mainly by clipboard operations
        MSOMETHOD_(BOOL, FSetRTF) (THIS_ void *pvDgs, MSOTXID txid, MSOHSP hsp,
                char *sz) PURE;

        // Save the textbox in the passed-in stream
        MSOMETHOD_(BOOL, FSaveTextbox) (THIS_ void *pvDgs, MSOTXID txid,
                MSOHSP hsp, MSOFBI fbi, MSOSVB *psvb) PURE;

        // Load the textbox from the passed-in stream
        MSOMETHOD_(BOOL, FLoadTextbox) (THIS_ void *pvDgs, MSOTXID *ptxid,
                MSOHSP hsp, MSOLDB *pldb) PURE;


        // ----- IMsoDrawingSite methods about OLE objects

        /*
        The host should fetch and return the presentation data of the specified
        OLE object, and its format, in the given parameters.
        */
        MSOMETHOD_(void, GetPictFromOid) (THIS_ void *pvDgs, MSOHSP hsp, MSOOID oid,
                HANDLE *phPict, LONG *pcf) PURE;

        /*
        The host should edit the specified picture object, whose presentation and
        presentation format are given, by creating a new OLE object. FALSE is
        returned if the edit is refused or no OLE object is created.
        */
        MSOMETHOD_(BOOL, FEditPictureHsp) (THIS_ void *pvDgs, MSOHSP hsp,
                HANDLE hPict, LONG cf) PURE;

        /*
        The host returns the main persistent storage of the specified OLE object.
        Office may clone it when copying drawing contents to the clipboard.
        */
        MSOMETHOD_(IStorage *, PistgFromOid) (THIS_ void *pvDgs, MSOHSP hsp,
                MSOOID oid) PURE;

        /*
        The host should create a new OLE object from the given storage, recognize
        that its attached HSP is the one given, and return its new host OID.
        Office will call this when pasting OLE shapes from the clipboard.
        */
        MSOMETHOD_(BOOL, FCreateOleObject) (THIS_ void *pvDgs, MSOHSP hsp,
                IStorage *pistg, MSOOID *poid, CLSID clsidSrc) PURE;

        /*
        The host should delete the specified OLE object. Office initiated the
        deletion of an OLE shape, the shape information of which should already
        have been registered as an undo record in the current command.
        */
        MSOMETHOD_(void, DeleteOid) (THIS_ void *pvDgs, MSOHSP hsp, MSOOID oid) PURE;

        /*
        The host needs to allocate and write out the object descriptor for
        the given ole shape. */
        MSOMETHOD_(HANDLE, HRenderObjDesc) (THIS_ void *pvDgs, MSOHSP hsp, MSOOID oid,
                BOOL fDragDrop) PURE;

        // ---- Constraints

        /*
        Does the host support rules? To support rules the FCreateAnchor and
        FLocateAnchor must never fail.
        */
        MSOMETHOD_(BOOL, FSupportsRules) (THIS_ void *pvDgs) PURE;

        /*
        Create the rule identified by the passed in id. Called when
        when a drawing is loaded rule that was added by the host.
        */
        MSOMETHOD_(BOOL, FCreateRule)   (THIS_ void *pvDgs, LONG rutType,
                interface IMsoRule **) PURE;


        // --- Misc

        /*
        Convert points from VBA units to host units.
        */
        MSOMETHOD (ConvertVBAToH) (THIS_ void *pvDgs, MSOPOINTF *pptvba, POINT *ppth,
                int cpt) PURE;

        /*
        Convert points from host units to VBA units.
        */
        MSOMETHOD (ConvertHToVBA) (THIS_ void *pvDgs, POINT *ppth, MSOPOINTF *pptvba,
                int cpt) PURE;

        /*
        Wrap the IDispatch pointer for VBA.  Returning TRUE means that you succeeded,
        and it is valid to simply return TRUE if you need not take any action.
        */
        MSOMETHOD_(BOOL, FWrapShapeForVBA) (THIS_ void *pvDgs, IDispatch **ppshape,
                MSODRWSK drwsk) PURE;

        /* Get the RGB color from a host scheme index. Return FALSE if color scheme is not supported
        or index is out-of-range */
        MSOMETHOD_(BOOL, FGetColorScheme) (THIS_ void *pvDgs, int index, long *pRGB) PURE;

        /* Return an IBindStatusCallback.  This interface is called when a blip
                is loaded asynchronously.  The arguments identify the shape to which
                the blip is attached and the anchor for the shape.  It is not safe to
                make changes to the drawing in this callback (e.g. deleting the shape
                or changing its properties would be very bad).  If the result is
                non-NULL callbacks will be made as a result of the hyperlink
                resolution - see msohl.h.  When the blip (picture) has been loaded
                a call is made to OnDataAvailable with the blip as the IUnknown.
                Note that the shape may not yet have an anchor - in this case the
                *ppvAnchor will be NULL. */
        MSOMETHOD_(IBindStatusCallback*, PiBindStatusCallback) (THIS_ void *pvDgs,
                void **ppvAnchor, MSOHSP hsp) PURE;

        /* Return the base for relative file names in this drawing, returns FALSE
                if there is no base name (all names must be absolute).  The interface
                is passed a buffer of length cwch (NOT including any extra for the
                terminating zero - there are only cwch WCHAR spaces in buffer!) */
        MSOMETHOD_(BOOL, FGetBaseName) (THIS_ void *pvDgs, WCHAR *pwzBase,
                int cwch) PURE;

        /* TODO asheshb: REMOVE this method.  it is never called!
                Clients should implement the method of the same name in
                IMsoDrawingViewSite */
        MSOMETHOD_(void, AfterCreateBdp) (THIS_ void *pvDgs, MSOBDRAWPARAM *pbdp,
                const IMsoBlip *pib, MSOHSP hsp, MSOPID pid) PURE;
        /* FCloneAnchorForReconcile is used to clone an pvAnchor for multiuser
                documents.  Currently only used by Word. */
        MSOMETHOD_(BOOL, FCloneAnchorForReconcile) (THIS_ void *pvDgs, MSOHSP hsp,
                MSOHSP hspClone, void *pvAnchorUndo, void **ppvAnchorUndoClone) PURE;

        /* FBackgroundShapeDrawing is used to ask the host which Drawing is to be
                used for the background shape. This is primarily needed for PowerPoint.*/
        MSOMETHOD_(BOOL, FBackgroundShapeDrawing) (THIS_ void *pvDgs,
                IMsoDrawing **ppidgBackground) PURE;

        // --- Clipboard

        /*
        Draws the shapes in the given drawing into a metafile and returns a handle to
        the created metafile.
        */
        MSOMETHOD_(HANDLE, HDrawShapesInMetafile) (THIS_ void *pvDgs,
                interface IMsoDrawing *pidg, RECT *prchBounds, BOOL fEmf) PURE;

        /* FGetTitleShape is used to ask the host for the title shape used
                in "Fade To Title" fill for PP.  The title shape will determine the
                focus for the fill */
        MSOMETHOD_(BOOL, FGetTitleShape) (THIS_ void *pvDgs, MSOHSP *phspTitle) PURE;

        /* Used to get the class id of an OLE object which is used to decide if
                we should bring up the picture toolbar or not */
        MSOMETHOD_(void, GetClsidFromHsp) (THIS_ void *pvDgs, MSOHSP hsp, MSOOID oid,
                CLSID *pclsid) PURE;

        MSOMETHOD_(HANDLE, FWriteOutHTML) (THIS_ void *pvDgs) PURE;
        /* Create an OLE object from data collected form HTML
                TODO barnwl (hailiu): IMsoDrawingImportSite may be a better place
                for this method*/
        MSOMETHOD_(BOOL, FCreateOleObjectFromHtml)(THIS_
                LPVOID pvClientDgs,
                MSOHSP hsp,
                REFCLSID rclsid,
                LPUNKNOWN punkData,
                MSOOID *poid,
                LPVOID pvClientHtml) PURE;

        /* Draws shapes in a supplied IMsoBitmapSurface (T-MARKW) (this just
                gets called from the corresponding drawing method.) */
        MSOMETHOD_(BOOL, HDrawShapesInBitmapSurface)(THIS_
                void *pvDgs,
                IMsoDrawing *pidg,
                interface IMsoBitmapSurface *pbms) PURE;

        /* Writes the XML for client data of a shape. This method is called
                before the SHAPE end tag. */
        MSOMETHOD_(BOOL, FWriteHspClientDataXML)(THIS_ void *pvDgs,
                MSOXMLW *pxmlw,
                void **ppvClient,
                MSOHSP hsp) PURE;

        /* This function is called when writing a Shape in XML to write the
                data hanging off its pvAnchor.  This should return TRUE iff it
                succeeds. If fAnchorIsRch (in MSODGSI) is TRUE we may just save
                the RCH as flat data without calling you. */
        MSOMETHOD_(BOOL, FWriteAnchorXML) (THIS_ void *pvDgs, void **ppvAnchor,
                MSOHSP hsp, MSOXMLW *pxmlw) PURE;

        // Write the attached text HTML.
        MSOMETHOD_(BOOL, FWriteTextboxHTML) (THIS_ void *pvDgs, MSOTXID txid,
                MSOHSP hsp, MSOXMLW *pxmlw) PURE;

        /* Queries the client data of a shape for XML content.*/
        MSOMETHOD_(BOOL, FHspClientDataHasXML)(THIS_ void *pvDgs,
                MSOXMLW *pxmlw,
                void **ppvClient,
                MSOHSP hsp) PURE;

        /* Allows the client to write app specific XML attributes for a shape. */
        MSOMETHOD_(BOOL, FWriteHspClientAttrXML)(THIS_ void *pvDgs,
                MSOXMLW *pxmlw,
                void **ppvClient,
                MSOHSP hsp) PURE;

        /* Get the color from a host scheme index. Return FALSE if color scheme is not supported
                or index is out-of-range. The ink color string is allocated on the heap by the host
                app. Office takes the ownership of this string from the host app if the function
                returns TRUE. */
        MSOMETHOD_(BOOL, FGetColorSchemeExt) (THIS_ void *pvDgs, int index,
                MSOCOLOREXT *pclr) PURE;
};


/****************************************************************************
        Drawing Group

        The Drawing Group (DGG) is the largest-scope Office Drawing object.
        It's supposed to parallel the host's notion of a document.  It's what
        a user saves.  It's the scope at which styles and master shapes live.
        Really it's not very big; it's pretty much just the head of a linked
        list of Drawings.  It's the one Office Drawing object the host will
        probably generate with a global API.  After you've made a Drawing Group
        you ask it to make you Drawings.
*****************************************************************************/

/* Split Menu ID - Index into the split menu colors in the MSODGGSI

        HEY!!! These are effectively in the file format, since we save and load
        split menu colors in this order.  Adding a new SMID to the end, though, is
        no problem (older builds will just discard its color).

        Office developers - If you add a SMID, be sure to provide a default color
        in vmpsmidclrInit. */
enum MSOSMID
        {
        msosmidMin=0,
        msosmidFillColor=msosmidMin,
        msosmidLineColor,
        msosmidShadowColor,
        msosmid3DColor,
        msosmidMax, msosmidNil = msosmidMax
        };

/* Ink Style ID - Index into ink style pen */
enum MSOISID
        {
        msoisidMin=0,
        msoisidInkStyle1 = msoisidMin,
        msoisidInkStyle2,
        msoisidInkStyle3,
        msoisidInkStyle4,
        msoisidInkStyle5,
        msoisidInkStyle6,
        msoisidInkStyle7,
        msoisidInkStyle8,
        msoisidInkStyle9,
        msoisidMax, msoisidNil = msoisidMax
        };

/* Ink Annotation Style ID - Index into ink annotation style pens */
enum MSOIASID
        {
        msoiasidMin=0,
        msoiasidInkAnnotationStyle1 = msoiasidMin,
        msoiasidInkAnnotationStyle2,
        msoiasidInkAnnotationStyle3,
        msoiasidInkAnnotationStyle4,
        msoiasidInkAnnotationStyle5,
        msoiasidInkAnnotationStyle6,
        msoiasidInkAnnotationStyle7,
        msoiasidInkAnnotationStyle8,
        msoiasidInkAnnotationStyle9,
        msoiasidMax, msoiasidNil = msoiasidMax
        };

// Define ink pen
typedef struct _MSOINKPENENTRY
        {
        MSOCLR             rgb;            // pen color
        int                rasOp;          // raster mode
        int                penTip;         // pen tip - ball or rectangel
        int                transparency;   // transparency
        float              penHeight;      // pen height
        float              penWidth;       // pen width
        BOOL               fPressure;      // pressure sensitive
        } MSOINKPENENTRY;

// fetch the hard-coded factory default pen style (for annotations or drawing and writing)
MSOAPI_(void) MsoInitDefaultInkPen(BOOL fAnnotation, MSOINKPENENTRY *pPenStyle);

typedef struct _MSOINKPEN
        {
        MSODGCID         dgcid;            // dgcid for specific pen
        MSOINKPENENTRY   info;			    // pen attributes
        } MSOINKPEN; 

// Define ink pen menu
typedef struct _MSONAMEDINKPEN
        {
        MSOINKPEN   pen;           // pen info
        WCHAR*      penTooltip;    // pen name
        } MSONAMEDINKPEN;

typedef struct _MSONAMEDPENSET
        {
        int inumPens;
        MSONAMEDINKPEN *pPenSet;
        } MSONAMEDPENSET;

MsoDEBDeclare(dgg, msodgeDggsFirst, msodgeDggsMax);

typedef struct _MSODGGSI
        {
        interface IMsoDrawingGroupSite *pidggs;
        void *pvDggs;
        HMSOINST hmsoinst;
        int cMaxMRUColors;
        ULONG cbMinAbortYieldSave; // minimum number of bytes to save before
                                                                                // calling the abort/yield callbacks
        MSOCOLOREXT mpsmidclrExt[msosmidMax];

        int   cxvPerInch;
        int   cyvPerInch;
                /* The "target DPI" of the expected output device.  This may not
                        necessarily correspond to the current screen dpi - it is a
                        document property which indicates the DPI of the output system
                        for which the document is designed.  This is used when interpreting
                        "pixel" values. */

        int   cxvTarget;
        int   cyvTarget;
                /* The total x and y pixel sizes of the target device.  Used when it
                        is necessary to format a document for a particular screen size. */

        MSOSPID spidHiLim;
                /* Biggest spid a shape in this DGG can have. For now we will rounded this
                   up to the multiple of 1024 minus 1. */

        ULONG rgGrfDgmtCaps[msodgmtMax]; // canvas/diagram extensibility caps

        MSODGID dgidHiLim;
                /* Largest dgid we have have in this DGG. Don't change or set this in the app
                unless you know what you are doing. */

        union
                {
                struct
                        {
                        ULONG fWantsEvents : 1;
                        ULONG fExportXMLDggWithDg : 1;
                                /* If true, when exporting DG level info out in XML for a DG,
                                   we will also export DGG level info together with it. */
                        ULONG fInline : 1; // this drawing group is for Word's inline layer
                        ULONG fReuseSpid : 1; // Reuse the spid when it's full.
                        ULONG fNoTempSpid : 1; // Do *NOT* use temp spid.

                        ULONG fNoRoteteBy90 : 1;
                                // fFlag = TRUE --> no need to support backward compatibility with
                                //   OfficeArt-2000 by rotating to closest 90 degree.

                        ULONG fAllowOleRotate : 1;      // Allow the rotation of OLE objects (used by Publisher), closely related to fNoRoteteBy90
                        ULONG fActivateTextDgmNode: 1;
                                /* PPT will turn an empty textbox into a shape after roundtrip through
                                HTML. This is not good for Diagrams. We will activate text if this bit
                                is set to TRUE. */
                        ULONG fPersistDgid : 1;
                           /* We will persist dgid is this is TRUE. */

                        ULONG fUndoInkParsing : 1; // dWxp
                           /* Find out if we're in Word before recording an undo for parsing shapes */
                        ULONG fFixInkSurfaceOverlap : 1; // dWxp
                           /* If true, fix ink surfaces that overlap */
                        ULONG fInkShapesDifferentPerView : 1; // dWxp
                           /* If true, ink shapes are different on different drawing views (Word).
                              Excel and Powerpoint are different per drawing. */
                        ULONG grfUnused : 20;
                        };
                ULONG grf;
                };

        MSOINKPEN      minkpen;
        MSOINKPEN      minkannotpen;
        MSONAMEDPENSET *mpsmidinkpen;
        MSONAMEDPENSET *mpsmidinkannotpen;

        MsoDEBDefine(dgg);

#if DEBUG
        LONG lVerifyInitDggsi; // Used to make sure people use MsoInitDggsi
#endif // DEBUG
        } MSODGGSI;
#define cbMSODGGSI (sizeof(MSODGGSI))
MSOAPI_(void) MsoInitDggsi(MSODGGSI *pdggsi);

#define msodggsicxyvPerInchDefault 96

/* MsoFCreateDrawingGroup makes a new Drawing Group. */
MSOAPI_(BOOL) MsoFCreateDrawingGroup(interface IMsoDrawingGroup **ppidgg,
        MSODGGSI *pdggsi);

/* IMsoDrawingGroup (idgg) */
#undef  INTERFACE
#define INTERFACE IMsoDrawingGroup
DECLARE_INTERFACE_(IMsoDrawingGroup, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingGroup methods (standard stuff)

        /* FDebugMessage method */
        MSODEBUGMETHOD

        /* Call Free to make an entire DrawingGroup go away.  This includes all
                the DGs, DGVGs, DGVs, and DGSLs that might be hanging off of it. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* PdggsiBeginUpdate returns a pointer to an MSODGGSI.  This will
                always succeed (never return NULL).  The caller may examine and
                (unless they passed in fReadOnly == TRUE) change the fields in
                the MSODGGSI.  The caller must in any case call EndDggsiUpdate
                when they're finished.  See comments by the declaration of
                MSODGGSI. */
        MSOMETHOD_(MSODGGSI *, PdggsiBeginUpdate) (THIS_ BOOL fReadOnly) PURE;

        /* After you call PdggsiBeginUpdate you have to call EndDggsiUpdate,
                passing back in the pointer you got from PdggsiBeginUpdate and
                TRUE for fChanged if you changed any fields in it.  See comment
                by PdggsiBeginUpdate. */
        MSOMETHOD_(void, EndDggsiUpdate) (THIS_ struct _MSODGGSI *pdggsi,
                BOOL fChanged) PURE;


        // ----- IMsoDrawingGroup methods (low-level editing)

        /* You may use FClone to make a new DGG with identical properties
                to this one except that it won't contain any DGs.  Because
                they each need their own DGSI you have to clone them over
                one at a time. */
        MSOMETHOD_(BOOL, FClone) (THIS_ MSODGGSI *pdggsi,
                interface IMsoDrawingGroup **ppidggClone) PURE;

        /* Call FFreePurgedStuff to get rid of purged stuff in the Drawings
                in this Drawing Group.  The reason for this method is that since
                Shapes in Drawings can reference one another we need to make a
                full linear pass through all the Shapes cleaning up referneces
                to Shapes we're about to dispose of.  We return TRUE if we finished,
                FALSE if we were interrupted.  At the moment we don't interrupt. */
        MSOMETHOD_(BOOL, FFreePurgedStuff) (THIS) PURE;


#if PR207308  // no'ed in O9
        /* FCloneFileBlocks clones the unknown file blocks in this DGG and
                stores the copy on pidggDest.  This function duplicates behavior
                in DGG::FClone; if you need unknown file blocks you must call
                either that function or this one. */
        MSOMETHOD_(BOOL, FCloneFileBlocks) (THIS_
                interface IMsoDrawingGroup *pidggDest) PURE;
#endif

        // ----- IMsoDrawingGroup methods (saving and loading)

        /* A call to InitSaveBlock to initialize an MSOSVB is _required_ before
                saving, including undo record saves.  If you pass in a non-NULL
                pistmImmediate, we'll set fSaveImmediate for you; if you pass in a
                non-NULL pistmDelay, we'll set fSaveDelay and fUseDelay. */
        MSOMETHOD_(void, InitSaveBlock) (THIS_ MSOSVB *psvb,
                IStream *pistmImmediate, IStream *pistmDelay) PURE;

        /* A call to InitLoadBlock to initialize an MSOLDB is _required_ before
                loading, including undo record loads.  If you pass in a non-NULL
                pistmDelay, we'll set fUseDelay for you. */
        MSOMETHOD_(void, InitLoadBlock) (THIS_ MSOLDB *pldb,
                IStream *pistmImmediate, IStream *pistmDelay) PURE;

        // TODO peteren: Comment
        MSOMETHOD_(BOOL, FSave) (THIS_ MSOSVB *psvb, IMsoDrawing *pidg) PURE;

        // TODO peteren: Comment
        MSOMETHOD_(BOOL, FLoad) (THIS_ MSOLDB *pldb, IMsoDrawing *pidg) PURE;

        /* Call FImportLink to update an exported link that was changed. */
        MSOMETHOD_(BOOL, FImportLink) (THIS_ struct _hlinkprop *phlinkprop) PURE;

        /* FFindShape finds the Shape in this Drawing Group with the
                specified MSOSPID.  It returns TRUE and fills out *phsp and *ppidg
                if it finds one and returns FALSE if it doesn't.  This is O(1) if
                we have a hash-table constructed (at load time) and O(n) thereafter. */
        MSOMETHOD_(BOOL, FFindShape) (THIS_ MSOHSP *phsp, IMsoDrawing **ppidg,
                MSOSPID spid) PURE;

        /* FFindDrawing finds the Drawing in this Drawing Group with the
                specified MSODGID.  It returns TRUE and fills out *ppidg if it finds
                one and returns FALSE if it doesn't. */
        MSOMETHOD_(BOOL, FFindDrawing) (THIS_ IMsoDrawing **ppidg,
                MSODGID dgid) PURE;

        // ----- IMsoDrawingGroup methods (other)

        /* FCreateDrawing generates a new Drawing in this Drawing Group.  It
                returns TRUE iff it succeeds.  You have to pass in an MSODGSI full
                of a description of the DGS.  If we succeed we'll have copied the
                MSODGSI into the DG; if we fail we won't have done anything with
                your MSODGSI. */
        MSOMETHOD_(BOOL, FCreateDrawing) (THIS_ IMsoDrawing **ppidg,
                MSODGSI *pdgsi) PURE;

        /* Check or set the dirtiness of the DGG */
        MSOMETHOD_(BOOL, FNeedsSave) (THIS) PURE;
        MSOMETHOD_(void, SetFNeedsSave) (THIS_ BOOL fNeedsSave) PURE;

#if BOGUS
        /*
        Return the HMSOINST Office instance reference for the drawing group.
        */
        MSOMETHOD_(HMSOINST, HinstGet) (THIS) PURE;
#endif // BOGUS

        /*
        This is another shape enumerator usable by the host. Given a generic
        MSOHSP, we return a different one. Passing in a pointer to NULL means
        start at the beginning. A return value of FALSE means the end of the list.
        Unlike DG::FGetRelatedShape, this may be called during load, and can only
        be called during load. The pidg parameter may be used to return only those
        shapes in the given DG, and may be NULL to go through all in the DGG.
        Note this goes through each shape only once, but in an undefined order.
        */
        MSOMETHOD_(BOOL, FGetRelatedShapeHash) (THIS_ MSOHSP *phsp,
                IMsoDrawing *pidg) PURE;

        /* This method is used to obtain a DG* from a DispShapes. */
        MSOMETHOD_(IMsoDrawing *, PidgFromPidisp) (THIS_ IDispatch *pidisp) PURE;

        // ----- Color MRU methods

        /* Call FAddColorToMRU (passing in a cr) to add a color to the
                MRU the DGG keeps track of for you.  If picr is non-NULL, returns
                the index of the added color in it.  Returns TRUE on success. */
        MSOMETHOD_(BOOL, FAddColorToMRU) (THIS_ MSOCR cr, int *picr) PURE;

        /* Call CColorsInMRU to see how many colors there are in the MRU. */
        MSOMETHOD_(int, CColorsInMRU) (THIS) PURE;

        /* Call FGetColorFromMRU to retrieve the i'th color from the MRU.
                If you pass in an invalid index we'll return FALSE. */
        MSOMETHOD_(BOOL, FGetColorFromMRU) (THIS_ MSOCR *pcr, int i) PURE;

        /* ClrOfSplitMenu returns the current color of the split menu
                corresponding to the given dgcid. Returns msoclrNil if
                there's no drawing group present. */
        MSOMETHOD_(MSOCLR, ClrOfSplitMenu) (THIS_ MSODGCID dgcid) PURE;

        /* SetSplitMenuClr sets the color of the split menu
                corresponding to the given dgcid. Most hosts should never need
                to call this method.  Note that this only does anything if
                there is a drawing group present.*/
        MSOMETHOD_(void, SetSplitMenuClr) (THIS_ MSODGCID dgcid, MSOCLR clr) PURE;

        //----- Default Access Methods

        /*
        Defaults for new objects are kept in the drawing group. They are saved
        with the drawing group. Like all properties, defaults are stored as delta's
        from a constant built-in set of default property values. In addition, to
        these methods, defaults can be altered by commands other property setting
        methods.
        */

        /* Fill the passed in RGSPP with the defaults for this drawing group.
                The flags of grffpp flags used in apply. See FFetchProperties for
                more details. */
        MSOMETHOD_(BOOL, FFetchDefaults) (THIS_ MSOSPP *rgspp, int cspp,
                ULONG grffspp) PURE;

        /* Set the default for this drawing. No undo records are
           generated. See FApplyProperties for more details on the
                parameters. */
        MSOMETHOD_(BOOL, FSetDefaults) (THIS_ MSOSPP *rgspp, int cspp,
                ULONG grfaspp) PURE;

        /* Convenient function to apply the defaults from the passed in DGG.
           If msofasppResetAll flag is passed in, the current defaults are
                thrown away. */
        MSOMETHOD_(BOOL, FApplyDefaults) (THIS_ interface IMsoDrawingGroup* pidgg,
                ULONG grfaspp) PURE;

        //--- Font Enumerations

        /*      Enumerate the fonts in the drawing group. These fonts include the
                those used for text effects and those found in pictures. The callback
                is called with the fonts Warning: demand loads the pictures
                in the drawing. */
        MSOMETHOD_(BOOL, FEnumFonts) (THIS_ MSOPFNELF pfnelf, void *pvParam, ULONG grfelf) PURE;

        // --- Write/Read XML/HTML
        MSOMETHOD_(BOOL, FWriteXML) (THIS_ MSOXMLW *pxmlw, IMsoDrawing *pidg,
                int hetnUser) PURE;
        MSOMETHOD_(void, EndHTMLImport) (THIS_ BOOL fCfhtml, interface IMsoReducedHTMLImport *pirhi) PURE;

        // --- DPI enquiry
        MSOMETHOD_(void, DrawingDPI)(THIS_ POINT *pptDpi) CONST_METHOD PURE;
                /* Return the DPI of all the drawings in the group, x and y DPI values
                        may be different, so a POINT is filled in.  This is really just
                        retriving the DGGSI information. */

        // --- Update Blips on shapes
        MSOMETHOD_(BOOL, FUpdateBlips) (THIS_ BOOL fEnvelope) PURE;
                /* Enumerate all Drawings in the Drawing Group and call FUpdateBlips
                        on each.  if fEnvelope is set, call FUpdateBlipForEnvelope
                        instead... */

        /* Call FAddColorExtToMRU (passing in a MSOCOLOREXT) to add a color to the
                MRU the DGG keeps track of for you.  If picr is non-NULL, returns
                the index of the added color in it. The ink color string is allocated
                on the heap by the host app. Office takes the ownership of this string
                from the host app regardless of success. Returns TRUE on success. */
        MSOMETHOD_(BOOL, FAddColorExtToMRU) (THIS_ MSOCOLOREXT *pclrExt, int *picr) PURE;

        /* Call FGetColorExtFromMRU to retrieve the i'th color from the MRU.
                If you pass in an invalid index we'll return FALSE. The ink color string
                is allocated on the heap by the office. The host app takes the ownership
                of this string if the function returns TRUE. */
        MSOMETHOD_(BOOL, FGetColorExtFromMRU) (THIS_ MSOCOLOREXT *pclrExt, int i) PURE;

        /* SetSplitMenuClr sets the color of the split menu
                corresponding to the given dgcid. Most hosts should never need
                to call this method.  Note that this only does anything if
                there is a drawing group present. The ink color string is allocated
                on the heap by the host app. Office takes the ownership of this string
                from the host app. */
        MSOMETHOD_(void, SetSplitMenuClrExt) (THIS_ MSODGCID dgcid, MSOCOLOREXT *pclrExt) PURE;

        /* FGetClrExtOfSplitMenu returns the current color of the split menu
                corresponding to the given dgcid. Returns msoclrNil if there's no
                drawing group present. The ink color string is allocated on the heap
                by the office. The host app takes the ownership of this string if
                the function returns TRUE.*/
        MSOMETHOD_(BOOL, FGetClrExtOfSplitMenu) (THIS_ MSOCOLOREXT *pclrExt, MSODGCID dgcid) PURE;

        /* BulletProof the DrawingGroup */
        MSOMETHOD_(BPSC, BpscBulletProof) (THIS_ MSOBPCB *pmsobpcb) PURE;

        /* Call ClearColorMRU to empty out the color MRU. */
        MSOMETHOD_(void, ClearColorMRU) (THIS) PURE;

        /* This method updates the flattened colors of all shapes in all
                the Drawings of this Drawing group. */
        MSOMETHOD_(BOOL, FUpdateColors) (THIS) PURE;

        // Ink Pen methods.
        MSOMETHOD_(BOOL, FGetInkNamedPen) (THIS_  MSODGCID dgcid, MSONAMEDINKPEN** pInkPen) PURE;
        MSOMETHOD_(void, GetInkNamedPens) (THIS_  BOOL fInkAnnot, MSONAMEDPENSET** pInkPens) PURE;
        MSOMETHOD_(void, SetInkNamedPens) (THIS_  BOOL fInkAnnot, MSONAMEDPENSET* pInkPens) PURE;
        MSOMETHOD_(BOOL, FGetInkPen) (THIS_  BOOL fInkAnnot, MSOINKPEN* pInkPen) PURE;
        MSOMETHOD_(void, SetInkPen) (THIS_  BOOL fInkAnnot, MSOINKPEN inkPen) PURE;
        MSOMETHOD_(void, ExitInkingModes) (THIS) PURE;
};
MSOAPI_(BOOL) MsoFValidDgg(IMsoDrawingGroup *pidgg);

/* MSODGGEB -- Drawing Group Event Block */
typedef struct MSODGGEB
        {
        /* These first fields are filled out for EVERY Drawing Event. */
        MSODGE dge;
        BOOL fResult;
        interface IMsoDrawingGroup *pidgg;
        MsoDEBPDefine(dgg);

        WPARAM wParam;
        LPARAM lParam;
        union
                {
                struct // for msodgeDoWritePidDefault
                        {
                        WORD opid;
                        ULONG *pop;
                        BOOL fNeedToWrite;
                        } dggexDoWritePidDefault;
                struct
                        {
                        MSODGCID dgcid;
                        MSOCLR   *pClrRGB;
                        } dggexInkPenColor; 
                };
        } MSODGGEB;
#define cbMSODGGEB (sizeof(MSODGGEB))

/* IMsoDrawingGroupSite (idggs).  This is the actual interface a
        Drawing Group Site must implement. */
#undef  INTERFACE
#define INTERFACE IMsoDrawingGroupSite
DECLARE_INTERFACE_(IMsoDrawingGroupSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingGroupSite methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* The Drawing will call DGGS::OnEvent whenever adrawing
                event (MSODGE) occurs.  See the definition of MSODGE for a more
                complete explanation of events.  This function will only be
                called at all if you set msodggsi.fWantsEvents to TRUE.
                The following events are passed to DGGS::OnEvent:
                        msodgeRequestAbortSave
                        msodgeRequestYieldSave
                        msodgeRequestAbortAfterYieldSave
                See comments by definition of MSODGE. */
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDggs, MSODGGEB *pdggeb) PURE;

        // ----- IMsoDrawingGroupSite methods (save stuff)

        /* When loading a bunch of Drawings that were saved along with a
                Drawing Group we'll need to call DGGS::FGetDrawingFromLabel to get
                the Drawings to load.  So when we save a bunch of Drawings with
                a Drawing Group  we call you just before we save each drawing to
                let you save some kind of label from which you'll (later, at load time)
                be able to pick the right Drawing.      We pass you an MSOSVB,
                and a pidg, pidgs, and pvDgs to identify the Drawing.  When saving,
                use psvb->pistmImmediate.  You should only actually save if
                psvb.fSaveImmediate is TRUE, and in all cases you should _add_ to
                psvb->dfoImmediate the number of bytes you saved (or would have saved).
                Return TRUE on success. */
        MSOMETHOD_(BOOL, FSaveDrawingLabel) (THIS_ void *pvDggs, MSOSVB *psvb,
                IMsoDrawing *pidg, IMsoDrawingSite *pidgs, void *pvDgs) PURE;

        /* When loading a DGG that had its DGs saved with it we'll need to call
                back into DGGS::FGetDrawingFromLabel to get IMsoDrawing *'s in which
                to load the DGs.  psvb->pistmImmediate will be positioned at whatever
                you saved under DGGS::FSaveDrawingLabel.  Return TRUE iff you succeed
                and _add_ to psvb->dfoLoaded the number of bytes you read. */
        MSOMETHOD_(BOOL, FGetDrawingFromLabel) (THIS_ void *pvDggs, MSOLDB *pldb,
                IMsoDrawing **ppidg) PURE;

        /* If a DGG is loaded incrementally, this method is called to get
           the delay stream when an object is loaded. It must be the same
           delay stream as the one passed into DGG::FLoad. ReleaseDelayStream
           will be called when the DGG is done with the stream. */
        MSOMETHOD_(BOOL, FGetDelayStream) (THIS_ void *pvDggs,
                IStream **ppistm) PURE;

        /* If a DGG is loaded incrementally, this method is called to release
           the delay stream when an object is loaded. The current position
           of the stream may have changed since the call to FGetDelaySteam. */
        MSOMETHOD_(void, ReleaseDelayStream) (THIS_ void *pvDggs,
                IStream *pistm) PURE;

        /* If the client stored a client type of IMsoBlip into the drawing
                group, then this method is called to create a blank blip. Normally,
                called during loads. This method will not be called for MSO blip
                types. */
        MSOMETHOD_(BOOL, FCreateBlip) (THIS_ void *pvDggs,
                IMsoBlip** ppib, MSOBLIPTYPE bt) PURE;

        /* Returns the IMsoEnvelope attached with this document.  If we aren't loading
                from an email message, this API will return NULL.  This API will call AddRef
                on the Pienv, and the caller must call Release when it is finished. */
        MSOMETHOD_(interface IMsoReducedHTMLImport *, PirhiGet)(THIS_ void *pvDggs) PURE;

        /* Used to retrieve the full path name of a relatively linked file.  pcchFull
                should be set to wzFull buffer length.  On return, it will indicate the length
                of the buffer actually used.  Set wzFull to NULL to retrieve the length of
                the buffer necessary.  Returns FALSE if the link cannot be resolved.
                If wzRel is absolute, it is merely copied into wzFull.  */
        MSOMETHOD_(BOOL, FGetFullPath) (THIS_ void *pvDggs, const WCHAR *wzRel,
                WCHAR *wzFull, int *pcchFull) PURE;

        /* Used to retrieve the codepage context of this document for server communication
                during delayed download of image files. */
        MSOMETHOD_(DWORD, DwGetCodePageContext) (THIS_ void *pvDggs) PURE;
};


/****************************************************************************
        Exported functions
*****************************************************************************/

/****************************************************************************
        MSODOT

        "DOT" is really an old name for some cases of a DRGHN, we store these
        numbers in the low word of an MSODRGH.

        A DOT is one of the eight "selection dots" that are often displayed and
        hit-tested around the boundary of rectangular objects. Or it could be one
        of the (up to) nine adjust handles around the boundary of a shape.

        These are loosely based on the DRGH notion from Word, although DOTs
        are quite a bit simpler.  They're just an enumeration; their only
        strangeness is that to make things work out well they're in a careful
        order and numbered 0 through 8 with no 4.  This is to make their
        value correspond in some useful ways to their position around
        the object's bounding rectangle.  If there were a 4 it would be right
        in the middle of the rectangle.

        BC = Bottom Center, CR = Center Right, TL = Top Left, etc.
        AH0 = Adjust Handle 0, AH1 = Adjust Handle 1, etc.
****************************************************************************/

// CAUTION:: (sreeramn)Keep the msodot* in the same order. They
// have a one-one correspondence with the values assigned to
// msocrsCrop* and also in MSOSP::GetHandles().
#define msodotTL                0
#define msodotTC                1
#define msodotTR                2
#define msodotCL                3
#define msodotCC                4 // this DOT never really exists
#define msodotCR                5
#define msodotBL                6
#define msodotBC                7
#define msodotBR                8
#define msodotMaxPerimeter      9

// TODO (SreeramN): Whenever you change this, go ahead and change the
// stupid drghIndex-msodotMaxPerimeter in DGV::AdjustShape and God knows
// where else.

#define msodotAH1               9
#define msodotAH2               10
#define msodotAH3               11
#define msodotAH4               12
#define msodotMaxAH             13

// note: these guys need to stay in this order!
#define msodotCropTL            13
#define msodotCropTR            14
#define msodotCropBR            15
#define msodotCropBL            16
#define msodotCropTC            17
#define msodotCropBC            18
#define msodotCropCL            19
#define msodotCropCR            20
#define msodotCropMax           21

#define msodotROT               21   // Rotate Knob


#define msodotMax               22
#define msodotNil               -1

// Define fake 'dots' for the msodrghkMove move handle
#define msodotMoveShape         0
#define msodotMoveFuzz          1


/******************************************************************************
        KINDS OF DRAG HANDLES.
****************************************************************** sreeramn **/

/* MSODRGH -- Drag Handle */
typedef ULONG MSODRGH;

/* MSODRGK -- Drag Handle Kind */
typedef ULONG MSODRGHK;
#define msodrghkNil            ((ULONG)0)   // Means that the nothing has been hit.
#define msodrghkSize           ((ULONG)1)   // Selection dot has been hit.
#define msodrghkVertex         ((ULONG)2)   // Vertex has been hit.
#define msodrghkAdjust         ((ULONG)3)   // Adjust Handle has been hit.
#define msodrghkMove           ((ULONG)4)   // Interior of the shape has been hit.
#define msodrghkControl        ((ULONG)5)   // Control point in reshape.
#define msodrghkSegment        ((ULONG)6)   // Segment of the shape.
#define msodrghkCrop           ((ULONG)7)   // Crop the shape
#define msodrghkRotate         ((ULONG)8)   // Rotate the shape
#define msodrghkButton         ((ULONG)9)   // Drag of a button (hotspot).
#define msodrghkText           ((ULONG)10)  // Textbox has been hit.
#define msodrghkConnector      ((ULONG)11)  // One end of a connector
#define msodrghkLineEnd        ((ULONG)12)  // Endpoint of a line.
#define msodrghkSelect         ((ULONG)13)  // Selects the shape only.
#define msodrghkJumpBack       ((ULONG)14)   //Jump buttons for Publisher
#define msodrghkJumpForth      ((ULONG)15)   //Jump buttons for Publisher
#define msodrghkTextOverflow   ((ULONG)16)   //text overflow indicator for Publisher
#define msodrghkGroup          ((ULONG)17)   //Group button for Publisher
#define msodrghkUngroup        ((ULONG)18)   //Ungroup button for Publisher
#define msodrghkSmartObject    ((ULONG)19)   //Smart Object button for Publisher
#define msodrghkGroupTL        ((ULONG)20)   //Group button in Top-Left position for Publisher
#define msodrghkWebComponent   ((ULONG)21)   //Web Component button for Publisher
#define msodrghkMax            ((ULONG)22)
        /* TODO peteren: These should be an enum */

/* MSODRGHN -- Drag Handle Number */
typedef ULONG MSODRGHN;

/* MsoDrghFromDrghkDrghn builds a DRGH from a DRGHK and a DRGHN. */
__inline MSODRGH MsoDrghFromDrghkDrghn(MSODRGHK drghk, MSODRGHN drghn)
        {return (MSODRGH)((((ULONG)drghk) << 24) | (ULONG)drghn);}

/* MsoDrghkOfDrgh pulls the DRGHK out of a DRGH. */
__inline MSODRGHK MsoDrghkOfDrgh(MSODRGH drgh)
        {return (MSODRGHK)(((ULONG)(drgh)) >> 24);}

/* MsoDrghnOfDrgh pulls the DRGHN out of a DRGH. */
__inline MSODRGHN MsoDrghnOfDrgh(MSODRGH drgh)
        {return (MSODRGHN)(((ULONG)(drgh)) & 0x00ffff);}

/* MsoDrghkDrghnOfDrgh pulls the DRGHK and DRGHN out of a DRGH. */
__inline void MsoDrghkDrghnOfDrgh(MSODRGH drgh, ULONG *pdrghk, ULONG *pdrghn)
        {
        *pdrghk = MsoDrghkOfDrgh(drgh);
        *pdrghn = MsoDrghnOfDrgh(drgh);
        }

#define msodrghNil      (MsoDrghFromDrghkDrghn(msodrghkNil, 0)) // == 0
#define msodrghMove     (MsoDrghFromDrghkDrghn(msodrghkMove, msodotMoveShape))
#define msodrghMoveFuzz (MsoDrghFromDrghkDrghn(msodrghkMove, msodotMoveFuzz))
#define msodrghButton   (MsoDrghFromDrghkDrghn(msodrghkButton, 0))
#define msodrghText     (MsoDrghFromDrghkDrghn(msodrghkText, 0))
#define msodrghTextAll  (MsoDrghFromDrghkDrghn(msodrghkText, 1))
#define msodrghSelect   (MsoDrghFromDrghkDrghn(msodrghkSelect, 0))


// MSOHTK: Hit text kinds: this is returned in MSODMHD.htk; it differentiates
// between true hits on a display element, and "fake" hits on the transparent
// region, resulting from the Pdgsi()->fNoFillHitTest flag.
typedef ULONG MSOHTK;
#define msohtkFail                 ((MSOHTK) 0x00000000)     // Hit test didn't hit anything
#define msohtkSuccess              ((MSOHTK) 0x00000001)     // Hit test hit something
#define msohtkTransparentInterior  ((MSOHTK) 0x00000011)     // Hit test hit the transparent interior of an object


/* Drag Set: an array of homogenous Drag Points */
typedef struct _MSODRGS
        {
        ULONG           drghk;                  // Drag handle kind
        int             cPts;                           // Number of points
        int             *rgDots;                        // dots or index
        POINT           *rgpti;                         // corresponding points
        }
        MSODRGS;

/* Drag Handle Info: Information about the kind of drag handles we have */
typedef struct _MSODRGHI
        {
        int             cDots;                  // Number of dots
        ULONG           grfDots;                        // bit flags indicating if a particular dot is present
        ULONG           grfLockedDots;  // bit flags indicating if a particular dot is locked
        }
        MSODRGHI;

/* Creates a MSODRGH with the given type and index. */
/* Use MsoDrghFromDrghkDrghn instead. */
#define MsoDrghCreate(drghk, idrgh) (0 | ((drghk) << 24) \
        | (idrgh))

//#define msodrghNil (MsoDrghCreate(msodrghkNil, 0))

/* Separates the type and index of a MSODRGH. */
//#define MsoDrghkOfDrgh(drgh) ((drgh) >> 24)
#define MsoTypeFromDrgh(drgh) MsoDrghkOfDrgh(drgh)
        /* TODO MsoTypeFromDrgh is an old name for MsoDrghkOfDrgh. Get rid of it. */
#define MsoIndexFromDrgh(drgh) ((drgh) & 0x00ffff)
/* Use MsoDrghnOfDrgh instead */

#define MsoBreakDrgh(drgh, drghk, idrgh) \
        (drghk) = MsoDrghkOfDrgh(drgh); \
        (idrgh) = MsoIndexFromDrgh(drgh)
/* MsoDrghkDrghnOfDrgh is better than MsoBreakDrgh because it doesn't
        evaluate drgh twice. */

/* Asserts if the Drgh is Valid or not. */
#define AssertValidDrgh(drgh) Assert((MsoDrghkOfDrgh(drgh) < msodrghkMax) \
        && (MsoDrghnOfDrgh(drgh) <= 0x00ffff))

/* msodgeConstrainSizeOfShape events get this to constrain the shape, if
        necessary. */
typedef struct _MSODGEBCSOS
        {
        // Input parameters for the Host.
        MSOHSP hsp;
        RECT rchShape;
        MSODRGH drgh;
        POINT pthHit;

        // Output parameters for the Host.
        int xhMinOut;
        int xhMaxOut;
        int xhMinIn;
        int xhMaxIn;
        BOOL fConstrainOutX;
        BOOL fConstrainInX;
        int yhMinOut;
        int yhMaxOut;
        int yhMinIn;
        int yhMaxIn;
        BOOL fConstrainOutY;
        BOOL fConstrainInY;
        } MSODGEBCSOS;

        // TODO davebu -- I know Kasia had a better number than this at some point.
#define cxyhClipboard 100000

// Exported functions.

/* Determines what cursor should we use for this drag handle (dot). */
/* TODO peteren: If this isn't declared with MSOAPI it isn't really exported
        and shouldn't be declared here. */
int MsoIhcFromDrgh(MSODRGH drgh);

#endif //$[VSMSO]
/*      MsoMapPoints transforms each point in an array through the
        transformation defined by two rectangles (*prcFrom and *prcTo) that
        should map to one another.

        pptFrom   Points to the array of points you want mapped.
        pptTo     Points to the array where you'd like the mapped points.
                  This can be the same as pptFrom.
        cpt       How many points are there in your array?
        prcFrom
        prcTo     These two rectangles define the mapping you'd like applied.n */
/* This replaces earlier APIs named MsoMapPt, MsoMapPts, and MsoMapRc */
/* I suspect this doesn't really need to be exported. */
MSOAPI_(void) MsoMapPoints(const POINT *pptFrom, POINT *pptTo, int cpt,
        const RECT *prcFrom, const RECT *prcTo);

/* MsoMapPoints2 is just like MsoMapPoints except for some funky behavior
        when prcTo has zero height or width.  Read the code. */
MSOMACAPI_(void) MsoMapPoints2(const POINT *pptFrom, POINT *pptTo, int cpt,
        const RECT *prcFrom, const RECT *prcTo);

#if 0 //$[VSMSO]
/* Return the direction based on dx and dy */
MSOMACAPI_(MSOCDIR) MsoCdirFromPt(int dx, int dy);

/* Return the direction based on an angle */
MSOMACAPIX_(MSOCDIR) MsoCdirFromAngle(LONG langle);


/*****************************************************************************
        SHAPE COMPARISON
****************************************************************** NeerajM **/

/*
        For comparing all properties of the shapes set opsid = msopsidNil
        and opid = msopidNil.

        For comparing all the properties of a property set, set opsid
        to the property set id and opid to msopidNil.

        For comparing a single property set opid
        to the property id of the property and opsid to msopsidNil.

*/
typedef struct _MSOHSPCMP
        {
        // first shape
        IMsoDrawing *pidg1;
        MSOHSP      hsp1;

        // second shape
        IMsoDrawing *pidg2;
        MSOHSP      hsp2;

        MSOPSID     opsid; // property set to compare
        MSOPID      opid;  // property to compare
        BYTE        fopid[msopidLast+1];
        BYTE        fresult[msopidLast+1];
        //MSOPID *popids;
        //BOOL   *presult;
        //int opidssize;
        } MSOHSPCMP;
#define cbMSOHSPCMP (sizeof(MSOHSPCMP))

MSOAPI_(void) MsoInitHspcmp(MSOHSPCMP *phspcmp);

MSOAPI_(BOOL) MsoFCompareShapes(MSOHSPCMP *phspcmp);


/****************************************************************************
        IMsoDisplayElementSet (ides)

        This is the interface you have to implement in order to get something
        to be drawn on the screen in amongst other stuff you don't know about.

        A Display Element is a 2-D object that can draw itself into a DC and
        hit-test itself.  It cannot be saved, loaded, or edited.  It does not
        contain high-resolution coordinates.  It exists   at a single Z-order
        position relative to all other Display Elements.

        Display Elements are not implemented as individual objects but by
        aggregate objects that speak the interface IMsoDisplayElementSet.
        For example, all of the Drawing Objects on a single Word "page" are
        displayed by a single implementation of IMsoDisplayElementSet, provided
        by a Drawing View object (DGV).

        All of the methods of IMsoDisplayElementSet take as their first
        argument (after the "this" pointer) a void * called pvDes.  This
        value is passed to the Display Manager when a Display Element Set
        is registered and then passed back as an argument whenever that the DM
        calls a method in that DES.  This allows a host to implement what amounts
        to several DESs with a single object.

        In addition, many of the methods of IMsoDisplayElementSet manipulate
        a single Display Element within the set.  These all take as their
        second argument another void * called pvDe, which the DES can use
        to distinguish one DE from another.
****************************************************************************/

// A MSOHDE identifies a Display Element relative to Display Manager
typedef void *MSOHDE;
#define msohdeNil ((MSOHDE)(0))

// A MSOHDES identifies a Display Element Set relative to Display Manager
typedef int MSOHDES;
#define msohdesMax ((MSOHDES)250)
#define msohdesNil ((MSOHDES)251)
        /* Internally, Office stores HDESs in 8 bits.  But it's only Office
                that ever generates them, so there's shouldn't any problem with them
                fitting. */

/* MSOIDEL -- Index to a display element layer */
/* When you create a Display Manager you use ask for some number of DELs
        (of various types).  Then when Display Element Sets register Display
        Elements into that Display Manager they pass in idel's specifying the
        layer they want to be in. */
#define msoidelMax (254)
#define msoidelNil (255)
        /* Internally Office stores idel's in 8 bits */

/* MSODELK -- Display Element Layer Kind */
/* Display Element Layers come in the following different kinds.  The DELs
        in a DM are always ordered, bottom to top, by DELK; that is,
        first come any "msodelkMain" layers, then any "msodelkCached",
        and so on. */
typedef enum
        {
        msodelkMain = 0,
                /* Normal Display Elements.  Shapes are drawn here. */
        msodelkCached = 1,
                /* Cached Display Elements; can be drawn however you like but
                        the DM will be keeping bitmaps behind them so they should probably
                        be small.  Selection Dots are drawn here. */
        msodelkXOR = 2,
                /* Display Elements drawn _entirely_ with XOR (such that drawing them
                        again erases them).  Selection fuzz is drawn here. */
        msodelkTop = 3,
                /* Display Elements that want to be on top of everyone else. */
        msodelkMax = 4,
        msodelkNil = 15,
                /* Hey! DELKs are stored in 4 bits in DELs. */
        } MSODELK;

/* MSODMHD == Display Manager Hit Description. */
/* TODO peteren: This is kind of an extensibility liabiliy.  Talk about
        it with someone smart. */
#define cbMSODESDMHDMax (32)
typedef struct _MSODMHD
        {
        /* These first three fields are filled out by the Display Manager
                before it starts calling Display Element Set "were you hit-test?"
                methods. */
        POINT ptw;
        POINT ptv;
        MSOMM mm;

        /* After a Display Element Set returns a hit (or none does) the Display
                Manager will set these next fields before returning. */
        BOOL fHit;
        MSOHTK htk;
        MSOHDES hdesHit;
        MSOHDE hdeHit;

        /* Finally the rgbDes is not used by the DM.  It may be filled by a
                Display Element Set while hit-testing itself.  The DES can even
                hang allocations off one of these, but only if it's returning
                fHit == TRUE (only in such cases will we properly call you
                back to let you free things). */
        BYTE rgbDes[cbMSODESDMHDMax];
        } MSODMHD;

#if FUTURE
/* GRFDESQDE == Display Element Set Query Display Element */
/* These flags are used by the DM ask yes/no questions of Display
        Elements Sets about particular Display Elements. */
#define msofdesqdeIsNonOpaque (1<<0)
#endif // FUTURE

/* GRFDESD - Arguments to FDrawDisplayElement */
#define msogrfdesdNil (0)

#undef  INTERFACE
#define INTERFACE IMsoDisplayElementSet
DECLARE_INTERFACE_(IMsoDisplayElementSet, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDisplayElementSet methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        // ----- IMsoDisplayElementSet methods (deal with whole set)

        /* SetHdes is called by the Display Manager to tell a Display Element Set
                how to identify itself. */
        MSOMETHOD_(void, SetHdes) (THIS_ void *pvDes, MSOHDES hdes) PURE;

        /* The Display Manager will call FixDisplayElements on any Display
                Element Sets that has been invalidated before it does most anything.
                The DES gets to register, unregister, move, or invalidate
                individual DEs.  It can't register new DESs or unregister itself,
                though. */
        MSOMETHOD_(void, FixDisplayElements) (THIS_ void *pvDes) PURE;

        /* TODO peteren: Comment? */
        MSOMETHOD_(BOOL, FHitTestSet) (THIS_ void *pvDes, MSODMHD *pdmhd,
                BOOL fAfterDe) PURE;


        // ----- IMsoDisplayElementSet methods (deal with single DE)

        /* SetDisplayElementHde is similar to SetHdes, except that it
                sets the "handle" of a single Display Element. */
        MSOMETHOD_(void, SetDisplayElementHde) (THIS_ void *pvDes, void *pvDe,
                MSOHDE hde) PURE;

        /* The Display Element Set is presumed to know the coordinates (relative
                to the Display Manager, V coordinates) of each Display Element.  To
                save space, therefore, the Display Manager does NOT store these
                rectangles, but instead asks the DES for them.  However, the DM may
                store some positional data, so you can't move a DE just be returning
                a new value in this routine. */
        MSOMETHOD_(void, GetDisplayElementBounds) (THIS_ void *pvDes, void *pvDe,
                RECT *prcvBounds) PURE;

        /* The image of a particular Display Element may be made up of a number
                of small rectangular pieces.  To avoid wasting time and memory
                the Display Manager will ask each DE how many pieces it consists of
                and (if it reports more than one) then ask for the bounding rectangle
                of each indivual piece.  You can return either 0 or 1 if you
                don't want to bother with this whole "piece" optimization; this
                will ensure that we never even call GetDisplayElementPieceBounds. */
        MSOMETHOD_(int, CountDisplayElementPieces) (THIS_ void *pvDes, void *pvDe)
                PURE;

        /* If you return a number > 1 from CountDisplayElementPieces the DM
                may call GetDisplayElementBounds to get the bounding rectangles of the
                individual pieces. */
        MSOMETHOD_(void, GetDisplayElementPieceBounds) (THIS_ void *pvDes,
                void *pvDe, int iPiece, RECT *prcvBounds) PURE;

        /* The DM calls FUnionDisplayElementCoverageRegion to ask the Display
                Element to union into a provided region one of it's "coverage"
                regions (as specified by cvrk).  On the way in *phrgnv may be
                msohrgnNil so if you want to add something to it you'll have to
                call MsoFEnsureRgn. Return TRUE if you succeed, FALSE if you run
                into an error. */
        MSOMETHOD_(BOOL, FUnionDisplayElementCoverageRegion) (THIS_ void *pvDes,
                void *pvDe, MSOCVRK cvrk, HRGN *phrgnv) PURE;

        /* FIsDisplayElementNonOpaque should return TRUE if the Display Element
                has stuff to draw that's not covered by the region it returns
                from FGetDisplayElementOpaqueRegion. */
        MSOMETHOD_(BOOL, FIsDisplayElementNonOpaque) (THIS_ void *pvDes,
                void *pvDe) PURE;

        /* FDrawDisplayElement should draw the DE specified by pvDes and
                pvDe into *pdc.  This method should return FALSE if display was not
                completed AND YOU HAVE SOME HOPE of completing the display if we
                call you back later.  In other words, if display is aborted you
                should return FALSE, but if you fail to allocate some region or
                other you should just return TRUE. *prcvClip will be a rectangle
                outside of which you don't have to bother drawing anything
                (clipping, etc.).  The last DES to draw in this MSODC will be
                identified by hdesPrev.  Various flags will be in grfdesd; see
                definitions thereof. */
        /* FUTURE peteren: Have this take an MSORGNX * instead of a RECT *.
                Also have it return group of flags. */
        MSOMETHOD_(BOOL, FDrawDisplayElement) (THIS_ void *pvDes, void *pvDe,
                MSODC *pdc, ULONG grfdesd, RECT *prcvClip, MSOHDES hdesPrev) PURE;

        /* FHitTestDisplayElement should check to see if the specified point
                "hits" the specified Display Element and return TRUE if it does.
                TODO peteren: Comment? */
        MSOMETHOD_(BOOL, FHitTestDisplayElement) (THIS_ void *pvDes, void *pvDe,
                MSODMHD *pdmhd) PURE;

        /* The DM will call FGetDisplayElementPieceInfo to get a bunch of
                information about this DE all at once.  This is intended to
                replace CountDisplayElementPieces and GetDisplayElementPieceBounds,
                which were something of a speed hit.  Really this should get the
                bounds of the DE itself while it's at it.  Return TRUE on success,
                FALSE on failure.  Return in *pcPieces the number of pieces this
                DE has.  If you return a value > 1 you must also in *prgrcvPiece
                return a pointer to an array of rectangles allocated with MsoFMarkMem. */
        MSOMETHOD_(BOOL, FGetDisplayElementPieceInfo) (THIS_ void *pvDes,
                void *pvDe, int *pcPieces, RECT **prgrcvPiece) PURE;

};


/****************************************************************************
        Display Manager Site

        TODO peteren: Comment?
****************************************************************************/

/* MSODMEB -- Display Manager Event Block
        These are passed as the argument to DMS::OnEvent */
MsoDEBDeclare(dm, msodgeDmsFirst, msodgeDmsMax);
typedef struct MSODMEB
        {
        /* These first fields are filled out for EVERY Display Manager Event */
        MSODGE dge;
        BOOL fResult;
        interface IMsoDisplayManager *pidm;
        MsoDEBPDefine(dm);

#if DEBUG
        /* Only one set of fields in this union will be
                filled out, depending on dge. */
        union
                {
                struct
                        {
                        int rgnPlaceholder[4];
                                /* Padding to avoid full builds when adding fields */
                        };
                };
#endif // DEBUG

        } MSODMEB;
#define cbMSODMEB (sizeof(MSODMEB))

/* DMSI = Display Manager Site Info.  This is a bunch of data that
        describes the Display Manager Site used by this DM.  This is passed
        as an argument to MsoFCreateDisplayManager and can subsequently be
        retrieved and modified by calling DM::PdmsiBeginUpdate and
        DM::EndDmsiUpdate.  This works just like the DGSI (Drawing Site Info)
        and DGVSI (Drawing View Site Info) -- it's primarly a push model.
        If the host wants to update these only when the DM actually needs
        them they should call these DM methods from DMS::OnEvent. */
typedef struct _MSODMSI
        {
        interface IMsoDisplayManagerSite *pidms; // DMS to use
        void *pvDms; // Client data to pass to all DMS methods
        POINT ptwOrigin; // At what W point is (0, 0) in V coordinates.
        RECT rcvBounds; // The bounds of the DM in V coordinates.
        RECT rcwVisible; // The visible portion of the window we live in.
        HBRUSH hbrBackground;  // The brush we should use to erase the background.
                /* TODO peteren: Get rid of hbrBackground! */

        /* The following fields are used by the host when creating the
                Display Manager to request a number of layers of different types. */
        int cdelMain;
        int cdelCached;
        int cdelXOR;
        int cdelTop;

        union
                {
                struct
                        {
                        ULONG fWantsEvents : 1;
                                /* Means that this DMS wants its OnEvent method
                                        called whenever a DisplayManager Event happens. */
                        ULONG fUpdatePostponed : 1;
                                /* Updates in this DM have been postponed.
                                        This means that FNeedsUpdate will lie and return
                                        FALSE. */
                        ULONG fDontUseRcvUpdate : 1;
                                /* WORD HACK! (JasonHi):
                                    Skip intersecting the clip region and rely on the
                                        hdc clip region being setup correctly.
                                        This should/probably could be done with changes to
                                        how we use/manage the various rects.  But for now
                                        non-Word apps probably never want to set this bit. */
                        ULONG grfUnused : 29;
#if BOGUS
                        /* Turned out to be terribly convienient to give the host
                                a few bits to remember stuff about how it's treating
                                this DM.  After initializing them to zero, Office will
                                never touch or look at the following bits. */
                        ULONG fSite1 : 1;
                        ULONG fSite2 : 1;
                        ULONG fSite3 : 1;
                        ULONG fSite4 : 1;
#endif // BOGUS
                        };
                ULONG grf;
                };

        MsoDEBDefine(dm);

#if DEBUG
        LONG lVerifyInitDmsi; // Used to make sure people use MsoInitDmsi
#endif // DEBUG
        } MSODMSI;
#define cbMSODMSI (sizeof(MSODMSI))

/* You should call MsoInitDmsi on any new DMSI you're filling out
        before passing it into MsoFCreateDisplayManager.  This lets us add
        new fields without changing code in all the hosts. */
MSOAPI_(void) MsoInitDmsi(MSODMSI *pdmsi);

/* IMsoDisplayManagerSite.
        This is the minimal interface you have to speak to contain a Display
        Manager.        To enable the host to simulate many "virtual" DMSs with a single
        actual object, all the methods in IMsoDisplayManagerSite take a void *
        named pvDms as their first argument (well, second after the "this"
        pointer).  When a DM is created it will be passed a IDMS interface
        pointer and a pvDms, which it will pass into all subsequent DMS
        method calls. */
#undef  INTERFACE
#define INTERFACE IMsoDisplayManagerSite
DECLARE_INTERFACE_(IMsoDisplayManagerSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDisplayManagerSite methods (standard stuff)

        /* FDebugMessage method */
        MSODEBUGMETHOD

        /* The Display Manager will call DMS::OnEvent when various interesting
                events take place.  See the definition of MSODGE (Drawing Event)
                for a more detailed description of events.  This method will only
                be called if you set msodmsi.fWantsEvents to TRUE. */
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDms, MSODMEB *pdmeb) PURE;

        // ----- IMsoDisplayManagerSite methods (DCs)

        /* The Display Manager needs to be able to call FGetDc at any time
                and be given an MSODC into which it may draw.  FGetDc should return
                FALSE if for some reason no HDC is available. */
        MSOMETHOD_(BOOL, FGetDc) (THIS_ void *pvDms, MSODC *pdc) PURE;

        /* In fact most of our clients are likely to have a DC lying around
                all the time.  Nevertheless the Display Manager won't hang onto
                DCs for excessive lengths of time and will call ReleaseHdc when
                it doesn't need the one it got any more. */
        MSOMETHOD_(void, ReleaseDc) (THIS_ void *pvDms, MSODC *pdc) PURE;

        /* The Display Manager will call RequestInvalidateRgn and
                RequestInvalidateRect from under calls to DM::InvalidateWindow.
                The host can implement these with simple calls to Windows'
                InvalidateRgn and InvalidateRect. */
        MSOMETHOD_(void, RequestInvalidateRect) (THIS_ void *pvDms, RECT *prcw)
                PURE;
        MSOMETHOD_(void, RequestInvalidateRgn) (THIS_ void *pvDms, HRGN hrgnw)
                PURE;

#if FUTURE
        /* DM calls RequestEraseRgnx from within DM::FDoUpdate to have the
                host erase exposed stuff. The MSODC's hdc will have already been
                set up to draw in V coordiantes. */
        MSOMETHOD_(void, RequestEraseRgnx) (THIS_ void *pvDms, MSODC *pdc,
                MSORGNX *prgnxv) PURE;
#endif // FUTURE
};


/****************************************************************************
        IMsoDisplayManager (idm)

        TODO peteren: Comment?

        One thing you'll probably have to know about to use a Display Manager is
        the difference between what we call "W" and "V" coordinates.  The "W"
        stands for "Window" and the "V" for View, although these terms are not
        terribly self-explanatory.
        What they actually boil down to is that W coordinates are external
        to the display manager and V are internal.  When we call FGetHdc in our
        site object we'll expect to be given a DC in which we might draw using W
        coordinates.  Before asking any Display Elements to display themselves,
        however, we'll modify the DC so that we can draw at V coordinates
        (relative to the Display Manager's origin).  The distinction means that
        we have very little adjusting to do when the Display Manager is scrolled
        around in its window.
****************************************************************************/

/* MsoFCreateDisplayManager attempts to create a new object speaking
        IMsoDisplayManager.  It returns TRUE if it succeeds.  A pointer to the
        new interface is returned in *ppidm.  You have to give us a filled-out
        MSODMSI describing the site of this DM.  Use MsoInitDmsi to initialize
        it before you change whatever fields you want to change and call us. */
MSOAPI_(BOOL) MsoFCreateDisplayManager(interface IMsoDisplayManager **ppidm,
        MSODMSI *pdmsi);


/* MSOFDMRDE -- Display Manager Register Display Element */
/* These flags are passed as arguments to
        IMsoDisplayManager::FRegisterDisplayElement.  Some of these describe
        the new Display Element, others give instructions on where to
        put it in the list. */
#define msogrfdmrdeNil           (0)

#define msofdmrdeLayerMain       (0)
#define msofdmrdeLayerCached     (1<<0)
#define msofdmrdeLayerXOR        (1<<1)
        /* FUTURE peteren: Remove these flags, as they don't do anything.
                The idel you pass in to RegisterDisplayElement specifies this. */

#define msofdmrdeInsertTop       (0)
#define msofdmrdeInsertBottom    (1<<2)
#define msofdmrdeInsertAbove     (1<<3)
#define msofdmrdeInsertBelow     (1<<4)
        /* These flags tell you where within the specified layer you want
                the new DE inserted (in the Z-order, that is).  Top (the default)
                and Bottom insert rather obviously at the top and bottom of the
                layer.  Above and Below insert relative to hdeNeighbor.  Inserting
                "above" msohdeNil means at the bottom (turns out to be useful). */

#define msofdmrdeVisible         (0)
#define msofdmrdeInvisible       (1<<5)
        /* These two options tell whether you want the Display Element to be
                initially visible or not. */

#define msofdmrdeInvalid         (0)
#define msofdmrdeValid           (1<<6)
        /* Ordinary new DEs are inserted "invalid" which means they will be
                drawn during the next update.  To suppress this you can register
                them with msofdmrdeValid. */

#define msofdmrdeBadClip         (1<<7)
        /* Annoying addition to handle XL charts, which are unable to
                stay inside their clipping regions. */

#define msogrfdmrdeInsert \
        (msofdmrdeInsertTop | msofdmrdeInsertBottom | \
        msofdmrdeInsertAbove | msofdmrdeInsertBelow)

/* DMUB = Display Manager Update Block */
/* Input and output parameters for DM::FInitUpdate, DM::FPlanUpdate,
        DM::FDoUpdate, and DM::FreeUpdate. */
typedef struct MSODMUB
        {
        RECT rcvErase; // Returned from DM::FPlanUpdate
        HRGN hrgnvErase; // Returned from DM::FPlanUpdate
        union
                {
                ULONG grf;
                struct
                        {
                        ULONG fNoBlink : 1;
                        ULONG fNoBlinkIfTop : 1;
                        ULONG fNoBlinkIfGuess : 1;
                        ULONG fAbortForBadClip : 1;
                                /* DM should abort if it needs to draw an "fBadClip" DE. */
                        ULONG : 12; // Space for future input
                        ULONG fError : 1;
                        ULONG fErase : 1;
                        ULONG fAbortedForBadClip : 1;
                                /* DM did abort for an "fBadClip" DE. */
                        ULONG fNothingToUpdate : 1;
                                /* fNothingToUpdate is returned from FPlanUpdate, TRUE means
                                        there turned out to be nothing to update after all.  You
                                        don't have to call DM::FDoUpdate, but you still need to
                                        call DM::FreeUpdate. */
                        ULONG fAborted : 1;
                                /* FDoUpdate aborted */
                        ULONG : 11; // Space for future output
                        };
                };
#if DEBUG
        ULONG rgPaddingToAvoidFullBuilds[10];
#endif // DEBUG
        } MSODMUB;
#define cbMSODMUB (sizeof(MSODMUB))


/* MSOFDMIDE flags are passed to DM::FInvalidateDisplayElement and
        it's core function, DM::FInvalidateDei */
#define msogrfdmideNil     0x00000000
#define msofdmideRcv       0x00000001 // *prcInval is an rcv
#define msofdmideRcw       0x00000002 // *prcInval is an rcw
#define msofdmideErase     0x00000004 // we need to erase this DEI
#define msofdmideDraw      0x00000008 // we need to draw this DEI
#define msofdmideDrawAbort 0x00000010
        /* We need to draw this DEI because we aborted redraw.  A few things
                are slightly different about such redraw.  I think this will probably
                only be used internally to Office. */
#define msofdmideDrawUp    0x00000020
#define msofdmideDrawDown  0x00000040
        /* We need to draw this DEI because it went up or down in the Z-Order.
                There's a few things weirdly different about such redraw. I think these
                will probably only be used from within Office. */


#define msogrfdmideRc (msofdmideRcv | msofdmideRcw)

/* MSOFDMPs are passed to IMsoDisplayManager::FPaint */
#define msogrfdmpNil          (0)
#define msofdmpHrgnw          (1<<0)
        /* Means hrgnv argument is really in W coordinates */
#define msofdmpRcw            (1<<1)
        /* Means rcv argument is really in W coordinates. */
#define msofdmpClip           (1<<2)
        /* Caller wants us to clip to hrgnv/rcv */
#define msofdmpOnlyMarked     (1<<3)
        /* We should only draw Marked Display Elements. */
#define msofdmpOmitOpaque     (1<<4)
        /* We can omit drawing anything covered by our opaque regions;
                for example the caller may have called FExcludeOpaqueRegion
                and then clipped to the remaining region and drawn something. */
#define msofdmpOtherDc        (1<<5)
        /* TODO peteren: comment */
#define msofdmpClipTop        (1<<6)
#define msofdmpValidate       (1<<7)
        /* Means that as we paint stuff we should try to validate any
                Display Elements covered by that area.  Not Yet Implemented!!! */
#define msofdmpValidateAll    (1<<8)
        /* Means we should validate all the DEs.  You should only pass
                this in if you're painting the entire DM.  Probably we should
                have a seperate DM function to mark all the DEs valid. */
#define msofdmpEraseXOR       (1<<9)
        /* Used internally when we're erasing XOR DEs (we pay attention to
                the "fVisibleNow" bit in the DEs). */
#define msofdmpStartAboveOpaque (1<<10)
        /* Pass this to have DM::FPaint spend a little extra time hunting
                for big opaque shapes that cover the entire area we're painting.
                If we find one we'll start the paint there in the Z-order (less
                flickering). */
#define msofdmpOnlyBelowCache (1<<11)
        /* This flag is a hack to make a particular XL case work.  It causes
                DM::FPaint to not make any new caches and to not draw anything
                above the cache.  The particular case is the call XL makes after
                it has XORed the cell hilite to redraw any bits of Shapes that might
                have just been bogusly XORed.  The reason this is problematic
                is that this may be happening because we just selected a Shape
                and the interference between the disappearing cell hilite and
                the appearing selection dots leads to bad caching and XOR scudge. */

/* Most callers of DM::FPaint should pass msogrfdmpDefault. */
#define msogrfdmpDefault (msogrfdmpNil)


/* MSOFDMHTs are passed to IMsoDisplayManager::FHitTest */
#define msogrfdmhtNil          (0)
#define msofdmhtDesBeforeDe (1<<0)
#define msofdmhtDe          (1<<1)
#define msofdmhtDesAfterDe  (1<<2)
#define msogrfdmhtNormal \
        (msofdmhtDesBeforeDe | msofdmhtDe | msofdmhtDesAfterDe)

/* MSOFDMI are flags to passed to DM::Invalidate */
#define msogrfdmiNil          (0)
#define msofdmiFromWindows    (1<<0)
        /* If you're giving us this invalidation because Windows told you
                a particular region needs painting then you need to pass
                msofdmiFromWindows to let the DM know that any caches it might
                have in this region are also suspect (there might have been some
                other window over top of ours, and our caches might have bits
                from that window in them) */
#define msofdmiRcvIsRcw        (1<<1)
        /* The RCV you passed us is really an RCW */
#define msofdmiHrgnvIsHrgnw    (1<<2)
        /* The HRGNV you passed us is really an HRGNW */

/* GRFDMQ flags are passed to and returned from DM::GrfdmqQuery */
#define msogrfdmqNil                 0
#define msofdmqEraseBeforeUnregister (1<<0)

/* GRFDMQDE flags are passed to DM::GrfdmqdeQueryDisplayElement */
#define msogrfdmqdeNil            0
#define msofdmqdeIsVisible        (1<<0)
#define msofdmqdeIsVisibleNow     (1<<1)
#define msofdmqdeIsMarked         (1<<2)

/* IMsoDisplayManager.  The actual interface spoken by DMs. */
#undef  INTERFACE
#define INTERFACE IMsoDisplayManager
DECLARE_INTERFACE_(IMsoDisplayManager, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDisplayManager methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Free tells a Display Manager to go away.  The caller has to know
                that no-one else is referencing it. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* PdmsiBeginUpdate returns a pointer to an MSODMSI.  The caller may
                examine and (if they didn't pass fReadOnly == TRUE) change the
                values therein.  They must then call EndDmsiUpdate.  See comment
                by the definition of MSODMSI. */
        MSOMETHOD_(MSODMSI *, PdmsiBeginUpdate) (THIS_ BOOL fReadOnly) PURE;

        /* EndDmsiUpdate should be called after PdmsiBeginUpdate.  You should
                pass back the pointer PdsiBeginUpdate gave you.  In fChanged you
                should pass back TRUE if you changed anything.  See comment by
                PdmsiBeginUpdate. */
        MSOMETHOD_(void, EndDmsiUpdate) (THIS_ struct _MSODMSI *pdmsi,
                BOOL fChanged) PURE;

        /* GrfdmqQuery can be used to ask various yes/no questions
                of the DM as a whole.  You pass in grfdmq full of the flags
                you want to ask and get back a grfdmq with those flags filled
                out. */
        MSOMETHOD_(ULONG, GrfdmqQuery) (THIS_ ULONG grfdmq) PURE;

        // ----- IMsoDisplayManager methods (coordinates)

        /* ConvertPtwToPtv converts an array of points from W coordinates
                (relative to the window the DM lives in) to V coordinates (relative
                 to the DM itself).  This works in-place (pptw == pptv). */
        MSOMETHOD_(void, ConvertPtwToPtv) (THIS_ const POINT *pptw, POINT *pptv,
                int cpt) PURE;

        /* ConvertPtvToPtw converts an array of points from V coordinates
                (relative to the DM) to W coordinates (relative to the window the
                DM lives in).  This works in-place (pptv == pptw). */
        MSOMETHOD_(void, ConvertPtvToPtw) (THIS_ const POINT *pptv, POINT *pptw,
                int cpt) PURE;

        /* GetPtwOrigin returns the ptw that corresponds to ptv = (0, 0).
                It's slightly redundant with ConvertPtvToPtw, but still useful. */
        MSOMETHOD_(void, GetPtwOrigin) (THIS_ POINT *pptwOrigin) PURE;

        /* GetRcvBounds returns the bounds in V's (relative to ptwOrigin) of
                this Display Manager. */
        MSOMETHOD_(void, GetRcvBounds) (THIS_ RECT *prcvBounds) PURE;

        /* Similar to GetPtwOrigin.  Tells you the visible rectangle of the DM.
                You can use this to avoid wasting time, but not much else. */
        MSOMETHOD_(void, GetRcwVisible) (THIS_ RECT *prcwVisible) PURE;

        /* FGetRcvVisible gives you back the intersection of rcvBounds
                and rcwVisible, in V coordinates.  Returns TRUE iff the rectangle
                is not empty. */
        MSOMETHOD_(BOOL, FGetRcvVisible) (THIS_ RECT *prcvVisible) PURE;

        // ----- IMsoDisplayManager methods (invalidation)

        /* You can call Invalidate to let the DM know it should redraw a
                particular region or rectangle next time it Updates. */
        MSOMETHOD_(void, Invalidate) (THIS_ ULONG grfdmi,
                RECT *prcv, HRGN hrgnv) PURE;

        /* Call InvalidateWindow to have the DM sort through the invalidation
                it's storing, call DMS::RequestInvalidateRect and
                DMS::RequestInvalidateRgn on appropriate rectangles and regions,
                and mark itself valid.  This function can be useful if for some
                reason you're about to process an WM_PAINT message. */
        /* FUTURE peteren: Rename this "Validate" and give it a group of
                flags including one to specify whether or not you want to get
                called back with invalid areas. */
        MSOMETHOD_(void, InvalidateWindow) (THIS) PURE;

        // ----- IMsoDisplayManager methods (update)

        /* FNeedsUpdate() returns TRUE iff this DM needs Updating. */
        MSOMETHOD_(BOOL, FNeedsUpdate) (THIS) PURE;

        /* Call DM::FInitUpdate to initialize a MSODMUB.  An MSODMUB is a
                Display Manager Update Block; these are used to contain a
                description of what needs to be done to get the screen up to date.
                We'll return FALSE if there's no updating to be done, this function
                cannot fail.  If we return TRUE you must eventually call
                DM::FreeUpdate on this MSODMUB. */
        MSOMETHOD_(BOOL, FInitUpdate) (THIS_ MSODMUB *pdmub) PURE;

        /* After initializing an MSODMUB the host should call FPlanUpdate to
                have the DM clean itself up and scan through all the invalidation it
                knows about, constructing in the MSODMUB a list of all the
                displaying that needs to happen to get the screen up to date. After
                you call this the DM will be free of invalidation, will have made
                any new bitmap caches it needs, etc.  This method does _not_
                actually draw anything on the screen.  After calling this you don't
                actually have to call FDoUpdate immediately; it's okay (for
                example) to call FPaint a few times.  FPlanUpdate returns TRUE if
                everything worked, FALSE if an error occurred that prevented it
                from constructing the MSODMUB it wanted to.  Usually this means it
                ran out of memory doing some region operation.  Even if it returns
                FALSE, though, you will have been give an MSODMUB that will work if
                passed to FDoUpdate (it will just update everything). */
        /* This used to be called FBeginUpdate */
        MSOMETHOD_(BOOL, FPlanUpdate) (THIS_ MSODMUB *pdmub) PURE;

        /* After calling FPlanUpdate (and getting back success) the DM will
                be all cleaned up internally (no invalidation, etc.), but the screen
                will still be out of date.  To have the necessary Display Elements
                redraw themselves, call FDoUpdate. */
        /* This used tobe call FEndUpdate */
        MSOMETHOD_(BOOL, FDoUpdate) (THIS_ MSODMUB *pdmub) PURE;

        /* After you've gotten back TRUE from FInitUpdate, you have to
                pass the MSODMUB you got back into FreeUpdate.  This is true
                regardless of any calls you might have made to FPlanUpdate or
                FDoUpdate */
        MSOMETHOD_(void, FreeUpdate) (THIS_ MSODMUB *pdmub) PURE;

        /* Call FDelayUpdate to see if this DM is delaying updates
                (that is, lying and returning FALSE from FNeedsUpdate). */
        MSOMETHOD_(BOOL, FDelayUpdate) (THIS) PURE;

        /* Call SetDelayUpdate to tell the DM to start or stop
                delaying updates. */
        MSOMETHOD_(void, SetDelayUpdate) (THIS_ BOOL fDelayUpdate) PURE;

        /* TODO peteren: Comment */
        /* TODO peteren: We should make this an option to GrfdmqQuery */
        MSOMETHOD_(BOOL, FNeedsPaint) (THIS_ const MSORECT *prcv) PURE;

        /* FPaint will cause the DM to draw all the DEs it knows about
                within both *prcv (can be NULL) and hrgnv (can be msohrgnNil)
                into the specified MSODC.  This will not clear any invalidation
                in the DM, for that you must call FPlanUpdate.
                It's OK to call FPlanUpdate and then, before getting around to
                calling FDoUpdate, call FPaint.  If while doing this you pass in
                a pointer to your MSODMUB we'll remove from it any regions
                that by painting you've corrected.  Even if the regions remaining
                to be updated shrink away to nothing, though, you still have to
                call FDoUpdate to do things like get rid of DEs we don't neeed
                anymore. */
        MSOMETHOD_(BOOL, FPaint) (THIS_ MSODC *pdc, RECT *prcv, HRGN hrgnv,
                ULONG grfdmp) PURE;

        // ----- IMsoDisplayManager methods (opaque regions)

        /* Call FGetOpaqueRegion to get from us a region that includes the union
                of all of the opaque regions of the Display Elements.  This is useful
                if you want to draw stuff under us and you'd like to clip out as much
                of the Shape layer as possible.  Pass us in a pointer to an
                uninitialized region, we'll give you a handle to our (cached) opaque
                region.  You must _not_ mess with the region we returned (since we
                have it cached).  We return TRUE if we actually give you a non-empty
                region, FALSE otherwise.  We'll return FALSE both in the case where
                an error happened and in the case where there just wasn't an opaque
                region.  You can check FErrorOpaqueRegion afterwards if you need
                to distinguish. */
        MSOMETHOD_(BOOL, FGetOpaqueRegion) (THIS_ HRGN *phrgnv) PURE;
        MSOMETHOD_(void, ResetOpaqueRegion) (THIS) PURE;
        MSOMETHOD_(void, SetErrorOpaqueRegion) (THIS) PURE;
        MSOMETHOD_(BOOL, FErrorOpaqueRegion) (THIS) PURE;

        // ----- IMsoDisplayManager methods (marking DEs)

        MSOMETHOD_(void, MarkDisplayElement) (THIS_ MSOHDE hde, BOOL fMark) PURE;
        /* Hey! Use GrfdmqdeQueryDisplayElement instead of FIsDisplayElementMarked */
        MSOMETHOD_(BOOL, FIsDisplayElementMarked) (THIS_ MSOHDE hde) PURE;
        MSOMETHOD_(void, MarkAllDisplayElements) (THIS_ BOOL fMark) PURE;
        MSOMETHOD_(BOOL, FMarkNonOpaqueDisplayElements) (THIS_ RECT *prcv,
                BOOL fMark) PURE;

        // ----- IMsoDisplayManager methods (hit-testing)

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FHitTest) (THIS_ MSODMHD *pdmhd, POINT *pptvTest,
                BOOL fPtwTest, MSOMM mmTest, MSOHDES *prghdesTest,
                int chdesTest, ULONG grfdmht) PURE;

        // ----- IMsoDisplayManager methods (affect a single DES)

        /* FRegisterDisplayElementSet() is how a new Display Element Set attaches
                itself to a Display Manager.  It returns TRUE on success, in which case
                it will have called IMsoDisplayElementSet::SetHdes().  The new DES
                starts out valid; to get us to call DES::FixDisplayElements you'll
                have to call DM::InvalidateDisplayElementSet. */
        MSOMETHOD_(BOOL, FRegisterDisplayElementSet) (THIS_
                interface IMsoDisplayElementSet *pides, void *pvDes) PURE;

        /* UnregisterDisplayElementSet() tells a Display Manager to let go
                of a particular DES.  The DM will neatly unregister all of the
                DEs in this DES first. */
        MSOMETHOD_(void, UnregisterDisplayElementSet) (THIS_ MSOHDES hdes) PURE;

        /* GetDisplayElementSetPvDes returns the pvDes the DM is using to identify
                this DES when it calls IMsoDisplayElementSet. */
        MSOMETHOD_(void, GetDisplayElementSetPvDes) (THIS_ MSOHDES hdes,
                void **ppvDes) PURE;

        /* SetDisplayElementSetPvDes may be used to change the pvDes the DM is
                using to identify a particular DES when it calls
                IMsoDisplayElementSet. */
        MSOMETHOD_(void, SetDisplayElementSetPvDes) (THIS_ MSOHDES hdes,
                void *pvDes) PURE;

        /* A Display Element Set should use InvalidateDisplayElementSet to tell
                the DM it wants it's FixDisplayElements method called before the next
                Update(). */
        MSOMETHOD_(void, InvalidateDisplayElementSet) (THIS_ MSOHDES hdes) PURE;

        /* To unregister all of its Display Elements at once without unregistering
                itself, a DES can call UnregisterAllDisplayElementsInSet.  Flags
                go in grfdmude */
        MSOMETHOD_(void, UnregisterAllDisplayElementsInSet) (THIS_
                MSOHDES hdes) PURE;

        // ----- IMsoDisplayManager methods (affect a single DE)

        /* FRegisterDisplayElement() is how a DES adds a new DE to a DM.
                The DES identifies itself with an MSOHDES and the particular DE
                with a pvThisDe (which doesn't actually have to be a pointer).
                Specify Z-order position and a bunch of other facts about this DE
                in grfdmrde; some flags use hdeNeighbor. */
        MSOMETHOD_(BOOL, FRegisterDisplayElement) (THIS_ MSOHDES hdes,
                void *pvDe, ULONG grfdmrde, int idel, MSOHDE hdeNeighbor) PURE;

        /* To get rid of a Display Element you pass in its MSOHDE to
                UnregisterDisplayElement.  Flags go in grfdmude. */
        MSOMETHOD_(void, UnregisterDisplayElement) (THIS_ MSOHDE hde) PURE;

        /* GrfdmqdeQueryDisplayElement can be used to ask various yes/no questions
                of a particular display element.  You pass in in grfdmqde flags
                corresponding to the questions you want to ask. */
        /* TODO peteren: Replace various methods with this. */
        MSOMETHOD_(ULONG, GrfdmqdeQueryDisplayElement) (THIS_ MSOHDE hde,
                ULONG grfdmqde) PURE;

        /* GetDisplayElementPvDe returns the pvDe the DM is using to identify
                this DE to its DES. */
        MSOMETHOD_(void, GetDisplayElementPvDe) (THIS_ MSOHDE hde,
                void **ppvDe) PURE;

        /* SetDisplayElementPvDe may be used to change the pvDe the DM is using
                to identify a particular DE to its DES. */
        MSOMETHOD_(void, SetDisplayElementPvDe) (THIS_ MSOHDE hde,
                void *pvDe) PURE;

        /* FIsDisplayElementVisible returns TRUE if the specified DE is
                currently marked "visible".  Note that the DM doesn't really make
                DEs appear and disappear; rather the client is expected to keep
                the DM up to date about which DEs are visible or not, this allows
                the DM to not waste time on spurious updates. */
        /* Hey! Use GrfdmqdeQueryDisplayElement instead!!!! */
        MSOMETHOD_(BOOL, FIsDisplayElementVisible) (THIS_ MSOHDE hde) PURE;

        /* SetDisplayElementVisible may be used to tell the DM that a particular
                Display Element has become visible or invisible. */
        /* TODO peteren: I changed this to set fVisible AND fVisibleNow.
                I think that's correct in all the places we're currently calling
                this.  Really we should this take another flag called fVisibleNow,
                but I'm trying to avoid interface changes. */
        MSOMETHOD_(void, SetDisplayElementVisible) (THIS_ MSOHDE hde,
                BOOL fVisible) PURE;

        /* Use InvalidateDisplayElement to tell the display manager we need
                to update a particular Display Element.  You describe the invalidation
                in grfdmide.  If you want to specify a particular rectangle of the DE
                you should pass msofdmideRcv or msofdmideRcw, depending on the
                coordinate space of the rectangle you've got, and pass in a pointer
                to your rectangle in prcInvalid.  You should also specify whether
                we need to erase or draw (or both) this rectangle by passing in
                msofdmideErase and/or msofdmideDraw. */
        MSOMETHOD_(BOOL, FInvalidateDisplayElement) (THIS_ MSOHDE hde,
                ULONG grfdmide, const RECT *prcInvalid) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FBeginDisplayElementUpdate) (THIS_ MSOHDE hde,
                MSODC *pdc) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(void, EndDisplayElementUpdate) (THIS_ MSOHDE hde,
                MSODC *pdc) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(void, MoveDisplayElementZOrder) (THIS_ MSOHDE hde,
                ULONG grfdmrde, int idel, MSOHDE hdeNeighbor) PURE;

        /* OffsetDisplayElement translates any "V" coordinates being held by
                a particular Display Element (and it's caches, etc.) along
                a specified vector.  This does no invalidation and causes no updating,
                it just changes coordinate values.  Used, for example, when scrolling
                in Excel (which annoyingly changes the V-space position of all the
                Shapes, instead of changing a V-to-W offset. */
        MSOMETHOD_(void, OffsetDisplayElement) (THIS_ MSOHDE hde, int dxv,
                int dyv) PURE;

#if BOGUS
        // ----- IMsoDisplayManager methods (doomed methods)

        /* Update() is the big important routine that you call to get
                the Display Manager's piece of screen up to date. */
        MSOMETHOD_(void, UpdateRemoveMe) (THIS_ BOOL fPeterEn) PURE;

        /* DrawMeta draws the shapes w/o checking invalid regions. */
        MSOMETHOD_(void, DrawMetaRemoveMe) (THIS_ MSODC *pdc) PURE;

        /* Call Draw to have the DM draw a rectangular portion of itself into
                a DC of your choice (must be compatible with the one this DM might
                get back from DMS::FGetDC).  If you pass in fClip == TRUE we'll
                clip to *prcvDraw, otherwise we'll just try to avoid drawing stuff
                outside *prcvDraw just to be faster.  If you'd rather pass in
                an RCW (relative to the window/DC) just do and pass fRcwDraw == TRUE. */
        MSOMETHOD_(void, DrawRemoveMe) (THIS_ MSODC *pdc, RECT *prcvDraw, BOOL fRcwDraw,
                BOOL fClip) PURE;
#endif // BOGUS

        // ----- IMsoDisplayManager methods stuck at the end to avoid full builds
};

#if DEBUG
MSOAPIX_(void) MsoDebugGetDcAndBlinkRgnRc(HRGN hrgn, const RECT *prcv);
#endif // DEBUG

/****************************************************************************
        Drawing View Site
****************************************************************************/


/* DGVSI (Drawing View Site Info).
        This is a bunch of data which describes this DGV's site.  Rather
        than calling through a bunch of functions to see such data, each DGV
        contains a DGVSI.  When you make a new DGV you have to pass in
        a DGVSI.  After creation you (the host) can use the methods
        DGV::PdgvsiBeginUpdate and DGV::EndDgvsiUpdate to examine and change
        this data.
        You should use MsoInitDgvsi to initialize one of these before
        changing the fields you care to and passing it into DGVG::FCreateView.
        If you (the host) want to be lazy about updating this you can
        use the DGVS::OnDrawingViewEvent method to get called in certain
        circumstances and fix the DGVSI in your DGV.  All of this stuff
        is entirely paralled with Drawings (DG, DGS, DGSI) and Display Managers
        (DM, DMS, DMSI).  */
#define msodxypSnapToGuideDefault 5

MsoDEBDeclare(dgv, msodgeDgvsFirst, msodgeDgvsMax);

/*      HEY!! If you change the fields in the MSODGVSI you probably need to
        change code in MsoInitDgvsi and AssertValidDgvsi. */
typedef struct _MSODGVSI
        {
        interface IMsoDrawingViewSite *pidgvs; // Actual DGVS interface pointer
        void *pvDgvs; // client data passed to all DGVS methods

        interface IMsoDrawingUserInterface *pidgui; // DGUI this DGV should use.

        interface IMsoDisplayManager *pidm;
                /* The Display Manager this DGV should use.  This may be NULL in
                        some cases when the DGV is being used for some strange-ish
                        purpose (Drawing Shapes into a host-provided DC, for instance,
                        or Word's HavePllr case). */

        int idelBackground;
        int idelMain;
        int idelCached;
        int idelXOR;
        int idelTop;
                /* These idel's tell the DGV into which DELs in its DM it should
                        register DEs of various kinds.  These all default to msoidelNil;
                        this is right only if your DGV has no DM and is not to be displayed.
                        You must give valid values for idelMain, idelCached, and idelXOR.
                        You may give a valid value for idelBackground; if you do we'll
                        display this DG's background shape in a special DE in that layer.
                        idelBackground and idelMain need to point to "msodelkMain" DELs,
                        idelCached must point to a "msodelkCached" DEL, and idelXOR must
                        point to a "msodelkXOR" DEL. */

        HWND hwnd;
        MSORCVI rcvi;
                /* The rcvi contains all kinds of useful coordinate information:
                        .rcv = bounding rectangle of DGV in V's (pixels)
                        .rci = bounding rectangle of DGV in I's
                        .cxiUnit = how many I's are there in an arbitrary unit (X-axis)
                        .cxvUnit = how many V's are there in the same unit (X-axis)
                        .cyiUnit = how many I's are there in an arbitrary unit (Y-axis)
                        .cyvUnit = how many V's are there in the same unit (Y-axis)
                        .cxiInch = how many I's are there in an inch (X-axis)
                        .cyiInch = how many I's are there in an inch (Y-axis) */
        POINT ptwOrigin;
                /* At what W point is (0, 0) in V coordinates.  This is redundant
                        in some cases with data the Display Manager has.  Trouble is,
                        not all DGVs have a display manager, and we need this anyway. */
        MSOGV gvPage;
                /* This is for Word.  The host can whatever data it wants here
                        when it creates the View.  We'll pass this data back into
                        DGS methods to create and locate anchors. */
        ULONG grfdraw;  // Flags indicating how to draw (see msofdraw)
        union
                {
                struct
                        {
                        ULONG fWantsEvents : 1;
                                /* Please call DGVS::OnDrawingViewEvent. */
                        ULONG fCanConvertPthToPti : 1;
                                /* DGVS::ConvertPthToPti can work.  If this is FALSE it
                                        means you can't do drags in this DGV. */
                        ULONG fPthEqualsPti : 1;
                                /* DGVS::ConvertPthToPti doesn't change the points.  We
                                        can use this to be faster. */
                        ULONG fUseIntegratedWndProc : 1;
                                /* Host promises to call Office from its WndProc with all the
                                        messages the DrawingView will need to directly
                                        manipulate shapes. */
                        ULONG fCanDrawAttachedText : 1;
                                /* If fCanDrawAttachedText is TRUE, Office will call
                                        DGVS::DrawText for each shape that has text over it
                                        immediately after drawing the shape.  If this is FALSE it
                                        means the host (probably Word) has it's own way of getting
                                        text displayed. */
                        ULONG fDrawShapes : 1;
                                /* This rather odd flag will default to TRUE.  Word will clear
                                        it and everyone else will probably leave it alone.
                                        If this is FALSE, this DGV will continue
                                        to generate SPVs for Shapes, but it will not register
                                        Display Elements for them.  It will continue to register
                                        Display Elements for selection dots, etc.  Word draws Shapes
                                        on it's own schedule using DGV::DrawShape. */
                        ULONG fOmitOffscreenShapes : 1;
                                /* This defaults to FALSE.  If this is TRUE we'll avoid generating
                                        SPVs for Shapes that don't intersect view, and then have to
                                        do work when we scroll.  Excel will probably set this. */
                        ULONG fCanEditTextRotated : 1;
                                /* If fCanEditTextRotated is TRUE, shape will remain rotated
                                        as in normal display when in text edit mode. Default is
                                        FALSE, where the shape will be unrotated but with rotated
                                        bounds while in text edit mode. */
                        ULONG fSelectionVisible : 1;
                                /* The selection (that is, the drag handles, etc.) in this
                                        DGV will be visible whenever this is TRUE.  There are
                                        DGV methods to query and set this (FIsSelectionVisible
                                        and SetSelectionVisible) but it's in the DGVSI so that
                                        you can specify an initial value.  The default is TRUE. */
                        ULONG fIgnoreHits : 1;
                                /* If this bit is TRUE then this DGV won't respond to hits.
                                        Word, which has two DGVs per page (one for the main text
                                        and another for the header/footer) uses this bit to make sure
                                        one of the DGVs is "out of the way" at all times. */
                        ULONG fDontShowHiddenShapes : 1;
                                /* Set this if you want us to not view fHidden shapes in this
                                        view.  This defaults to TRUE. */
                        ULONG fOnlyShowPrintingShapes : 1;
                                /* Set this if you want us to only view fPrinting Shapes.  This
                                        defaults to FALSE. */
                        ULONG fConstrainDirectManipulationToPage : 1;
                                /* Set this if you want direct manipulation to always constrain
                                        shapes to be entirely inside the view. */
                        ULONG fSetNoHiliteInDc : 1;
                                /* If this is TRUE this DGV will always set the msofdrawNoHilite
                                        bit in MSODCs before drawing Shapes into them. */
                        ULONG fShowNoShapes : 1;
                                /* Don't show _any_ Shapes.  Defaults to FALSE. */
                        ULONG fShowShapePlaceholders : 1;
                                /* Show "placeholder" versions of slow-to-draw Shapes.
                                        Defaults to FALSE. */
                        ULONG fShowBackgroundShape : 1;
                                /* Defaults to FALSE.  If this is TRUE we'll check the DG
                                        to see if it has a background Shape and if there is one
                                        we'll generate an SPV for the Background Shape. */
                        ULONG fIgnoreInvalidation : 1;
                                /* This reasonably horribly bit defaults to FALSE.  Word sets
                                        it in the DGVSI used for inline pictures, which tends to
                                        get invalidated and then never validated. */
                        ULONG fScrollChangesViewCoordinates : 1;
                                /* This bit indicates that scrolling will screw up all coordinates.
                                        Currently used only by Excel, as this is pretty evil and to
                                        be avoided if possible. */
                        ULONG fNoEraseXOR : 1; // dwxp
                                /* If this is TRUE we'll not try to erase XOR DEs when
                                        you call DGV::Validate.  This should just be an
                                        argument to DGV::Validate next time (in a grf). */
                        ULONG fNoActivateTextInChildShape :1; //dwXp
                                /* If this is TRUE we'll not try to activate text
                                        on a child shape        */
                        ULONG fLimitRciTo16Bits : 1; // DWXP
                                /* Slightly misleading name.  If this is TRUE we'll pin the rciAnchor
                                        of each SPV in a DGV so that the resulting rcvAnchor is limited
                                        to 16-bit coordinates so that when we draw stuff in the
                                        various lame "32-bit" OSs we run on we don't run into weird
                                        16-bit limitations.  This is only set to FALSE by XL when it's
                                        determining the print area of the sheet. */
                                /* FUTURE peteren: I'd like this better defaulting to FALSE */
                        ULONG fSnapCentersOfShapesToGuides : 1; //dwxP
                                /* If this is TRUE we'll snap shape CENTERs to host-provided
                                        guide points, in addition to (as normally) the left and
                                        right edge. */
                        ULONG fSavingHtml : 1;
                                /* This drawing view is for Saving to HTML */
                        ULONG fDrawWithinClip : 1;
                                /* If set doesn't draw shapes outside of clip rectangle.
                                        Currently only used by PPT (AGIFs) */
                        ULONG fRefCount : 1;
                                /* Adding refcount ability, but the way we handle this must
                                   depend on whether or not the host will honor it. */
                        ULONG fCancelCreateShapeCleanup : 1;
                                /* If set, the host can handle the cleanup when we cancel
                                        the creation of shape creations and return false. */
                        ULONG fInkingView : 1;
                                /* If set, this view can support ink surfaces */
                        ULONG grfUnused : 4;
                        };
                ULONG grf;
                };

        int dxScrollInset;
        int dyScrollInset;
        MsoDEBDefine(dgv);

#if DEBUG
        LONG lVerifyInitDgvsi; // Used to make sure people use MsoInitDgvsi
#endif // DEBUG
        } MSODGVSI;
#define cbMSODGVSI (sizeof(MSODGVSI))

/* You should call MsoInitDgvsi on any new DGVSI you're filling out
        before passing it into DGG:FCreateView.  This lets us add
        new fields without changing code in all the hosts. */
MSOAPI_(void) MsoInitDgvsi(MSODGVSI *pdgvsi);

/* MSOSNPI (MSO SNaP Information)
        This structure holds all the information
*/
typedef struct _MSOSNPI
        {
        struct
                {
                /* Should we snap to grid by default? */
                ULONG fSnapToGrid : 1;

                /* Should we snap to guides by default? */
                ULONG fSnapToGuides : 1;

                /* Should we snap/nudge by on pixel ? */
                ULONG fNudgeByOne : 1;

                /* Should we snap to guides when the grid key is pressed? */
                ULONG fSnapToGuidesOnGridKey : 1;

                /* Should we add the boundaries of our shapes currently in the view to
                        our list of guides? */
                ULONG fMakeGuidesFromShapes : 1;

                /* Does the host want to specify a distance to snap by? */
                ULONG fSnapDistance : 1;

                /* Does the hose only want snapping on shape moves? */
                ULONG fSnapOnMoveOnly: 1;

                /* Ignore grid on nudge, default is grid values take precedence if both set */
                ULONG fIgnoreGridOnNudge: 1;

                ULONG grfUnused : 24;
                };

        /* What is the grid resolution factor relative to the D coordinates? */
        USHORT gridResFactor;

        /* What is the grid origin in HiRes D coordinates? */
        int xdGridOrigin;
        int ydGridOrigin;

        /* What is the grid size in HiRes D coordinates? */
        int dxdGrid;
        int dydGrid;

        /* How close do we need to be to the guide before we snap to it? */
        int dxypSnapToGuide;

        /* px is the pointer to the X-coordinate of the first guide */
        long *px;
        /* cx is the number of X-coordinate guides we have in this direct
                manipulation */
        int cx;
        /* cbx is the number of bytes we have between X-coordinate guides.  If
                you have an array of fixed size structures which contain your guides,
                and the X-coordinate sits in the same place in every structure, then
                you can make px point to the first X-coordinate in the structure,
                cx the number of structures in your array and cbx the size of the
                structure, and this will save you an allocation.  Otherwise, you
                should just allocate the array of ints, point px to the beginning,
                set cx to the number of elements and set cbx to sizeof(int) */
        ULONG cbx;

        /* py is the pointer to the Y-coordinate of the first guide */
        long *py;
        /* cy is the number of Y-coordinate guides we have in this direct
                manipulation */
        int cy;
        /* cby is the number of bytes we have between Y-coordinate guides.  If
                you have an array of fiyed size structures which contain your guides,
                and the Y-coordinate sits in the same place in every structure, then
                you can make py point to the first Y-coordinate in the structure,
                cy the number of structures in your array and cby the size of the
                structure, and this will save you an allocation.  Otherwise, you
                should just allocate the array of ints, point py to the beginning,
                set cy to the number of elements and set cby to sizeof(int) */
        ULONG cby;

        /* Distance to nudge by if fSnapDistance is TRUE and the grid is turned off */
        ULONG dxdNudge;
        ULONG dydNudge;
        } MSOSNPI;

/* You should call MsoInitSnpi on any new SNPI you're filling out.
        This way        we can add new fields to the structure without adding
        new code in all the hosts; just a recompile will be sufficent. */
MSOAPIX_(void) MsoInitSnpi(MSOSNPI *psnpi);


/* MSOCRS = Cursor
        These are an enumeration for all the cursors we'd like our host to be
        able to display.  Office Drawing doesn't actually contain any cursors
        or ever set the cursor, it just passes these values back to the host. */
/* TODO peteren: Move these somewhere more general in Office? */
/* TODO peteren: Review these? */
typedef enum
        {
        msocrsNil = 0,
        msocrsArrow,
        msocrsHourGlass,
        msocrsSizeWestEast,
        msocrsSizeNorthSouth,
        msocrsSizeNWSE,
        msocrsSizeNESW,
        msocrsCrossHair,
        msocrsAdjust,
        msocrsIBeam,
        msocrsRotateTool,
        msocrsRotateMode,
        msocrsRotPerf,
        msocrsSelMove,
        msocrsMovePerf,
        msocrsConnectTool,
        msocrsCropTool,
        msocrsCropMode, // this one might not actually be used anymore
        // CAUTION:: Keep the crop cursors in the same order. They
        // have a one-one correspondence with the values assigned to msodot*.
        msocrsCropTL,
        msocrsCropTR,
        msocrsCropBR,
        msocrsCropBL,
        msocrsCropTC,
        msocrsCropBC,
        msocrsCropCL,
        msocrsCropCR,
        msocrsCropCC, // what is this for?
        msocrsMovePicture,
        msocrsTarget, // Cross hair with a hole
        msocrsButton,
        msocrsFreehand,
        msocrsReshapeMode,
        msocrsEditVertex,
        msocrsDragVertex,
        msocrsMoveVertex,
        msocrsAddVertex,
        msocrsDelVertex,
        msocrsCopy,
        msocrsFormatPainter,
        msocrsMagicWand,
        msocrsMagicWandHit,
        msocrsToolText,
        msocrsNo,
        msocrsInk
        } MSOCRS;

/* TODO #define's for old names, replace with new names */
#define msocrsRotTool msocrsRotateTool
#define msocrsRotMode msocrsRotateMode

/* We used to call MSOCRSs MSODGCRSs.  I haven't gotten
        rid of all the references.  So for now here's some
        #defines to make references to the old values work. */
#define MSODGCRS MSOCRS
#define msodgcrsNil msocrsNil
#define msodgcrsArrow msocrsArrow
#define msodgcrsHourGlass msocrsHourGlass
#define msodgcrsSizeWestEast msocrsSizeWestEast
#define msodgcrsSizeNorthSouth msocrsSizeNorthSouth
#define msodgcrsSizeNWSE msocrsSizeNWSE
#define msodgcrsSizeNESW msocrsSizeNESW
#define msodgcrsCrossHair msocrsCrossHair
#define msodgcrsAdjust msocrsAdjust
#define msodgcrsIBeam msocrsIBeam
#define msodgcrsRotTool msocrsRotTool
#define msodgcrsRotMode msocrsRotMode
#define msodgcrsRotPerf msocrsRotPerf
#define msodgcrsSelMove msocrsSelMove
#define msodgcrsButton  msocrsButton
#define msodgcrsCropTL  msocrsCropTL
#define msodgcrsCropTC  msocrsCropTC
#define msodgcrsCropTR  msocrsCropTR
#define msodgcrsCropCL  msocrsCropCL
#define msodgcrsCropCC  msocrsCropCC
#define msodgcrsCropCR  msocrsCropCR
#define msodgcrsCropBL  msocrsCropBL
#define msodgcrsCropBC  msocrsCropBC
#define msodgcrsCropBR  msocrsCropBR
#define msodgcrsNo      msocrsNo

/* MSODGVHD -- Drawing View Hit Description.
        These are returned from GetHitDescription to allow the host to pick
        a good cursor or otherwise process a hit on a Shape.  They can
        also use this data to help decide whether or not to call
        FDoMouseDown. */
typedef struct MSODGVHD
        {
        MSOCRS crs; // The cursor we think you should use
        MSOHSP hsp; // The shape we hit
        MSOHSPV hspv; // Handle to the shape we hit relative to the view
        // TODO peteren: Put more stuff hosts want here.
        } MSODGVHD;
#define cbMSODGVHD (sizeof(MSODGVHD))

/* MSODGVEB -- Drawing View Event Block */
typedef struct MSODGVEB
        {
        /* These first fields are filled out for EVERY
                Drawing View Event. */
        MSODGE dge;
        BOOL fResult;
        interface IMsoDrawingView *pidgv;
        MsoDEBPDefine(dgv);

        /* Only one set of fields in this union will be
                filled out, depending on .dge. */
        union
                {
                struct // fields for all old-style events
                        {
                        WPARAM wParam;
                        LPARAM lParam;
                        };
                struct
                        {
                        RECT *pdrch;
                        interface IMsoDrawingSelection *pidgsl;
                        RECT *prchOld;
                        int cShapesSelOrig;
                        } dgexDoAdjustDrchForPaste;
                struct
                        {
                        MSOHSP hsp;
                        MSOHSPV hspvOld;
                        MSOHSPV hspvNew;
                        RECT rcvBoundsOld;
                        RECT rcvBoundsNew;
                        } dgexAfterChangeShapeOnView;
                struct
                        {
                        /* Fields for msodgeRequestCancelButtonDrag and
                                msodgeAfterButtonShapeClick */
                        MSOHSP hspButton;
                        MSOHSPV hspvButton;
                        MSOMM mmButton;
                        // TODO (sreeramn): Maybe later we need to send in the point
                        // representing the click.
                        };
                struct // fields for msodgeDragShape*
                        {
                        int cShapesBeingDragged;
                        MSOGV gvHost;
                        MSODRGH drgh;

                        // The following two fields are only used for message on a single shape
                        RECT rcvShape;
                        POINT ptvDrag;

                        MSOHSP hspDrag;
                        MSODGCID dgcid;
                        union
                                {
                                struct
                                        {
                                        BOOL fOleDrag : 1;
                                        BOOL fInsertDrag : 1;
                                        };
                                ULONG grf;
                                };

                        long fxpPercentX;
                        long fypPercentY;

                        // Donot change the order of these crop params.
                        long fxpCropFromLeft;
                        long fxpCropFromRight;
                        long fxpCropFromTop;
                        long fxpCropFromBottom;
                        POINT pthDrag; // emu coordinates ( this is only used for messages on a single shape)
                        };
                struct // fields for msodgeAfterHitTestShape // was msodgeRequestHitTestShape
                        {
                        MSOHSP hsp; // was hspHitTest
                        MSOHSPV hspv; // was hspvHitTest
                        MSOSVI *psvi; // was psviHitTest
                        MSOHSP hspRoot; // new
                        MSOHSPV hspvRoot; // new
                        MSOPOINT ptv; // was ptvHitTest
                        MSOMM mm; // was mm
                        MSODGCID dgcidTool; // new
                        MSODRGH drgh; // DGVS may change this, msodrghNil means no hit
                        } dgexAfterHitTestShape;
                struct // fields for msodgeDoClickShape
                        {
                        MSOHSP hspRoot; // was hsp
                        MSOHSP hspDeep; // new
                        MSOHSP hspButton; // was hspWithButton
                        MSODGCID dgcidTool; // was dgcid
                        MSODRGH drgh;
                        MSODMHD *pdmhd;
                        } dgexDoClickShape;
                struct // fields for msodgeDoScrollOrDragDrop
                        {
                        BOOL fScroll;
                        BOOL fDragDrop;
                        MSOPOINT ptwMouse;
                        HWND hwnd;
                        BOOL fCanDragDrop;
                        } scrollordragdrop;
                struct // fields for msodgeDoPaste
                        {
                        MSOCSD *pcsd;
                        interface IMsoDrawingSelection *pidgsl;
                        };
                struct
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        MSOSVI *psvi;
                        } dgexRequestPositionSelection;
                struct
                        {
                        /* msodgeRequestPositionText is fired from under DGV::Validate
                                after Office has decided where a shape goes in the view
                                and figured out where it's text frame is.  The text frame
                                isn't necessarily where the text should really go, it's just
                                a starting point.  The DGS should look at all these fields
                                and if the text belongs elsewhere it should change
                                *prcvTextBounds to the pixel boundaries of the text attached
                                to this object. The DGV won't pass this event if the shape
                                doesn't have text attached.
                                Set dgveb.fResult to TRUE if you want the DGV to respond
                                to any new values. */
                        /* FUTURE peteren: Explain prciTextFrame. */
                        MSOHSP hsp;
                        void *pvClient;
                        MSOTXID txid;
                        MSOSVIT *psvit;
                        BOOL fForDiagram;// TODO: gussand review
                        RECT *prcvTextBounds;
                        RECT *prciTextFrame;
                        } dgexRequestPositionText;
                struct // fields for msodgeRequestUnionShapeCoverage
                        {
                        MSOHSP hsp;
                        MSOHSPV hspv;
                        MSOSVI *psvi;
                        MSOCVRK cvrk;
                        HRGN hrgnv; // return
                        BOOL fComplete; // return
                        BOOL fError; // return
                        } dgexRequestUnionShapeCoverage;
                struct // for msodgeRequestCancelViewShape
                        {
                        MSOHSP hspCancelView;
                        void *pvClientCancelView;
                        };
                union
                        {
                        /* Hey! Until we get rid of the unnamed version of this struct
                                be careful to keep the next two structures lined up! */
                        struct // for msodgeAfterGetShapeBounds
                                {
                                MSOHSP hsp;
                                void *pvClient;
                                MSOSVI *psvi;
                                RECT *prcv;
                                } dgexAfterGetShapeBounds;
                        struct // for msodgeAfterGetShapeBounds
                                {
                                MSOHSP hspGetBounds;
                                void *pvClientGetBounds;
                                MSOSVI *psviGetBounds;
                                RECT *prcvGetBounds;
                                };
                        };
                struct // for msodgeAfterGetTopLevelGroupShapeBounds
                        {
                        MSOHSP  hsp;
                        void    *pvClient;
                        MSOSVI  *psvi;
                        RECT    *prcv;
                        } dgexAfterGetTopLevelGroupShapeBounds;
                union
                        {
                        /* Hey! Until we get rid of the unnamed version of this struct
                                be careful to keep the next two structures lined up! */
                        struct // for msodgeDoDrawShape, msodgeAfterDrawShape
                                {
                                MSOHSP hsp;
                                MSOHSPV hspv;
                                void *pvClient;
                                MSOSVI *psvi;
                                MSODC *pdc;
                                BOOL fAborted;
                                BOOL fSavingHtml;
                                } dgexDrawShape;
                        /* The following unnamed struct are the OLD FIELDS.  We should
                                get the hosts to use dgexDrawShape instead and then get rid
                                of this. */
                        struct // old field names for msodgeDoDrawShape, msodgeAfterDrawShape
                                {
                                MSOHSP hspDraw;
                                MSOHSPV hspvDraw;
                                void *pvClientDraw;
                                MSOSVI *psviDraw;
                                MSODC *pdcDraw;
                                };
                        };
                struct // for msodgeQueryDragCopyShape
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        } dgexQueryDragCopyShape;
                struct
                        {
                        ULONG grfdgiDg;
                                /* All invalid bits in DG. */
                        ULONG grfdgiSp;
                                /* Invalid bits collected from SPs on view in this DGV. */
                        } dgexAfterDgvCopeWithDgInval;
                struct
                        {
                        MSOHSP hsp;
                        MSOHSPV hspv;
                        BOOL fOpaque; // Host may change this field
                        } dgexAfterIsShapeOpaque;
                struct
                        {
                        MSOHSP hsp;
                        MSOHSPV hspv;
                        ULONG grfspvs;
                        } dgexAfterDgvChooseGrfspvs;
                struct // for msodgeCalcRchPage
                        {
                        RECT *prchPage;
                        } dgexCalcRchPage;
                struct // for msodgeRequestSelectionForDrag
                        {
                        interface IMsoDrawingSelection *pidgslUI;
                        interface IMsoDrawingSelection *pidgslDrag;
                        MSODRGHK drghk;
                        } dgexRequestSelectionForDrag;
                struct // for msodgeAfterPaste
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        int cShapes;
                        } dgexAfterPaste;
                struct // for msodgeDoSelectShape
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        MSOHSP hsp;
                        } dgexDoSelectShape;
                struct // for msodgeDoUnselectShape
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        MSOHSP hsp;
                        } dgexDoUnselectShape;
                struct // for msodgeDoChangeDgslMode
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        MSODGSLK dgslkOld;
                        MSODGSLK dgslkNew;
                        } dgexDoChangeDgslMode;
                struct // for msodgeAfterChangeDgsl
                        {
                        interface IMsoDrawingSelection *pidgsl;
                        int cspBefore;
                        int cspAfter;
                        MSOHSP hspFocusBefore;
                        MSOHSP hspFocusAfter;
                        MSODGSLK dgslkBefore;
                        MSODGSLK dgslkAfter;
                        BOOL fSetChanged; // TRUE if different shapes are selected now
                        BOOL fDrillDownBefore;
                        BOOL fDrillDownAfter;
                        } dgexAfterChangeDgsl;
                struct // for msodgeBeforeRegisterShapeDEs
                        {
                        MSOHSP hsp;
                        void *pvClient;
                        ULONG grfdmrdeShape;
                        ULONG grfdmrdeText;
                        } dgexBeforeRegisterShapeDEs;
                struct
                        {
                        IMsoDrawing *pidgSrc;
                        IMsoDrawing *pidgDest;
                        POINT pptvDrop;
                        MSOMM mm;
                        int   dxh;        // Difference in X coordinate from beginning to end of drag (Pub needs)
                        int   dyh;        // Difference in Y coordinate from beginning to end of drag (Pub needs)
                        BOOL  fOleDrop;   // Is this a cross-window OLE drop or a normal drop?
                        BOOL  fCloning;   // Are we going to clone shapes?
                        } dgexRequestCancelDrop;
                struct
                        {
                        MSOHSPV hspv;
                        } dgexAfterRequestSpvRedraw;
                struct // for msodgeRequestPosition
                        {
                        MSOHSP hsp;
                        MSOHSPV hspv;
                        RECT rcv; // in view coordinates
                        POINT ptImgOffset;
                        BOOL fAlignTable; // if true do left alignment of the table if this
                                                                        // shape is positioned in a table.
                        BOOL fMayNeedWrapBlock; // if true the VML for the shape may need to be
                                                                                        // within a o:wrapblock containing element
                        BOOL fPageBreak;        // if true the VML for the shape needs to be
                                                                                        // within a o:wrapblock containing element with
                                                                                        // pagebreak = "t"
                        BOOL fUseSpaceInFlow;   // if true the shape takes up space in the flow
                        LONG lZIndex;
                        } dgexRequestPosition;
                struct
                        {
                        interface IMsoBitmapSurface *pibms;
                        MSOHSPV hspv;
                        } dgexAddTextboxColorsBMS;
                struct
                        {
                        MSOHSPV hspv;
                        MSODC *pmsodc;
                        } dgexDrawTextboxBMS;
                struct
                        {
                        MSOHSP hsp;                                // Scroll this shape to view
                        RECT   rcw;             // bound of the shape in w unit
                        interface IMsoDrawingView *pidgv; // current DGV (Word might set it on a page scroll)
                        } dgexScrollToShape;
                struct  // for msodgeDoDisplayRightClickDragMenu
                        {
                        int                     dxh;        // Difference in X coordinate from beginning to end of drag
                        int                     dyh;        // Difference in Y coordinate from beginning to end of drag
                        POINT           ptvMouse;      // Where the mouse went up in the view -- for menu positioning.
                        } dgexRightClickMenu;
                struct  // for msodgeBeforeShapeFill
                        {
                        IMsoDrawing **ppidg;    // background drawing
                        } dgexBeforeShapeFill;
                struct // for msodgeRequestDrag
                        {
                        IDataObject *piDataObject;    // [out] the Escher IDataObject we were going to use
                        DWORD dwEffect;            // [in] the effect of the drag
                        BOOL    fSameWindow;            // [in] did we end up in the same window (resume Escher Drag)?
                        BOOL    fLocalDrop;                     // We did a drop in the local window (not used currently)
                        } dgexRequestDrag;
                struct // for msodgeBeforeSnappingShape
                        {
                        POINT ptTop;
                        POINT ptBotRight;
                        } dgexSnappingShape;
                struct // for msodgeRequestShapeDragRch
                        {
                        MSOHSP hsp;
                        RECT *prchDrag;
                        } dgexShapeDragRch;
                struct  // for msodgeBeforeColorScheme
                        {
                        IMsoDrawing **ppidg;    // color scheme drawing
                        } dgexBeforeColorScheme;
                struct // for msodgeRequestCancelShapeRotateKnob
                        {
                        MSOHSP hsp;                     // Shape requesting knob
                        BOOL   fDisplay;        // wheter to display knob (defaults to true).
                        } dgexCancelShapeRotateKnob;
                struct  //msodgeCanDisplayPlaceholderText
                        {
                        MSOHSP hsp;
                        MSOTXID txid;
                        }  dgexDisplayPlaceholderText;
                 struct // for msodgeRequestInputRcv
                        {
                        RECT rcvInput;
                        } dgexeRequestInputRcv;
                 struct // for msodgeRequestInkSelectionUpdate
                        {
                        MSOHSP hsp;
                        RECT rcv;
                        } dgexRequestInkSelectionUpdate;
                struct // for msodgeRequestRcviFromZoom
                       {
                       MSORCVI rcvi;
                       ULONG cxiInch;
                       ULONG cxiInchN;
                       ULONG cxiInchD;
                       ULONG cyiInch;
                       ULONG cyiInchN;
                       ULONG cyiInchD;
                       } dgexRequestRcviFromZoom;
                };
        } MSODGVEB;
#define cbMSODGVEB (sizeof(MSODGVEB))

/* some structure passed as parameter to IMsoDrawingView::FDrawShapes and
        a very few of other IMsoDrawingView methods.  The author never documented
        what the use of ths structure is... */
// TODO davebu(asheshb): document this structure
typedef struct MSODGVFDSC
        {
        BOOL fContinue;
        int ispv;
        int ispvRoot;
        struct SPV *pspv;
        MSOTXID txid;
        } MSODGVFDSC;

/* initializer to fill in starting value */
MSOAPI_(void) MsoInitDgvfdsc(MSODGVFDSC *pdgvfdsc);

/* a structure that holds miscellaneous (infrequently used) parameters for
        IMsoDrawingView::FDrawShapes.  If you want to add infrequently-used
        parameters to FDrawShapes that have obvious default values, change the
        typedef below to a struct and add your new fields.  THis will be easier
        thn adding the new param directly to fdrawshapes.  */
typedef MSOGECACHEHINT MSODGVFDSPARAM;

#define msofdrgsdBottom  0x1
#define msofdrgsdTop     0x2
#define msofdrgsdLeft    0x4
#define msofdrgsdRight   0x8
#define msogrfdrgsdTL    (fdrgsdTop | fdrgsdLeft)
#define msogrfdrgsdTR    (fdrgsdTop | fdrgsdRight)
#define msogrfdrgsdBL    (fdrgsdBottom | fdrgsdLeft)
#define msogrfdrgsdBR    (fdrgsdBottom | fdrgsdRight)

/* IMsoDrawingViewSite must be implemented by anyone who wants to make
        and use a Drawing View.  Fill out a MSODGVSI (use MsoInitDgvsi)
        and add a pointer to your IMsoDrawingViewSite, and then pass that
        off to IMsoDrawingViewGroup::FCreateView.
        To enable the host to simulate many "virtual" DGVSs with a single
        actual object, all the methods in IMSoDrawingViewSite take a void *
        named pvDgvs as their first argument (well, second after the "this"
        pointer).  When you create a DGV you also pass (in the MSODGVSI)
        a pvDgvs, which will be passed back to all DGVS methods. */
#undef  INTERFACE
#define INTERFACE IMsoDrawingViewSite
DECLARE_INTERFACE_(IMsoDrawingViewSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingViewSite methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* The Drawing View will call DGVS::OnEvent whenever a drawing view
                event (MSODGE) occurs.  See the definition of MSODGE for a more
                complete explanation of events.  This function will only be called
                if you set msodgvsi.fWantsEvents to TRUE. */
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDgvs, MSODGVEB *pdgveb) PURE;

        // ----- IMsoDrawingViewSite methods (other)

        /* TODO peteren: These methods are only here so that we can
                implement IMsoDragSite for doing drags through the DGV.
                We should perhaps think about these.  It's certainly more
                correct that we get these functions in DGVS than that the host
                give us a whole pre-packaged DRGS. */
        MSOMETHOD_(BOOL, FGetDc) (THIS_ void *pvDgvs, MSODC *pdc) PURE;
        MSOMETHOD_(void, ReleaseDc) (THIS_ void *pvDgvs, MSODC *pdc) PURE;
        MSOMETHOD_(void, ConvertPtiToPth) (THIS_ void *pvDgvs,
                const POINT *ppti, POINT *ppth, int cpt) PURE;
        MSOMETHOD_(void, ConvertPthToPti) (THIS_ void *pvDgvs,
                const POINT *ppth, POINT *ppti, int cpt) PURE;
#if FUTURE
        MSOMETHOD_(BOOL, FBeginDrag) (THIS_ void *pvDgvs) PURE;
                /* Called when a drag is beginning in the DGV.  This host
                        should capture the mouse on behalf of the DGV's hwnd
                        and, if fWillCallFWndProc is TRUE, begin sending
                        the DGV messages through DGV::FWndProc.  Replaces
                        FSetHostCapture. */
        MSOMETHOD_(void, EndDrag) (THIS_ void *pvDgvs) PURE;
                /* Called when a drag is ending (at least as far as a particular
                        DGV is concerned).  The host should release the capture
                        and stop sending the DGV mouse messages.  Replaces
                        ReleaseHostCapture. */
#endif // FUTURE

        // ----- IMsoDrawingViewSite methods (display)

        /* Office Drawing calls DGVS::RequestUpdate when it thinks
                a particular view should be updated.  Your implementation
                of this should be fast in cases when there's nothing that needs
                updating.
                Calling this replaces all the cases where we used to
                directly tell a DM to update.  It turns out we need to
                let the host in there. */
        MSOMETHOD_(void, RequestUpdate) (THIS_ void *pvDgvs) PURE;


        // ----- IMsoDrawingViewSite methods (coordinates)

        /* FLocateAnchorInView should look at the passed in pvAnchor and
                determine the appropriate RCI (upright bounding rectangle)
                for this Shape in this View. An RCI is possibly higher-resolution
                version of an RCV; when a Drawing View is created the host passes
                in some scale factors to convert between PTIs and PTVs.  This
                function is pretty similar to DGS::FLocateAnchor, but that
                one locates the pvAnchor in global "document" coordinates, these might
                be used, say, to implement the "Group" command.  FLocateAnchorInView
                is used only to display the Shape in a particular view.  Note that
                a host could decide to make a particular shape invisible in a
                particular view by returning FALSE from this.  But I think we should
                have a more efficient mechanism for a host's deciding which Shapes
                should show in a particular View. */
        MSOMETHOD_(BOOL, FLocateAnchorInView) (THIS_ void *pvDgvs,
                void **ppvAnchor, RECT *prci, MSOHSP hsp) PURE;

        /* Get the rectangle of the current view. This is used when doing
                a drag to determine if we need to scroll or not */
        /* Hey! This method is remarkably ill-defined.  DaveBu's using
                it in the drag code to be the bounds of the window in which the
                view lives, basically the rectangle which if we get outside
                we should scroll.  I'm not sure what units you're supposed to
                be passing in either, although I think "v" is correct. */
        /* TODO davebu(peteren): We might could be a little more careful here.
                Maybe this should be called GetViewWindowRect and expect an rcw? */
        MSOMETHOD_(void, GetViewRect) (THIS_ void *pvDgvs, RECT *prcv) PURE;


        // ----- IMsoDrawingViewSite methods (textboxes)

#if BOGUS
        /* Get textbox's bounding rect given the txid, hsp, and rciTextFrame (inscribed rect). */
        MSOMETHOD_(void, GetRciTextBounds) (THIS_ void *pvDgvs,
                MSOTXID txid, MSOHSP hsp, RECT *prciTextFrame, RECT *prciTextBounds) PURE;
#endif //BOGUS

        /* Draw the textbox pointed to by ptxbx into this HDC and RECT. */
        /* This method is flawed; it takes a pdgs and no pvDgs, without which
                the host may not be able to identify their pdgs. */
        /* TODO peteren: Also this should take an MSOSVIT so that the host can
                get rotation right without calling back some expensive function. */
        MSOMETHOD_(void, DrawAttachedText) (THIS_ void *pvDgvs, MSODC *pdc, MSORCVI *prcvi,
                interface IMsoDrawingSite *pdgs, MSOTXID txid, MSOHSP hsp) PURE;

        /* Return TRUE if ptv hits text */
        MSOMETHOD_(BOOL, FHitTestText) (THIS_ void *pvDgvs, MSOHSP hsp,
                MSOTXID txid, const POINT *pptv, const MSOSVIT *psvit, MSOHTK *phtk) PURE;
#if BOGUS
        /* Returns TRUE if the point is in the textbox. */
        MSOMETHOD_(BOOL, FIsPtiInText) (THIS_ void *pvDgvs,
                MSOTXID txid, MSOHSP hsp, POINT *ppti) PURE;
#endif // BOGUS

        /* Returns TRUE if the textbox takes the mouse down. */
        /* TODO!!! The fact that this takes a pointer to a drawing site and NOT
                a pvDgvs leads me to believe that this is wack.  Also it should take
                a psvit. */
        MSOMETHOD_(BOOL, FDoMouseDownOnText) (THIS_ void *pvDgvs,
                interface IMsoDrawingSite *pdgs, MSOTXID txid, MSOHSP hsp, MSODMHD *pdmhd) PURE;


        // ----- IMsoDrawingViewSite methods (OLE objects)

        /*
        The host should activate the specified OLE object with the given verb.
        Office has initiated (e.g. double clicked on) the edit of an OLE shape.
        */
        MSOMETHOD_(BOOL, FActivateOleObject) (THIS_ void *pvDgvs, MSOOID oid,
                MSOHSP hsp, int iVerb) PURE;

        /*
        The given event flags at the given point have been done on the specified
        OLE object. If the host absorbs the event, it should return TRUE.
        Excluding OCX controls, most objects should just return FALSE and let
        Office deal with the event.
        */
        MSOMETHOD_(BOOL, FDoEventOnOleObject) (THIS_ void *pvDgvs, MSOOID oid,
                MSOHSP hsp, const POINT *pptw, int grfem) PURE;

        /*
        The given event flags at the given point have been done within the bounds
        of the specified OLE object. The host should return TRUE if the point
        actually hit on the object. In general only OLE 3 objects with regions
        where the point is outside of its region will ever want to return FALSE.
        */
        MSOMETHOD_(BOOL, FHitTestOleObject) (THIS_ void *pvDgvs, MSOOID oid,
                MSOHSP hsp, POINT *pptw, int grfem) PURE;

        /* FDoContextMenu replaces FDoContextOnShape */
        /* Office Drawing will call FDoContext when it sees a the correct
                sort of click (right-click on Windows) for a context menu.
                dgcxm is an enumeration telling you which context menu to bring up.
                pptw tells you the coordinates (in view-relative coordinates)
                of the click, hsp is the Shape that was hit, and mm contains
                the various modifiers of the click.
                You should return TRUE if you put up the menu, FALSE otherwise. */
        MSOMETHOD_(BOOL, FDoContextMenu) (THIS_ void *pvDgvs, MSODGCXM dgcxm,
                POINT *pptv, MSOHSP hsp, MSOMM mm) PURE;


        // ----- IMsoDrawingViewSite methods (direct manipulation)

        /* Convert hi-res view coordinates into the coordinate system that you'd
                like us to do our direct manipulation math in.  This coordinate system
                needs to work with FChangeViewAnchor */
        MSOMETHOD_(void, ConvertPtiToPtd) (THIS_ void *pvDgvs,
                const POINT *ppti, POINT *pptd, int cpt) PURE;
        MSOMETHOD_(void, ConvertPtdToPti) (THIS_ void *pvDgvs,
                const POINT *pptd, POINT *ppti, int cpt) PURE;

        /* Tell me all about the kind of snap-to-grid information for this drag.
                This includes any and all guide information.  If this routine
                returns TRUE, we MUST call ReleaseSnapInfo afterwards. */
        MSOMETHOD_(BOOL, FGetSnapInfo) (THIS_ void *pvDgvs, MSOSNPI *psnpi,
                interface IMsoDrawingSelection *pidgsl) PURE;

        /* I'm done with the snap-to-grid information, and you can free any
                allocations you had to do to give it to me. */
        MSOMETHOD_(void, ReleaseSnapInfo) (THIS_ void *pvDgvs, MSOSNPI *psnpi) PURE;

        /* Please set the cursor to the following enum. */
        MSOMETHOD_(void, RequestSetCursor) (THIS_ void *pvDgvs, MSODGCRS dgcrs) PURE;

        /* Please scroll for me.  I'll pass the direction I want you to scroll
                in the grfdrgsc.  Return TRUE if you scroll and FALSE if you don't.
                In most cases, I don't care how far you scroll, but if you have a
                coordinate system that is zero based at the window origin, then I
                need to know how much you scrolled, and you should put that result in
                pptwScroll.  Otherwise, you can just leave it blank. */
        MSOMETHOD_(BOOL, FScrollView) (THIS_ void *pvDgvs, POINT *pptwMouse,
                POINT *pptwScroll, int grfdrgsd) PURE;

        // ----- IMsoDrawingViewSite methods (blips)

        /* AfterCreateBdp is called after an MSOBDRAWPARAM has been created to
                control the rendering of a particular blip.  The client can set
                the client callback data in the MSOBDRAWPARAM before it is used.  In
                fact the client can update any information in the MSOBDRAWPARAM, however
                the main intention is to support client recoloring of the blip.
                A default implementation will just return. */
        MSOMETHOD_(void, AfterCreateBdp) (THIS_ void *pvDgvs, MSOBDRAWPARAM *pbdp,
                const IMsoBlip *pib, MSOHSP hsp, MSOPID pidBlip) PURE;

        /* Return TRUE if the offset from from Window coordinates to the coordinates
                in your message pump is different from the number stored in the MSODGVSI
                you passed to the DGV.  In general, this is only true if you don't have
                a WLM HWND wrapper for each window you create on the Mac. In other
                words, most people return FALSE and don't care. */
        MSOMETHOD_(BOOL, FGetMessageOffset) (THIS_ void *pvDgvs, SIZE *pdpt) PURE;
};


/****************************************************************************
        Drawing View

        All DGVs are created by IMsoDrawingViewGroup::FCreateView.  They're
        freed by calling IMsoDrawingView::Free().
        TODO peteren: Comment?
****************************************************************************/

/* grfdgvr flags are passed to DGV::HspvRelated */
#define msogrfdgvrNil    0
#define msofdgvrTopLevel (1<<0)
        /* Return the top-level Shape above hspv, or hspv if it's already
                top-level. */
#define msofdgvrParent   (1<<1)
        /* Return the immediate parent of hspv, or msohspvNil if hspv
                has no parent. */

/* MSOFSPVM = flags describing weird behavior we want some shapes to
        show in some views. */
#define msogrfspvmNil    0
#define msofspvmAbove    (1<<0)
#define msofspvmTop      (1<<1)
#define msofspvmRotate0  (1<<2)
#define msofspvmRotate90 (1<<3)
#define msogrfspvmAll \
        (msofspvmAbove|msofspvmTop|msofspvmRotate0|msofspvmRotate90)

typedef struct _MSODGVEN
        {
        MSOHSPV hspv;
        MSOHSP hsp;

        void *pv1;
        void *pv2;
        void *pv3;
        void *pv4;
        void *pv5;
        } MSODGVEN;

#define msogrfdgvdsNil 0
#define msofdgvdsRcwDraw  (1<<0)
#define msofdgvdsClip     (1<<1)
#define msofdgvdsValidate (1<<2)

/* MSODGVDMHD == DGV extenstion to DMHD.  We stuff one of these into
        the rgb field at the end of the DMHD. */
// HEY!! Keep the fields in the DGVDMHD and MSODGVDMHD in ssync.
typedef struct _MSODGVDMHD
        {
        MSODGCID dgcidTool;
        MSODRGH drgh;
        MSOHSPV hspv;
        } MSODGVDMHD;
#define cbMSODGVDMHD (sizeof(MSODGVDMHD))

#define MsoPdgvdmhdOfPdmhd(pdmhd) ((MSODGVDMHD *)&(pdmhd)->rgbDes[0])

#define msofhtsWindowCoordinates 1
#define msofhtsForceHit          2
#define msofhtsNoChildren        4
#define msofhtsHitText           8


/* IMsoDrawingView (idgv).  This is the interface spoken by DGVs. */
#undef  INTERFACE
#define INTERFACE IMsoDrawingView
DECLARE_INTERFACE_(IMsoDrawingView, IUnknown)
{
        // ----- IUnknown methods
        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(ULONG, AddRef) (THIS) PURE;
        MSOMETHOD_(ULONG, Release) (THIS) PURE;

        // ----- IMsoDrawingView methods

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Free() tells the drawing view to go away.  We're leaving this puppy in
           even though the preferred way to handle this interface is to use AddRef
           and Release.  Eventually, all the host apps should migrate to using the
           safer refcounting mechanism. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* PdgvsiBeginUpdate returns a pointer to an MSODGVSI.  The caller
                may examine or (if they passed FALSE for fReadOnly) change it's
                contents.  In any case the caller must then call EndDgvsiUpdate.
                See comments by the declaration of MSODGVSI. */
        MSOMETHOD_(MSODGVSI *, PdgvsiBeginUpdate) (THIS_ BOOL fReadOnly) PURE;

        /* EndDgvsiUpdate should be called only after FBeginDgvsiUpdate has
                succeeded.  You should pass back the pointer FBeginDgvsiUpdate gave
                you. In fChanged you should pass back TRUE if you changed anything.
                See comment by FBeginDgvsiUpdate. */
        MSOMETHOD_(void, EndDgvsiUpdate) (THIS_ struct _MSODGVSI *pdgvsi,
                BOOL fChanged) PURE;

        /* Call PidgGetDrawing to get the Drawing viewed by this Drawing View. */
        MSOMETHOD_(interface IMsoDrawing *, PidgGetDrawing) (THIS) PURE;

        /* PidgslGetSelection returns an IMsoDrawingSelection pointer
                to the selection being used by this Drawing View.  It will NOT
                (so there official COM rules) have been AddRef()'d. */
        MSOMETHOD_(interface IMsoDrawingSelection *, PidgslGetSelection)
                (THIS) PURE;

        /* SetSelection tells this Drawing View to begin using the
                DGSL you specify as its selection.  A single DGSL can be shared
                between several DGVs.  We'll AddRef() the DGSL you pass in
                and Release() the one we used to be holding onto.
                If you pass in fOtherViews == TRUE we'll also change all the
                other views sharing the same selection as this one to use the
                new one.  We can do this faster (O(n) instead of O(n^2)) than
                you can calling us in a loop.
                It's possible we could make it work to set this to NULL if someone
                comes up with a good reason, but we'd have to put a bunch of checks
                for NULL in the DGV code. */
        MSOMETHOD_(void, SetSelection)
                (THIS_ interface IMsoDrawingSelection *pidgsl, BOOL fOtherViews) PURE;

        /* Call PidgvgGetViewGroup to get the Drawing View Group used by
                this View. */
        MSOMETHOD_(interface IMsoDrawingViewGroup *, PidgvgGetViewGroup) (THIS) PURE;

        /* Call FSelectionVisible to determine whether or not the Drawing Selection
                in this DGV is visible or not. */
        MSOMETHOD_(BOOL, FSelectionVisible) (THIS) PURE;

        /* Call SetSelectionVisible to change whether or not the Drawing Selection
                in this DGV is visible. */
        MSOMETHOD_(void, SetSelectionVisible) (THIS_ BOOL fSelectionVisible) PURE;

        /* Call GrfspvmSetShapeModes to change the per-view shape options
                on a particuar Shape.  Pass in the msohsp for the Shape you want
                to be weird and in grfspvmOn the flags you want turned on
                and in grfspvmOff the flags you want turned off.  You can
                pass in msohspNil, 0 for grfspvmOn, and anything you want for
                grfspvmOff to turn some flags off for all the Shapes.  This
                method returns the flags the shape in question ended up with. */
        MSOMETHOD_(ULONG, GrfspvmSetShapeModes) (THIS_ MSOHSP hsp,
                        ULONG grfspvmOn, ULONG grfspvmOff) PURE;

        /* Call GrfspvmGetShapeModes to retrieve the modes of a particular
                Shape in this view. */
        MSOMETHOD_(ULONG, GrfspvmGetShapeModes) (THIS_ MSOHSP) PURE;

        /* FTrackingComponent returns TRUE if this DGV has set itself up
                as the Tracking Component with the Component Manager. */
        MSOMETHOD_(BOOL, FTrackingComponent) (THIS) PURE;

        /* TODO peteren: Real comment.  For Word. */
        MSOMETHOD_(BOOL, FCopyCompatibleView) (THIS_ IMsoDrawingView *pidgvToCopy)
                PURE;

        /* Invalidate the DGV's display elements based on the given flags. */
        MSOMETHOD_(void, Invalidate) (THIS_ ULONG grfdgvi, MSOHSPV hspv) PURE;

        /* Have the DGV validate itself (rebuild SPVs) */
        MSOMETHOD_(void, Validate) (THIS) PURE;

        // ----- IMsoDrawingView methods (mouse events, editing, etc.)

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FOwnsHdes) (THIS_ MSOHDES hdes) PURE;

        /* TODO peteren: Comment. */
        MSOMETHOD_(BOOL, FHitTest) (THIS_ MSODMHD *pdmhd, POINT *pptvTest,
                BOOL fPtwTest, MSOMM mmTest) PURE;

        /* TODO peteren: Comment. */
        MSOMETHOD_(BOOL, FHitTestShape) (THIS_ MSODMHD *pdmhd, POINT *pptvTest,
                int grfhts, MSOMM mmTest, MSOHSPV hspvTest) PURE;

        /* TODO peteren: Comment. */
        MSOMETHOD_(BOOL, FHitTestShapeSelection) (THIS_ MSODMHD *pdmhd,
                POINT *pptvTest, BOOL fPtwTest, MSOMM mmTest, MSOHSPV hspvTest) PURE;

        /* Call FDoMouseDown to have the DGV handle a mouse-down.  This will
                return TRUE if it processes the event (even if it does something
                like start a drag which the user then cancels).  If we return FALSE
                it means we did _nothing_; the host may want to handle the click
                in this case. */
        MSOMETHOD_(BOOL, FDoMouseDown) (THIS_ MSODMHD *pdmhd) PURE;

        /* Call FDoMouseDownOnShape if you already know what kind of hit you want
                to happen on a Shape.  Basically this fakes up a DMHD based on the
                arguments you supply and calls FDoMouseDown on it. */
        MSOMETHOD_(BOOL, FDoMouseDownOnShape) (THIS_ MSOHSPV hspv,
                POINT *pptv, BOOL fPtw, MSOMM mm, MSODRGH drgh,
                MSODGCID dgcidTool) PURE;

        /* DoMouseOnShape has been replaced by FDoMouseDownOnShape. */
        MSOMETHOD_(void, DoMouseOnShape) (THIS_ MSOHSP hsp, POINT *pptv, MSOMM mm,
                MSODRGH drgh) PURE;

        /* Call FDoMouseFeedback to have the DGV handle the feedback for a
           mouse move while not dragging.

                Note that this method has been entirely decoupled from hit-testing.
                This is because there are now several different valid ways to
                get an MSODMHD. FHitTest is the most probable way to get this method.

                Also, this method should be called at least once when fHitTest is false
                to remove any feed back from previous hit. You still need to call
                GetHitDescription       to get the cursor information. */
        MSOMETHOD_(BOOL, FDoMouseFeedback) (THIS_ MSODMHD *pdmhd)       PURE;

        /* GetHitDescription looks at the "hit" in a DMHD and fills out an
                MSODGVHD with a more detailed description of the hit.  The host can use
                this data to choose an appopriate cursor, or in other weird cases
                (like XL's built-in controls) actually process the hit. */
        MSOMETHOD_(void, GetHitDescription) (THIS_ MSODGVHD *pdgvhd,
                MSODMHD *pdmhd) PURE;


        // ----- IMsoDrawingView methods (drag and drop)

        MSOMETHOD (DragEnter) (THIS_ IDataObject *pDataObject, DWORD grfKeyState,
                POINTL *ppt, DWORD *pdwEffect) PURE;
        MSOMETHOD (DragOver) (THIS_ DWORD grfKeyState, POINTL *ppt,
                DWORD *pdwEffect) PURE;
        MSOMETHOD (DragLeave) (THIS) PURE;
        MSOMETHOD (Drop) (THIS_ IDataObject *pDataObject, DWORD grfKeyState,
                POINTL *ppt, DWORD *pdwEffect) PURE;


        // ----- IMsoDrawingView methods (enumeration)

        /* BeginEnumShapes is used in conjunction with FEnumShapes and
                EndEnumShapes to enumerate the shapes in a DGV, as follows:

                MSODGVEN dgven;
                pidgv->BeginEnumShapes(&dgven, fGroups, fChildren);
                while (pidgv->FEnumShapes(&dgven))
                        {
                        ... do stuff with this shape ...
                        pidgv->SomeMethodOrOther(dgven.hspv);
                        }

                Pass FALSE for fGroups and we'll skip groups.  Pass FALSE for
                fChildren and we'll not dive into groups. */
        /* FUTURE peteren: Rename BeginEnumerateSpv */
        MSOMETHOD_(void, BeginEnumShapes) (THIS_ MSODGVEN *pdgven, BOOL fGroups,
                BOOL fChildren) PURE;

        /* FEnumShapes is used with FBeginEnumShapes to enumerate the Shapes
                in a DGV. */
        /* FUTURE peteren: Rename FEnumerateSpv */
        MSOMETHOD_(BOOL, FEnumShapes) (THIS_ MSODGVEN *pdgven) PURE;


        // ----- IMsoDrawingView methods (data access)

        /* CSpv returns the number of shapes in this DGV. Ordinarily we
                return the "deep" number of Shapes, including both groups AND the
                children of groups.  If you pass in FALSE for fGroups we don't count
                groups.  If you pass in FALSE for fChildren we omit the children of
                groups.  Passing in FALSE for both is very weird, but it works. */
        MSOMETHOD_(int, CSpv) (THIS_ BOOL fGroups, BOOL fChildren) PURE;

        /* FFindSpv finds the SPV in a DGV that's viewing a particular shape.
                This is not terribly fast; at the worst it does a linear scan of
                all the SPVs in the DGV.  Returns TRUE if it actually finds a SPV
                showing the shape.  This replaces HspvFromHsp. */
        MSOMETHOD_(BOOL, FFindSpv) (THIS_ MSOHSPV *phspv, MSOHSP hsp) PURE;

        /* HspOfSpv pulls the MSOHSP out of an SPV.  This is fast and
                cannot fail.  This replaces (renames) HspFromHspv. */
        MSOMETHOD_(MSOHSP, HspOfSpv) (THIS_ MSOHSPV hspv) PURE;

        /* GetSviOfSpv fills out an MSOSVI with the position of the
                shape specified by hspv (including rotation, etc.). */
        MSOMETHOD_(void, GetSviOfSpv) (THIS_ MSOSVI *psvi, MSOHSPV hspv) PURE;

        /* GetSvitOfSpv fills out an MSOSVIT with the position of the text
                on the shape specified by hspv. I believe this works even for Shapes
                that don't have text attached to them. */
        MSOMETHOD_(void, GetSvitOfSpv) (THIS_ MSOSVIT *psvit, MSOHSPV hspv) PURE;

        /* GetRcvBoundsOfSpv returns the bounding rectangle (in pixels,
                relative to the DGV) of a particular Shape, identified with
                an HSPV.  This is the absolute bounding rectangle (outside
                arrowheads, thick lines, etc.) and my in fact be a little bigger
                than the object. */
        MSOMETHOD_(void, GetRcvBoundsOfSpv) (THIS_ RECT *prcv, MSOHSPV hspv) PURE;

        /* In a few terribly weird cases, the host needs a way to tell a
                particular SPV that it's visible on screen even though it doesn't
                necessarily think it is.  Chart objects do this; there's a seperate
                window that appears when the chart is active that then invisibly
                vanishes leaving bits on the screen that the chart object is supposed
                to take over responsibility for. */
        MSOMETHOD_(void, SetSpvVisible) (THIS_ MSOHSPV hspv, BOOL fVisible) PURE;

        /* TODO peteren: Comment? */
        MSOMETHOD_(void, SetRciAnchorOfSpv) (THIS_ MSOHSPV hspv, RECT *prciAnchor) PURE;

        /* HspvRelated takes an hspv identifying a Shape in a View and a group
                of flags describing a "relationship" (first child, last child, parent,
                root-level-parent, etc." and returns the hspv of that Shape.
                At the moment the only valid flag is msofdgvrTopLevel, which returns
                the top-level Shape of which hspv is a child, or just hspv if
                hspv is already top-level. */
        MSOMETHOD_(MSOHSPV, HspvRelated) (THIS_ MSOHSPV hspv, ULONG grfdgvr) PURE;

        /* This finds the MSODGVFDSC for a continued textbox from the group's
                MSOHSPV and the MSOHSP of the textbox we would continue on. */
        MSOMETHOD_(BOOL, FGetPositionFromGroupTextbox) (THIS_ MSODGVFDSC *pdgvfdsc,
                MSOHSPV hspvGroup, MSOHSP hspTextbox) PURE;

        /* FGetSviOfSp fills out an MSOSVI with the position of the shape
                specified by hsp.  If you know the shape is actually in the view you
                should use GetSviOfSpv instead.  This function works even if the SP
                was way outside the bounds of the view. */
        MSOMETHOD_(BOOL, FGetSviOfSp) (THIS_ MSOSVI *psvi, MSOHSP hsp) PURE;

        /* FGetSvitOfSp fills out an MSOSVIT with the position of the text
                on the shape specified by hsp.  If you know the shape is actually
                in the view you should use GetSvitOfSpv instead.  This function
                works even if the SP was way outside the bounds of the view. */
        MSOMETHOD_(BOOL, FGetSvitOfSp) (THIS_ MSOSVIT *psvit, MSOHSP hsp) PURE;

        // ----- IMsoDrawingView methods (display)

        /* PigeFromHspv returns the single top level IMsoGE for the given MSOHSPV.
                The IMsoGE is created (and cached) if required. */
        MSOMETHOD_(IMsoGE*, PigeFromHspv) (THIS_ MSODC *pdc, MSOHSPV hspv) PURE;

        /* Clear out the pinkSet plex then fill it in with the ink description
                for every component of the rendering of the give hspv.  The input
                plex is initialized by the caller (though it may be set to be zero
                length).  Each element <iMax must have a correctly set up pinks
                pointer - either it is NULL or it points to an array of BYTEs of
                length cplates.  The output plex iMac is equal to the number of
                separate components rendered by the HSPV.  The components are
                rendered in the order given (i.e. pinkSet[0] is lowest).  The
                ucomponent values will be unique:

                        pinkSet[n].ucomponent != pinkSet[m].ucomponent
                                {for all values of} n != m

                The caller is responsible for the memory in the output.  The
                MSOINKSET::pinks array will be allocated from the same data group
                as the plex.  The MSODC *must* have correct rendering information
                (resolution, device, etc) for the actual print job, but it can be
                set up for any plate (even for a plate which does not actually exist,
                such as the RGB "plate" msoplateRGB).

                The result includes components for all fragments of a shape with
                fragments.  The ucomponent values will be distinguished by some
                arbitrary fragment index.

                NOTE: InkSets are not the same as the Ink from tablet PC's
                */
        MSOMETHOD_(BOOL, FInkSetFromHspv)(THIS_
                MSOPX *pinkSet, //USE: MSOTPX<MSOINKSET> *pinkSet /*in/out*/
                MSODC *pdcBase, int cplates, MSOPLATEDESCRIPTION *pplateDescription,
                MSOHSPV hspv) PURE;

        /* FDrawShapes draws either all of the Shapes in this Drawing View
                into a specified DC or else one of them (if hspvJustMe is not
                msohspvNil).  In pdc pass a pointer to the DC you want us to
                draw into.  In pptwOrigin you can pass the origin of the V coordiate
                system we should use for this Draw or else pass NULL and we'll
                use the ptwOrigin in the DGV.  pextraparam allows just rendering,
                just caching to speed later operations(no render), or both.  If no
                optional parameters are required, set pextraparam = NULL.

                We'll return FALSE if for any reason we didn't finish, whether that be
                that an error occurred or we were interrupted.  You have to check the
                last error to know which. */
        MSOMETHOD_(BOOL, FDrawShapes) (THIS_ MSODC *pdc, POINT *pptwOrigin,
                RECT *prcvDraw, BOOL fRcwDraw, BOOL fClip, MSOHSPV hspvJustMe,
                MSODGVFDSC *pdgvfdsc, CONST MSODGVFDSPARAM *pextraparam) PURE;

        /* FDrawShapeSelection draws the selection dots of a specified Shape
                into a specified DC.  It shares a bunch of arguments with
                FDrawShapes. */
        MSOMETHOD_(BOOL, FDrawShapeSelection) (THIS_ MSODC *pdc,
                POINT *pptwOrigin, RECT *prcvDraw, ULONG grfdgvds, MSOHSPV hspv) PURE;

        /* Sets/Fetchs properties on an HSPV which act as view-specific overrides on
                properties set into the shape.  FSetPropHSPV() will return false if it
                fails to set the property.  FFetchPropHspv() will return false if the property
                does not exist in the override set.  It does NOT search the shape, master,
                defaults like the IMsoDrawingVersion */
        MSOMETHOD_(BOOL, FSetPropHspv)(MSOHSPV hspv, MSOPID opsid, const void* pSrc) PURE;
        MSOMETHOD_(BOOL, FFetchPropHspv)(MSOHSPV hspv, MSOPID opid, void* pDest) PURE;

        /* Union the coverage region of the shape for a particular view
                and scale to the passed in region. 'cvrk' tells which region the caller
                is interested in.

                msocvrkModified -- includes all the pixels which FDraw will modify
                msocvrkOpaque   -- includes only the pixels which FDraw will modify
                msocvrkTransparent -- include only pixels that FDraw will not modify

                Regions don't have pixel perfect coverage. They include more or less
                pixels depending on the kind of calculation done. Returns the same
                values as the WIN32 CombineRegion call.

                The hrgn is an in/out parameter - the region must exist before the
                method is called.  Note that this method is expensive to call and
                wasteful if the shape has not been drawn - typical use is to find out
                whether an existing shape must be redrawn when another object changes.
                */
        MSOMETHOD_(MSORGNK, RgnkUnionCoverage) (THIS_ HRGN hrgn, MSOHSPV hspv,
                MSOCVRK cvrk) PURE;

        /* PALETTE MANAGEMENT API:

                Return the optimal palette to use for this view.  The method returns
                the number of palette entries required and fills these entries in in
                the *ppe array.  In the event of error the routine returns
                msopaletteError, if there are too many and too various colors to make
                palette optimization possible the routine returns msopaletteHalftone
                and fills in no entries - both of these values are negative, use an
                "int" to receive the result of the interface.  In any case the routine
                fills in *pdcPrecision with the precision required to reduce the colors
                in the desired fashion - this MUST be set into the MSODC passed to the
                view drawing procedures to get the correct results.

                The routine assumes that display is for an 8 bit per pixel (256 color)
                device, however the result can be used to determine the best palette
                for storage in bitmaps with more than 256 colors with the caveat that
                colors are not returned in order of importance.

                The caller must specify whether the system colors will be available.
                In all cases the interface assumes that black and white will always
                be available and does not return these colors, if the caller declares
                that the system (VGA) colors will be available they will not be added
                to the result - the caller must do that.  In addition the caller may
                specify additional "fixed" colors which will be added to the palette,
                again these are not returned in the generated palette and must be
                added (somehow) by the caller, these colors are in addition to the
                system colors, if specified.

                The additional flags allow the caller more control over how any palette
                reduction is done.  Note that there is no point calling this interface
                on a gray scale device - the internal color precision is only 8 bits,
                so all gray scales can be represented in a 256 color palette (so long
                as the system colors are not added!), the low order 8 bits of these
                flags are a "tolerance" which may be applied to the "secondary" colors
                in attempt to reduce the number of colors required by bitmaps.  A
                value of 26 or more should typically be supplied - 26 corresponds to
                the "tolerance" implied by using the halftone palette.  The default
                value of 32 is sufficient to limit bitmaps to 125 colors.  Use the
                formula (255/tol+1)^3 to find out how many colors correspond to a
                particular tolerance tol.

                NOTE: ppe[] is used as working store and is changed even if the method
                fails to find a palette.  ppe[] and ppeFixed[] must not overlap. */
#define msopaletteError         (-1)
#define msopaletteHalftone      (-2)
#define msopaletteToleranceMask 0xFFU
#define msopaletteJunkBitmaps   0x100U  // Disregard bitmap colors if required
#define msopaletteReducePrimary 0x200U  // Reduce shade colors if necessary
#define msopaletteFlagsDefault  0x220U
        MSOMETHOD_(int, NPalette) (THIS_ PALETTEENTRY *ppeOut, int cpeMax,
                              MSOGELDCINFO *pdcPrecision,
                              const PALETTEENTRY *ppeFixed, int cpeFixed,
                              BOOL fAssumeSystemColors, ULONG grf,
                              interface IMsoDrawingView *pidgvMaster,
                              interface IMsoBlip **ppib,
                              int cppib) PURE;

        // ----- IMsoDrawingView methods TO GET RID OF!!!!!

        MSOMETHOD_(MSOHSP, HspFromHspv) (THIS_ MSOHSPV hspv) PURE;
                // Call HspOfSpv instead
        MSOMETHOD_(MSOHSPV, HspvFromHsp) (THIS_ MSOHSP hsp) PURE;
                // Call FFindSpv instead

        // ----- IMsoDrawingView methods (stuck at end to avoid interface change)

        MSOMETHOD_(BOOL, FGetTextFrameRcvi) (THIS_ MSOHSP hsp, MSORCVI *prcviText) PURE;
                // I believe GetSvitOfSpv is better than this.
                // I believe FGetSvitOfSp is also better than this.

        MSOMETHOD_(BOOL, FSetTextFrameRci)(THIS_ MSOHSP hsp, RECT *prciText) PURE;
                // Set the text frame in I unit's. Use this function instead of FSetTextFrameRch
                // if the text is measured in I unit's. Note: it is preferable for clients
                // to measure in H units. I units are not view independent, therefore, problems
                // can arrise in multiple view scenerios.

        /* FGetShapePolygon converts the shape to the single polygon that is most representive
           of the shape. All shapes have a polygon representation. At best it is exactly the
           shape, at worst it could be the bounding box. For complex shape, a wrap polygon is
           returned The current adjustment, rotation, position of the shape is taken into account.

           hspv           -- Shape [in]
           dg             -- data group to allocate the array [in]
           prgpt          -- return the Pointer to Array of points in V units [out]
           prgply         -- return the Pointer to Array of points in a polygon [out]
           pcpt           -- number of points [out]
           pcply          -- number of polygons [out]
                                msospIsRect if the the polygon is a rect
                                msospIsCircle if the the polygon is a circle

           A polygon is allocated in 'dg'. See msospqCanConvertToPolygon to check
           if shape can be converted well to a polygon.
        */

#define msospIsRect        (-1)
#define msospIsCircle      (-2)
        MSOMETHOD_(BOOL, FGetShapePolygon) (THIS_ MSOHSPV hspv, int dg,
                _Out_ POINT **prgpt, _Out_ int *pcpt, _Out_ LONG **prgply, _Out_ int *pcply) PURE;

        /* Like SetRciAnchorOfSpv except that it also allows the user to set the angle
                and frees any cached GEL effects. */
        MSOMETHOD_(void, SetRciAnchorAndAngleOfSpv) (THIS_ MSOHSPV hspv, RECT *prciAnchor,
                MSOANGLE langle) PURE;

        MSOMETHOD_(BOOL, FRequestDoDgvInkCommand) (THIS) PURE;
        MSOMETHOD_(BOOL, FInInkMode) (THIS_ int) PURE;
        MSOMETHOD_(BOOL, FInkAnnotationMode) (THIS) PURE;
        MSOMETHOD_(BOOL, FMadeInkOrEraseEdits) (THIS_ BOOL fCheckSurface) PURE;
        MSOMETHOD_(void, UpdateInputRc) (THIS) PURE;
        // Update the input rect for this view.
        // Sucks that we need this, but due to the way the wisp api's work, it seems we
        // need this so we can update the input areas atomically for inking.
};


MSOAPIX_(int) MsoNPalette(PALETTEENTRY *ppeOut, int cpeMax,
                              MSOGELDCINFO *pdcPrecision,
                              const PALETTEENTRY *ppeFixed, int cpeFixed,
                              BOOL fAssumeSystemColors, ULONG grf,
                              IMsoDrawingView **ppdgv,
                              int cpdgv,
                              IMsoBlip **ppib,
                              int cppib);
/* Call NPalette on each of the cpdgv views ppdgv[cpdgv] and combine the
        results into a single ppeOut.  This should be used if the caller needs to
        generate a single palette for more than just the view and the master. */

#if 0 //$[VSMSO]
MSOAPI_(BOOL) MsoFCanHandleCagDrop(IDataObject *pido);
/* Call MsoFCanHandleCagDrop to determine if Office can handle the "MSClipGallery" format
   (msoszCagClipboardFormat) data objects. Clients typically will iterate through
   the formats available from a data object and delegate the handling of that object
   to Office if data object has an Office format available. For the "MSClipGallery"
   format, a further test is needed to see if Office can handle that format.
   Hence, the need for this API. */

typedef enum {msocftPicture = 0, msocftMovie, msocftSound} MSOCFT;

MSOAPI_(BOOL) MsoFIsCagFile(const WCHAR* wzFileName, MSOCFT *pcftReturn);
/* MsoFIsCagFile looks at the extension of the passed in filename and
   determine if it is a filename that ClipArt Gallery supports. Returns
   the type of file extension found. */

MSOAPI_(void) MsoCloseCagInstances();
/* CloseCagInstances closes all open instances of ClipArt Gallery.
   Instances are created using the msodgcidRunCag command */
#endif //$[VSMSO]

/****************************************************************************
        Drawing View Group (DGVG)

        A Drawing View (DGV) corresponds to a single rectangular area in which
        a portion of a Drawing can be viewed.  Within a DGV each Shape can
        appear at most once.  Both Word and Excel have a higher-level notion of
        a "group" of such Views, which we need to support.

        In Excel you can split a View window into two panes.  These two panes
        are always viewing the same worksheet, and therefore always the same
        DG.  They also share a single selection, which is visible in both panes
        simulateously.

        In Word when you split a window into panes it's really just like
        opening a new window; each pane has a different selection, only one of
        which is visible at a time.  However, you can in multi-page page view
        look at several pages at once.  These, very similarly, are a number of
        different DGVs all viewing the same DG and sharing a single selection
        which is visible in all of them simultaneously.

        A Drawing View Group (DGVG), then, is a set of DGVs all attached to the
        same DG and sharing a single DGSL.

        The interface to DGVG isn't very interesting, all the host can do is
        create DGVs in it and look at the DGSL.  A DGVG requires no site
        interface.

        DGVGs inherit from ISimpleUnknown instead of IUnknown; they keep
        no ref count.

        We should think carefully about whether word really needs this or if
        there's some other way to share a selection across a set of DGVs.

        To create a DGVG you call IMsoDrawing::FCreateViewGroup.  To free it
        you call IMsoDrawingViewGroup::Free().  You used to be able to create
        DGVGs implicitly by calling IMsoDrawing::FCreateView, but that method
        didn't really add much and made the life-span of DGVGs more complicated,
        so I took it out.
****************************************************************************/

/* IMsoDrawingViewGroup (idgvg) */
#undef  INTERFACE
#define INTERFACE IMsoDrawingViewGroup
DECLARE_INTERFACE_(IMsoDrawingViewGroup, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingViewGroup methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Free() tells a DrawingViewGroup to go away. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        // ----- IMsoDrawingViewGroup methods (other)

        /* FCreateView() makes a new drawing view in this drawing view group.
                It copies the MSODGVSI you pass in in the new view, thereby
                connecting it to a Drawing View Site and a Display Manager. */
        MSOMETHOD_(BOOL, FCreateView) (THIS_ interface IMsoDrawingView **ppidgv,
                MSODGVSI *pdgvsi, BOOL fValidate) PURE;

        /* PidgGetDrawing() returns an IMsoDrawing pointer.  This cannot fail.
                Since IMsoDrawing does not inherit from IUnknown there aren't any
                ref-counts involved and you don't have to Release() the returned
                pointer when you're finished with it. */
        MSOMETHOD_(interface IMsoDrawing *, PidgGetDrawing) (THIS) PURE;

        /* PidgslGetSelection() returns IMsoDrawingSelection pointer.
                DGVGs always have selections, so this cannot fail.  The DGSL
                will NOT have been AddRef()'d. */
        MSOMETHOD_(interface IMsoDrawingSelection *, PidgslGetSelection)
                (THIS) PURE;

        /* Update tells the DGVG to update the Display Managers it shows in.
                This is a pretty unsubstantial method, but it's nicely
                orthogonal with DG::Update, DGV::Update, and DGSL::Update.
                The arguments to this should be identical to DM::Update. */
        MSOMETHOD_(void, Update) (THIS) PURE;

        /* Drag-drop methods */
        MSOMETHOD (DragEnter) (THIS_ IMsoDrawingView *pidgv, IDataObject *pDataObject,
                DWORD grfKeyState, POINTL *ppt, DWORD *pdwEffect) PURE;
        MSOMETHOD (DragOver) (THIS_ DWORD grfKeyState, POINTL *ppt,
                DWORD *pdwEffect) PURE;
        MSOMETHOD (DragLeave) (THIS) PURE;
        MSOMETHOD (Drop) (THIS_ IDataObject *pDataObject, DWORD grfKeyState,
                POINTL *ppt, DWORD *pdwEffect) PURE;

        /* SetSelection() tells all the DGVs in this DGVG (and the DGVG itself,
                which means DGVs later created in this DGVG) to abandon their
                current DGSLs and use the one the caller specify. */
        MSOMETHOD_(void, SetSelection) (THIS_
                interface IMsoDrawingSelection *pidgsl) PURE;
};


/****************************************************************************
        Graphics XML Import/Site interfaces

        IMPORTANT NOTE: By design these interfaces are not aware of
        IMsoDrawingGroup, IMsoDrawing, or MSOHSP's.

        These interfaces are currently only used by Chart for filling the
        interior of Series data points and legends.
****************************************************************** NeerajM */

/* MSOGRXMLISI = GRaphics XML Import Site Info.  Every GRXMLI is connected
        to a GRXMLIS object, which provides necessary call-back functions to
        the host for it.  A GRXMLIS is also defined by some relatively
        constant pieces of data.  These are defined in a single structure
        to contain all of this data, the GRXMLISI.  You have to pass a
        pointer to one of these to any methods that create a GRXMLI.
        Probably the most important thing it contains is the pointer to
        the site object.*/
typedef struct _MSOGRXMLISI
        {
        interface IMsoGraphicsXMLImportSite *pigrxmlis;
                /* GRXMLIS interface pointer. Created by host. */
        void *pvgrxmlis;
                /* client data pased to all GRXMLIS methods. Can be NULL */
        interface IMsoReducedHTMLImport *pirhi;
        } MSOGRXMLISI;
#define cbGRXMLISI (sizeof(MSOGRXMLISI))


/**************************************************************************
        Graphics XML Import Interface

        To create a GRXMLI you call MsoFCreateGraphicsXMLImport.
        To free it you call IMsoGraphicsHTMLImport::Free().
**************************************************************** NeerajM */

/* IMsoGraphicsXMLImport (igrxmli) */
#undef  INTERFACE
#define INTERFACE IMsoGraphicsXMLImport
DECLARE_INTERFACE_(IMsoGraphicsXMLImport, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoGraphicsXMLImport methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Free() tells a GraphicsXMLImport to go away. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* Process and parse the XML data encoded in 'phisd'. We are either in the
           the VG (Vector Graphics) / OA (OfficeArt)  namespace or are switching to it.
           On entry, the XML data are encoded as 1 of the following cases:
              1) Start-Tag
              2) End-Tag
              3) Name/Value pair
              4) Text

           It will callback to host site, using IMsoGraphicsXMLImportSite interface,
           to set the appropriate properties.

           Return 'fEnd' TRUE if we're done with parsing. */
        MSOMETHOD_(BOOL, FProcessData) (THIS_ const MSORHISD *prhisd, BOOL *fEnd, interface IMsoReducedHTMLImport *pirhi) PURE;
};


/* You should call MsoInitGrxmlisi on any new MSOGRXMLISI you're filling out
        before passing it into a method that creates a GRXMLI object. */
MSOAPI_(void) MsoInitGrxmlisi(MSOGRXMLISI *pgrxmlisi);

/* MsoFCreateGraphicsXMLImport makes a new Graphics XML Import. */
MSOAPI_(BOOL) MsoFCreateGraphicsXMLImport(interface IMsoGraphicsXMLImport
        **ppigrxmli, MSOGRXMLISI *pgrxmlisi);


/****************************************************************************
        Graphics XML Import Site (GRXMLIS)

        GRXMLIS inherits from ISimpleUnknown instead of IUnknown; it keeps
        no ref counts.

        IMsoGraphicsXMLImportSite must be implemented by anyone who wants to
        make and use IMsoGraphicsXMLImport.  Fill out a MSOGRXMLISI (use
        MsoInitGrxmlisi) and add a pointer to your IMsoGraphicsXMLImportSite,
        and then pass that off to the MsoFCreateGraphicsXMLImport function
        that creates a IMsoGraphicsXMLImport object.

        All the methods in IMsoGraphicsXMLImportSite take a void *
        named pvGrxmlivs as their second argument after the "this"
        pointer.  This opaque client data is defined when the host creates a
        MSOGRXMLISI.
***************************************************************** NeerajM */

#undef  INTERFACE
#define INTERFACE IMsoGraphicsXMLImportSite
DECLARE_INTERFACE_(IMsoGraphicsXMLImportSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoGraphicsXMLImportSite methods

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* The GRXMLI will call GRXMLIS::FSetFill() to pass the Fill Property
                set to the host app. */
        MSOMETHOD_(BOOL, FSetFill) (THIS_ void *pvDgxmlivs,
                MSOPSFillStyle *popsFill) PURE;

        /* The GRXMLI will call GRXMLIS::FSetLine() to pass the Line Property
                set to the host app. */
        MSOMETHOD_(BOOL, FSetLine) (THIS_ void *pvDgxmlivs,
                MSOPSLineStyle *popsLine) PURE;
};


/* Fixup the fill for Chart.  It is called at the end Web/Chart tag to fix
        out of bound scheme index or gradient fill with focus % not equal to
        0, 50 or 100.  Any other fix in future can be fixed here */
MSOAPI_(BOOL) MsoFFixupFill(MSOPSFillStyle *popsFill, COLORREF *pcrScheme,
        int ccrScheme);

/* Fixup the stroke (line) for PUB's Table Diagonal Border*/
MSOAPI_(BOOL) MsoFFixupLine(MSOPSLineStyle *popsLine, COLORREF *pcrScheme,
        int ccrScheme);


/****************************************************************************
        Drawing XML Import/Site interfaces
******************************************************************** TuanN */

/* MSODGXIE = Drawing Xml Event.  These are passed to DGXMLI/DGXMLIS::OnEvent.*/
typedef enum
        {
        msodgxiNil = 0,

        /* ----- Events passed to the DGXMLIS, the Drawing Xml Import Site. */
        msodgxiDgxisFirst = 100,
        msodgxiDoReanchor,
        msodgxiDoInline,
        msodgxiBeforeCreateHsp,
        msodgxiNotifyCreateHsp,
        msodgxiNotifyCreateBackground,
        msodgxiDoCreateAnchor,
        msodgxiNotifyTextLink,
        msodgxiNotifyDefaultProp,
        msodgxiApplyCSS,
        msodgxiDgxisMax,
        } MSODGXIE;

/* MSODGXMLISI = Drawing XML Import Site Info.  Every DGXMLI is connected to
        a DGXMLISI object, which provides necessary call-back functions to the
        host for it.  A DGXMLISI is also defined by some relatively
        constant pieces of data.  These are defined in a single structure
        to contain all of this data, the DGXMLISI.  You have to pass a
        pointer to one of these to any methods that create a DGXMLI.
        Probably the most important thing it contains is the pointer to
        the site object.*/
MsoDEBDeclare(dgxi, msodgxiDgxisFirst, msodgxiDgxisMax);
typedef struct _MSODGXMLISI
        {
        interface IMsoDrawingXMLImportSite *pidgxmlis;
                /* DGXMLISI interface pointer. Created by host*/
        void *pvDgxmlivs;
                /* client data pased to all DGXMLISI methods. Can be NULL */
        interface   IMsoDrawingGroup *pidgg;
        interface   IMsoDrawing *pidg;
        MSOHSP      hsp;
        interface   IHlinkSite *pihlSite;
        union
                {
                struct
                        {
                        ULONG fSaveAnchorInTempSpace : 1; // store anchor in shape tempspace
                        ULONG fWantHlinkProp : 1; // parse the shape hyperlink property?
                        ULONG fCfHtml : 1; // are we doing cfhtml?  (we need to keep undo)
                        ULONG fFireDoCreateAnchorEvent : 1;
                        ULONG fNoButtonHlinks : 1; // don't set msopidFIsButton property for href shape attr.
                        ULONG fIgnoreFLockText : 1;
                                /* True if the app prefers to ignore the msopidFLockText. */
                        ULONG fIgnoreHR : 1;
                                /* True if the app prefers to ignore Horizontal Rule related properties. */
                        ULONG fWantsCSSBoxEvents : 1; // Most apps don't want full box model info.  FP does.
                        ULONG fNoHIPathResolution : 1; // Don't use HI to resolve paths, hosts (like FP) will do their own resolution.
                        ULONG grfUnused : 24;
                        };
                ULONG grf;
                };

                interface IMsoReducedHTMLImport *pirhi;
                MSOHISD *phisdCache; // MAY BE NULL: HISD on which to cache self

        MsoDEBDefine(dgxi);
        } MSODGXMLISI;
#define cbDGXMLISI (sizeof(MSODGXMLISI))


/* MSODGXIEB -- Drawing Xml Import Event Block */
typedef struct MSODGXIEB
        {
        /* These first fields are filled out for EVERY
                Drawing Xml Import Event. */
        MSODGXIE dgxie;
        BOOL fResult;
        interface IMsoDrawingXMLImport *pidgxmli;
        MsoDEBPDefine(dgxi);
        union
                {
                struct // for msodgxiDoReanchor
                        {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        RECT rch;
                        BOOL fRelativeLeft;
                        BOOL fRelativeTop;
                        BOOL fSeenLeft;
                        BOOL fSeenTop;
                        BOOL fPositionAbsolute;
                        BOOL fRelHPosSet;  //used by Word
                        BOOL fRelVPosSet;  //used by Word
                        } dgxixDoReanchor;
                struct // for msodgxiDoInline
                        {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        } dgxixDoInline;
                struct // for msodgxiBeforeCreateHsp
                        {
                        BOOL fGroup;
                        MSOSPT spt;
                        } dgxixBeforeCreateHsp;
                struct // for msodgxiNotifyCreateHsp
                        {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        } dgxixNotifyCreateHsp;
                struct // for dgxixDoCreateAnchor
                        {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        RECT rch;
                        void *pvAnchor;
                        } dgxixDoCreateAnchor;
                struct // for msodgxiNotifyDefaultProp
                        {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        MSOPID opid;
                        } dgxiNotifyDefaultProp;
                struct // msodgxiApplyCSS
                        {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        MSOSPCSSANCHOR anchor;
                        } dgxiApplyCSS;
                };
        } MSODGXIEB;
#define cbMSODGXIEB (sizeof(MSODGXIEB))


/**************************************************************************
   Drawing XML Import Interface

        To create a DGXMLI you call MsoFCreateDrawingXMLImport.
        To free it you call IMsoDrawingHTMLImport::Free().
****************************************************************** TuanN */

/* IMsoDrawingXMLImport (idgxlmi) */
#undef  INTERFACE
#define INTERFACE IMsoDrawingXMLImport
DECLARE_INTERFACE_(IMsoDrawingXMLImport, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingXMLImport methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Free() tells a DrawingXMLImport to go away. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        /* Process and parse the XML data encoded in 'phisd'. We are either in the
           the VG (Vector Graphics) / OA (OfficeArt)  namespace or are switching to it.
           On entry, the XML data are encoded as 1 of the following cases:
              1) Start-Tag
              2) End-Tag
              3) Name/Value pair
              4) Text

           It will callback to host site to get info on DGG, DG, SP and set the appropriate
           Escher properties.

           Return 'fEnd' TRUE if we're done with parsing. */
        MSOMETHOD_(BOOL, FProcessData) (THIS_ const MSORHISD *prhisd, BOOL *fEnd, interface IMsoReducedHTMLImport *pirhi) PURE;
};


/* You should call MsoInitDgxmlisi on any new MSODGXMLISI you're filling out
        before passing it into a method that creates a DGXMLI object. */
MSOAPI_(void) MsoInitDgxmlisi(MSODGXMLISI *pdgxlmisi);

/* MsoFCreateDrawingXMLImport makes a new Drawing XML Import. */
MSOAPI_(BOOL) MsoFCreateDrawingXMLImport(interface IMsoDrawingXMLImport **ppidgxmli,
        MSODGXMLISI *pdgxmlisi);


/****************************************************************************
        Drawing XML Import Site (DGXMLIS)

        DGXMLIS inherits from ISimpleUnknown instead of IUnknown; it keep
        no ref count.

        IMsoDrawingXMLImportSite must be implemented by anyone who wants to make
        and use IMsoDrawingXMLImport.  Fill out a MSODGXMLISI (use MsoInitDgxmlisi)
        and add a pointer to your IMsoDrawingXMLImportSite, and then pass that
        off to the MSO function that creates a IMsoDrawingXMLImport.

        All the methods in IMsoDrawingXMLImportSite take a void *
        named pvDgxmlivs as their second argument after the "this"
        pointer.  This opaque client data is defined when the host creates a
        MSODGXMLISI.
****************************************************************** TuanN */

#undef  INTERFACE
#define INTERFACE IMsoDrawingXMLImportSite
DECLARE_INTERFACE_(IMsoDrawingXMLImportSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingXMLImportSite methods

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* The DGXMLI will call DGXMLIS::Pidgg() iff a DGG is not passed in MSODGXMLISI
                upon this interface creation. Returning a NULL DGG will force failure in
                the caller, DGXMLI::FProcessData. */
        MSOMETHOD_(IMsoDrawingGroup *, Pidgg) (THIS_ void *pvDgxmlivs) PURE;

        /* The DGXMLI will call DGXMLIS::Pidg() iff a DG is not passed in MSODGXMLISI
                upon this interface creation. Returning a NULL DG will force failure in
                the caller, DGXMLI::FProcessData. */
        MSOMETHOD_(IMsoDrawing *, Pidg) (THIS_ void *pvDgxmlivs) PURE;

        /* The DGXMLI will call DGXMLIS::Hsp() iff a MSOHSP is not passed in MSODGXMLISI
                upon this interface creation. Returning a NULL MSOHSP will force failure in
                the caller, DGXMLI::FProcessData. */
        MSOMETHOD_(MSOHSP, Hsp) (THIS_ void *pvDgxmlivs) PURE;

        /* The DGXMLI will call DGXMLIS::OnEvent whenever a drawing xml import
                event (MSODGXIE) occurs.  See the definition of MSODGXIE for a more
                complete explanation of events.*/
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDgxmlivs, MSODGXIEB *pdgxieb) PURE;

        /* The DGXMLI will call DGXMLIS::SetCurHsp() when VGPSITE creates a new
                shape during EndSp. This is a way of letting the client know which
                shape is the one most recently (persumably, most relevent) created */
        MSOMETHOD_(VOID, SetCurHsp) (THIS_ void *pvDgxmlivs, MSOHSP *phsp,
                BOOL *pfMoved) PURE;
};


/****************************************************************************
        Drawing Export interfaces
****************************************************************** NeerajM */

/* MSODGHE = Drawing Html Event.  These are passed to DGHIS/DGHES::OnEvent.*/
typedef enum
        {
        msodgheNil = 0,

        /* ----- Events passed to the DGHIS, the Drawing Html Import Site. */
        msodgheDghisFirst = 100,
        msodgheDghisMax,

        /* ----- Events passed to the DGHES, the Drawing Html Export Site. */
        msodgheDghesFirst = 200,
        msodgheDghesRequestFilter,
        msodgheDghesRequestAddTD,
        msodgheDghesRequestEndTD,
        msodgheDghesRequestCSS, // this name should be NotCSS!!
        msodgheDghesRequestNoMerge,
        msodgheDghesRequestNoRender,
        msodgheDghesRequestImgAttributes,
        msodgheDghesRequestBuffer,
        msodgheDghesRequestHasHlink,
        msodgheDghesRequestHlinkImageMap,
        msodgheDghesDoRelativeOffset,
        msodgheDghesRequestHVSpace,
        msodgheDghesRequestMailHackImg,
        msodgheDghesRequestParaAlignNotLeft,
        msodgheDghesRequestParaAlignNotTop,
        msodgheDghesRequestCellBounds,
        msodgheDghesRequestHostRender,
        msodgheDghesBeforeWriteVML,
        msodgheDghesAfterWriteVML,
        msodgheDghesRequestSpecialObj,
        msodgheDghesRequestNoOffsetTable,
        msodgheDghesExpandHTMLFragmentToken,
        msodgheDghesRequestHsp,
        msodgheDghesRequestIStream,
        msodgheDghesRequestInline,
        msodgheDghesMax,
        } MSODGHE;

MsoDEBDeclare(dghe, msodgheDghesFirst, msodgheDghesMax);

/* MSODGHEEB -- Drawing Html Export Event Block */
typedef struct MSODGHEEB
        {
        /* These first fields are filled out for EVERY
                Drawing Html Export Event. */
        MSODGHE dghe;
        BOOL fResult;
        interface IMsoDrawingHTMLExport *pidghe;
        MsoDEBPDefine(dghe);

        union
        {
                struct
                {
                /*      msodgheDghesRequestFilter
                        This passes in a RECT in the host's native coordinates and
                        requests the host to check if the position is legal.  Later,
                        we may add in the hsp so the host can filter out that stuff too.
                        The shape is filtered out if fResult is FALSE. */
                /*      msodgheDghesRequestNoRender
                        This is to allow the host to not render a shape after layout, but
                        still requires round-trip info. For example, XL checks for hidden
                        rows and columns.  This is different than filtering the shape, which
                        would not go through layout code and not do round-trip info */
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        RECT rcView;
                        RECT rcvShape;
                        BOOL *pfAttachHLink;
                        RECT *prcvPicture;   //inline Picture bounds in Word
                };
                struct
                {
                /*      msodgheDghesRequestCSS
                        This passes in a RECT in the host's native coordinates and
                        requests the host to check if the position is legal.  Later,
                        we may add in the hsp so the host can filter out that stuff too.
                        The shape is filtered out if fResult is FALSE. */
                        RECT rcCSS;
                        RECT rcvView;
                        RECT rcvCSSShape;
                };
                struct
                {
                /*      msodgheDghesRequestAddTD
                        This requests the host to add the <TD> for the shape.  This is
                        called as an event back to the host, because different hosts may
                        want to do different things, for instance XL needs to specify
                        rowspan and colspan stuff, while Word may do nothing at all.  The
                        rectangle is the host rect. */
                        RECT rcImage;
                        BOOL fCSS;
                        BOOL fNoRender;
                        BOOL fRoundtrip;
                        int cxiUnit; // how many I's are there in an arbitrary unit (X-axis)
                        int cxvUnit; // how many V's are there in the same unit (X-axis)
                        int cyiUnit; // how many I's are there in an arbitrary unit (Y-axis)
                        int cyvUnit; // how many V's are there in the same unit (Y-axis)
                } dghexRequestAddTd;
                struct // for msodgheDghesRequestNoMerge
                {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                } dghexRequestNoMerge;
                struct // for msodgheDghesRequestSpecialObj
                {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        BOOL fSpecialObj; // excel class=shape form controls
                        BOOL fNotTopAlign; // misnomer
                        BOOL fHostRendered; // publisher form controls
                        BOOL fWantsVML;
                } dghexRequestSpecialObj;
                struct // for msodgheDghesRequestImgAttributes
                {
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        MSOPH *pph; //position horizontal
                        MSOPV *ppv; //position vertical
                } dghexRequestImgAttributes;
                struct // for msodgheDghesRequestBuffer
                {
                        RECT rcAnchor;
                        POINT ptBuffer;
                        BOOL fWidth;
                        int cxiUnit; // how many I's are there in an arbitrary unit (X-axis)
                        int cxvUnit; // how many V's are there in the same unit (X-axis)
                        int cyiUnit; // how many I's are there in an arbitrary unit (Y-axis)
                        int cyvUnit; // how many V's are there in the same unit (Y-axis)
                } dghexRequestBuffer;
                struct
                /*      msodgheDghesRequestHasHlink
                        This asks the host if there are any attached text hyperlinks in this
                        shape.  If so, then result should be TRUE, or FALSE otherwise.
                        msodgheDghesRequestHlinkImageMap
                        This asks the host to export all the image maps for this shape with
                        respect to rcvShapeHlink(bounds of view for shape) in rcvImageHlink
                        (bounds of image), using piheHlink. */
                        {
                        MSOHSP hspHlink;
                        MSOTXID txidHlink;
                        interface IMsoHTMLExport *piheHlink;
                        RECT rcvShapeHlink;
                        RECT rcvImageHlink;
                };
                struct // for msodgheDghesRequestMailHackImg
                {
                        MSOHSP hspMailHack;
                } dghexRequestMailHackImg;
                struct
                /* msodgheDghesRequestParaAlignNotLeft
                        This asks the host if we need to write out left:0 for CSS. */
                {
                        MSOHSP hsp;
                } dghexRequesParaAlignNotLeft;
                struct
                /* msodgheDghesRequestParaAlignNotTop
                        This asks the host if we need to write out top:0 for CSS. */
                {
                        MSOHSP hsp;
                } dghexRequesParaAlignNotTop;
                struct
                /* msodgheDghesRequestCellBounds
                        This asks the host if we need to write out left:0 for CSS. */
                {
                        RECT rcAnchor;
                        const MSORCVI *prcvi;
                        ULONG lwidth;
                        ULONG lheight;
                } dghexRequestCellBounds;
                struct
                /* msodgheDghesRequestHostRender
                        This allows the host to write out whatever representation of HTML
                        for a shape, instead of an image.  This is used by Publisher's form
                        controls.  Set fHostRendered in msodgheDghesRequestNoMerge to get
                        this callback for a particular shape. */
                {
                        interface IMsoHTMLExport *pihe;
                        IMsoDrawing *pidg;
                        MSOHSP hsp;
                        const MSORCVI *prcvi;
                        ULONG lwidth;
                        ULONG lheight;
                        BOOL fWantsVML;
                        MSOXMLW *pxmlw;
                } dghexRequestHostRender;
                struct
                /* msodgheDghesBeforeWriteVML
                        This notifies the host before writing XML on a shape.  This event
                        is currently used by Excel for chart objects.  */
                {
                        MSOHSP hsp;
                        MSOXMLW *pxmlw;
                        DWORD hetnUser;
                        BOOL fWantVML;
                } dghexBeforeWriteVML;
                struct
                /* msodgheDghesAfterWriteVML
                        This notifies the host after writing XML on a shape.  This event
                        is currently used by Publisher.  */
                {
                        MSOHSP hsp;
                        MSOXMLW *pxmlw;
                } dghexAfterWriteVML;
                struct
                /* msodgheDghesExpandHTMLFragmentToken
                        Allows the host to expand ${TOKENS} found in an HTML fragment or
                        web component.  This event is currently used by Publisher. */
                {
                        MSOHSP hsp;
                        WCHAR *wzToken;
                } dghexExpandHTMLFragmentToken;
                struct
              /* msodgheDghesRequestHsp
                        This passes the current hsp that's being exported to the host. */
                {
                        MSOHSP hsp;
                } dghexRequestHsp;
                struct
              /* msodgheDghesRequestIStream
                        This passes the just generated IStream to the host. */
                {
                        IStream *pistm;
                        DWORD hetn;
                        DWORD hetnPfx;
                } dghexRequestIStream;

        };
        } MSODGHEEB;
#define cbMSODGHEEB (sizeof(MSODGHEEB))


/* DGHESI = DrawinG Html Export Site Info.  Every DGHE is connected to
        a DGHES object, which provides necessary call-back functions to the
        host for it.  A DGHES is also defined by some relatively
        constant pieces of data.  These are defined in a single structure
        to contain all of this data, the DGHESI.  You have to pass a
        pointer to one of these to any methods that create a DGHE.
        Probably the most important thing it contains is the pointer to
        the site object.*/
typedef struct _MSODGHESI
        {
        interface IMsoDrawingHTMLExportSite *pidghes;
                /* DGHES interface pointer. */
        void *pvDghevs; /*client data pased to all DGHES methods.*/
        union
                {
                struct
                        {
                        /* HTML file format types */
                        ULONG fRelyOnVML : 1;
                                // minimized down-level content
                        ULONG fOnlySave: 1;
                                // only roundtrip (Excel's Hide All Objects)
                        ULONG fRtl:1;
                                // right-to-left layout (Excel's BIDI)

                        /* Export format options */
                        ULONG fWriteXML : 1;
                                // export shape VML along with down-level data
                        ULONG fAllowPNG : 1;
                                // can create PNG image files
                        ULONG fNoMsoIgnore : 1;
                                // mso-ignore vglayout, vglayout2, and colspan-rowspan

                        /* Down-level HTML export options for image */
                        ULONG fImgAttributes : 1;
                                // fire msodgheDghesRequestImgAttributes for <img> attributes

                        /* Down-level HTML export options for layout */
                        ULONG fTabular : 1;
                                // we are inside a table (Excel)
                        ULONG fAlignTable : 1;
                                // align tables
                        ULONG fWriteWrapBlock: 1;
                                // host supports using inline space in the flow
                        ULONG fWantsDirectionLTR : 1;
                                // write direction:ltr on shapes (Excel BIDI)
                        ULONG fAbsoluteOffsets : 1;
                                // we'll write left, top instead of margin-left, margin-top

                        /* Options for attached text */
                        ULONG fWrapTextPassThruWithTable : 1;
                                // enclose <div> with extra VGX-only <table> (Word needs this)
                        ULONG fForceTextPassThru : 1;
                                // text pass-thru for attached text (for Excel comments)
                        ULONG fNoTextPassThru: 1;
                                // turn off optimization (for sendto mail)
                        ULONG fNeedsTextPassThruHeader : 1;
                                // text passthru header info (Excel sometimes marks fonts)
                        ULONG fAddSlopToTextboxes : 1;
                                // adds two pixel allowance to rcv

                        /* Host site information */
                        ULONG fEarlyRequestPosition : 1;
                                // want to get host anchor before flattening (Word)
                        ULONG fWantsSpecialObjComments : 1;
                                // class=shape conditionals (for charts etc in publish)
         
                        ULONG fNoInkPassThru : 1;
                               // No special case for ink. Flatten ink with other shapes

                        ULONG grfUnused : 12;
                        };
                ULONG grf;
                };

        // these strings are used for the conditional comments of sepcial objs
        WCHAR *pwchSpecial;
        WCHAR *pwchNotSpecial;

        MsoDEBDefine(dghe);
        } MSODGHESI;
#define cbDGHESI (sizeof(MSODGHESI))

/* You should call MsoInitDghesi on any new MSODGHESI you're filling out
        before passing it into a method that creates a DGHE object. */
MSOAPI_(void) MsoInitDghesi(MSODGHESI *pdghesi);

/* GRFDGHE -- Arguments for various DGHE methods */
/* In all cases msogrfdgheNil (0) is a reasonable default. */
#define msogrfdgheNil              (0)
/* The following options are for use by FExportShape method. */
#define msofdgheIsmap              (1<<0)
#define msofdgheForceTextPassThru  (1<<1)
#define msofdgheNoXMLForPassThru   (1<<2)
#define msofdgheAttachHLink        (1<<3) // TODO NeerajM(BarnWL): get rid of this
#define msofdgheHorizontalRule     (1<<4)
#define msofdgheInlinePicture      (1<<5)
#define msofdgheInHyperlink        (1<<6)
#define msofdgheWriteExcelConditionalComment (1<<7)
#define msofdgheBlockVML           (1<<8)


/****************************************************************************
        Drawing HTML Export (DGHE)

        DGHEs inherit from ISimpleUnknown instead of IUnknown; they keep
        no ref count.

        To create a DGHE you call IMsoHTMLExport::FCreateDrawingHTMLExport.
        To free it you call IMsoDrawingHTMLExport::Free().
****************************************************************** NeerajM */

/* IMsoDrawingHTMLExport (idghe) */
#undef  INTERFACE
#define INTERFACE IMsoDrawingHTMLExport
DECLARE_INTERFACE_(IMsoDrawingHTMLExport, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingHTMLExport methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* Free() tells a DrawingHTMLExport to go away. */
        MSOMETHOD_(void, Free) (THIS) PURE;

        // ----- IMsoDrawingHTMLExport methods (other)
        // basic initialization
        MSOMETHOD_(BOOL, FInit) (THIS) PURE;
        // add a view, will enumerate shapes and insert in internal layout
        MSOMETHOD_(BOOL, FAddView) (THIS_ IMsoDrawingView *pidgv) PURE;
        // after inserting all views, will allow secondary processsing
        MSOMETHOD_(BOOL, FEndAddView) (THIS) PURE;
        // gets the anchor value for the next shape anchor, does not pop
        MSOMETHOD_(BOOL, FGetNextAnchor) (THIS_ POINT *ptAnchor) PURE;
        // processes the shape anchor, pops it off, plnskips is optional
        MSOMETHOD_(BOOL, FProcessShapeAnchor) (THIS_ LONG *plnskips, DWORD hetnUser) PURE;
        // gets the full anchor spanning of all the layout collection
        MSOMETHOD_(BOOL, FGetRcSpan) (THIS_ RECT *prcSpan) PURE;
        // checks if the next anchor is just to skip columns, no real processing
        MSOMETHOD_(BOOL, FIsAnchorSkip) (THIS) PURE;
        // gets the width in anchor units of the next anchor
        MSOMETHOD_(BOOL, FGetAnchorSpan) (THIS_ int *pcolspan, int *prowspan) PURE;
        // checks if the next anchor is a CSS
        MSOMETHOD_(BOOL, FCSSNextAnchor) (THIS) PURE;
        // create image, optionally write shape xml, write <img> for single shape
        MSOMETHOD_(BOOL, FExportShape) (THIS_ IMsoDrawingView *pidgv,
                MSOHSPV hspv, RECT *prcv, DWORD hetnUser, ULONG grfdghe) PURE;
        // return the MSOXMLW used for writing shape XML
        MSOMETHOD_(MSOXMLW*, Pxmlw) (THIS) PURE;
        // return the IMsoBitmapSurface
        MSOMETHOD_(interface IMsoBitmapSurface *, PibmsGet) (THIS) PURE;
        MSOMETHOD_(BOOL, FSetShapeZIndex) (THIS_ MSOHSP hsp, LONG zIndex) PURE;

        MSOMETHOD_(BOOL, FExportHspAlignAttribute)(THIS_ MSOHSP hsp);

        MSOMETHOD_(BOOL, FWantsXML) (THIS) PURE;

        MSOMETHOD_(interface IMsoDrawingHTMLExportSite *, GetClientSite) (THIS) PURE;
};


/****************************************************************************
        Drawing HTML Export Site (DGHES)

        DGHESs inherit from ISimpleUnknown instead of IUnknown; they keep
        no ref count.

        IMsoDrawingHTMLExportSite must be implemented by anyone who wants to make
        and use IMsoDrawingHTMLExport.  Fill out a MSODGHESI (use MsoInitDghesi)
        and add a pointer to your IMsoDrawingHTMLExportSite, and then pass that
        off to IMsoHTMLExport::FCreateDrawingHTMLExport.

        All the methods in IMsoDrawingHTMLExportSite take a void *
        named pvDghivs as their second argument after the "this"
        pointer.  When you create a DGHI you also pass (in the MSODGHESI)
        a pvDghivs, which will be passed back to all DGHES methods.
****************************************************************** NeerajM */

#undef  INTERFACE
#define INTERFACE IMsoDrawingHTMLExportSite
DECLARE_INTERFACE_(IMsoDrawingHTMLExportSite, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- IMsoDrawingHTMLExportSite methods (standard stuff)

        /* FDebugMessage */
        MSODEBUGMETHOD

        /* The DGHE will call DGHES::OnEvent whenever a drawing html export
                event (MSODGHEE) occurs.  See the definition of MSODGHEE for a more
                complete explanation of events.  This function will only be called
                if you set msodghesi.fWantsEvents to TRUE. */
        MSOMETHOD_(void, OnEvent) (THIS_ void *pvDghevs, MSODGHEEB *pdgheeb) PURE;

        // ----- IMsoDrawingHTMLExportSite methods (other)
        MSOMETHOD_(BOOL, FTextPassThruHeader) (THIS_ MSOHSP hspTextbox,
                MSOTXID txid, void *pvDghevs) PURE;
        MSOMETHOD_(BOOL, FCanPassThruText) (THIS_ MSOHSP hspTextbox,
                void *pvDghevs) PURE;
        MSOMETHOD_(BOOL, FGetTextPassThruParams) (THIS_ MSOHSP hspTextbox,
                int *havH, int *havV, int *icvkH, int *icvkV, RECT *prcvTextBounds) PURE;
        MSOMETHOD_(BOOL, FExportPassThruText) (THIS_ interface IMsoHTMLExport *pihe,
                MSOHSP hspTextbox, MSOTXID txid, void *pvDghevs, BOOL fXMLOnly) PURE;
        /* fBeforeImg set when called before DGHE::FExportImage, fBeforeImg reset
                when called after DGHE::FExportImage */
        MSOMETHOD_(BOOL, FExportHspClientData)(THIS_ LPVOID pvClient,
                interface IMsoHTMLExport *pihe,
                interface IMsoHTMLExportUser *piheu,
                IMsoDrawing *pidg,
                LPVOID pvDgs,
                MSOHSP hsp,
                BOOL fBeforeImg) PURE;
        /* NOTE; only word implement this methods, all other apps can just return
                FALSE. This because, only word has inline ocx shape, which has spt of
                msosptPictureFrame, not msosptHostControl */
        MSOMETHOD_(BOOL, FOcxHsp)(THIS_ MSOHSP hsp, void* pvDgs, INT msoxocx) PURE;
};




/****************************************************************************
        IMsoDrawingSelection (idgsl)

        FUTURE PeterEn: Comment?
****************************************************************************/

/* IMsoDrawingSelectionSite */
#undef  INTERFACE
#define INTERFACE IMsoDrawingSelectionSite
DECLARE_INTERFACE_(IMsoDrawingSelectionSite, ISimpleUnknown)
{
        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSODEBUGMETHOD

        /* The Selection will call DGSLS::OnEvent when various interesting
                events take place.  See the definition of MSODGE (Drawing Event)
                for a more detailed description of events. */
        MSOMETHOD_(void, OnEvent) (THIS_ MSOGV gvDgsls, MSODGE dge,
                BOOL *pfResult, WPARAM wParam, LPARAM lParam) PURE;
};

/* DGSLSI -- Drawing Selection Site Info */
typedef struct _MSODGSLSI
        {
        interface IMsoDrawingSelectionSite *pidgsls; // DGSLS to use
        MSOGV gvDgsls;
        union
                {
                struct
                        {
                        ULONG fDefaults : 1;
                           /* TRUE iff this is an empty selection just for the
                              purposes of applying properties to the defaults.*/
                        ULONG fKeepChildren : 1;
                           /* TRUE iff client wants to keep group children selected.
                           Automation code needs this feature.
                           Use carefully since it implies that client is now responsible
                           for keeping track of children Z-Order and cross-grouping. */
                        ULONG fNoUI : 1;
                           /* TRUE iff client uses this dgsl for NON UI-related purpose like
                              in Ole Automation. It will not trigger a DG validation event even
                              if the current DG is invalidated. */
                        ULONG fTable : 1;
                           /* TRUE iff the selection contains elements of a PowerPoint
                              native table. If also fKeepChildren is TRUE, it is a cell
                              selection, otherwise at least one native table is selected. */
                        ULONG fChildrenOnly : 1;
                           /* TRUE iff child shapes are the only ones that are selected.
                              Automation code needs this feature.*/
                        ULONG grfUnused : 27;
                        };
                ULONG grf;
                };
#if DEBUG
        LONG lVerifyInitDgslsi; // Used to make sure people use MsoInitDgslsi
#endif // DEBUG
        } MSODGSLSI;
#define cbMSODGSLSI (sizeof(MSODGSLSI))

/* You should call MsoInitDgslsi on any new DGSLSI you're filling out
        before passing it into any functions that create DGSLs.  This lets us add
        new fields without changing code in all the hosts. */
MSOAPI_(void) MsoInitDgslsi(MSODGSLSI *pdgslsi);



#define msofdgslqAll                            0xFFFFFFFF
#define msofdgslqIsEmpty                        (1<<0)
#define msofdgslqIsOneShapeWithText             (1<<1)
#define msofdgslqHasPicture                     (1<<2)  // pib property is NULL:
                                                        // WARNING: false for delay loaded blips!
#define msofdgslqHasTextbox                     (1<<3)
#define msofdgslqHasConnector                   (1<<4)
#define msofdgslqHasCallout                     (1<<5)  // Only works for root-level callouts!
#define msofdgslqCanHaveFill                    (1<<6)
#define msofdgslqCanHaveArrow                   (1<<7)
#define msofdgslqCanHaveCompoundLine            (1<<8)
/* Used by the Pattern Line fill dialog to allow/disallow entry to the UI.*/
#define msofdgslqCanHaveLineFill                (1<<9)
#define msofdgslqCanRotate                      (1<<10)
#define msofdgslqCanHave3D                      (1<<11)
#define msofdgslqCanHaveTextEFfect              (1<<12)
#define msofdgslqCanHaveTextEffect              (1<<12)
#define msofdgslqCanHaveShadow                  (1<<13)
#define msofdgslqHasLine                        (1<<14)
#define msofdgslqCanReroute                     (1<<15)
// will be TRUE if the polygon can be reshaped
#define msofdgslqCanReshape                     (1<<16)
// the CanHaveLine query should actually be called CanHaveLineWidth
#define msofdgslqCanHaveLine                    (1<<17)
#define msofdgslqHasFill                        (1<<18)
#define msofdgslqCanChangeShape                 (1<<19)
#define msofdgslqCanHaveDashing                 (1<<20)
#define msofdgslqCanSetTransparentColor         (1<<21)
#define msofdgslqHasStartArrow                  (1<<22)
#define msofdgslqHasEndArrow                    (1<<23)
#define msofdgslqIsOnePicture                   (1<<24)
#define msofdgslqCanHaveTransparency            (1<<25)
#define msofdgslqCanHaveLineColor               (1<<26)
#define msofdgslqCanHavePerspectiveShadow       (1<<27)
#define msofdgslqCanHaveGradientFills           (1<<28)
#define msofdgslqCanHaveDoubleShadow            (1<<29) // actually covers double and embossed/engraved
#define msofdgslqNeedRasterMode                 (1<<30) //Raster mode needed for correct printed
                                                        //output on Win95 using MS UNIVERSAL
                                                        //PCL driver
#define msofdgslqSuggestRasterMode              (1<<31) //Raster mode may be necessary for correct
                                                        //printed output on Win95 using
                                                        //MS UNIVERSAL PCL driver
typedef enum //TODO ilanb: take this enumeration out!
        {
        msodgslqIsEmpty,
        msodgslqIsOneShapeWithText,
        msodgslqHasPicture,
        msodgslqHasTextbox,
        msodgslqHasConnector,
        msodgslqHasCallout,
        msodgslqCanHaveFill,
        msodgslqCanHaveArrow,
        msodgslqCanHaveCompoundLine,
        msodgslqCanHaveLineFill,
        msodgslqCanHave3D,
        msodgslqCanHaveTextEffect,
        msodgslqCanHaveShadow,
        } MSODGSLQ;

#define msofssNormal         0
#define msofssStayInTextEdit 1

/* GRFDGSL -- Arguments for various DGSL methods */
/* In all cases msogrfdgslNil (0) is a reasonable default. */
#define msogrfdgslNil              (0)
/* The first half of these are reasonably general purpose options
        that work on several DGSL APIs. */
#define msofdgslRoot               (1<<0)
#define msofdgslDeep               (1<<1)
#define msofdgslNoUI               (1<<2)
#define msofdgslNoDoSelectShape    (1<<3)
#define msofdgslNoDoUnselectShape  (1<<4)
#define msofdgslNoDoChangeDgslk    (1<<5)
#define msofdgslNoAfterChange      (1<<6)
#define msofdgslNoInvalidate       (1<<7)
#define msofdgslDgmLevel           (1<<8)
#define msofdgslDgmBranch          (1<<9)
#define msofdgslDgmAssistants      (1<<10)
#define msofdgslDgmConnectors      (1<<11)
#define msofdgslUser               (1<<12) // for CSelectedShapes, real count
#define msofdgslNoCaretUpdate      (1<<13) // for EndChange
#define msofdgslDgmOnlyNode        (1<<14) // only select node in Diagram, no connector.
#define msofdgslDelayInkSelUpdate  (1<<15) // Delay ink selection update till EndChange.
/* The second half are terribly specific to particular DGSL APIs. */
#define msofdgslUpdateCaret        (1<<29) // Accessibility: hint to update caret
#define msofdgslNotSelected        (1<<30)
#define msofdgslCancelTextEditMode (1<<31) // TODO peteren: Lose this?


//MSOFGRP --- Used for adding/deleting shapes to /from a group
// Used in DGSL functions FAddShapeUndo, FRemoveShapeUndo

#define msofgrpNil                      0x00000000
#define msofgrpDelete                   0x00000001
#define msofgrpRemove                   0x00000002
#define msofgrpLastShape                0x00000004

#define msofgrpDelLast (msofgrpDelete | msofgrpLastShape)
#define msofgrpRemLast (msofgrpRemove | msofgrpLastShape)

//MSOFGRPTANGLE --- Used for adding/deleting shapes to /from a selection
// Used for the groupTangle properties

#define msofgrpTangleNil                0x00000000
#define msofgrpTangleFill               0x00000001


/****************************************************************************
    IMsoShapeIter

    This is the Next-Generation(tm) shape enumeration interface, meant to
    replace the clunky and confusing *BeginEnum* APIs.  The potentially bad
    thing about returning an interface is the risk of memory failure, but
    please try to use this way of iterating when possible, I think
    you'll like it.

    IMsoShapeIter can currently be created by IMsoDrawing and
    IMsoDrawingSelection.  You must pass in msogrfShapeIterFlags to specify
    what kinds of shapes you want to iterate.
*****************************************************************************/

/* msogrfShapeIterFlags options useful for both IMsoDrawing::FCreateShapeIter
   and IMsoDrawingSelection::CreateShapeIter */

#define msofShapeIterNone         (0)
#define msofShapeIterSafe         (1 << 0) // iter source (ie. DGSL) may change
#define msofShapeIterWantGroups   (1 << 1) // include group shapes
#define msofShapeIterWantChildren (1 << 2) // include all shapes inside groups
#define msofShapeIterOnlyShapes   (1 << 3) // !group, children, canvas bkgd

/* Options additonally useful for IMsoDrawingSelection::CreateShapeIter */

#define msofShapeIterZOrder       (1 << 4) // selection returns in z-order
#define msofShapeIterToplevelOnly (1 << 5) // includes only top-level root
        // shapes (w/ host anchors) and root groups, ignores drill-down and
        // selections in canvas
#define msofShapeIterNoDrillDown  (1 << 6) // ignores group drill-down
        // selections, but includes selections within canvas

/* Options for IMsoDrawing::FCreateShapeIter to drill into a canvas */
#define msofShapeIterCanvasContents   (1 << 7)
		// get the shapes within a canvas
#define msofShapeIterCanvasBackground (1 << 8)
		// give us back the single canvas background


#undef  INTERFACE
#define INTERFACE IMsoShapeIter
DECLARE_INTERFACE_(IMsoShapeIter, ISimpleUnknown)
{
    // there's nothing to QI
    MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

    // fetch the next HSP and increment the internal counter
    MSOMETHOD_(MSOHSP, HspNext) (THIS) PURE;
    // reset the counter to the beginning of the iteration
    MSOMETHOD_(void, Reset) (THIS) PURE;

    MSOMETHOD_(void, Release) (THIS) PURE; // free this instance
};


#undef  INTERFACE
#define INTERFACE IMsoDrawingSelection
DECLARE_INTERFACE_(IMsoDrawingSelection, ISimpleUnknown)
{
        // ----- ISimpleUnknown methods

        MSOMETHOD(QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

        // ----- FDebugMessage method

        MSODEBUGMETHOD

        // ----- IMsoDrawingSelection methods

        /* Hey, look, even though it inherits from ISimpleUnknown (like all
                the other Office Drawing interfaces), IMsoDrawingSelection actually
                has AddRef and Release methods.  We didn't inherit from IUnknown
                because I really hate the idea that every function that returns
                an IMsoDrawingSelection pointer is supposed to addref it.  Functions
                that create new DGSLs actually do addref them, functions that just
                pick one out of some other structure do not. */
        MSOMETHOD_(ULONG, AddRef) (THIS) PURE;
        MSOMETHOD_(ULONG, Release) (THIS) PURE;

        /* Like many Office Drawing objects, DGSLs have a "public" structure
                (an MSODGSLSI) embedded in them.  PdgslsiBeginUpdate and
                EndDgslsiUpdate allow the host to push new values into this structure. */
        MSOMETHOD_(MSODGSLSI *, PdgslsiBeginUpdate) (THIS_ BOOL fReadOnly) PURE;
        MSOMETHOD_(void, EndDgslsiUpdate) (THIS_ MSODGSLSI *pdgslsi,
                BOOL fChanged) PURE;

        /* PidgGetDrawing() returns a pointer to this DGSL's Drawing. */
        MSOMETHOD_(interface IMsoDrawing *, PidgGetDrawing) (THIS) PURE;

        /* FGetView is used to retrieve the DGVs referencing a particular
                DGSL.  Either ppidgv or pcdgv can be NULL, but not both.
                If you pass in pcdgv we'll set *pcdgv to the number of DGVs
                using this DGSL (and if ppidgv is NULL return TRUE if this number
                is > 0).  If you pass in ppidgv we'll set *ppidgv to point to the
                idgv-th DGV using this DGSL and return TRUE iff that DGV is non-NULL.
                Remember that DGSLs can have NO DGVs. */
        MSOMETHOD_(BOOL, FGetView) (THIS_ IMsoDrawingView **ppidgv, int idgv,
                int *pcdgv) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(BOOL, FGetFocusView) (THIS_ IMsoDrawingView **pidgv) PURE;

        /* TODO peteren: Comment */
        MSOMETHOD_(void, SetFocusView) (THIS_ IMsoDrawingView *pidgv) PURE;

        /* Call FVisible to see whether or not the selection has it's fVisible
                bit set.  A DGSL will show in a particular DGV only if both
                the DGSL's fVisible and the DGV's fSelectionVisible bits are set. */
        MSOMETHOD_(BOOL, FVisible) (THIS) PURE;

        /* SetVisible lets you specify whether or not a selection is "visible".
                This affects whether DGVs using this DGSL as their selection will
                show it or not. */
        MSOMETHOD_(void, SetVisible) (THIS_ BOOL fVisible) PURE;

        /* FClone makes another DGSL just like this one.  The new DGSL
                will have been AddRef()'d for you.  The DGSL's MSODGSLSI will
                be duplicated in the new DGSL; if the caller wants to do any more
                complicated cloning of fields therein they should just call
                PdgslsiBeginUpdate/EndDgslsiUpdate after calling FClone. */
        MSOMETHOD_(BOOL, FClone) (THIS_ interface IMsoDrawingSelection **ppidgsl)
                PURE;

        MSOMETHOD_(void, BeginChange) (THIS_ ULONG grfdgsl) PURE;
        MSOMETHOD_(void, EndChange) (THIS_ ULONG grfdgsl) PURE;

        /* FSelectShape() adds the shape specified by hsp to the selection.
                Returns TRUE if the shape is already.  Can fail for
                various reasons, including actual errors (OOM) and because
                the Shape was for some reason unselectable.  Handles any of
                msofdgslUI, msofdgslAskSite, msofdgslNotifySite,
                msofdgslIsNotSelected, msofdgslCancelTextEditMode. */
        MSOMETHOD_(BOOL, FSelectShape) (THIS_ ULONG grfdgsl, MSOHSP hsp) PURE;

        /* FUnselectShape() removes the specified Shape from the selection.
                It returns TRUE if the Shape was actually removed from the
                selection, FALSE otherwise. */
        MSOMETHOD_(BOOL, FUnselectShape) (THIS_ ULONG grfdgsl, MSOHSP hsp) PURE;

        /* FSelectAllShapes() selects every top level shape in the drawing. */
        MSOMETHOD_(BOOL, FSelectAllShapes) (THIS_ ULONG grfdgsl) PURE;

        /* FUnselectAllShapes() unselects all the Shapes. */
        MSOMETHOD_(BOOL, FUnselectAllShapes) (THIS_ ULONG grfdgsl) PURE;

        /* FSelectOneShape() is the same as FSelectShape except that
                it ALSO deselects any Shapes that might already be selected. */
        MSOMETHOD_(BOOL, FSelectOneShape) (THIS_ ULONG grfdgsl, MSOHSP hsp) PURE;

        /*Add a shape to a group*/
        MSOMETHOD_(BOOL, FAddShapeUndo) (THIS_ MSOHSP hspGroup, MSOHSP hsp, ULONG grfgrp) PURE;

        MSOMETHOD_(BOOL, FRemoveShapeUndo) (THIS_ MSOHSP hspGroup, MSOHSP hsp, ULONG grfgrp) PURE;

        MSOMETHOD_(BOOL, RecalGroupBounds) (THIS_ MSOHSP hspGroup) PURE;

        // Please see msogrfShapeIterFlags for a description of its use
        MSOMETHOD_(interface IMsoShapeIter*, CreateShapeIter) (THIS_
                ULONG msogrfShapeIterFlags) PURE;

#if BOGUS
        /* SetChangeNotify() effects whether or not the host receives high level
                selection change events. Pass in FALSE to start remembering from the
                current state of the selection, and TRUE to return to normal while
                firing any needed events for the change since the last call. */
        MSOMETHOD_(void, SetChangeNotify) (THIS_ BOOL fNotify) PURE;
#endif // BOGUS

        /* FIsShapeSelected() */
        MSOMETHOD_(BOOL, FIsShapeSelected) (THIS_ ULONG grfdgsl, MSOHSP hsp) PURE;

        /* CSelectedShapes() returns the number of selected shapes. */
        MSOMETHOD_(int, CSelectedShapes) (THIS_ ULONG grfdgsl) PURE;

        /* FGetSelectedShape() */
        MSOMETHOD_(BOOL, FGetSelectedShape) (THIS_ ULONG grfdgsl, int isp,
                MSOHSP *phsp) PURE;

        /* Multiple selections maintain knowledge of a single focus shape.
                You can use SetFocusShape() to change the focus to the shape specified
                by MSOHSP.  If that shape is not currently selected we'll assert
                and return leaving the focus unchanged. */
        MSOMETHOD_(void, SetFocusShape) (THIS_ MSOHSP hsp) PURE;

        /* FGetFocus() returns TRUE if there is a focus Shape (and returns
                the MSOHSP of that Shape in phsp) or FALSE if there isn't one. */
        MSOMETHOD_(BOOL, FGetFocusShape) (THIS_ MSOHSP *phsp) PURE;

        /* FSetSelectionMode() sets the "mode" of the current selection.
                Each in selection is always in exactly one "mode", weirdly
                enumerated via the MSODGSLK type. */
        MSOMETHOD_(BOOL, FSetSelectionMode) (THIS_ ULONG grfdgsl,
                MSODGSLK dgslk) PURE;

        /* DgslkGetSelectionMode returns the current mode (dgslk) of the DGSL. */
        MSOMETHOD_(MSODGSLK, DgslkGetSelectionMode) (THIS_ ULONG grfdgsl) PURE;

        /* Use BeginEnumerateSelectedShapes to initiate an enumeration of the
                Shapes selected by a Selection.  This enumerates them in the order
                they are stored in the selection, which is the order in which they
                were selected. If you want to enumerate them in the order they
                appear in the Drawing (Z order) you'll have to mark them and
                then use a DG enumerator to cycle through them. */
        MSOMETHOD_(void, BeginEnumerateSelectedShapes) (THIS_ MSODGEN *pdgen) PURE;

        /* Call FEnumerateSelectedShapes to continue an enumeration begun by a
                call to BeginEnumerateSelectedShapes.  We return TRUE if there were
                any more Shapes over which to enumerate and FALSE if you're done.
                If we return TRUE you can examine pdgen->hsp and pdgen->grfdgeno to
                learn the state of the enumeration. */
        MSOMETHOD_(BOOL, FEnumerateSelectedShapes) (THIS_ MSODGEN *pdgen) PURE;

        /* FMarkSelectedShapes() sets the "fSelected" bit in all the
                selected shapes.  Since multiple selections can reference
                the same drawing, we use this only temporarily.  We return
                TRUE if we actually marked shome Shapes, FALSE if the selection
                in question was empty. */
        MSOMETHOD_(BOOL, FMarkSelectedShapes) (THIS) PURE;

        /* GrfdgslqQuery can be used to ask yes/no questions of the DGSL.
                msogrfdgslq's are an #defined questions we support. */
        MSOMETHOD_(ULONG, GrfdgslqQuery) (THIS_ ULONG grfdgslq) PURE;


        // --------- Specific Command Methods, affecting selected shapes.


        /* Fetch the property identified by 'MSOPSID' of the shapes in the
                selection.'pDest' may return a pointer to an object. The client
                should never alter the object returned. Instead, it should copy
                the object, modify the object, and then do FSetProp on the
                property. FetchProp will never fail. pDest must point to a
                value that is 32bits in length. */
        MSOMETHOD_(void, FetchProp) (THIS_ MSOPID, void *) PURE;

        /* Fetch a property set identified by 'MSOPSID' of the shapes in the
           selection. It either retrieves a property set from the shape's
           property set cache, or it fills in the passed property set structure.
           NOTE: If the value of the property is not the same for all the
           selected shapes then the NINCH value for the 'opid' is
           returned.
                FetchPropSet may fail for lack of memory. */
        MSOMETHOD_(BOOL, FFetchPropSet) (THIS_ MSOPSID, void *) PURE;

        /* Set the property set identified by 'MSOPSID' of the shapes in the
           selection to the pointed value. pSrc must point to a value that
           can fit in a long. Depending on the property type, 'pSrc'
           may point to a pointer to an object. This object must be
           allocated on the heap using 'new'. The shape will assume ownership
           of the object. 'fSetDefaults' indicates whether the defaults of
                the properties should be set. FSetPropSet may fail for lack of memory. */
        MSOMETHOD_(BOOL, FSetPropSet) (THIS_ MSOPSID, const void *,
                BOOL fSetDefaults, BOOL fUpdate) PURE;

        /* Reset the property set identified by 'MSOPSID' to follow the value
                of its master or style for all of the shapes in the selection */
        MSOMETHOD_(BOOL, FResetPropSet) (THIS_ MSOPSID) PURE;

        /* Apply the preset properties to the selection. */
        MSOMETHOD_(BOOL, FApplyPreset) (THIS_ MSOPRE pre, BOOL fSetDefaults) PURE;

        /* Find the first preset in the range of presets passed that match the
                properties of the selection. Return msopreNone if no match is found.
                Ignore pidIgnore in the matching process. */
        MSOMETHOD_(MSOPRE, PreMatchPreset) (THIS_ MSOPRE preFirst, MSOPRE preLast,
                                        MSOPID pidIgnore) PURE;

        /* Use FApplyRgspp to apply a set of Shape Properties (SPPs) to
                the selected Shapes.  If the Drawing in question is making an
                Undo Record, then we'll add information to it; otherwise        we'll apply
                the properties non-undoably.  We returns TRUE on success.  Whether
                or not we succeed, we consume any "complex" properties you pass us
                (those with allocations, references, etc.). See definitions of
                msofaspp*. */
        MSOMETHOD_(BOOL, FApplyProperties) (THIS_ MSOSPP *rgspp, int cspp,
                ULONG grfaspp) PURE;

        /* Call FFetchProperties to fetch a set of Shape Properties (SPPs) from
                the selected Shapes.  You should fill out the sppid's in the MSOSPPs
                before you call us using MsoInitSpp.  If we find Shapes to fetch
                from we'll return TRUE and the MSOSPPs all filled out.  If we fetched
                properties from more than one Shape and they differed we'll return
                that MSOSPP with it's fNinch bit set.  See definitions of
                msoffspp*. */
        MSOMETHOD_(BOOL, FFetchProperties) (THIS_ MSOSPP *rgspp, int cspp,
                ULONG grffspp) PURE;

        // --------- Other IMsoDrawingSelection methods.

        /* GetDrawing is yucky.  Use PidgGetDrawing instead. */
        MSOMETHOD_(void, GetDrawing) (THIS_ interface IMsoDrawing **ppdg) PURE;

        /* Update() is a pretty shallow method; it just calls Update() on
                all of the Display Managers that are showing this selection. */
        MSOMETHOD_(void, Update) (THIS) PURE;

        /* FActivateText() attaches/activates text to the shape in selection. */
        MSOMETHOD_(BOOL, FActivateText) (THIS) PURE;

        MSOMETHOD_(BOOL, FSave) (THIS_ MSOSVB *psvb) PURE;
        MSOMETHOD_(BOOL, FLoad) (THIS_ MSOLDB *pldb) PURE;

        /* This method creates a ShapeRange which represents the shapes selected
                by this object. */
        MSOMETHOD_(BOOL, FGetDispShapeRange) (THIS_ IMsoDrawingUserInterface *pidgui, IDispatch **ppidsr) PURE;

        MSOMETHOD_(BOOL, FInCanvas) (THIS) PURE;
        MSOMETHOD_(BOOL, FTrueDrillDown) (THIS) PURE;
        MSOMETHOD_(BOOL, FTextDrillDown) (THIS) PURE;

        /* FQuery can be used to ask yes/no questions of the DGSL.
                MSODGSLQ's are an enumeration of questions we support.
                NOTE: This might go away soon! Use the cooler GrfdgslqQuery() */
        MSOMETHOD_(BOOL, FQuery) (THIS_ MSODGSLQ dgslq) PURE;

        /* TODO peteren: Replace these methods with queries */

        /* CSelectedGroups() returns the number of selected Groups. */
        MSOMETHOD_(int, CSelectedGroups) (THIS) PURE;

        /* CSelectedPictures() returns the number of selected pictures. */
        MSOMETHOD_(int, CSelectedPictures) (THIS_ BOOL fPicture) PURE;

        /* CSelectedPictures() returns the number of selected shapes we consider
        to be pictures. Here shapes with picture fills are also considered pictures*/
        MSOMETHOD_(BOOL, CSelectedOneTruePicture) (THIS) PURE;

        /* CSelectedPolygons() returns the number of selected polygons. */
        MSOMETHOD_(int, CSelectedPolygons) (THIS_ BOOL fCountLines) PURE;


        /* Use this method to find out if any child shape is selected. */
        MSOMETHOD_(BOOL, FChildShapeSelected) (THIS) PURE;

        /* This method creates a ShapeRange of the child shapes selected
                by this object. */
        MSOMETHOD_(BOOL, FGetDispChildShapeRange) (THIS_
                IMsoDrawingUserInterface *pidgui, IDispatch **ppidsr) PURE;
};

#endif //$[VSMSO]
/****************************************************************************
        DRAG stuff
****************************************************************************/
#include <drdrag.h>


#if 0 //$[VSMSO]
/*****************************************************************************
        Office Text
******************************************************************* DAVEBU **/

/****************************************************************************
        Exported functions
*****************************************************************************/
MSOAPIX_(BOOL) MsoFConvertRTF(const char *pchRTF, WCHAR **ppwchText);


/******************************************************************************
        IMsoRule

        A rule with the help of a solver enforces contraints on the properties
        of a set of shapes. Each rule governs a set of shapes that belong to
        single drawing. It acts on the shapes through shape proxies. The shape
        proxies actually apply the property changes to the shapes the rule
        governs.

        A rule is constructed with a list of shapes that it governs and the
        drawing the shapes belong in. At construction time, the rule installs
        itself in the solver of the drawing. It also registers the proxy shapes
        for the shapes that it governs.

        Rules can change. If a caller changes a rule or the number of shapes
        the rule governs, the rule should call OnRuleChange so that the solver
        mark the rule for testing. The caller should then call FSolve to
        ensure the constraint system can be solved. The caller should be
        prepared to backout changes to a rule, if the constraint system becomes
        unsolvable.

        The solver solves the contraints imposed by a set of rules in 3 steps.
        First, when some external code adjusts the properties of a proxy, the
        solver marks all the rules that govern that shape. Rules keep this
        dependency information in singly linked list. Second, the solvers
        calls FTest for all the rules that are marked. A rule is expected to
        test the properties of a shape proxy and report if they satisfy its
        constraint. Third, for all unsatisfied rules, the FTry method is called.
        When FTry is called a rule is expected to change the properties of a
        the shapes it governs so that it is satisfied (i.e. FTest returns TRUE).

********************************************************************** RICKH **/

// SCTX: The current state of the solver
typedef enum                                            // Solver context
        {
        msosctxNormal,                                  // Normal: edits, vba, and other operations
        msosctxDrag,                                    // Drag: Temporary UI
        msosctxUndo,                                    // Undo: Rolling back state
        msosctxInit,                                    // Just added. First time initialalize
        } MSOSCTX;


// RUle Block: Information about the context of the solver
typedef struct _MSORUB
        {
        MSOSCTX sctx;
        MSODRGH drgh;
        MSODGCID dgcidTool;
        } MSORUB;

#define cbMSORUB (sizeof(MSORUB))

#undef  INTERFACE
#define INTERFACE IMsoRule

DECLARE_INTERFACE_(IMsoRule, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        //--- Rule Methods

        /* Return TRUE if the conditions of the rule are met.        The rule is expected
           to query the properties of the shapes it governs */
        MSOMETHOD_(BOOL, FTest) (THIS_ const MSORUB* prub) PURE;

        /* Try to alter properties of the shapes the rule governs so they conform
                to the rule. Return TRUE if it was possible to make the shapes conform.
                On entry FTest(THIS) == FALSE. On Exit FTest(THIS) == TRUE.

                To determine what was changed the rule can look at
                        1) the MSORUB which will give the context and drag handle
                        2) call DG::FHasPropChanges to what properties are changing.
                        3) call SP::FInitator to see if the shape was the original shape
                           that was change. */
#ifndef VSPRUNEBUILD
   MSOMETHOD_(BOOL, FTry) (THIS_ const MSORUB* prub) PURE;
#endif//VSPRUNEBUILD

        /* Return TRUE if the rule governs the specified shape and property set */
        MSOMETHOD_(BOOL, FGoverns) (THIS_ MSOHSP hsp, MSOPSID osid,
                                                                MSOPID opidFirst, const void* pSrc) PURE;

        /* Set the owning shape of a rule. Only rules owned by a shape
           have this method called. */
        MSOMETHOD_(void, SetOwner) (THIS_ MSOHSP hsp) PURE;

        /* Mark shapes that this rule governs and that depend on selected
                on the selected shapes. */
        MSOMETHOD_(void, MarkDependents) (THIS) PURE;


        //--- Shape Proxies Access Methods

        /* The count of shapes that this rule governs */
        MSOMETHOD_(int,  CProxy) (THIS) PURE;

        /* The i'th (O-based) subject shape that this rule governs. */
        MSOMETHOD_(void, GetProxy) (THIS_ int iProxy, MSOHSP* phsp) PURE;

        /* The i'th (O-based) subject shape that this rule governs. */
        MSOMETHOD_(void, SetProxy) (THIS_ int iProxy, MSOHSP hsp) PURE;

        /* The proxy has been removed from the drawing. In response to
           this call, the rule may set pfRemove to TRUE. It should not call
                not RemoveRule. */
        MSOMETHOD_(void, OnProxyRemove) (THIS_ MSOHSP hsp, BOOL *pfRemove) PURE;


        //--- Marking Methods

        /* Set and get the flags on the rule. The flags on a rule allows a
                solver to associate bits of information with each rule. The
                interpretation of flags are only known to the solver. The rule does
                not have to serialize these flags. */
        MSOMETHOD_(BOOL, FFlag) (THIS_ ULONG mask) PURE;
        MSOMETHOD_(void, SetFlag) (THIS_ ULONG mask, BOOL f) PURE;

        /* Set and get the LID for the rule. The LID is a persistent id for the
                rule. The rule should serialize this id. */
        MSOMETHOD_(MSORUID, Ruid) (THIS) PURE;
        MSOMETHOD_(void, SetRuid) (THIS_ MSORUID id) PURE;


        //--- Serialize Methods

        /* Load and save the rule into the stream */
#ifndef VSPRUNEBUILD
        MSOMETHOD_(BOOL, FSave) (THIS_ MSOSVB *psvb) PURE;
#endif//VSPRUNEBUILD

#if NewParse
        MSOMETHOD_(BOOL, FLoad) (THIS_ MSOFBH fbh, const void* pdata) PURE;
#else
        MSOMETHOD_(BOOL, FLoad) (THIS_ MSOLDB *pldb) PURE;
#endif

        /* Get the RUT or rule type of the rule. Negative numbers are used
           by rules defined by Office. Positive numbers are for rules defined
           by the host.  */
        MSOMETHOD_(void, GetRuleType) (THIS_ MSORUT *prut) PURE;

        /* Duplicate a rule. If fCheckMarks is set, the new rule contains hsp's
                for marked shapes. All other proxies are nil. The proxies for the new rule
                must be fixed up by the caller. */
#ifndef VSPRUNEBUILD
        MSOMETHOD_(BOOL, FDuplicate) (THIS_ IMsoRule** ppiru,
                interface IMsoDrawing* pidgTo, BOOL fCheckMarks) PURE;
#endif//VSPRUNEBUILD

        /* Given the shapes marked for copying, should this rule be copied as well? */
        MSOMETHOD_(BOOL, FShouldCopy) (THIS) PURE;

        /* Compare the properties of two rules */
        MSOMETHOD_(BOOL, FIsEqual)      (THIS_ IMsoRule* piru) PURE;

        /* Given an undo record that has been registered with the Drawing Site
                Execute, Free, Save, or Load the record */
        MSOMETHOD_(BOOL, FExecuteUndoRecord) (THIS_ MSOHRUR hrur) PURE;
        MSOMETHOD_(void, FreeUndoRecord)  (THIS_ MSOHRUR hrur) PURE;
        MSOMETHOD_(BOOL, FSaveUndoRecord) (THIS_ MSOSVB *psvb, MSOHRUR hrur) PURE;
        MSOMETHOD_(BOOL, FLoadUndoRecord) (THIS_ MSOLDB *pldb, MSOHRUR* phrur,
                                           interface IMsoDrawing* pidg) PURE;
        MSOMETHOD_(BOOL, FWriteBeHRUR) (THIS_ MSOHRUR hrur, LPARAM lparam, BOOL fSaveObj) PURE;
        // The following API is for internal use only.
#ifndef VSPRUNEBUILD
        MSOMETHOD_(BOOL, FWriteXMLStack) (THIS_ void *pxmlw, const char *pchId,
                int cchId, const char *pchType, int cchType) PURE;
        MSOMETHOD_(BOOL, FReadXML) (THIS_ struct XMLRule *pxlRule) PURE;
        MSOMETHOD_(BPSC, BpscBulletProof) (THIS_ MSOBPCB *pmsobpcb, void *pbstore) PURE;
        MSOMETHOD_(BOOL, FIsValid) (THIS) PURE;
#endif//VSPRUNEBUILD
        };

// Masks for SetFlag
#define msoruTestMask           ((ULONG)0x01)
#define msoruChangedMask        ((ULONG)0x08)

/****************************************************************************
        Drawing Dialog functions.
****************************************************************************/

/* These are sdm escher dialog initialization.  This are only
   needed for Macintosh implementation.  These are going away. */

MSOAPIX_(BOOL) MsoFInitDlgs();

/* Free up things needed for the hacks we use to make Mac SDM
        work correctly with WLM-based code.
   The lazy Initialization will happen automatically when the first Office/Escher
        dialog is opened.  However, the host should call this when they are
        terminating. */

MSOAPI_(VOID) MsoFTermDlgs();

/* The is the enum value returned for the Picture tab on the
   Fill Effect dialog which is used by Chart.  These values
        are used when the MSOFILLPICTURE is specified for the
        creation flags on the fill dialog. */

typedef enum
        {
        msofillformatStretch,
        msofillformatStack,
        msofillformatStackScale,
        } MSOPICFORMATTYPE;

/* This extension structure is used by the MSOFILLPICTURE dialog
   to pass to and from the fill effect dialog the information
        needed for the picture tab used by Chart. */

#define MSOSCALEWZMAX                     255

typedef struct _EXFILLCONTROL
        {
        MSOPICFORMATTYPE        iFormat;
        BOOL                    fSides;
        BOOL                    fFront;
        BOOL                    fEnd;
        WCHAR                   wzScale[MSOSCALEWZMAX];
        WCHAR*                  wzPath;
        ULONG                   ulFlags;
        } MSOEXFILLCONTROL;

/* These flag values are used to tell the picture tab in the
   Fill Effect dialog which controls or options are  not
   allowed.  Either the option will not show up in the
   dropdown or the control itself will be disabled.  */

#define MSOFILLNOSTRETCH                        0x00000001
#define MSOFILLNOTILE                           0x00000002
#define MSOFILLNOSTACK                          0x00000004
#define MSOFILLNOSTACKSCALE                     0x00000008
#define MSOFILLAPPLYNOSIDES                     0x00000010
#define MSOFILLAPPLYNOFRONT                     0x00000020
#define MSOFILLAPPLYNOEND                       0x00000040
#define MSOFILLPICTUREMASK                      0x000000FF
/* Shows the picture tab. For Chart/Graph client. */
#define MSOFILLPICTURE                          0x00000100
/* Shows the Gradient/Shaded fill tab */
#define MSOFILLSHADE                            0x00000200
/* Shows the textured fill tab */
#define MSOFILLTEXTURE                          0x00000400
/* Shows the Pattern fill tab */
#define MSOFILLPATTERN                          0x00000800
/* Shows the picture tab without extra controls.  Regular office tab. */
#define MSOFILLPICTURE2                         0x00001000
/* Passed to MsoGetShadeInfoFromVariant as a switch to determine if the
   fill shade type is concentric shapes or just a rectangle for fill From
        Center.  */
#define MSOFILLSHADESHAPE                       0x00002000
/* Used to have the Gradient fill style to show From Center,
   instead of From Title (which is used by Powerpoint). */
#define MSOFILLDLGAPPLY                         0x00004000
/* Hide the "Semi-Transparnet" checkbox */
#define MSOFILLDLGHIDETRANSPARENT               0x00008000
/* This new HIDE transparent flag is used only by Charting to hide the transparency
   checkbox on the texture and picture tab only when they pass it.*/
#define MSOFILLHIDETRANSPARENT2                 0x00040000
#define MSOLINEPATTERN                          0x00080000       // used for line patterns in dialog
/* Used to enable the Control for Embed ot Link to File in the Insert Picture
   dialog from the Texture or Picture tab */
#define MSOFILLBLIPEMBEDORLINK                  0x00100000
/* Shows the tints tab. For Publisher. */
#define MSOFILLTINTS                            0x00200000
/* Should the Fill dialog have special behavior because
        it is in spot mode? Used by Publisher when the publication
        is setup for spot printing */
#define MSOFILLSPOTMODE                         0x02000000


/* The following modal modeless flags are used for the
   fill dialog and other escher dialogs (Color dialog,
   Insert Text Effect dialog, and Edit Your Text Dialog).*/
#define MSODLGMODELESS                          0x00010000
#define MSODLGUIMODAL                           0x00020000


/* Show CMYK in the Colors dialog. For Publisher. */
#define MSOCOLORSCMYK                           0x00400000
/* Show Registered CMSs in the Colors dialog. May also
        cause "Unknown" to appear if this is set and the current
        color uses an unregistered CMS. For Publisher. */
#define MSOCOLORSCMS                            0x00800000
/* Should the More Colors dialog hide the RGB and
        HSL options. Used by Publisher when the publication
        is setup for Process printing */
#define MSOCOLORSPROCESSMODE                    0x01000000
/* Overrides the default HLS range in the color dialog.  The default
        range for hue, sat, and lum is [0,255].  Using this flag will
        setup the range to be hue=[0,240) and sat,lum=[0,240] */
#define MSOCOLORSOVERRIDEHLSMAX                 0x04000000
/* The following 3 defines shows which tabs to show in colors dialog */
#define MSOCOLORSTABSTANDARD                    0x08000000
#define MSOCOLORSTABCUSTOM                      0x10000000
#define MSOCOLORSTABUSERDEFINED                 0x20000000


/* These are defines for Used Defined tab in Colors dialog */
// show scroll control to select tint
#define MSOCOLORSUSETINTS                       0x00000001
// show radiobuttons/checkbox to be able to keep luminosity
#define MSOCOLORSKEEPLUMINOSITY                 0x00000002
// hide scroll control for transparency (for inks in Pub)
// #define MSOCOLORSHIDETRANSPARENCY  MSOFILLDLGHIDETRANSPARENT
// show tints in color preview
#define MSOCOLORSSHOWTINTSPREVIEW               0x00000008

typedef void  (CALLBACK *MSOPFNPREVIEW)(int msg, LONG lParam);
#define msopfnPreviewNil ((MSOPFNPREVIEW)0)

/****************************************************************************
        Fill Effect Drawing Dialog

        This multi-tab dialog contains the following tab:

        Gradient fill
        Texture fill
        Pattern fill
        Picture fill

****************************************************************************/

#undef  INTERFACE
#define INTERFACE IMsoFillsDlg

DECLARE_INTERFACE_(IMsoFillsDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        /* This is only useable when the dialog is run modelessly.
           This is used to move the Office assistant out of the way of the
                dialog rectangle.  */
        MSOMETHOD_(void, AvoidDialog) (THIS_ interface IMsoDrawingUserInterface *pidgui) PURE;

        /*
         * This returns the extra structure MSOEXFILLCONTROL filled if provided.
         * the pchPicPath must be copied by the caller to their own data
         * heap or stack.  After Free is called on this object, that pchPicPath
         * is invalid.
         */
        MSOMETHOD_(void, Get) (THIS_ MSOPSFillStyle *popsFill, MSOEXFILLCONTROL *pefc) PURE;
        MSOMETHOD_(void, Set) (THIS_ MSOPSFillStyle *popsFill) PURE;

        /* This runs the dialog either modal or modeless depending on how
           the dialog was created. */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;
        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;

        /* To set the filename of the browsed image for either the texture or
                picture tab.
           This is currently only used by PowerPoint for their UI modal dialogs.
           The filename length max should be MAX_PATH.*/
        MSOMETHOD_(void, SetFilename) (THIS_ const WCHAR *wzFilename, BOOL fTexture) PURE;

        MSOMETHOD_(void, SetColor) (THIS_ COLORREF color, unsigned int tmc) PURE;

        /* This sets this dialog as the current dialog */
        MSOMETHOD_(BOOL, FSetCurrent) (THIS) PURE;

        MSOMETHOD_(void, SetPictFill) (THIS_ BOOL fPictureFills, BOOL fAspect) PURE;
        MSOMETHOD_(BOOL, GetPictFill) (THIS) PURE;
        MSOMETHOD_(BOOL, GetPFillsChanged) (THIS) PURE;
        MSOMETHOD_(void, SetAutoCheck) (THIS_ BOOL fAutoCheck) PURE;
        MSOMETHOD_(BOOL, GetAutoCheck) (THIS) PURE;
        MSOMETHOD_(BOOL, GetAspectRatio) (THIS) PURE;
        MSOMETHOD_(void, SetBackgroundFlag) (THIS_ BOOL fBack) PURE;
        MSOMETHOD_(BOOL, GetBackgroundFlag) (THIS) PURE;
        MSOMETHOD_(void, SetOpacity) (THIS_ LONG lOpacityFrom, LONG lOpacityTo);
        MSOMETHOD_(void, GetOpacity) (THIS_ LONG* lOpacityFrom, LONG* lOpacityTo);
        MSOMETHOD_(void, SetSpecialLockFlag) (THIS_ BOOL fLock) PURE;
        };


/* To create the Escher fill effect dialog.  The pidgui parameter
   is now needed for the office assistant and also to get the
   document's list of Blips.

        The pib that is passed through here is for the Format Object dialog
        which may call the fill effect dialog and may select a custom texture.
        But it may call this dialog again before applying that custom texture
        to the selection.  Thus, the custom texture needs to be passed in addition
        to looking at all the shapes in the drawing.
        */
MSOAPI_(BOOL) MsoFCreateFillsDlg(HMSOINST pinst, IMsoToolbarSet *pitbs,
        HWND hwndParent, ULONG flFlags, MSOEXFILLCONTROL *pefc,
        IMsoDrawingUserInterface *pidgui,
        MSOPFNPREVIEW pfnPreview, IMsoBlip *pib, const WCHAR *wzBlipName,
        IMsoFillsDlg **ppidlg);

/* FILL EFFECT FLAGS:  Pass these as creation flag so that the
   appropriate tab will be shown.  Otherwise, that tab is
        hidden.
   */



/* Gradient Fill effect dialog control values.  */
typedef enum
        {
   msoccNone = -1,
        msoccOne,
        msoccTwo,
        msoccPreset,
        } MSOCOLORDLGCTL;

/* Gradient fill style value for the style radiobutton group. */
typedef enum
        {
        msossNone = -1,
        msossHorizontal,
        msossVertical,
        msossDiagUp,
        msossDiagDown,
        msossFromCorner,
        msossFromCenter,
        msossMaxShadeStyles
        } MSOCOLORDLGSTYLE;

/* Gradient fill variant enum. */
typedef enum
        {
        msovarNone = -1,
        msovar0,
        msovar1,
        msovar2,
        msovar3,
        msovarMax
        } MSOVARIANT;



/* Return the focus number, fillType, fillToRect, and
        angle for the given variant number and style.  The opposite is MsoVarGetVariant */
MSOAPI_(void) MsoGetShadeInfoFromVariant(MSOVARIANT var,
                                                                                                          MSOCOLORDLGSTYLE style,
                                                                                                          BOOL fShadeShape,
                                                                                                          MSOFILLTYPE *pfillType,
                                                                                                          LONG *pfillFocus,
                                                                                                          RECT *prcFillTo,
                                                                                                          LONG *pfillAngle);

/*      Compute the fill and back color for the preset gradient fill. */
MSOAPI_( BOOL ) MsoFSetForeBackForPresetGradient(MSOPRESET pre, COLORREF *pcrFore,
                                                                                                  COLORREF *pcrBack);

/* This API is to get the brightness (0->65535) from a system back color */
MSOAPI_(LONG) MsoLGetBrightness(COLORREF clrSysBack);
/* This API is to set the brightness (0->65535) and return a system back color */
MSOAPI_(COLORREF) MsoCrSetBrightness(LONG brightness);
/* This API is to derive the rgb back color from an rgb fore color and syscolor
        containing the brightness value from one-color gradient fills */
MSOAPI_(COLORREF) MsoCrGetBackColorValue(COLORREF clrFore, COLORREF clrSysBack);
/* Gets the fore or fill Color of a texture. */
MSOAPI_(COLORREF) MsoCrGetTextureForeColor( int tag );
/* Get fore and background color for preset gradients */
MSOAPI_(BOOL) MsoFGetForeBackForPresetGradient(MSOPRESET pre, COLORREF *pcrFore, COLORREF *pcrBack);
/* Returns the angle for the specified style and variant.  */
MSOAPIX_(LONG) MsoLGetShadeAngle(MSOCOLORDLGSTYLE style, int ivar);
/*      Return the shaded fill style - one of the radiobutton styles
        on the gradient/shaded fill tab -*/
MSOAPI_(MSOCOLORDLGSTYLE) MsoLGetStyle(MSOFILLTYPE fillType, LONG lAngle, const RECT *prcFillTo);
/*      Return the shaded fill style variant - one of the radiobutton style variant
        on the gradient/shaded fill tab -*/
MSOAPI_(MSOVARIANT) MsoLGetVariant(MSOFILLTYPE fillType, LONG lFocus, LONG lAngle,
                                         RECT *prcFillTo, MSOCOLORDLGSTYLE style);


/* Returns the MSOSHADECOLOR array for the msopreShadePreset1 through
        msopreShadePreset16.  This calls MsoFCopyProp and passes that storage to the
        caller.  The caller is responsible for deallocating and tracking that storage.*/
MSOAPI_(IMsoArray *)MsoRgCrGradientPreset(MSOPRESET pre);
/* Return the msopre value.  The valid range is msopreShadePreset1 through
        msopreShadePreset16 or msopreNone.  It will try to match the MSOSHADECOLOR array
        passed via the IMsoArray with the array used for the preset gradient fills.
        **prgshclr must be a valid pointer to a *IMsoArray.  If null, will return msopreNone.*/
MSOAPI_(MSOPRESET) MsoPreMatchGradient(IMsoArray **prgshclr);
/* Return the string of the blip name for the built in textures and patterns.  This
   will allocate the storage using the MsoPvAlloc.  The caller is responsible for
   deallocating the storage.  */
MSOAPI_(BOOL) MsoFCopyBlipName(LONG tag, WCHAR **ppwz);
/* Retrieve the preset properties values that are within the MsoPSGeoText set which
   are used from the Insert WordArt (TextEffect) dialog.  */
MSOAPI_(void) GetGeoTextFromPreset(int index, MSOPSGeoText *popsGeoText);


/****************************************************************************
        User Defined tab in More Color Drawing Dialog

****************************************************************************/
#undef  INTERFACE
#define INTERFACE IMsoUserDefList

//these bits are passed in flags when calling IMsoUserDefList.DrawItem
#define MSODRAWFLAGSELECTED flbeSelected
#define MSODRAWFLAGCURSOR   flbeCursor

DECLARE_INTERFACE(IMsoUserDefList)
        {
        MSOMETHOD_(BOOL, DrawItem) (THIS_ HDC hdc, int iItem, RECT *prc, int iFlags) PURE;
        MSOMETHOD_(BOOL, GetWctInfo) (THIS_ int iItem, WCHAR *wz, int cchMax) PURE;
        MSOMETHOD_(BOOL, GetColor) (THIS_ int iItem, MSOCOLOREXT *pClr, LONG *plOpacity) PURE;
        };


/****************************************************************************
        More Color Drawing Dialog

        This multi-tab dialog contains the following tabs:

        Standard Color
        Custom Color
        User Defined Colors

****************************************************************************/

#undef  INTERFACE
#define INTERFACE IMsoColorsDlg

DECLARE_INTERFACE_(IMsoColorsDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        /* This sets and gets the color to and from the color
           dialog.  Set should be called before FRun.  */
        MSOMETHOD_(void, Get) (THIS_ COLORREF *pcolor, LONG *plOpacity) PURE;
        MSOMETHOD_(void, Set) (THIS_ BOOL fSolid, COLORREF *pcolor, LONG *plOpacity) PURE;

        /* This is only useable when the dialog is run modelessly.
           This is used to move the Office assistant out of the way of the
                dialog rectangle.  */
        MSOMETHOD_(void, AvoidDialog) (THIS_ interface IMsoDrawingUserInterface *pidgui) PURE;

        /* This runs the dialog either modal or modeless depending on how
           the dialog was created. */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;
        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;

        /* This sets this dialog as the current dialog */
        MSOMETHOD_(BOOL, FSetCurrent) (THIS) PURE;

        /* Extended color versions of Get and Set. GetExt may give the caller back
                an allocated string that the caller is then responsible for freeing. SetExt
                will make its own copy of any string passed in so the caller maintains
                ownership of any string passed in. */
        MSOMETHOD_(void, GetExt) (THIS_ MSOCOLOREXT *pclrExt, LONG *plOpacity) PURE;
        MSOMETHOD_(void, SetExt) (THIS_ BOOL fSolid, const MSOCOLOREXT *pclrExt, LONG *plOpacity) PURE;


        /* Set additional info about Custom tab */
        MSOMETHOD_(BOOL, SetUserDefinedTabInfo) (
                THIS_ ULONG flFlags,
                IMsoUserDefList *pUserDefList,
                const WCHAR *pwzTitle,
                const WCHAR *pwzLabel,
                int iCurrentColor,
                int iCurrentTint,
                int iEntryHeight,
                int cColors
                ) PURE;

        MSOMETHOD_(BOOL, GetUserDefinedColor) (THIS_ int *piIndex, int *piTint) PURE;

        };

/* The MsoFCreateColorsDlg is used to create the color dialog.
   It returns the IMsoColorsDlg pointer as non-null if it succeeds and
        return TRUE from the call.  The IMsoDRawingUserInterface is needed
        to properly coordinate the dialog's rectangle with the Office
        Assistant and is used for sticky or centering of the dialog.

        The MSOFILLDLGTRANSPARENT is used to hide or show the transparent
        checkbox.  If the pwzTitle is non null and is a valid string,
        it is used in the dialog caption instead of the one from the
        dialog built in template. */

MSOAPI_(BOOL) MsoFCreateColorsDlg(HWND hwndParent, ULONG flFlags,
                                                                                         IMsoDrawingUserInterface *pidgui,
                                                                                         MSOPFNPREVIEW pfnPreview,
                                                                                         const WCHAR *pwzTitle,
                                                                                         IMsoColorsDlg **ppidlg);


/* The following structure is used to set the color data retrieved
   from the dialog and is used to execute an escher command to
   set the color into the appropriate property slot.  */

typedef struct
        {
        MSOPID pidColor;
        MSOPID pidOpacity;
        COLORREF clr;
        LONG lOpacity;
        } MSODLGColorStyle;

typedef struct
        {
        MSOPID pidColor; // pid of the 'main' color.
        MSOPID pidOpacity;
        MSOCOLOREXT clrExt;
        LONG lOpacity;
        } MSODLGColorExtStyle;


#define msotmcFmtObjTabSwitcher (tmcUserMin + 399)
//These numbers are arbitrarily chosen, if we need another or a larger range,
//we'll just change these.
#define msotmcMin (tmcUserMin + 400)
#define msotmcMax (msotmcMin + 1000)
#define MsoFIsOfficeTmc(tmc)    (((tmc) >= msotmcMin) && ((tmc) < msotmcMax))

/* DO NOT MODIFY THE ORDER of these tmc's, a few hacks which I (ilanb) plan to
 remove depend on the ordering. If you need to shuffle things around, please
 contact me! */

// tmc's for office Fill & Line tab
#define msotmcFillMin                           (tmcUserMin + 400) //1424
#define msotmcFillLineColor                     (msotmcFillMin + 0)
#define msotmcFillLineBackColor                 (msotmcFillMin + 1)
#define msotmcFillLinePattern                   (msotmcFillMin + 2)
#define msotmcFillLineTransparent               (msotmcFillMin + 3)
#define msotmcFillLineColorPic                  (msotmcFillMin + 4)
#define msotmcFillLineStyle                     (msotmcFillMin + 5)
#define msotmcFillLineStylePic                  (msotmcFillMin + 6) //1430
#define msotmcFillLineWidthPic                  (msotmcFillMin + 7)
#define msotmcFillLineDash                      (msotmcFillMin + 8)
#define msotmcFillLineDashPic                   (msotmcFillMin + 9)
#define msotmcFillLineConnect                   (msotmcFillMin + 10)
#define msotmcFillLineConnectPic                (msotmcFillMin + 11)
#define msotmcFillArrowStyleL                   (msotmcFillMin + 12)
#define msotmcFillArrowStyleLPic                (msotmcFillMin + 13)
#define msotmcFillArrowStyleR                   (msotmcFillMin + 14)
#define msotmcFillArrowStyleRPic                (msotmcFillMin + 15)
#define msotmcFillArrowSizeL                    (msotmcFillMin + 16)    //1440
#define msotmcFillArrowSizeLPic                 (msotmcFillMin + 17)
#define msotmcFillArrowSizeR                    (msotmcFillMin + 18)
#define msotmcFillArrowSizeRPic                 (msotmcFillMin + 19)
#define msotmcFillFillPic                       (msotmcFillMin + 20)
#define msotmcFillMoreColorsButton              (msotmcFillMin + 21)
#define msotmcFillEffectsButton                 (msotmcFillMin + 22)
#define msotmcFillFillColor                     (msotmcFillMin + 23)
#define msotmcFillBackColor                     (msotmcFillMin + 24)
#define msotmcFillShadeStyle                    (msotmcFillMin + 25)
#define msotmcFillVariant                       (msotmcFillMin + 26)    //1450
#define msotmcFillShadeBrightness               (msotmcFillMin + 27)
#define msotmcFillTransparent                   (msotmcFillMin + 28)
#define msotmcFillCColor                        (msotmcFillMin + 29)
#define msotmcFillPreset                        (msotmcFillMin + 30)
#define msotmcFillBlip                          (msotmcFillMin + 31)
#define msotmcFillType                          (msotmcFillMin + 32)
#define msotmcFillFillColorExt                  (msotmcFillMin + 33)
#define msotmcFillBackColorExt                  (msotmcFillMin + 34)
#define msotmcFillLineColorExt                  (msotmcFillMin + 35)
#define msotmcFillLineBackColorExt              (msotmcFillMin + 36)    //1460
#define msotmcFillLineWidth                     (msotmcFillMin + 37) //first non hidden edit
#define msotmcFillNoShade                       (msotmcFillMin + 38)
#define msotmcFillHR                            (msotmcFillMin + 39)
#define msotmcFillTransparency                  (msotmcFillMin + 40)
#define msotmcFillTransparencyPer               (msotmcFillMin + 41)
#define msotmcFillMax                           (msotmcFillMin + 42)    //1466

//tmc's for office Size tab
#define msotmcSizeMin                           (msotmcFillMax) //1466
#define msotmcSizeHeight                        (msotmcSizeMin + 0)
#define msotmcSizeWidth                         (msotmcSizeMin + 1)
#define msotmcSizeHeightPer                     (msotmcSizeMin + 2)
#define msotmcSizeWidthPer                      (msotmcSizeMin + 3)
#define msotmcSizeRotation                      (msotmcSizeMin + 4) //1470
#define msotmcSizeRelative                      (msotmcSizeMin + 5)
#define msotmcSizeLockAspect                    (msotmcSizeMin + 6)
#define msotmcSizeHeightOrig                    (msotmcSizeMin + 7)
#define msotmcSizeWidthOrig                     (msotmcSizeMin + 8)
#define msotmcSizeReset                         (msotmcSizeMin + 9)
#define msotmcSizeBestScale                     (msotmcSizeMin + 10)
#define msotmcSizeResolution                    (msotmcSizeMin + 11)
#define msotmcSizeMax                           (msotmcSizeMin + 12) //1478

//tmc's for office Position tab
#define msotmcPosMin                            (msotmcSizeMax) //1478
#define msotmcPosHorPos                         (msotmcPosMin + 0)
#define msotmcPosVerPos                         (msotmcPosMin + 1)
#define msotmcPosHorFrom                        (msotmcPosMin + 2) //1480
#define msotmcPosVerFrom                        (msotmcPosMin + 3)
#define msotmcPosMovWText                       (msotmcPosMin + 4)
#define msotmcPosLockAnchor                     (msotmcPosMin + 5)
#define msotmcPosLockPos                        (msotmcPosMin + 6)
#define msotmcPositionLabel                     (msotmcPosMin + 7)
#define msotmcPosMax                            (msotmcPosMin + 8) //1486

//tmc's for office Picture tab
#define msotmcPicMin                            (msotmcPosMax) //1486
#define msotmcPicCropLeft                       (msotmcPicMin + 0)
#define msotmcPicCropRight                      (msotmcPicMin + 1)
#define msotmcPicCropTop                        (msotmcPicMin + 2)
#define msotmcPicCropBottom                     (msotmcPicMin + 3)
#define msotmcPicColor                          (msotmcPicMin + 4) //1490
#define msotmcPicBrightness                     (msotmcPicMin + 5)
#define msotmcPicContrast                       (msotmcPicMin + 6)
#define msotmcPicBrightnessPer                  (msotmcPicMin + 7)
#define msotmcPicContrastPer                    (msotmcPicMin + 8)
#define msotmcPicReset                          (msotmcPicMin + 9)
#define msotmcPicRecolor                        (msotmcPicMin + 10)
#define msotmcCompress                          (msotmcPicMin + 11)
#define msotmcPicMax                            (msotmcPicMin + 12) //1497

//tmc's for office Textbox tab
#define msotmcTboxMin                           (msotmcPicMax) //1497
#define msotmcTboxMarginLeft                    (msotmcTboxMin + 0)
#define msotmcTboxMarginRight                   (msotmcTboxMin + 1)
#define msotmcTboxMarginTop                     (msotmcTboxMin + 2)
#define msotmcTboxMarginBottom                  (msotmcTboxMin + 3) //1500
#define msotmcTboxVerticalAlignment             (msotmcTboxMin + 4)
#define msotmcTboxWrapText                      (msotmcTboxMin + 5)
#define msotmcTboxResizeObjToFitText            (msotmcTboxMin + 6)
#define msotmcTboxRotateTextWithObject          (msotmcTboxMin + 7)
#define msotmcTboxTextFlow                      (msotmcTboxMin + 8)
#define msotmcTboxAutoMargin                    (msotmcTboxMin + 9)
#define msotmcTboxMax                           (msotmcTboxMin + 10) //1507

// tmc's for office paragraph formatter
#define msotmcParaMin                           (msotmcTboxMax) //1507
#define msotmcParaBefore                        (msotmcParaMin+0x01)
#define msotmcParaAfter                         (msotmcParaMin+0x02)
#define msotmcParaLineSpacing                   (msotmcParaMin+0x03) //1510
#define msotmcParaIndentBy                      (msotmcParaMin+0x04)
#define msotmcParaMax                           (msotmcParaMin+0x05) //1512

#define msotmcCanvasMin                         (msotmcParaMax)//1512
#define msotmcCanvasHeight                      (msotmcCanvasMin + 0x01)
#define msotmcCanvasWidth                       (msotmcCanvasMin + 0x02)//1514

#define msotmcFillDlgMin                        (msotmcCanvasMin + 3)//1515
#define msotmcFillDlgShadedColor1               (msotmcFillDlgMin+1)
#define msotmcFillDlgShadedColor2               (msotmcFillDlgMin+2)
#define msotmcFillDlgPatternForeColor           (msotmcFillDlgMin+3)
#define msotmcFillDlgPatternBackColor           (msotmcFillDlgMin+4)
#define msotmcFillDlgShadedColorHide1           (msotmcFillDlgMin+5) //1520
#define msotmcFillDlgShadedColorHide2           (msotmcFillDlgMin+6)
#define msotmcFillDlgPatternColorHide1          (msotmcFillDlgMin+7)
#define msotmcFillDlgPatternColorHide2          (msotmcFillDlgMin+8)
#define msotmcFillDlgTransparent                (msotmcFillDlgMin+9)
#define msotmcColorsHatchedState                (msotmcFillDlgMin+10)
#define msotmcIndexHoneyComb                    (msotmcFillDlgMin+11)
#define msotmcFontFace                          (msotmcFillDlgMin+12)
#define msotmcTextEditDlg                       (msotmcFillDlgMin+13)
#define msotmcTextEditChange                    (msotmcFillDlgMin+14)
#define msotmcTextEditFontSizeHidden            (msotmcFillDlgMin+15) //1530
#define msotmcFillDlgPatternID                  (msotmcFillDlgMin+16)
#define msotmcFillDlgTextureID                  (msotmcFillDlgMin+17)
#define msotmcFillDlgTextureRealID              (msotmcFillDlgMin+18)
#define msotmcFillDlgShadeVariantIndex          (msotmcFillDlgMin+19)
#define msotmcLTxGallery                        (msotmcFillDlgMin+20)
#define msotmcFillDlgclrSave1                   (msotmcFillDlgMin+21)
#define msotmcFillDlgclrSave2                   (msotmcFillDlgMin+22)
#define msotmcFillDlgPictureFlags               (msotmcFillDlgMin+23)
#define msotmcFillDlgschsh1                     (msotmcFillDlgMin+24)
#define msotmcFillDlgschsh2                     (msotmcFillDlgMin+25) //1540
#define msotmcFillDlgschpat2                    (msotmcFillDlgMin+26)
#define msotmcFillDlgschpat1                    (msotmcFillDlgMin+27)
#define msotmcFillDlgOrigSab                    (msotmcFillDlgMin+28)
#define msotmcFillDlgTintBaseColor              (msotmcFillDlgMin+29)
#define msotmcFillDlgTintBaseColorHide          (msotmcFillDlgMin+30)
#define msotmcFillDlgTintBaseColorExtHide       (msotmcFillDlgMin+31)
#define msotmcFillDlgShadedColorExtHide1        (msotmcFillDlgMin+32)
#define msotmcFillDlgShadedColorExtHide2        (msotmcFillDlgMin+33)
#define msotmcFillDlgPatternColorExtHide1       (msotmcFillDlgMin+34)
#define msotmcFillDlgPatternColorExtHide2       (msotmcFillDlgMin+35) //1550
#define msotmcFillDlgschTintBase                (msotmcFillDlgMin+36)
#define msotmcFillDlgTintColorExtMod            (msotmcFillDlgMin+37)
#define msotmcFillDlgHiddenMax                  (msotmcFillDlgMin+38) //1553


#define msotmcColorDlgMin                       (msotmcFillDlgHiddenMax) //1553
#define msotmcColorsHue                         (msotmcColorDlgMin+0)
#define msotmcColorsSat                         (msotmcColorDlgMin+1)
#define msotmcColorsRed                         (msotmcColorDlgMin+2)
#define msotmcColorsGreen                       (msotmcColorDlgMin+3)
#define msotmcColorsBlue                        (msotmcColorDlgMin+4)
#define msotmcColorsLum                         (msotmcColorDlgMin+5)
#define msotmcColorsCyan                        (msotmcColorDlgMin+6)
#define msotmcColorsMagenta                     (msotmcColorDlgMin+7) //1560
#define msotmcColorsYellow                      (msotmcColorDlgMin+8)
#define msotmcColorsBlack                       (msotmcColorDlgMin+9)
#define msotmcColorsClrOrig                     (msotmcColorDlgMin+10)
#define msotmcColorsCmykOrig                    (msotmcColorDlgMin+11)
#define msotmcColorsColorModelPrev              (msotmcColorDlgMin+12)
#define msotmcColorsMax                         (msotmcColorDlgMin+13) //1566

// the following are the text items with the rendered line separator
// near them in the format object dialog.
#define msotmcSTMin                             (msotmcColorsMax) //1566
#define msotmcFillFillST                        (msotmcSTMin + 0)
#define msotmcFillLineST                        (msotmcSTMin + 1)
#define msotmcFillArrowsST                      (msotmcSTMin + 2)
#define msotmcSizeSizeRotST                     (msotmcSTMin + 3)
#define msotmcSizeScaleST                       (msotmcSTMin + 4) //1570
#define msotmcSizeOrigSizeST                    (msotmcSTMin + 5)
#define msotmcTboxMarginST                      (msotmcSTMin + 6)
#define msotmcPicCropFromST                     (msotmcSTMin + 7)
#define msotmcPicImageCtrlST                    (msotmcSTMin + 8)
#define msotmcSTMax                             (msotmcSTMin + 9) //1575


// tmc for Description Tab
#define msotmcAltTextMin                        (msotmcSTMax)
#define msotmcAltTextEdit                       (msotmcAltTextMin + 0)
#define msotmcAltTextTitle                      (msotmcAltTextMin + 1)
#define msotmcAltTextNotes                      (msotmcAltTextMin + 2)
#define msotmcAltTextChangeHidden               (msotmcAltTextMin + 3)
#define msotmcAltTextMax                        (msotmcAltTextMin + 4)


// TODO seankang fix the definition name

// Picture Container Specifiers for the GCC's in the format object dialog
#define msopcsMin 0x00008000 //32768
#define msopcsFillLineColor                     (msopcsMin + 0)
#define msopcsFillLineStyle                     (msopcsMin + 1)
#define msopcsFillLineWidth                     (msopcsMin + 2)
#define msopcsFillLineDash                      (msopcsMin + 3)
#define msopcsFillLineConnect                   (msopcsMin + 4)
#define msopcsFillArrowStyleL                   (msopcsMin + 5)
#define msopcsFillArrowStyleR                   (msopcsMin + 6)
#define msopcsFillArrowSizeL                    (msopcsMin + 7)
#define msopcsFillArrowSizeR                    (msopcsMin + 8)
#define msopcsFillFill                          (msopcsMin + 9)
#define msopcsFillDlgShadedColor1               (msopcsMin + 10)
#define msopcsFillDlgShadedColor2               (msopcsMin + 11)
#define msopcsFillDlgPatternForeColor           (msopcsMin + 12)
#define msopcsFillDlgPatternBackColor           (msopcsMin + 13)
#define msopcsFillHrColor                       (msopcsMin + 14)
#define msopcsFillDlgTintBaseColor              (msopcsMin + 15)


/* the following structs are used to pass information between the host's
        format object dlg and the office code that initializes and applies the
        tabs.
        TT is hungarian for Tab Template */
// Fill and Line tab

//
// MSOTTFLLID - line id in FillNLine tab template (is used to index MSOTTFL::lineProps[] array)
//
typedef enum
        {
        msottfllidLine,
        msottfllidLeftLine,
        msottfllidRightLine,
        msottfllidTopLine,
        msottfllidBottomLine,
        msottfllidColumnLine,
        msottfllidMax,
        } MSOTTFLLID;
//
// Fill and Line tab: line properties
//
#define MSOTTFLLINEPROPS                \
struct                                  \
        {                               \
        UINT crLineBackColor;           \
        LONG iLinePattern;              \
        UINT uLineTransparency;         \
        UINT crLineColor;               \
        UINT preLineStyle;              \
        UINT xyaLineWidth;              \
        UINT preLineDash;               \
        UINT iLineConnect;              \
        UINT preArrowStyleL;            \
        UINT preArrowStyleR;            \
        UINT preArrowSizeL;             \
        UINT preArrowSizeR;             \
        MSOCOLOREXT *pclrExtLine;       \
        MSOCOLOREXT *pclrExtLineBack;   \
        BOOL fInsetPen;                 \
        BOOL fInsetPenOK;               \
        }

//
// Fill and Line tab: fill properties
//
#define MSOTTFLFILLPROPS                        \
struct                                          \
        {                                       \
        MSOCOLORDLGSTYLE ssFillShadeStyle;      \
        MSOVARIANT varFillVariant;              \
        MSOCOLORDLGCTL ccFillColor;             \
        BOOL fFillTransparent;                  \
        UINT lFillBrightness;                   \
        MSOPRESET preFillPreset;                \
        BOOL fFillBlip;                         \
        MSOFILLTYPE iFillType;                  \
        COLORREF crFillColor;                   \
        COLORREF crFillBackColor;               \
        UINT uFillTransparency;                 \
        UINT uFillTransparencyPer;              \
        MSOCOLOREXT *pclrExtFill;               \
        MSOCOLOREXT *pclrExtFillBack;           \
        UINT fDefaultNew;                       \
        }

typedef struct _MSOTTFL
        {
        //
        // Fill properties
        //
        MSOTTFLFILLPROPS;

        //
        // Line properties
        //
        union
                {
                MSOTTFLLINEPROPS;
                MSOTTFLLINEPROPS linesProps[msottfllidMax];
                };

        } MSOTTFL;

// Size tab
typedef struct _MSOTTSZ
        {
        UINT yHeight;
        UINT xWidth;
        UINT uHeightPer;
        UINT uWidthPer;
        UINT uRotation;
        UINT fRelative;
        UINT fLockAspect;
        UINT fBestScale; // powerpoint only
        UINT iResolution;       // powerpoint only
        } MSOTTSZ;

// Picture tab
typedef struct _MSOTTPC
        {
        UINT xCropLeft;
        UINT xCropRight;
        UINT yCropTop;
        UINT yCropBottom;
        UINT iColor;
        UINT uBrightness;
        UINT uContrast;
        UINT uBrightnessPer;
        UINT uContrastPer;
        } MSOTTPC;

// Position tab
typedef struct _MSOTTPS
        {
        UINT xHorPos;
        UINT yVerPos;
        UINT iHorFrom;
        UINT iVerFrom;
        UINT fMovWText;
        UINT fLockAnchor;
        UINT fLockPos;
        } MSOTTPS;

// Textbox tab
typedef struct _MSOTTTB
        {
        BOOL fAutoMargin;               // for excel only
        UINT xMarginLeft;
        UINT xMarginRight;
        UINT yMarginTop;
        UINT yMarginBottom;
        UINT iVertAlign;                // for powerpoint only
        UINT fWrapText;                 // for powerpoint only
        UINT fResizeObjFitText;         // for powerpoint only
        UINT fRotateTextWithObject;     // for powerpoint only
        BOOL fVerticalText;             // for powerpoint only
        } MSOTTTB;

// Web tab
typedef struct _MSOTTWB
        {
        WCHAR  *pwzAltText;
        BOOL    fAltTextChanged;  // no real use. Reauired to match up Excel's rgoper
        BOOL    fNoDefault;       // If set, do not return default alt text, used by PPT
        } MSOTTWB;

// the whole format object dialog, like at tab template for all
typedef struct _MSOTTFO
        {
        MSOTTFL *pttfl;
        MSOTTSZ *pttsz;
        MSOTTPC *pttpc;
        MSOTTPS *pttps;
        MSOTTTB *ptttb;
        MSOTTWB *pttwb;
        } MSOTTFO;

/* Process sdm dialog messages sent to office owned items in the host's
        format object dialog. If the message is absorbed, store the value that
        the host's dlg proc should return in pfResult and return TRUE,
        else return FALSE. */

MSOAPI_(BOOL) MsoFProcessDlgMsg(BOOL *pfResult, unsigned int dlm, unsigned int tmc,
        unsigned int wNew, unsigned int wOld, unsigned int wParam,
        IMsoDrawingSelection *pidgsl, IMsoDrawingUserInterface *pidui);

/* Process sdm item messages */
MSOAPI_(BOOL) MsoFProcessItemMsg(unsigned int dlm, unsigned int tmc,
        unsigned int wNew, unsigned int wOld, unsigned int wParam,
        IMsoDrawingSelection *pidgsl, IMsoDrawingUserInterface *pidui);

// Exported sdm callback functions
// TODO ilanb: use SDM types!
MSOAPI_(unsigned int) MsoWParseProc (unsigned int tmm, WCHAR *wz, DWORD cchMax, VOID **ppv,
        unsigned long bArg, unsigned int tmc, unsigned int wParam);

MSOAPI_(UINT_PTR)MsoWRenderLine(unsigned int tmm, void *prds, UINT_PTR ftmsNew,
        UINT_PTR ftmsOld, UINT_PTR tmc, UINT_PTR wParam);

MSOAPI_(unsigned int)MsoWListStt(unsigned int tmm, WCHAR *wz, DWORD cchMax, unsigned int isz,
        unsigned int filler, unsigned int tmc, unsigned int stt, IMsoDrawingSelection *pidgsl);

// this definition is also in helpid.h, but it needs to be exported here so that
// the apps can put it in their .des files for the format object dialog.
#define msohidDlgDrawFormatObj 3012

#undef  INTERFACE
#define INTERFACE IMsoTextEditDlg

DECLARE_INTERFACE_(IMsoTextEditDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD


        /*      Callers of this can only expect the pch to be fill and only
                selected fields of the LOGFONT, specifically the height,
                weight, lfItalics, lfPitchAndFamily, and lfFaceName     */
        MSOMETHOD_(BOOL, FGet) (THIS_ _Out_cap_(cb) WCHAR *wz, __no_count LOGFONTW *plf, int cb,
                                                                        ULONG *pflChange) PURE;
        /* The dialog only uses the height, weight, italics and facename
                fields of the logfont.  The pch contents are copied and stored
                in an internally allocated string buffer in the dialog.  */
        MSOMETHOD_(void, Set) (THIS_ const WCHAR *wz, const LOGFONTW *plf) PURE;

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        /* This is only useable when the dialog is run modelessly.
           This is used to move the Office assistant out of the way of the
                dialog rectangle.  */
        MSOMETHOD_(void, AvoidDialog) (THIS_ interface IMsoDrawingUserInterface *pidgui) PURE;

        /* The dialog can be expected to run either modally or modelessly,
                depending on how this dialog was created via the MsoFCreateTextEditDlg
                api.  */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;
        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;

        /* This sets this dialog as the current dialog */
        MSOMETHOD_(BOOL, FSetCurrent) (THIS) PURE;
        };

#undef  INTERFACE
#define INTERFACE IMsoInsertTextEffectDlg

DECLARE_INTERFACE_(IMsoInsertTextEffectDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        /* This is only useable when the dialog is run modelessly.
           This is used to move the Office assistant out of the way of the
                dialog rectangle.  */
        MSOMETHOD_(void, AvoidDialog) (THIS_ interface IMsoDrawingUserInterface *pidgui) PURE;

        MSOMETHOD_(void, Get) (THIS_ MSOSPT *pspt, MSOPRE *ppre) PURE;
        MSOMETHOD_(void, GetIndex) (THIS_ int *pindex) PURE;
        MSOMETHOD_(void, Set) (THIS_ int index) PURE;

        /* This runs the dialog either modal or modeless depending on how
           the dialog was created. */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;
        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;

        /* This sets this dialog as the current dialog */
        MSOMETHOD_(BOOL, FSetCurrent) (THIS) PURE;
        };


// change flags for the Text Effect Edit Text dialog.
#define msotdTextChange         0x00000001
#define msotdFaceChange         0x00000002
#define msotdSizeChange         0x00000004
#define msotdBoldChange         0x00000008
#define msotdItalicChange       0x00000010



/*----------------------------------------------------------------------------
        Called by the client to create the Text Effect "Enter Your Text Here"
        dialog.
-----------------------------------------------------------------*/

MSOAPI_(BOOL) MsoFCreateTextEditDlg(HWND hwnd, ULONG flFlags,
                                                                                                IMsoDrawingUserInterface *pidgui,
                                                                                                MSOPFNPREVIEW pfnPreview,
                                                                                                IMsoTextEditDlg **ppted);


/*---
   Creates the Insert Text Effect Dialog.
  -*/

MSOAPI_(int) MsoFCreateInsertTextEffectDlg(HWND hwnd, ULONG flFlags,
                                                                                                                 IMsoDrawingUserInterface *pidgui,
                                                                                                                 MSOPFNPREVIEW pfnPreview,
                                                                                                                 IMsoInsertTextEffectDlg **ppited);

#define MSODLGWORDARTFORMAT       0x00000001

MSOAPI_(UINT) MsoWParseSCProc(UINT iTmm, WCHAR *wz, UINT cch, VOID **ppv,
        ULONG ulArg, UINT ulTmc, UINT iParam);

MSOAPI_(BOOL) MsoFSCItemProc(UINT dlm, UINT tmc, UINT wNew,
        UINT wOld, UINT wParam);


#define MsoSCSTYLE_NORMAL   0x0000
#define MsoSCSTYLE_REVERSE  0x0001
#define MsoSCSTYLE_AUTO     0x0002    // 0 means Auto
#define MsoSCSTYLE_NOLIM    0x0004    // 0 means No Limit
#define MsoSCSTYLE_UNITS    0x0008    // 0 means no units


/* Units */
// If you change the order of the following units, you should change
// the order of the Units in ostrman.pp too.

#define msoutFirst         -2
#define msoutNone          -2
#define msoutUndefined     -1
#define msoutInch          0 /* used for '"' string */
#define msoutPt            1
#define msoutCm            2
#define msoutLine          3
#define msoutInch2         4 /* used for "in" string */
#define msoutPercent       5
#define msoutDegree        6
#define msoutPica          7
#define msoutMm            8
#define msoutInch3         9 /* used for the double byte inch character */
#define msoutPixel         10 /* used for "px" string */
#define msoutPixel2        11 /* used for "pixel" string */
#define msoutPixel3        12 /* used for "pixels" string */

#define msoutPtEnu         13
#define msoutCmEnu         14
#define msoutLineEnu       15
#define msoutInch2Enu      16 /* used for "in" string */
#define msoutPicaEnu       17
#define msoutMmEnu         18
#define msoutPixelEnu      19 /* used for "px" string */
#define msoutPixel2Enu     20 /* used for "pixel" string */
#define msoutPixel3Enu     21 /* used for "pixels" string */
#define msoutMax           22
#define msoutLast          21

enum
        {
        msoscValid,
        msoscError,
        msoscBlank,
        };


#undef  INTERFACE
#define INTERFACE IMsoSpinControl

DECLARE_INTERFACE_(IMsoSpinControl, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        MSOMETHOD_(void, SetUnit) (THIS_ int unit) PURE;
        MSOMETHOD_(void, SetDecimalPlaces) (THIS_ int cDecimal) PURE;
        MSOMETHOD_(void, SetIncrement) (THIS_ int inc) PURE;
        MSOMETHOD_(void, SetMinMax) (THIS_ int min, int max) PURE;
        MSOMETHOD_(void, GetMinMax) (THIS_ int *pmin, int *pmax) PURE;
        MSOMETHOD_(BOOL, HasUnit) (THIS) PURE;

        MSOMETHOD_(void, SetState) (THIS_ int state) PURE;
        MSOMETHOD_(int, GetState) (THIS) PURE;
        MSOMETHOD_(LONG, EMUGetConversion) (THIS) PURE;
        MSOMETHOD_(VOID, EMUSetConversion) (THIS_ LONG emuper) PURE;
        MSOMETHOD_(void, SetDecimalSymbol) (THIS_ WCHAR wchDecimal) PURE;
        MSOMETHOD_(BOOL, FReportError) (THIS) PURE;
        MSOMETHOD_(void, SetReportError) (THIS_ BOOL fReportError) PURE;

        /*These two following interfaces are used to set the spinner to use a
          given value when the spinner is in a blank or ninch state and the
          user has clicked on the up or down arrow.  */
        /* If we return True, then *plVal will have the value to use*/
        MSOMETHOD_(BOOL, FUseValWhenBlank) (THIS_ LONG *plVal) PURE;
        MSOMETHOD_(void, SetUseValWhenBlank) (THIS_ LONG lVal) PURE;
        MSOMETHOD_(LONG, EMUGetPxConversion) (THIS) PURE;
        MSOMETHOD_(VOID, EMUSetPxConversion) (THIS_ LONG emuper) PURE;
        };



typedef struct
   {
        MSOTTFL ttfl;
        MSOTTSZ ttsz;
        MSOTTPC ttpc;
        MSOTTPS ttps;
        MSOTTTB tttb;
        MSOTTWB ttwb;
        MSOTTFL ttflMask;
        MSOTTSZ ttszMask;
        MSOTTPC ttpcMask;
        MSOTTPS ttpsMask;
        MSOTTTB tttbMask;
        MSOTTWB ttwbMask;
  }
   MSOTTOBJPARAM;


/* PPT only. Needed for 'best scale for slideshow' option */
#define  SLIDESHOWBORDER  0

#undef  INTERFACE
#define INTERFACE IMsoFormatObjectDlg

DECLARE_INTERFACE_(IMsoFormatObjectDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        MSOMETHOD_(void, Set) (THIS_ MSOTTOBJPARAM *pttobj) PURE;
        MSOMETHOD_(void, Get) (THIS_ MSOTTOBJPARAM *pttobj) PURE;
        MSOMETHOD_(BOOL, FSetup) (THIS_ IMsoDrawingSelection *pidgsl,
                IMsoDrawingUserInterface *pidgui) PURE;

        /* This is only useable when the dialog is run modelessly.
           This is used to move the Office assistant out of the way of the
                dialog rectangle.  */
        MSOMETHOD_(void, AvoidDialog) (THIS_ interface IMsoDrawingUserInterface *pidgui) PURE;

        MSOMETHOD_(BOOL, FRun) (THIS) PURE;
        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;

        /* Ability to set the fill color.  Used by Powerpoint UI Modal dlg.
           The dgcid is the one used to bring up the color dialog.  It
      can be called from either the More Color button or from the
           more Line Color button from the dropdown.  */
        MSOMETHOD_(void, SetFillColor) (THIS_ MSODLGColorStyle *pcds,
                                                                                          MSODGCID dgcid) PURE;
        /* Ability to set the fill effects.  Used by Powerpoint UI Modal dlg.
      This function assumes that the MSOPSFillStyle was first ninched,
                and the the values were loaded using the Get function of the
                IMsoFormatObjectDlg.  */
        MSOMETHOD_(void, SetFillEffect) (THIS_ MSOPSFillStyle *popsFill) PURE;

        /* Ability to set the line fill pattern style into the format object dialog.
           Used by powerpoint UI modal dialog models.  */
        MSOMETHOD_(void, SetLineFillPattern) (THIS_ MSOPSLineStyle *popsLineStyle) PURE;

        /* This returns the zero based index of the tab which was used last
           in the dialog.  This can only be used after the dialog is closed with
                OK. If this is called anyother time, it returns the value passed
           to MsoFCreateFormatObjectDlg.  Best time to use this function is right
                after the call to FEndDlg (modelessly) or FRun (modally) */
        MSOMETHOD_(int, GetTabIndex) (THIS) PURE;

        /* This sets this dialog as the current dialog */
        MSOMETHOD_(BOOL, FSetCurrent) (THIS) PURE;

        /* Similar to SetFillColor except that it handles extended colors. The ink
                color string is allocated on the heap by the host app. Office takes the
                ownership of this string from the host app. */
        MSOMETHOD_(void, SetFillColorExt) (THIS_ MSODLGColorExtStyle *pdlgclrExt,
                                                                                          MSODGCID dgcid) PURE;
};

/*
 *  This creates the drawing format object dialog.  This is currently
 *  used only by Powerpoint.  However, this interface reuses
 *  common core format object code which are used by non-Powerpoint
 *  office clients.  The prcDoc is the current doc size or in the
 *  case of Powerpoint, the current slide rect which is used for the
 *  Best Scale and Resolution controls.
 *
 *  The MSOPFNRECOLOR is the callback function for the caller to
 *  bring up their recolor dialog. If the caller passes NULL as
 *  the pfnRecolor, the recolor button is disabled.
 *
 */

typedef VOID (CALLBACK *MSOPFNRECOLOR)();
#define msopfnrecolorNil ((MSOPFNRECOLOR)0)

/*This function creates the format object dialog that is specialized
  for Powerpoint.  The flFlags are the common escher office flags and are
  currenly only supports the modal/modeless flags only now.  The
  pfnRecolor parameter is the function pointer used by powerpoint to
  receive the event that the Recolor button being pressed.  The iSabIndex
  is used to control which tab is initially shown at the top.  The value
  is zero through and including 4.  */

MSOAPI_(int) MsoFCreateFormatObjectDlg(HWND hwnd,
                                                                                                        HMSOINST hmsoinst,
                                                                                                        ULONG flFlags,
                                                                                                        RECT *prcDoc,
                                                                                                        MSOPFNRECOLOR pfnRecolor,
                                                                                                        IMsoDrawingSelection *pidgsl,
                                                                                                        UINT iTabIndex,
                                                                                                        MSOPFNPREVIEW pfnPreview,
                                                                                                        IMsoFormatObjectDlg **ppdlg);


#undef  INTERFACE
#define INTERFACE IMsoOptPictDlg

DECLARE_INTERFACE_(IMsoOptPictDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* The dialog can be expected to run either modally or modelessly,
                depending on how this dialog was created via the MsoFCreateInsDgmDlg
                api.  */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;

        MSOMETHOD_(void, SetSelection) (THIS_ BOOL fSelectedShapes) PURE;
        MSOMETHOD_(void, SetResampleType) (THIS_ int resampleType) PURE;
        MSOMETHOD_(void, SetPrintDPI) (THIS_ int printDPI) PURE;
        MSOMETHOD_(void, SetWebScreenDPI) (THIS_ int webscrDPI) PURE;
        MSOMETHOD_(void, SetCompress) (THIS_ BOOL fCompress) PURE;
        MSOMETHOD_(void, SetCrop) (THIS_ BOOL fCrop) PURE;
        MSOMETHOD_(int,  GetResampleType) (THIS) PURE;
        MSOMETHOD_(BOOL, FSelectedShapes) (THIS) PURE;
        MSOMETHOD_(BOOL, FCompress) (THIS) PURE;
        MSOMETHOD_(BOOL, FCrop) (THIS) PURE;
   };


MSOAPI_(BOOL) MsoFCreateOptPictDlg(HWND hwnd,
                                  ULONG flFlags,
                                                                                         IMsoDrawingUserInterface *pidgui,
                                                                                         BOOL fSaveAsMode,
                                                                                         IMsoOptPictDlg **ppidd);

/****************************************************************************

  OrgChart and Diagramming Dialogs

 ****************************************************************************/

#undef  INTERFACE
#define INTERFACE IMsoDgmStyleDlg

DECLARE_INTERFACE_(IMsoDgmStyleDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        /* Get the index selected */
        MSOMETHOD_(int, GetIndex) (THIS) PURE;

        /* The dialog can be expected to run either modally or modelessly,
                depending on how this dialog was created via the MsoFCreateOrgChartStyleDlg
                api.  */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;

        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;

        /* Get the states of the check boxes */
        MSOMETHOD_(UINT, GetCheckBoxStates) (THIS) PURE;

   };


MSOAPIX_(BOOL) MsoFCreateDiagramStyleDlg( HWND hwnd,
                                         ULONG flFlags,
                                         MSODGMT dgmt /*diagram type*/,
                                                                                                          IMsoDrawingSelection *pdgsl,
                                                                                                          IMsoDrawingUserInterface *pidgui,
                                                                                                          IMsoDrawingSite *pidgs,
                                                                                                          IMsoDrawingGroup *pdgg,
                                                                                                          IMsoDrawing *pdg,
                                                                                                     IMsoDgmStyleDlg **ppsd);


#undef  INTERFACE
#define INTERFACE IMsoInsDgmDlg

DECLARE_INTERFACE_(IMsoInsDgmDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* This is only useable when the dialog is run modelessly */
        MSOMETHOD_(HWND, GetWindow) (THIS) PURE;

        /* Get the index selected */
        MSOMETHOD_(int, GetIndex) (THIS) PURE;

        /* The dialog can be expected to run either modally or modelessly,
                depending on how this dialog was created via the MsoFCreateInsDgmDlg
                api.  */
        MSOMETHOD_(BOOL, FRun) (THIS) PURE;

        /* This is used when the dialog is to run modeleslly.  This return
           TRUE  if the dialog has ended and returns TRUE in the pfOK
                if the OK button was pressed. */
        MSOMETHOD_(BOOL, FEndDlg) (THIS_ BOOL *pfOK) PURE;
   };


MSOAPIX_(BOOL) MsoFCreateInsDgmDlg(HWND hwnd,
                                  ULONG flFlags,
                                                                                         IMsoDrawingUserInterface *pidgui,
                                                                                         IMsoInsDgmDlg **ppidd);


#undef  INTERFACE
#define INTERFACE IMsoDgmAlertDlg

DECLARE_INTERFACE_(IMsoDgmAlertDlg, ISimpleUnknown)
        {
        //--- ISimpleUnknown    methods
        MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
        MSOMETHOD_(void, Free) (THIS) PURE;

        //--- Debug
        MSODEBUGMETHOD

        /* Get the "don't show again" checkbox value */
        MSOMETHOD_(BOOL, FDoNotShow) (THIS) PURE;

        MSOMETHOD_(BOOL, FRun) (THIS) PURE;
   };


MSOAPI_(BOOL) MsoFCreateDgmAlertDlg(HWND hwnd,
        ULONG flFlags, int idsDlgCaption, int idsAlertText, int idsOk,
        int idsCancel, IMsoDrawingUserInterface *pidgui, IMsoDgmAlertDlg **ppidd);


/****************************************************************************

 ****************************************************************************/



/* This flag allows the format object dialog to know that the current
   selection(s) is a picture object.  */
#define MSOFMTOBJDLGPICTURE                     0x00000001

#endif //$[VSMSO]
/****************************************************************************
        Escher Dialog Help to bring up help.
        usCommand is the WinHelp first parameter ie. HELP_CONTEXTPOPUP.
        ulData is SDM specific.  It should be the ContextHidOfDlg.
        idsHelpFile is the office international defined string.

 ****************************************************************************/
MSOAPI_(BOOL) MsoFHelp(UINT_PTR tmc, USHORT usCommand, ULONG_PTR ulData);

#if 0 //$[VSMSO]
/*****************************************************************************
        File name utilities - these are used internally when trying to decide
        whether two picture file names are the same.  The client may, but need
        not, use these interfaces to expand a file name which is attached to
        a blip.  The advantage is that the file name is then absolute and won't
        change with the current directory, the disadvantage is that copying the
        document and the linked picture files somewhere else will not fix up the
        links.  If relative path names are used the client must ensure that the
        working directory corresponds to the that of the document (the client may
        find that too difficult).
******************************************************************* JohnBo **/
/* MsoCchGetFullPathName

        This interface takes a wide character string file name and converts it
        into the corresponding narrow full path name in the given buffer.  0 is
        returned on error (and the destination buffer will not have been filled),
        >cb is returned if the buffer is not big enough.

        WARNING: like GetFullPathName this API returns the number of characters
        in the path EXCLUDING THE TERMINATING NULL if (and only if) the thing fits
        in the buffer  like GetFullPathName if the string will not fit it returns
        the size of the buffer INCLUDING THE TERMINATING NULL. */
MSOAPIX_(ULONG) MsoCchGetFullPathName(const WCHAR *pwzFile, ULONG cch, _Out_z_cap_(cb)char *psz);

/* Stash a wide character file name, converting it into a full path name in
        single byte characters (as required by the Win95 OS interfaces).  The
        data group in which to stash the file must be supplied.  The interface can
        handle arbitrary length files - the result may be bigger than MAX_PATH.
        The memory is allocated using the counted memory manager (can be freed
        with delete[]). */
MSOMACAPI_(char *) MsoSzFileName(const WCHAR *wzFile, int dg);

/* MsoFSameFile

        Return true if two names refer to the same file - this returns FALSE on
        error (out of memory) as well as if the file names are not the same.  The
        interface cannot handle full path names longer than MAX_PATH - in this
        case it returns FALSE. */
MSOAPIX_(BOOL) MsoFSameFile(const WCHAR *wzFile1, const WCHAR *wzFile2);



//*****************
#define msocbMaxDlgTitle   128
#define msocbMaxItemTitle  25

/* Custom item types. Copied from dm96.h
        Currently only msodmotpCommand, msodmotpCheckbox and msodmotpRadioButton
        are allowed as a custom control in Insert Pict dialogs. */
typedef int MSOOTP;

#if 0
enum
        {
        msodmotpNil = 0,
        msodmotpEdit,                           // Edit control.
        msodmotpCommand,                        // Command (push) button.
        msodmotpDescrText,                      // Descriptive text.
        msodmotpCheckbox,                       // Check box.
        msodmotpRadioButton,                    // Radio button in single radio group.
        msodmotpGroupBox,                       // Group box.
        msodmotpStdButton,                      // Standard push button.
        msodmotpPushButton,                     // Normal push button.
        msodmotpListBox,                        // List box.
        msodmotpDropList,                       // Drop down.
        msodmotpComboBox,                       // Combo box.
        msodmotpConvertButton,                  // Convert button.
        msodmotpMax
        };
#endif

#ifndef OFFOPEN_CALLBACK
#define OFFOPEN_CALLBACK        __cdecl
#endif

// prototype of FindFile event hook.
typedef HRESULT (OFFOPEN_CALLBACK* PFNPICTDLGEVENT) (int,
        interface IMsoFindFile*, int, int, int, int*);

// Used to specify a custom control item to be put in the
// Insert/SaveAs/Browse Picture dialog
typedef struct _MSOPICTDLGCUSTCNTL
        {
        BOOL   fDisable;                    // disable if TRUE
        MSOOTP kind;                        // what kind of control
        int    value;                       // the value of the control
        WCHAR  cwzTitle[msocbMaxItemTitle]; // title of custom control
        }
MSOPICTDLGCUSTCNTL;


/***************************************************************************
        MSOPICTDLGPARAM

        Control structure for the insert picture dialog, allows customisation of
        the dialog.
*****************************************************************************/
typedef struct _MSOPICTDLGPARAM
        {
        union
                {
                struct
                        {
                        ULONG fAddLink: 1;
                        /* Add the "Link" item to the open button dropdown - otherwise
                                the only choice is "Embed". */
                        ULONG fAddLinkAndEmbed: 1;
                        /* Add the "Link and Embed" item to the open button dropdown.
                                This requires that fAddLink also be specified, if not then
                                the code really will add only the link-and-embed option. */

                        ULONG fRetEmbed: 1;     // Embed the picture in the document
                        ULONG fRetLink: 1;      // Store a link to the picture in the document
                        /* NOTE: !fRetLink implies fRetEmbed is required! */

                        ULONG fNoList: 1;       // Disable the list dialog if import fails
                        ULONG fSetDirectory: 1; // Set initial directory to wzDirName

                        ULONG fNoLoad: 1;       // Don't load the pib unless required

                        ULONG fOnlyFileSys: 1;
                        ULONG fMultiSelect: 1;

                        ULONG fAllowODMA: 1;

                        //NOTE: 10 bits used above.  Insert new ones here (shifts the
                        // check code quite deliberately.)

                        ULONG lInitVal: 13;     // Combined version and check code
                        };
                ULONG grf;
                };


        IMsoBlip               *pib;
                /* Blip of the file the user wants to insert. The host must not
                        forget to Free this blip if it is not used.  May be NULL if the
                        blip was not found. */

        PFNPICTDLGEVENT         pfnHandleEvent; // eventhook
        interface IMsoFindFile *piFindFile;     // pointer to IFindFIle
        WCHAR                   wzDlgTitle[msocbMaxDlgTitle];   // dialog title
        WCHAR                   wzDirName[INTERNET_MAX_URL_LENGTH];

        /* Multi-select for Insert Picture Dialog: The host must set the
           fMultiSelect flag above.  The above pib contains the "primary" blip that
           is selected.  Extra blips are placed in the plex ppxPibMulti, and their
           complementary filenames in ppxPwchFile.  The caller host is responsible
           for freeing the plexes via MsoFreePx().  In addition, one must call
           MsoFreePv() on all the allocated entries in ppxPwchFile.
           There used to be a message here saying new stuff shouldn't be added, to
           avoid exceeding a 4kB page size, but since this struct is already bigger
           than that (probably due to a INTERNET_MAX_URL_LENGTH change), I guess we
           can forget it.
           WARNING: MsoFreePx() must be used on ppxPibMulti and ppxPwchFile, even
           for C++ clients!  Same goes for MsoFreePv().
        */
#ifdef __cplusplus
        MSOTPX<IMsoBlip *> *ppxPibMulti;
        MSOTPX<WCHAR *> *ppxPwchFile;
#else
        MSOPX *ppxPibMulti;
        MSOPX *ppxPwchFile;
#endif
        }
MSOPICTDLGPARAM;

#define cbMSOPICTDLGPARAM (sizeof(MSOPICTDLGPARAM))
#define MAXFILEEXTLEN 8

/* You must call MSOInitPictDlgParam on any new MSOPICTDLGPARAM you're filling
        out before passing it into MsoFXXXXPictureDlg. */
MSOAPI_(void) MsoInitPictDlgParam(MSOPICTDLGPARAM* pdlgparam);

/* A limit on the number of characters in a description.  Descriptions should
        be short (short, concise and to the point), so this limit is not great. */
#define MsoCchMaxDescription 256

/* API for bringing up Insert Picture dialog. The selected file name is in
        rgwch, rgwchDescription is filled in with the description, if any. picmd
        points to the command value returned from the dialog (refer to dm96.h,
        FF_icmdXXX), could be NULL and will be ignored. Returns TRUE if no error,
        otherwise returns FALSE. */
MSOAPI_(BOOL) MsoFInsertPictureDlg(HMSOINST hinst,HWND hwnd,
        MSOPICTDLGPARAM *pdlgpm, WCHAR *rgwch, DWORD cchMax,
        WCHAR * rgwchDesc, DWORD cchMaxDesc, int *picmd);

//  Some math routines that we export
MSOAPI_(double) MsoLog(double);
MSOAPI_(double) MsoPow(double, double);
MSOAPI_(double) MsoExp(double);

/****************************************************************************
    PhotoDraw editing APIs
****************************************************************************/
MSOAPI_(void) MsoSetAppNamePhdGraphics(const WCHAR* szAppName);
MSOAPI_(BOOL) MsoFStartPhotoDrawInFrontPageMode(const WCHAR* wzFilename,
                                                const WCHAR* wzURL);

#if 0 //$[VSMSO]
/****************************************************************************
        Online Meeting functions.
****************************************************************************/
/*This function starts a conference using NetMeeting application conf.exe as a
  server, and shares out the application whose top window is hwndApp.
  Returns TRUE if successful.*/
MSOAPI_(BOOL) MsoFStartConference(HWND hwndApp, HMSOINST m_hmsoinst, BOOL fromOutlook);

/* Sends file indicated by filename to all participants in a conference.
   filename is a fully qualified path name of the file to be sent. */
#ifndef _RTC_TOOLBAR_
MSOAPI_(BOOL) MsoFSendFileW(const WCHAR* pwzDocPath);
#else
MSOAPI_(BOOL) MsoFSendFileW(const WCHAR* pwzWho, const WCHAR* pwzDocPath);
#endif // _RTC_TOOLBAR_

/* Returns TRUE if application sharing is active.  To determine if you should 
   allow the user to send a file, use MsoFIsFileTransferEnabled instead. */
MSOAPI_(BOOL) MsoFIsAppSharing();

/* Returns TRUE if the file transfer feature is enabled, which is an indication that
   MsoFSendFile should work.*/
MSOAPI_(BOOL) MsoFIsFileTransferEnabled();

/* Returns TRUE if a conference is currently active. This is used to determine
   state of related UI controls. */
MSOAPI_(BOOL) MsoFIsConferencing();

/* End the current conference. Returns TRUE if ended successfully. FALSE means
   user canceled the action so the conference continues to be active. */
MSOAPI_(BOOL) MsoFEndConference();

/* Call Outlook to schedule an Online Meeting. Returns TRUE if success. FALSE if fail. */
MSOAPI_(BOOL) MsoFScheduleMeeting(HWND hwndApp, const WCHAR* pwzDocPath);
#endif //$[VSMSO]

/* For Word9, bug 72950, focus always shifts to first doc window . Need to constantly change top window handle
   as user shifts focus to different doc window */
MSOAPIX_(void) MsoSetHwnd(HWND newTopHwnd);

#if 0 //$[VSMSO]
/* Pop the Place-a-Call dialog. Returns NOERROR if successful, error from NM's CallDialog if fail. */
MSOAPI_(HRESULT) MsoHrPlaceCall(HWND parentHwnd);

/* Data Services (ODSO)
 * From mmselect.des: funky things are done with these values, so they
 * must stay in the same order.  Talk to EBailey to get the full story.
 */
#define msotmcQueryMin          (msotmcMin+0)
#define msotmcSelectField1      (msotmcQueryMin+0)
#define msotmcComp1             (msotmcQueryMin+1)
#define msotmcCompTo1           (msotmcQueryMin+2)
#define msotmcCond1             (msotmcQueryMin+3)
#define msotmcSelectField2      (msotmcQueryMin+4)
#define msotmcComp2             (msotmcQueryMin+5)
#define msotmcCompTo2           (msotmcQueryMin+6)
#define msotmcCond2             (msotmcQueryMin+7)
#define msotmcSelectField3      (msotmcQueryMin+8)
#define msotmcComp3             (msotmcQueryMin+9)
#define msotmcCompTo3           (msotmcQueryMin+10)
#define msotmcCond3             (msotmcQueryMin+11)
#define msotmcSelectField4      (msotmcQueryMin+12)
#define msotmcComp4             (msotmcQueryMin+13)
#define msotmcCompTo4           (msotmcQueryMin+14)
#define msotmcCond4             (msotmcQueryMin+15)
#define msotmcSelectField5      (msotmcQueryMin+16)
#define msotmcComp5             (msotmcQueryMin+17)
#define msotmcCompTo5           (msotmcQueryMin+18)
#define msotmcCond5             (msotmcQueryMin+19)
#define msotmcSelectField6      (msotmcQueryMin+20)
#define msotmcComp6             (msotmcQueryMin+21)
#define msotmcCompTo6           (msotmcQueryMin+22)
#endif //$[VSMSO]

/* Used in WRenderTextInBar / WDrawTextInBar */
#define msoftibNil              0x00
#define msoftibThin             0x01
#define msoftibAbsLeft          0x02
#define msoftibNoText           0x04
#define msoftibGray             0x08
#define msoftibOnTabSheet       0x10

// Defines for Query Options Dialog
#define wListSelectFieldsU      0
#define wListSelectFieldsS      1

/*-----------------------------------------------------------------------------
        MsoAcquireImages
        Call to Twain and WIA support to collect images

        wzPath - path to images, allocated by callee and deleted by caller
        nImages - number of images returned
        wzDeviceName - if non-NULL, the name of the device to use for images
        bWIA - If TRUE, the device specified by wzDeviceName is a WIA device
        phinst - pointer to App instance
        hWndApp - App window
        hWndStatus - Status Window
        pMSOParams->bSupressClipArt - if TRUE, do not show the "Add to Clip Art" check box.
        pMSOParams->nPrintDPI - if nonzero, the DPI value to use for Print

        MsoCheckSTIEvent
        Call to see if it was launched due to an STI event

        wzDeviceName - If return value is TRUE, the name of the device which caused the event
        bWIA - If TRUE, the device is a WIA device
        phinst - pointer to App instance
        lpxszCmdLine - pointer to command line
-------------------------------------------------------------------- v-brianf -*/
typedef struct _MSOTWPARAMS   /*Typedef for parameter passing */
        {
        BOOL bSupressClipArt;
        INT nPrintDPI;
        } MSOTWPARAMS;

MSOAPI_(int) MsoAcquireImages(WCHAR** wzPath, unsigned int *nImages, const WCHAR* wzDeviceName, BOOL bWIA, HMSOINST phinst, HWND hWndApp, HWND hWndStatus, MSOTWPARAMS* pMSOParams);
MSOAPI_(BOOL) MsoCheckSTIEvent(WCHAR* wzDeviceName, int cchMax, BOOL* bWIA, HMSOINST phinst, const WCHAR* lpxszCmdLine);
MSOAPI_(BOOL) MsoFRemoveHspFromShapeKeys(IMsoDrawing *pidg, MSOHSP hsp);

#endif //$[VSMSO]
#endif // MSODR_H
