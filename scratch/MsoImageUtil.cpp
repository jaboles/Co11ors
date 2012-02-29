/*----------------------------------------------------------------------------
%%File: MsoImageUtil.cpp
%%Unit:
%%Contact: rmolden

The various MSO image loading/retrival functions that are relied on by
VS to deal with MSO images that packages may be using.

----------------------------------------------------------------------------*/

#include "offpch.h"

//the following includes will suck in the generated headers which define such things as vrgBStripInfo
#if VSMSO_DEVDIV_BUILD	//$[VSMSO]
#include "tbbstrip.h"	// Generated include file
#include "tbbstrip.i"	// Generated include file
#else //$[VSMSO]
#include "btndrp/tbbstrip.h"
#include "btndrp/tbbstrip.i"
#endif //$[VSMSO]

#include "msotbfor.h"
#include "lzbmp.h"
#include "msocrit.hpp"

const int vrgtbcentry[] =
{
#include "tbbtnent.h" // Generated list of entries
};

inline int BStripEntry(int id, int bf)
{
    if (bf < 0 || bf >= cIconFormat)
    {
        return -1;
    }

    if (vrgBstripInfo[bf].rgtbcentry == NULL)
    {
        return vrgtbcentry[id] % vrgBstripInfo[bf].cIcons;
    }

    if (vrgBstripInfo[bf].rgtbcentry[vrgtbcentry[id]] < 0)
    {
        return -1;

    }

    return vrgBstripInfo[bf].rgtbcentry[vrgtbcentry[id]] % vrgBstripInfo[bf].cIcons;
}

inline int BStripStrip(int id, int bf)
{
    if (id < 0 || id >= centryMax)
    {
        return -1;
    }

    if (bf < 0 || bf >= cIconFormat)
    {
        return -1;
    }

    if (vrgBstripInfo[bf].rgtbcentry == NULL)
    {
        return vrgtbcentry[id] / vrgBstripInfo[bf].cIcons;
    }

    if (vrgBstripInfo[bf].rgtbcentry[vrgtbcentry[id]] < 0)
    {
        return -1;
    }

    return vrgBstripInfo[bf].rgtbcentry[vrgtbcentry[id]] / vrgBstripInfo[bf].cIcons;
}
const int centryMax	= sizeof(vrgtbcentry) / sizeof(int);

/* Return if the bitmap pointed to by pbi is a PM bitmap or not */
__inline BOOL FIsPMDIB(const BITMAPINFO* pbi)
{
#if DEBUG
    if (pbi->bmiHeader.biSize == sizeof(BITMAPCOREHEADER))
    {
        MsoReportSz("PMDIB Detected.  This is unexpected.  Contact DMORTON.");
    }
#endif

    return pbi->bmiHeader.biSize == sizeof(BITMAPCOREHEADER);
}

/* Return the pbi BITMAPINFO parameter cast as a BITMAPCOREINFO pointer */
__inline BITMAPCOREINFO* PBCI(const BITMAPINFO* pbi)
{
    return (BITMAPCOREINFO*)pbi;
}

/* Return the pbi BITMAPINFO parameter cast as a BITMAPCOREHEADER pointer */
__inline BITMAPCOREHEADER* PBC(const BITMAPINFO* pbi)
{
    return (BITMAPCOREHEADER*)pbi;
}

/*---------------------------------------------------------------------------
CbGetDibSize

Given a pointer to a BitmapInfo in parameter pbi, this function will
return the size in bytes of the bitmap
------------------------------------------------------------------ JIMMUR -*/
#pragma OPT_PCODE_OFF
int CbGetDibSize(const BITMAPINFO * pbi)
{
    MsoAssertTag(NULL != pbi, ASSERTTAG('eqpp'));

    // FUTURE: Add support for  compressed Bitmaps
    //  Assert(pbi->bmiHeader.biCompression == BI_RGB);

    int cb = 0;

    // FUTURE: can we use PvFindDIBBits to do some of this work?
    if (FIsPMDIB(pbi))
    {                                // It is a PM DIB
        cb = PBC(pbi)->bcSize;

        if (PBC(pbi)->bcBitCount != 24)
        {
            cb += (1 << PBC(pbi)->bcBitCount) * sizeof(RGBTRIPLE);
        }

        // When adding the size of the bitmap ensure that it is WORD aligned
        cb += ((PBC(pbi)->bcWidth * PBC(pbi)->bcBitCount + 31) & ~31) / 8 * PBC(pbi)->bcHeight;
    }
    else
    {                             // It is a Windows DIB
        cb =  pbi->bmiHeader.biSize;
        if (pbi->bmiHeader.biClrUsed != 0)
        {
            cb += pbi->bmiHeader.biClrUsed * sizeof(RGBQUAD);
        }
        else if (pbi->bmiHeader.biBitCount < 24)
        {
            cb += (1 << (pbi->bmiHeader.biBitCount)) * sizeof(RGBQUAD);
        }

        // When adding the size of the bitmap ensure that it is WORD aligned
        cb += ((pbi->bmiHeader.biWidth * pbi->bmiHeader.biBitCount + 31) & ~31) / 8 * pbi->bmiHeader.biHeight;
    }

    return cb;
}
#pragma OPT_PCODE_ON

/*---------------------------------------------------------------------------
MsoFLoadBStrip

Given an instance handle (hinst), the resource id to load (idb), the
number of sub bitmaps in the bitmap (one based), and a set of flags
(olb), The size of an individual subbitmap (idx, idy), this funciton 
will load in a bitmap and return a handle to an Office bitmap object in 
the phbstrip out parameter.  Returns Success.
------------------------------------------------------------------ JIMMUR -*/
#pragma OPT_PCODE_OFF
BOOL MsoFLoadBStripEx(HINSTANCE hinst, int idb, int cbmp, int olb, int idx, int idy, HBSTRIP * phbstrip)
{
    MsoAssertTag(NULL != phbstrip, ASSERTTAG('eqqa'));

    void* pvSh = NULL;
    BITMAPINFO * pbi = NULL;
    BOOL fSuccess = FALSE;
    int cbDIB;
    BITMAPINFOHEADER bmih;
    HBSTRIP hbstrip = NULL;

    // We are allocating into local variable.
    // We should NOT set *phbstrip until we have valid BSTRIP
    if (!MsoFAllocMem(reinterpret_cast<void**>(&hbstrip), sizeof(BSTRIP), msodgMisc))
    {
        return FALSE;
    }

    /* find the resource and figure out how big it should be in memory. */
    MsoLoadBitmapResourceEx(hinst, idb, (BITMAPINFO *)&bmih, sizeof(BITMAPINFOHEADER));

    if (bmih.biSize == 0)
    {
        goto Fail;
    }

    cbDIB = CbGetDibSize((BITMAPINFO *)&bmih);

    /* unshared, normal DIB, just load it. */
    pbi = MsoLoadBitmapResource(hinst, idb);
    if (pbi == NULL)
    {
        goto Fail;
    }

FillStrip:
    if (MsoFFillInHBitmap(pbi, cbmp, olb, idx, idy, hbstrip))
    {
        fSuccess = TRUE;

        // set *phbstrip now that we have valid BSTRIP
        *phbstrip = hbstrip;
    }
    else
    {
Fail:
        MsoFreeMem(hbstrip, sizeof(BSTRIP));
    }

    return fSuccess;
}
#pragma OPT_PCODE_ON


#ifndef STATIC_LIB_DEF
/*---------------------------------------------------------------------------
MsoLoadBitmapResourceEx

Loads a bitmap resource from an the resource module handle hinst with
the resource ID if "lpName".  If the bitmap is stored as an LZBITMAP, the
bitmap is decompressed first and the uncompressed bitmap is returned,
otherwise the resource "as is" is returned.

If pbi is NULL, then space for the bitmap is allocated on the Office
heap, and must be freed by calling MsoFreeBitmapResource.

If pbi is not NULL, then cbbi is the size of the buffer, and the bitmap
is created in this buffer.  No space is allocated and MsoFreeBitmapResource
does not need to be called (and doing so is a bug).  If the buffer is
too small, then NULL is returned and only the BITMAPINFOHEADER is copied
into the buffer.

Returns NULL on failure.
*--------------------------------------------------------MICHMARC----------*/
MSOAPI_(BITMAPINFO *) MsoLoadBitmapResourceEx(HINSTANCE hinst, int idb,
                                              BITMAPINFO *pbi, DWORD cbbi)
{
    MsoAssertTag((!pbi) || (cbbi >= sizeof(BITMAPINFOHEADER)), ASSERTTAG('eqpx'));

    if (pbi)
    {
        MsoDebugFill(pbi, cbbi, msomfBuffer);
        pbi->bmiHeader.biSize = 0;
    }

    HRSRC hrsrc = FindResource(hinst, MAKEINTRESOURCE(idb), RT_LZBITMAP);
    if (hrsrc == NULL)
    {
        // Ack -- look for normal bitmap
        hrsrc = FindResource(hinst, MAKEINTRESOURCE(idb), RT_BITMAP);

        // Assert(hrsrc);
        if (hrsrc == NULL)
        {
            return NULL;
        }

        HGLOBAL hg = LoadResource(hinst, hrsrc);
        if (hg == NULL)
        {
            return NULL;
        }

        BITMAPINFO *pbiRes = reinterpret_cast<BITMAPINFO *>(LockResource(hg));

        MsoAssertTag(pbiRes, ASSERTTAG('xczna'));
        MsoAssertTag(pbiRes->bmiHeader.biCompression == BI_RGB, ASSERTTAG('xczn'));

        if (!pbiRes || pbiRes->bmiHeader.biCompression != BI_RGB)
        {
            return NULL;
        }

        int cbBitmap = CbGetDibSize(pbiRes);

        BITMAPINFO *pbiRet = NULL;
        if (pbi)
        {
            pbiRet = pbi;
            if (cbbi < cbBitmap || cbBitmap < 0)
            {
                MsoMemcpy(pbi, pbiRes, sizeof(BITMAPINFOHEADER));
                return NULL;
            }
        }
        else if (!MsoFAllocMem((void **)&pbiRet,cbBitmap, msodgMisc))
        {
            return NULL;
        }

        MsoMemcpy(pbiRet, pbiRes, cbBitmap);
        pbiRet->bmiHeader.biSizeImage = cbBitmap;

        return pbiRet;
    }

    HGLOBAL hg = LoadResource(hinst, hrsrc);
    if (hg == NULL)
    {
        return NULL;
    }

    LZBITMAPHEADER *plzb = reinterpret_cast<LZBITMAPHEADER *>(LockResource(hg));
    MsoAssertTag(plzb, ASSERTTAG('xcznb'));
    if (!plzb)
    {
        return NULL;
    }

    int cbBitmap = 0;
    int ccol = 0;
    WORD biBitCount = 0;
    if (plzb->type == 1)  // 256 color paletted image
    {
        ccol = *(BYTE *)plzb->data;
        if (ccol == 0)
        {
            ccol = 256;
        }

        cbBitmap = plzb->dy*(((plzb->dx)+3)&~3) + sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*ccol;
        biBitCount = 8;
    }
    else if (plzb->type == 2) // Monochrome image
    {
        cbBitmap = plzb->dy*((((plzb->dx+7)/8)+3)&~3) + sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*2;
        biBitCount = 1;
    }
    else if (plzb->type == 32 || plzb->type == 24)
    {
        ccol = 0;
        cbBitmap = plzb->dy*(((plzb->dx*plzb->type/8)+3)&~3) + sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*0;
        biBitCount = plzb->type;
    }
    else
    {
        MsoAssertTag(plzb->type == 0 || plzb->type == 3, ASSERTTAG('eqpy'));  // 16 color system palette image		

        cbBitmap = plzb->dy*((((plzb->dx+1)/2)+3)&~3) + sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*16;
        biBitCount = 4;
    }

    BITMAPINFO *pbiRet = NULL;

    if (pbi)
    {
        pbiRet = pbi;
    }
    else if (!MsoFAllocMem(reinterpret_cast<void **>(&pbiRet), cbBitmap, msodgMisc))
    {
        return NULL;
    }

    pbiRet->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    pbiRet->bmiHeader.biWidth = plzb->dx;
    pbiRet->bmiHeader.biHeight = plzb->dy;
    pbiRet->bmiHeader.biPlanes = 1;
    pbiRet->bmiHeader.biBitCount = biBitCount;
    pbiRet->bmiHeader.biCompression = BI_RGB;
    pbiRet->bmiHeader.biSizeImage = cbBitmap;
    pbiRet->bmiHeader.biXPelsPerMeter = 0;
    pbiRet->bmiHeader.biYPelsPerMeter = 0;
    pbiRet->bmiHeader.biClrUsed = ccol;
    pbiRet->bmiHeader.biClrImportant = 0;

    if (pbi && (cbbi < cbBitmap))
    {
        return NULL;
    }

    if (ccol)
    {
        MsoMemcpy(&pbiRet->bmiColors[0], plzb->data+1, sizeof(RGBQUAD)*ccol);

        const int sizeOfPalette = ccol*sizeof(RGBQUAD);   // don't know why this shuts up PREfast, but it does.  There's no issue:  ccol is at most 256....

        MsoUncompressLZW(plzb->data+1+sizeOfPalette, (BYTE *)pbiRet->bmiColors + sizeOfPalette, 
            SizeofResource(hinst, hrsrc) - sizeof(LZBITMAPHEADER), 
            cbBitmap - sizeof(BITMAPINFOHEADER) - sizeof(RGBQUAD)*ccol);
    }
    else if (biBitCount == 1)
    {
        MsoMemcpy(&pbiRet->bmiColors[0], rgrgbqBWPalette, sizeof(rgrgbqBWPalette));

        MsoUncompressLZW(plzb->data, ((BYTE *)pbiRet->bmiColors) + sizeof(rgrgbqBWPalette), 
            SizeofResource(hinst, hrsrc) - sizeof(LZBITMAPHEADER), 
            cbBitmap - sizeof(BITMAPINFOHEADER) - sizeof(RGBQUAD)*2);
    }
    else
    {
        if (plzb->type == 3)
        {
            MsoMemcpy(&pbiRet->bmiColors[0], plzb->data+1, sizeof(RGBQUAD)*16);

            MsoUncompressLZW(plzb->data+1+sizeof(RGBQUAD)*16, ((BYTE *)pbiRet->bmiColors) + sizeof(RGBQUAD)*16, 
                SizeofResource(hinst, hrsrc) - sizeof(LZBITMAPHEADER), 
                cbBitmap - sizeof(BITMAPINFOHEADER) - sizeof(RGBQUAD)*16);
        }
        else if (plzb->type == 32 || plzb->type == 24)
        {
            MsoUncompressLZW(plzb->data, ((BYTE *)pbiRet->bmiColors), 
                SizeofResource(hinst, hrsrc) - sizeof(LZBITMAPHEADER), 
                cbBitmap - sizeof(BITMAPINFOHEADER));
        }
        else
        {
            MsoMemcpy(&pbiRet->bmiColors[0], rgrgbqColorPalette, sizeof(rgrgbqColorPalette));

            MsoUncompressLZW(plzb->data, ((BYTE *)pbiRet->bmiColors) + sizeof(rgrgbqColorPalette), 
                SizeofResource(hinst, hrsrc) - sizeof(LZBITMAPHEADER), 
                cbBitmap - sizeof(BITMAPINFOHEADER) - sizeof(RGBQUAD)*16);
        }		
    }

    return pbiRet;
}
#endif // STATIC_LIB_DEF

/*---------------------------------------------------------------------------
MsoFFillInHBitmap(BITMAPINFO *pbi, int icbm, int olb, int idx, int idy, 
HBSTRIP  hbstrip)

Given a pointer to a BITMAPINFO in parameter pbi,  the count of
subbitmaps (icbm), the HBitmap flags (olb) int cbmp, int olb,
the size of an individual bitmap (idx, idy). this function will
fill in correct values in the suppiled hbstrip parameter
------------------------------------------------------------------ JIMMUR -*/

#pragma OPT_PCODE_OFF
MSOAPI_(BOOL) MsoFFillInHBitmap(BITMAPINFO *pbi, int cbmp, int olb, int idx, int idy, HBSTRIP  hbstrip)
{		
    hbstrip->olb = olb;
    hbstrip->cbmp = cbmp;

    hbstrip->psm = pbi;

    // Set up the fields in the BStrip struct

    if (FIsPMDIB(pbi))
    {
        hbstrip->dx = PBC(pbi)->bcWidth;
        hbstrip->dy = PBC(pbi)->bcHeight;

    }
    else
    {
        hbstrip->dx = pbi->bmiHeader.biWidth;
        hbstrip->dy = pbi->bmiHeader.biHeight;
    }

    if ( (0 == idx) || (0 == idy) )
    {
        if (hbstrip->olb & msoolbVertical)
        {
            hbstrip->dxSub = hbstrip->dx;
            hbstrip->dySub = hbstrip->dy/hbstrip->cbmp;
        }
        else
        {
            hbstrip->dxSub = hbstrip->dx/hbstrip->cbmp;
            hbstrip->dySub = hbstrip->dy;
        }
    }
    else
    {
        hbstrip->dxSub = idx;
        hbstrip->dySub = idy;
    }

    return TRUE;
}
#pragma OPT_PCODE_ON

/*---------------------------------------------------------------------------
PvFindDIBits

Given the pbi BitmapInfo pointer parameter this funciton returns a pointer
to the start of the DIB bit.
------------------------------------------------------------------ JIMMUR -*/
#pragma OPT_PCODE_OFF
const void* PvFindDIBits(const BITMAPINFO* pbi)
{
    const BYTE* pbBits;

    /* Calcluate the pointer to the Bits information - skip over header
    and color table entries */

    if (FIsPMDIB(pbi))
    {                    // This is a PM DIB
        pbBits = (BYTE*)&PBCI(pbi)->bmciColors[0];
        if (PBC(pbi)->bcBitCount != 24)
        {
            pbBits += (1 << PBC(pbi)->bcBitCount) * sizeof(RGBTRIPLE);
        }
    }
    else
    {                    // This is a Windows DIB
        pbBits = reinterpret_cast<const BYTE*>(&pbi->bmiColors[0]);

        if (pbi->bmiHeader.biClrUsed != 0)
        {
            pbBits += pbi->bmiHeader.biClrUsed * sizeof(RGBQUAD);
        }
        else
        {
            if (pbi->bmiHeader.biBitCount <= 8)
            {
                pbBits += (1 << (pbi->bmiHeader.biBitCount)) * sizeof(RGBQUAD);
            }
        }
    }

    return pbBits;
}
#pragma OPT_PCODE_ON

/*---------------------------------------------------------------------------
MsoGetSubBitmap

Extract a HBITMAP indexed idb from hbstrip in the phb out parameter.
Returns Success.
------------------------------------------------------------------ JIMMUR -*/
/* FUTURE: phb is NOT an HBITMAP pointer -- it's a BITMAPINFO pointer with
bits hanging off of it, like a DIB resource.  Should fix the declaration
some day */
MSOAPI_(BOOL) MsoFGetSubBitmap(HBSTRIP hbstrip, int idb, HBITMAP * phb)
{
    MsoAssertExpTag(MsoFValidBStrip(hbstrip), ASSERTTAG('eqqs'));
    if (!hbstrip)
    {
        return FALSE;
    }

    MsoAssertTag((idb >= 0) && (idb < hbstrip->cbmp) && (NULL != phb), ASSERTTAG('eqqt'));

    int cbDIB = 0;             // size of the new DIB being created
    BITMAPINFO * pbi = NULL;
    BYTE *pch = NULL;
    BITMAPINFO * pbiNew = NULL;
    int crgb = 0;
    BOOL fSuccess = FALSE;

    int dyHeight = hbstrip->dy;
    int dxWidth = hbstrip->dx;

    if (hbstrip->cbmp > 1)
    {
        dyHeight = hbstrip->dySub;
        dxWidth = hbstrip->dxSub;
    }

    pbi = (BITMAPINFO *)hbstrip->psm;

    if (FIsPMDIB(pbi))
    {
        BITMAPCOREHEADER bc;
        bc = *PBC(pbi);
        crgb = 1 << PBC(pbi)->bcBitCount;
        bc.bcHeight = (WORD)dyHeight;
        bc.bcWidth = (WORD)dxWidth;
        cbDIB = CbGetDibSize((BITMAPINFO*)&bc);
    }
    else
    {
        BITMAPINFOHEADER bih;
        bih = pbi->bmiHeader;
        crgb = (pbi->bmiHeader.biClrUsed != 0) ? pbi->bmiHeader.biClrUsed : (1 << pbi->bmiHeader.biBitCount);

        if (pbi->bmiHeader.biBitCount > 16)
        {
            crgb = 0;
        }

        bih.biHeight = dyHeight;
        bih.biWidth = dxWidth;
        cbDIB = CbGetDibSize((BITMAPINFO*)&bih);
    }

    *phb = reinterpret_cast<HBITMAP>(GlobalAlloc(GMEM_MOVEABLE, cbDIB));

    if (NULL == *phb)
    {
        goto Return;
    }

    pbiNew = reinterpret_cast<LPBITMAPINFO>(GlobalLock(*phb));

    MsoAssertTag(pbiNew, ASSERTTAG('eqqv'));

    MsoMemset(pbiNew, 0, sizeof(BITMAPINFOHEADER));

    pbiNew->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    pbiNew->bmiHeader.biWidth = dxWidth;
    pbiNew->bmiHeader.biHeight = dyHeight;
    pbiNew->bmiHeader.biPlanes = 1;

    // FUTURE: wrong for !msoolbShared
    pbiNew->bmiHeader.biBitCount = (FIsPMDIB(pbi)) ? PBC(pbi)->bcBitCount :  pbi->bmiHeader.biBitCount;
    pbiNew->bmiHeader.biCompression = BI_RGB;

    // FUTURE: fix for !msoolbShared
    if (FIsPMDIB(pbi))
    {
        RGBQUAD *prgbQ = pbiNew->bmiColors;
        const RGBTRIPLE *prgbT = (PBCI(pbi))->bmciColors;

        for (int irgb = 0; irgb < crgb; irgb++, prgbQ++, prgbT++)
        {
            prgbQ->rgbBlue = prgbT->rgbtBlue;
            prgbQ->rgbGreen = prgbT->rgbtGreen;
            prgbQ->rgbRed = prgbT->rgbtRed;
        }
    }
    else
    {
        MsoAssertSzTag(sizeof(RGBQUAD) * crgb < cbDIB, "Oops we're going to crash. Chances are we're trying to mess with a true color bitmap", ASSERTTAG('eqqw'));

        memcpy(&pbiNew->bmiColors, pbi->bmiColors, sizeof(RGBQUAD) * crgb);
    }

    pch = reinterpret_cast<BYTE *>(&(pbiNew->bmiColors[crgb]));

    if (hbstrip->olb & msoolbVertical)
    {
        int cbLength = 0;
        BYTE * pSrcBits = NULL;

        cbLength = (pbi->bmiHeader.biBitCount > CHAR_BIT) ? (pbi->bmiHeader.biBitCount / CHAR_BIT) * dyHeight : dyHeight / (CHAR_BIT / pbi->bmiHeader.biBitCount);

#if DEBUG
        // assert to catch 119200
        if (pbi->bmiHeader.biBitCount <= CHAR_BIT)
        {
            MsoAssertSzTag(cbLength*(CHAR_BIT / pbi->bmiHeader.biBitCount) == dyHeight, "I can't handle this height. The extracted bitmap will have some flaws.", ASSERTTAG('eqqx'));
        }
#endif

        pSrcBits =  (reinterpret_cast<BYTE*>(const_cast<void*>(PvFindDIBits(pbi)))) + (idb * dxWidth * cbLength);

        for (int i = 0; i < dxWidth; i++ )
        {
            memcpy(pch, pSrcBits,  cbLength);
            pch += cbLength;
            pSrcBits += cbLength;
        }
    }
    else
    {
        int cbWidth = 0;
        int cbBitOffset = 0;
        BYTE * pSrcBits = NULL;

        cbWidth = (pbi->bmiHeader.biBitCount > CHAR_BIT) ? (pbi->bmiHeader.biBitCount / CHAR_BIT) * dxWidth : dxWidth / (CHAR_BIT / pbi->bmiHeader.biBitCount);

        cbBitOffset = cbWidth * pbi->bmiHeader.biWidth / pbi->bmiHeader.biHeight;
        pSrcBits =  (reinterpret_cast<BYTE*>(const_cast<void*>(PvFindDIBits(pbi)))) + (idb * cbWidth);

        for (int i = 0; i < dyHeight; i++ )
        {
            memcpy(pch, pSrcBits,  cbWidth);
            pch += cbWidth;
            pSrcBits += cbBitOffset;
        }
    }

    fSuccess = TRUE;

Return:
    if (*phb != NULL)
    {
        GlobalUnlock(*phb);
    }

    return fSuccess;
}

//fwd decl
BOOL FFindBtnFaceEx(IMsoButtonUser *pibu, void **ppvPibu, int id, HBSTRIP *phbstrip, int *piBtn, BOOL *pfBlankIcon, int *msobf);

/*---------------------------------------------------------------------------
FFindBtnFace

Set the values of phbstrip and piBtn which point to the strip and offset
for the buttonface given by id.	pibu is used if we need to ask the
user about the strip. Optionally returns whether the icon is blank in
pfBlankIcon.
----------------------------------------------------------------- JEFFJO -*/
BOOL FFindBtnFace(IMsoButtonUser *pibu, void **ppvPibu, int id, HBSTRIP *phbstrip, int *piBtn, BOOL *pfBlankIcon)
{
    int msobf = (msobfLoResLoClr | msobfFailIfNotFound);

    return FFindBtnFaceEx(pibu, ppvPibu, id, phbstrip, piBtn, pfBlankIcon, &msobf);
}

/*---------------------------------------------------------------------------
FFindBtnFaceEx

Set the values of phbstrip and piBtn which point to the strip and offset
for the buttonface given by id.	pibu is used if we need to ask the
user about the strip. Optionally returns whether the icon is blank in
pfBlankIcon.

msobf is an IN/OUT paramater that specifies the desired format of the
buttonface, eg. high/low color and high/low size.  It will be set to the format
returned.

----------------------------------------------------------------- MCirello -*/
BOOL FFindBtnFaceEx(IMsoButtonUser *pibu, void **ppvPibu, int id, HBSTRIP *phbstrip, int *piBtn, BOOL *pfBlankIcon, int *msobf)
{
    if (msobf == NULL)
    {
        return FALSE;
    }

    const BOOL fFailIfNotFound = ((*msobf & msobfFailIfNotFound) != 0);
    *msobf &= ~msobfFailIfNotFound;

    MsoAssertTag(id >= 0, ASSERTTAG('ewig'));

#ifdef DEBUG
    if (!(id < centryMax || id >= msotcidUser))
    {
        char szMsg[256] = "";

        // Debug - only code
        OACR_REVIEWED_CALL(rolandr,wnsprintf(RgC(szMsg), "FFindBtnFace: tcid %d is out of range, using msotcidNil.\n", id));
        ReportMsgTag(FALSE, szMsg, ASSERTTAG('wuml'));
    }
#endif

    if (id < 0 || id >= centryMax)
    {
        if (id < msotcidUser)
        {
            id = msotcidNil;	// Fall thru with Nil (covers out of range.)
        }
        else
        {
            if (pibu == NULL)
            {
                return FALSE;
            }

            // Ask the app for the ID
            if (pibu->FFindCustomBtnface(NULL, ppvPibu, FALSE, id, phbstrip, piBtn, FALSE))
            {
                goto LExit;
            }

            id = msotcidNoIcon;	// Fall thru with blank
        }
    }

    const int iStrip = BStripStrip(id, *msobf);
    const int iEntry = BStripEntry(id, *msobf);

    if (iStrip == -1 || iEntry == -1)
    {
        if (fFailIfNotFound)
        {
            return FALSE;
        }

        switch(*msobf)
        {
        case msobfLoResHiClr:
            {
                *msobf = msobfLoResLoClr;
                break;
            }
        default:
            {
                return FFindBtnFace(pibu, ppvPibu, msotcidAwsPaneBullet, phbstrip, piBtn, pfBlankIcon);
            }
        }

        return FFindBtnFaceEx(pibu, ppvPibu, id, phbstrip, piBtn, pfBlankIcon, msobf);
    }

    MsoAssertTag(BStripStrip(id, *msobf) == iStrip && BStripEntry(id, *msobf) == iEntry, ASSERTTAG('kjrg'));
    MsoAssertTag((iStrip >= 0) && (iStrip < cBtnStrip), ASSERTTAG('ewih'));

    if ( (iStrip >= 0) && (iStrip < cBtnStrip)  &&  (*msobf >= 0) && (*msobf < _countof(vrgBstripInfo)) )
    {
        // Load resource if never loaded
        if (vrgStrip[iStrip + vrgBstripInfo[*msobf].iStripStart] == NULL)
        {
            // enter Office critical section
            OFFICE_CRITICAL_SECTION ocr;

            // Load resource if never loaded. 
            // Note that we need to check again while under critical section,
            // this ensure that bstrip is only loaded once even if multiple threads
            // try to do so at the same time.
            if (vrgStrip[iStrip + vrgBstripInfo[*msobf].iStripStart] == NULL)
            {
                HBSTRIP hbstrip;

                const int olb = (msoolbShared | msoolbVertical | msoolbNoDither);

                // load into local variable and assign array slot only after load is complete,
                // this way we guarantee that only fully initialized BSTRIPs are ever stored
                // in the vrgBstripInfo array.
                if (!MsoFLoadBStripEx(vhinstMsoIntl, vrgBstripInfo[*msobf].btnstripbase + iStrip, vrgBstripInfo[*msobf].cIcons, olb, 0, 0, &hbstrip))
                {
                    return FALSE;
                }

                vrgStrip[iStrip + vrgBstripInfo[*msobf].iStripStart] = hbstrip;
            }
        }

        *phbstrip = vrgStrip[iStrip + vrgBstripInfo[*msobf].iStripStart];

        *piBtn = iEntry;
    }

LExit:
    if (pfBlankIcon)
    { 
        const int iStripNoIcon = BStripStrip(msotcidNoIcon, *msobf);
        const int iEntryNoIcon = BStripEntry(msotcidNoIcon, *msobf);

        MsoAssertTag(BStripStrip(msotcidNoIcon, *msobf) == iStripNoIcon && BStripEntry(msotcidNoIcon, *msobf) == iEntryNoIcon, ASSERTTAG('kjrh'));

        *pfBlankIcon = (*phbstrip == vrgStrip[iStripNoIcon + vrgBstripInfo[*msobf].iStripStart]) && (*piBtn == iEntryNoIcon);
    }

    return TRUE;
}
