#pragma once

/*****************************************************************************
	msoem.h

	Owner: DaleG
	Copyright (c) 1997 Microsoft Corporation

	Typedef file for Rules Engine of Event Monitor.

*****************************************************************************/

#ifndef MSOEM_H
#define MSOEM_H

#ifndef MSO_H
#pragma message ("MsoEM.h file included before Mso.h.  Including Mso.h.")
#include "mso.h"
#endif

MSOEXTERN_C_BEGIN	// ***************** Begin extern "C" ********************


#if 0 //$[VSMSO]
#include "msoemtyp.h"
#endif //$[VSMSO]

#if STANDALONE_WORD
#error
#endif


//---------------------------------------------------------------------------
// Define useful macros (mostly stolen from Word)
// REVIEW daleg: these need to be Mso-ized
//---------------------------------------------------------------------------

#if !WORD_BUILD


///#define FNeNcLpxch(lpch1, lpch2, cch) FNeNcRgxch(lpch1, lpch2, cch)
#define SzFromStz(stz) 		 	((stz)+1)
#define XszFromXstz(xstz)	 	((xstz)+1)
#if !defined(ACCESS_BUILD) || defined(ZENSTAT_LIB_DEF)
#define CchSz(sz)				MsoCchSzLen(sz)
#ifndef CchWz
#define CchWz(wz)				MsoCchWzLen(wz)
#endif /* !CchWz */
#endif /* !ACCESS_BUILD */
#define XszFromXstz(xstz)	 	((xstz)+1)
#define RgxchFromXstz(xstz)	 	((xstz)+1)
#define CchXstz(xstz)			CchXst(xstz)
#define CchXst(xst) 			(*xst)
#define cbMaxSt		256
#define cbMaxSz		256
#define cbMaxStz 	257
#define cchMaxSz	255
#ifndef WORD_BUILD
#ifndef fTrue
#define fTrue 1
#endif
#ifndef fFalse
#define fFalse 0
#endif
#define tYes	1      /* much like fTrue */
#define tNo		0      /* much like fFalse */
#define tMaybe	(-1)   /* the "different" state */
#define iNil (-1)
#endif /* !WORD_BUILD */
#undef wValue
#define STATIC

// Return "low" value of a split value
#define W1OfPsv(psv) 			((psv)->wValue1)

// Set "low" value of a split value
#define SetW1OfPsv(psv, w)		((psv)->wValue1 = (short)(w))

// Increment "low" value of a split value
#define IncrW1OfPsv(psv, w)		((psv)->wValue1 += (w))

// Return "high" value of a split value
#define W2OfPsv(psv) 			((psv)->wValue2)

// Set "high" value of a split value
#define SetW2OfPsv(psv, w)		((psv)->wValue2 = (short)(w))

// Increment "high" value of a split value
#define IncrW2OfPsv(psv, w)		((psv)->wValue2 += (w))

// Merge two shorts into a long, compatible with SVL
#define SvlFromWW(wValue2, wValue1) \
			(((long) (wValue2) << 16) + ((long) (wValue1)))


#define PbCopyRgb(pbFrom, pbTo, cb)	\
			(((unsigned char *) memmove((pbTo), (pbFrom), (cb))) + (cb))
#define PbCopyRgbNo(pbFrom, pbTo, cb) \
    		(memcpy((pbTo), (pbFrom), (cb)), (pbTo) + (cb))
#define CopyRgb(pbFrom, pbTo, cb)		memmove(pbTo, pbFrom, cb)
#define CopyRgbNo(pbFrom, pbTo, cb)		memcpy(pbTo, pbFrom, cb)
#define CopyRgxch(pxchFrom, pxchTo, cch) \
			CopyRgb(pxchFrom,pxchTo,(cch)*cbXchar)

#define ClearLp(lpv, type) \
			MsoMemset((lpv), '\0', sizeof(type))

#define ClearLprg(lpv, type, iMax) \
			MsoMemset((lpv), '\0', sizeof(type) * (iMax))

#define CopyLprg(lpvFrom, lpvTo, type, iMax) \
			CopyRgb((lpvFrom), (lpvTo), (unsigned int)(sizeof(type) * (iMax)))

#define IMaxRg(dcl, type) \
			(sizeof(dcl) / sizeof(type))


#define ClearLprgBlocked(lpv, type, iMax) \
			MsoMemset((lpv), '\0', (sizeof(type) * (iMax) + sizeof(void *)))

#define IMaxRgBlocked(dcl, type) \
			((sizeof(dcl) / sizeof(type)) - sizeof(void *))


#endif /* !WORD_BUILD */

// Allocate storage for a given type, return pointer
#define MsoPNewEm(type) \
			MsoPNewCbDg(type, (unsigned int) sizeof(type), msodgMisc)

// Allocate storage for a given type and data group, return pointer
#define MsoPNewDg(type, dg) \
			MsoPNewCbDg(type, (unsigned int) sizeof(type), (dg))

// Allocate storage for an array of a given type, return pointer
#define MsoPNewEmRg(type, iMax) \
			MsoPNewRgDg(type, iMax, msodgMisc)

// Allocate storage for an array of given type and data group, return pointer
#define MsoPNewRgDg(type, iMax, dg) \
			MsoPNewCbDg(type, (unsigned int)(sizeof(type) * (iMax)), (dg))

// Allocate storage for an array of type, plus a pointer to next alloc block
#define MsoPNewEmRgBlocked(type_blk, type, iMax) \
			MsoPNewRgBlockedDg(type_blk, type, iMax, msodgMisc)

// Allocate storage for an array of type, plus a pointer to next alloc block
#define MsoPNewRgBlockedDg(type_blk, type, iMax, dg) \
			MsoPNewCbDg(type_blk, MsoCbRgBlocked(type_blk, type, iMax), (dg))

// Return the number of bytes used by a block allocation
#define MsoCbRgBlocked(type_blk, type, iMax) \
			((sizeof(type) * ((iMax) - 1)) + sizeof(type_blk))

// Allocate storage for a given type, but specifying size, return pointer
#define MsoPNewEmCb(type, cb) \
			MsoPNewCbDg(type, cb, msodgMisc)

// Allocate storage for a given type, but specifying size, return pointer
#define MsoPNewCbDg(type, cb, dg) \
			((type *) MsoPvAlloc((cb), (dg)))

// Zero storage for a type, given pointer
#define MsoClearPv(pv, type) \
			MsoClearPvCb((pv), sizeof(type))

// Zero storage for number bytes given
#define MsoClearPvCb(pv, cb) \
			MsoMemset((pv), '\0', (cb))

// Zero storage for an array of a type, given pointer
#define MsoClearRg(pv, type, iMax) \
			MsoMemset((pv), '\0', sizeof(type) * (iMax))

// Copy storage for an array of a type, given pointers, checking for overlaps
#define MsoCopyRg(rgFrom, rgTo, type, iMax) \
			MsoMemmove((rgTo), (rgFrom), sizeof(type) * (iMax))

// Copy storage for an array of a type, given pointers, no overlap checking
#define MsoCopyRgNo(rgFrom, rgTo, type, iMax) \
			MsoMemcpy((rgTo), (rgFrom), sizeof(type) * (iMax))

// Copy an array of Unicode characters, checking for overlaps
#define MsoCopyRgwch(rgFrom, rgTo, cch) \
			MsoCopyRg((rgFrom), (rgTo), WCHAR, (cch))

// Copy an array of Unicode characters, no overlap checking
#define MsoCopyRgwchNo(rgFrom, rgTo, cch) \
			MsoCopyRgNo((rgFrom), (rgTo), WCHAR, (cch))

// Return maximum number of bytes used by array, given *definition*
#define MsoIMaxRg(rg) \
			(sizeof(rg) / sizeof(rg[0]))

#define MsoPReallocEm(pv, type, cb) \
			((type *) MsoPvReallocEmPvCb((pv), (unsigned int) (cb)))

#define MsoPReallocEmRg(pv, type, iMax) \
			((type *) \
				MsoPvReallocEmPvCb \
					((pv), (unsigned int) (sizeof(type) * (iMax))))

#define MsoPvReallocEmPvCb(pv, cb) \
			(void *)MsoPvRealloc(pv, cb)

#define MsoPvAllocEmCb(cb) 		(void *)MsoPvAlloc((cb), msodgMisc )
#define MsoFreeEmPv(pv) 	{ if (pv) MsoFreePv(pv); }
#define MsoFreeEmPvClear(pv)	MsoFreeEmPpv(&(pv))

#define MsoFreeEmPpv(ppv)	{	\
								MsoFreeEmPv(*(void **) (ppv));	\
								*(void * UNALIGNED *) (ppv) = NULL;	\
								}


#define MsoPbCopyRgb(pbFrom, pbTo, cb)	\
			(((unsigned char *) MsoMemmove((pbTo), (pbFrom), (cb))) + (cb))
#define MsoPbCopyRgbNo(pbFrom, pbTo, cb) \
    		(MsoMemcpy((pbTo), (pbFrom), (cb)), (pbTo) + (cb))

// Return whether x is between lo and hi values, inclusive.
#define MsoFBetween(x, lo, hi)		((x) >= (lo)  &&  (x) <= (hi))
// REVIEW: consider the following minutely faster but more dangerous version
//#define MsoFBetween(x, lo, hi)		((unsigned)((x) - (lo)) <= (hi) - (lo))


#include "msodbglg.h"

#ifdef STANDALONE

#define MsoPvAlloc(cb, dg)		malloc(cb)
#define MsoFreePv(pv)			free(pv)
#define MsoPvRealloc(pv, cb)	realloc((pv), (cb))

#endif /* STANDALONE */


#if !defined(WORD_BUILD)  &&  !defined(WORD_H)

#ifndef CommSz
#define CommSz(sz)          OutputDebugStringA(sz)
#endif /* !CommSz */

#ifndef AVOID_MSOEM_DMUOPEN_CONFLICT
#ifndef ACCESS_BUILD
typedef void		   *DOC;							// Define object
#define docNil 	((DOC) NULL)
#else /* ACCESS_BUILD */
#define pdocNil 	((DOC *) NULL)
#endif /* !ACCESS_BUILD */
#endif //AVOID_MSOEM_DMUOPEN_CONFLICT

#define cpNil ((MSOCP) -1)
#define cp0 ((MSOCP) 0)
///typedef unsigned short BF;

typedef unsigned char uchar;
typedef unsigned int uint;
typedef unsigned short ushort;
typedef unsigned long ulong;

#endif /* !WORD_BUILD  &&  !WORD_H */


#if 0 //$[VSMSO]
#ifdef OFFICE_BUILD
MSOAPIX_(BOOL) MsoFInitOfficeEm(struct MSOINST *pinst);	// Init MSO Em

// Query type in FEmNotifyAction(msoemssAppEm,..)
#define msonaAppEmUseMso		1						// App needs MSO EM?
#endif /* !OFFICE_BUILD */

MSOAPI_(long) MsoFValidURLPwchEx(
	WCHAR				*pwch,
	int					cch,
	MSOCP				*pcpFirst,
	MSOCP				*pcpLim,
	int					*pbhf
	);
MSOAPI_(long) MsoFValidURLPwch(                         // URL string valid?
	WCHAR			   *pwch,
	int					cch,
	MSOCP              *pcpFirst,                       // RETURN
	MSOCP              *pcpLim                          // RETURN
	);
MSOAPIX_(int) MsoLRuleParsePwch(							// Parse string w/rules
	WCHAR              *pwch,
	int                 cch,
	int					rulevt,
	int					rulg
	);

#ifdef DEBUG
MSOAPI_(int) MsoLRuleParseFile(							// Parse file w/rules
	__no_ecount // debug only
	char               *rgchPath,
	int					rulevt,
	int					rulg
	);
#ifdef OFFICE_BUILD
BOOL FWriteEmBe(LPARAM lParam, struct MSOINST *pinst);	// Mark MSO Em memory
#endif /* OFFICE_BUILD */
#endif /* DEBUG */

#endif //$[VSMSO]
MSOEXTERN_C_END		// ****************** End extern "C" *********************

#endif /* !MSOEM_H */


