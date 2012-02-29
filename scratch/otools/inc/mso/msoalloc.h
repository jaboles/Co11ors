#pragma once

/*************************************************************************
 	msoalloc.h

 	Owner: rickp
 	Copyright (c) 1994 Microsoft Corporation

	Standard memory manager routines for the Office world.
*************************************************************************/

#if !defined(MSOALLOC_H)
#define MSOALLOC_H

#if DEBUG
#define DEBUGHEAP 1
#endif
#ifndef STATIC_LIB_DEF
// Uncomment this if you need to dump the heap in the ship build:
//#define DEBUGHEAP 1
#endif

// VSMSO uses regular HeapAlloc/HeapFree always
#define NTHEAP

#if !defined(MSODEBUG_H)
#include <msodebug.h>
#endif

#if defined(__cplusplus)
extern "C" {
#endif

/* data groups - objects allocated out of the same data group will tend
   to have some memory locality, which can reduce paging */
enum
{
	msodgMask = 0x07ff,	// all dgs must fall in this range

	/* standard data groups */

	msodgNil      = 0x0,
	msodgTemp     = 0x1,	// fast, temporary allocations
	msodgMisc     = 0x2,	// miscellaneous allocations
	msodgBoot     = 0x3,	// boot time allocations
	msodgCore     = 0x4,	// core, commonly used allocations
	msodgPpv      = 0x5,	// core, double-deref allocations
#ifdef DEBUG
	msodgDebug    = 0x7,	// Debug only allocations should go here
#endif

	/* start of Office data groups */

	msodgToolBar = 0x100,
	msodgSDI = msodgToolBar,
	msodgDraw = 0x101,
	msodgDr = msodgDraw,
	msodgEscherLarge = msodgDraw, // old name for msodgDraw TODO peteren: Remove
	msodgEscherSmall = msodgDraw, // old name for msodgDraw TODO peteren: Remove
#ifdef DEBUG
	msodgDrDebug = msodgDraw,
#endif
	msodgDrDoc = msodgDraw,
	msodgDrView = msodgDrDoc,
	msodgGEffect = 0x102,     // Graphic effects go here, NOTHING ELSE!!!
	msodgGEL3D = msodgGEffect,       // GEL 3D graphics effects go here
	msodgGELLarge = msodgDraw,    // GEL creates really big things here
	msodgGELCache = msodgDraw,    // Bitmap/blip cache overhead date
	msodgGELTemp = msodgDraw,     // Appropriate because other msodgDraw does
											// not overlap with this.
	msodgGELMisc = msodgMisc,     // Used only where API does not identify dg
	msodgAssist = 0x103,
	msodgFC = msodgAssist,
	msodgBalln = msodgAssist,
	msodgAnswerWizard = msodgAssist,
	msodgCompInt = msodgAssist,
	msodgWebToolbar = msodgToolBar,
	msodgSdm = msodgAssist,
	msodgMso95 = msodgMisc,
	msodgAutoCorr = msodgMisc,
	msodgUserDef = msodgMisc,
	msodgStmIO = msodgMisc,
	msodgDocSum = msodgMisc,
	msodgPropIO = msodgMisc,
	msodgPostDoc = msodgMisc,
	msodgLoad = msodgMisc,
	msodgWebClient = 0x104,
	msodgWcHtml = msodgWebClient,
	msodgTCO = 0x105,
	msodgProg = 0x106,
	msodgVse = 0x107,
	msodgWebTheme = 0x108,
	msodgConference = 0x109,     // used by conference feature
	msodgSpeech = 0x110,         // used by Office Speech objects
    msodgSow = 0x111,
    msodgDigSig = 0x112,         // Digital signatrue
	msodgSearch = 0x113,
	msodgIWPC	= 0x114,
	msodgCrypt = 0x115,          // Password encryption
	msodgThumbView = 0x116,		 // Thumbnail view in File.Open
	msodgActivation = 0x117,	 // Office Activation (for animations)
	msodgLic = 0x118,		 // Office licensing
	msodgIM = 0x119,		// Office Instant Messaging
	msodgSqm        = 0x11A,     // Service Quality Monitoring
	msodgFS = 0x11B,        // Office Foundation Services
	msodgXml = 0x11C,
	msodgDss = 0x11D,
	msodgXDM = 0x11E,	// XML Display Mapper
	msodgRef = 0x11F,   // Research and Reference interface
	msodgSharing = 0x120,
	msodgDWS = 0x121,	// Document Worksace
	msodgOString = 0x122, // OString allocation
	msodgDSP = msodgMisc, // DSP Connection info objects
	msodgVersions = 0x123,
	msodgSplitDropdown = 0x124,
	msodgRefContent = 0x125,
	/* Add new Office data groups here */

    /* Word data groups */
    msodgWordPerm = 0x200,
    msodgWordTemp = 0x201,

	/* Data groups reserved for GDI Plus - data allocated *within*
		GDI Plus comes out of these groups.  Other data must not be
		allocated from these (doing so would damage our internal memory
		checking.)  At present data groups 0x400 to 0x4ff are reserved
		for GDI Plus. */
	msodgGDIPlus = 0x0400,
	msodgExtMask = 0x00FF, // used to check for a GDIPlus dg

	/* shared memory data groups */
	
	msodgShMask = 0x007F,
	msodgShBase = 0x0780, // all shared dgs must be larger than this
	msodgShMisc = 0x0780,
	msodgShCore = 0x0781,
	msodgShToolbar = 0x0782,
	msodgShDr = 0x0783,
	msodgShCompInt = 0x0784,
	msodgShFcNtf = 0x0785,

	msodgShNstb    = 0x7f0,  // Nstb block in the RSG

	msodgExec = 0x0800,	// executeable sb
	msodgNonPerm = 0x1000,	// non-permanent 
	msodgPerm = 0x2000,	// permanent 
	
	/* optimization flags to avoid attempting allocations out of a full sb */

	msodgOptRealloc = 0x4000, // failed a realloc in this sb
	msodgOptAlloc = 0x8000, // failed an alloc in this sb
	
};
#define msocdgShBitmap (msodgShBitmapMax - msodgShBitmapBase) // Number of shared Bitmap groups

/*************************************************************************
	The other memory manager, for those who don't like to keep track
	of the sizes of things.
*************************************************************************/

#if DEBUG
	__maybe_null MSOAPI_(void*) MsoPvAllocCore(int cb, int dg, const CHAR* szFile, int li);
	#define MsoPvAlloc(cb, dg) MsoPvAllocCore((cb), (dg), __FILE__, __LINE__)
#else
	__maybe_null MSOAPI_(void*) MsoPvAllocCore(int cb, int dg);
	#define MsoPvAlloc(cb, dg) MsoPvAllocCore((cb), (dg))
#endif

MSOAPI_(void) MsoFreePv(void* pv);

__maybe_null MSOAPI_(void*) MsoPvRealloc(__no_count void* pv, int cb);

MSOAPI_(long) MsoCbSizePv(const void* pv);


/*	Our low-level memory allocator.  Does not keep track of the size
	of the object, so the caller is responsible for knowing it.  Supports
	data groups, which can be used to force some allocation locality.
	Allocates cb bytes out of the datagroup dg.  Returns the new pointer
	at *ppv, and returns TRUE if successful */

// Under the NT heap, which does keep track of sizes; this is merely
// a wrapper to the "other" memory manager.
#if DEBUG
	#define MsoFAllocMem(ppv, cb, dg) MsoFAllocMemCore(ppv, cb, dg, __FILE__, __LINE__)
	MSOAPI_(BOOL) MsoFAllocMemCore(void** ppv, int cb, int dg, const CHAR* szFile, int li);
	#define MsoFAllocShMem(ppv, cb, dg) MsoFAllocShMemCore(ppv, cb, dg, __FILE__, __LINE__)
	MSOAPI_(BOOL) MsoFAllocShMemCore(void** ppv, int cb, int dg, const CHAR* szFile, int li);
#else
	#define MsoFAllocMem(ppv, cb, dg) MsoFAllocMemCore(ppv, cb, dg)
	MSOAPI_(BOOL) MsoFAllocMemCore(void** ppv, int cb, int dg);
	#define MsoFAllocShMem(ppv, cb, dg) MsoFAllocShMemCore(ppv, cb, dg)
	MSOAPI_(BOOL) MsoFAllocShMemCore(void** ppv, int cb, int dg);
#endif

/*	Our low-level memory free routine, the opposite bookend of
	MsoFAllocMem.  Note that the caller is responsible for keeping track
	of the size of the object freed here.  Frees the pointer pv, which must
	have been allocated with the size cb. */
#if DEBUG
	MSOAPI_(void) MsoFreeMem(void* pv, int cb);
	MSOAPI_(void) MsoFreeShMem(void* pv, int cb);
#else
	#define MsoFreeMem _MsoPvFree
	#define MsoFreeShMem _MsoPvShFree
#endif

MSOAPI_(void) _MsoPvFree(void* pv, int cb);
MSOAPI_(void) _MsoPvShFree(void* pv, int cb);

MSOAPI_(BOOL) MsoFReallocMemCore(void** ppv, int cbOld, int cbNew, int dg);
#define MsoFReallocMem(ppv, cbOld, cbNew, dg) MsoFReallocMemCore(ppv, cbOld, cbNew, dg)

#if DEBUG
	MSOAPI_(void) MsoGetAllocCount(int * pnAlloc, int * pnFree);
#endif

typedef struct MSOMEMSTATS
{
	long cbCur;
	long cbHigh;
	long cbHeap;
} MSOMEMSTATS;

/*************************************************************************
	The double-deref memory manager
*************************************************************************/

#ifdef DEBUG
	MSOAPI_(void **) MsoPpvAllocCore(long, const char *, int);
	#define MsoPpvAlloc(cb) MsoPpvAllocCore(cb, __FILE__, __LINE__)
#else
	MSOAPI_(void **) MsoPpvAllocCore(long);
	#define MsoPpvAlloc MsoPpvAllocCore
#endif

MSOAPI_(long) MsoCbSizePpv(void **ppv);

MSOAPI_(void) MsoFreePpv(void **ppv);

MSOAPI_(BOOL) MsoFPpvRealloc(void **ppv, long cb);

#if DEBUG
	MSOAPIXX_(BOOL) MsoFWritePpvBe(LPARAM lparam);
#else
	#define MsoFWritePpvBe(lparam) (1)
#endif

/*************************************************************************
	Other miscellaneous heap functions
*************************************************************************/

MSOAPI_(long) MsoCbHeapUsed();
MSOAPI_(BOOL) MsoFCompactHeap();

MSOAPI_(BOOL) MsoFHeapHp(const void *pv);
MSOAPI_(BOOL) MsoFAllocMem64(void** ppv, int cb, int dg, BOOL fComp);

MSOAPI_(void) MsoFreeAllNonPerm();
MSOAPI_(BOOL) MsoFCanQuickFree();

#if !defined(_ALPHA_) && !defined(_IA64_)
/* All sub-allocations are made on 4-byte boundaries for X86 */
#define MsoMskAlign()	(4-1)
#else
/* All sub-allocations are made on 8-byte boundaries for Alpha */
#define MsoMskAlign()	(8-1)
#endif

/* Round cb up to the next 4-byte boundary */
#define MsoCbHeapAlign(cb)		(((cb) + MsoMskAlign()) & ~MsoMskAlign())
#define MsoCbHeapAlignDown(cb)	((((cb) - 1) + MsoMskAlign()) & ~MsoMskAlign())

/*************************************************************************
	Mark and Release allocations, used for LIFO-pattern (e.g., stack) 
	allocations.  Especially useful for replacing local string buffers.
*************************************************************************/

#if DEBUG
	__maybe_null MSOAPI_(BOOL) MsoFMarkMemCore(_Out_ void** ppv, int cb, const CHAR* szFile, int li);
	#define MsoFMarkMem(ppv, cb) MsoFMarkMemCore(ppv, cb, __FILE__, __LINE__)
#else
	__maybe_null MSOAPI_(BOOL) MsoFMarkMemCore(_Out_ void** ppv, int cb);
	#define MsoFMarkMem(ppv, cb) MsoFMarkMemCore(ppv, cb)
#endif
#ifdef __cplusplus
inline BOOL MsoFMarkRgwchMem(_Out_ WCHAR** prgwch, DWORD cwch )
{
   return MsoFMarkMem( reinterpret_cast<void**>(prgwch), cwch*sizeof(WCHAR));
}
#else
#define MsoFMarkRgwchMem( rgwch, cwch ) MsoFMarkMem((void**)rgwch, (cwch)*sizeof(WCHAR))
#endif

#if DEBUG
MSOAPI_(BOOL) MsoFReallocMarkMemCore(void** ppv, int cb, const CHAR* szFile, int li);
#define MsoFReallocMarkMem(ppv, cb) MsoFReallocMarkMemCore(ppv, cb, __FILE__, __LINE__)
#else
MSOAPI_(BOOL) MsoFReallocMarkMemCore(void** ppv, int cb);
#define MsoFReallocMarkMem(ppv, cb) MsoFReallocMarkMemCore(ppv, cb)
#endif

MSOAPI_(void) MsoReleaseMemCore(void* pv);
#define MsoReleaseMem(pv) MsoReleaseMemCore(pv)

#ifdef DEBUG
MSOAPI_(void) MsoCheckMarkMem(void);
#endif


/*************************************************************************
	Integrity check for the memory manager.
*************************************************************************/

// Used during crash recovery -- frees do nothing; allocs go to new HBs
// Minimizes the chance of 2nd failure due to a corrupt heap.
MSOAPI_(void) MsoEnterHeapProtectMode(void);

#if DEBUGHEAP
	MSOAPI_(BOOL) MsoCheckHeap(void);
	MSOAPI_(BOOL) MsoFCheckAlloc(void* pv, int cb, const char * pszTitle);
	MSOAPI_(void) MsoDumpHeap(void);
	MSOAPI_(BOOL) MsoFPvLegit(void *);
	MSOAPI_(DWORD) MsoGetPvAllocatorLineCb(const void *);
	MSOAPI_(const char *) MsoGetPvAllocatorFileCb(const void *);
	MSOAPI_(DWORD) MsoGetPvAllocatorLineNoCb(const void *);
	MSOAPI_(const char *) MsoGetPvAllocatorFileNoCb(const void *);
#else
	#define MsoCheckHeap() (1)
	#define MsoFCheckAlloc(pv, cb, pszTitle) (1)
	#define MsoDumpHeap()
	#define MsoFPvLegit(pv) (1)
#endif
	MSOAPI_(BOOL) MsoFPvRangeOkay(const void *pv, int cb);

#if defined(__cplusplus)
}
#endif


/*************************************************************************
	Heap integrity checks
*************************************************************************/

enum
	{
	msobtInvalid = 0xFFFFFFFF,
	};

/* The MsoSaveBe API is used to write out a records for an allocation that 
	was made using the office allocation API.  This record is used to determine
	if any memory has leaked.  The pinst parameter should be the HMSOINST that
	is returned from the call to MsoFInitOffice API.  If the allocation was
	done by the Office DLL, then the pinst parameter should be NULL.  The
	lparam parameter provides a way for an application to have information 
	passed back to it when it writes out is be records.  The value of the 
	lparam parameter will be the same as was pass to the MsoFChkMem API.  The
	fAllocHasSize parameter specifies if the allocation used the MsoFAllocMem
	API or the MsoPvAlloc API.  If MsoPvAlloc is used then the fAllocHasSize
	parameter should be true otherwise it should be false.  Operator new uses 
	the MsoPvAlloc API so all object should set the fAllocHasSize to False. 
	The hp parameter points to the allocation.  The cb parameter is the size
	of the allocation in bytes.  The bt parameter is the type of the allocation.
	Each different type should have its own bt type.  This type is maintained
	in the kjalloc.h file in the office enum.  It also must be added to this file
	in the vmpbtszOffice	array. */
#if DEBUG
	MSOAPI_(BOOL) MsoFSaveBe(HMSOINST hinst, LPARAM lparam, BOOL fAllocHasSize,
				VOID *hp, unsigned cb, int bt);
#else
	#define MsoFSaveBe(hinst, lparam, fAllocHasSize, hp,  cb, bt) (TRUE)
#endif

/* Check the Office heap for consistency and report errors. The phinst
   parameter contains the office instance needed for call backs.  The lparam
   parameter provides a way for an application to supply some context
   information which the FCheckAbort method is called.  The fMenu parameter
	specifies if this function should be interruptible. */
#if DEBUG
	MSOAPI_(BOOL) MsoFChkMem(HMSOINST hinst, BOOL fMenu, LPARAM lparam);
#else
	#define MsoFChkMem(hinst, fMenu, lparam) (TRUE)
#endif
	
/* This function returns back the current iteration number asscociated with
	a call to MsoChkMem.  This value can be used to reduce the change of 
	writing a duplicate BE record.  */
#if DEBUG
	MSOAPI_(DWORD) MsoDWGetChkMemCounter(void);
#endif
	
/* Given a BE this function will return back the pointer return back from the
	allocation */
#if DEBUG
	MSOAPI_(void*) MsoPvFromBe(MSOBE* pbe);
#endif
	
/* Given a BE this function will return back the size of the original
	allocation */	
#if DEBUG
	MSOAPIXX_(int) MsoCbFromBe(const MSOBE* pbe);
#endif
	
/* Given a BE this function will return back the bt field.  If the HMSOINST
	file is NULL (an office bt) then msobtInvalid will be returned */
#if DEBUG
	MSOAPI_(int) MsoBtFromBe(const MSOBE* pbe);
#endif
	
/* write out to a file all of the BE that have been written. It will also
	write out the msissing entries. */		
#if DEBUG
	MSOAPI_(BOOL) MsoFDumpBE(HMSOINST pinst, const char* szFileName,
	   	BOOL fJustErrors, LPARAM lParam);
#endif

#if DEBUG
// writes out BE for pv - pv should have been allocated using MsoPvAlloc or one of its derivatives
	MSOAPI_(void) MsoSaveBeMsoAlloc(void *pv);
#endif

#if DEBUG && defined(STATIC_LIB_DEF)
	BOOL MsoFCheckStaticLibForLeaks(UINT *cLeaks);
#endif

// DM has its own operator new.
// REVIEW KirkG: This sucks!  We turn this code on when #include'd into
//   mso files, but not when included into the dm files.
#if defined(__cplusplus) && !(defined(DM96) && defined(DEBUG))

/*************************************************************************
	C++ operator new and variants
*************************************************************************/

#if !defined(MSO_NO_OPERATOR_NEW)

	#ifndef _INC_NEW
	#include <new.h>
	#endif

	__maybe_null static inline void* __cdecl operator new(size_t cb, int dg)
		{
		//MsoAssertTag(!FDgIsSharedDg(dg), ASSERTTAG('epyl'));     // Cannot be a shared DG
		void *pv = MsoPvAlloc(cb, dg);
		#ifdef DEBUG
			if (pv != NULL)
				MsoDebugFill(pv, cb, msomfClass);
		#endif
		return pv;
		}

	__maybe_null static inline void* __cdecl operator new(size_t cb)
		{
		return operator new(cb, msodgMisc);
		}

	__maybe_null static inline void * __cdecl operator new[](size_t cb)
		{
		return operator new(cb, msodgMisc);
		}

	__maybe_null static inline void* __cdecl operator new[](size_t cb, int dg)
		{
		return operator new(cb, dg);
		}

	static inline void __cdecl operator delete(void* pv)
		{
		/* handle NULL pointers for sloppy people */
		if (pv)
			MsoFreePv(pv);
		}

	static inline void __cdecl operator delete[](void* pv)
		{
		return operator delete(pv);
		}

	#if DEBUG

		__maybe_null static inline void* _cdecl operator new(size_t cb, int dg, const CHAR* szFile, int li)
			{
			//MsoAssertTag(!FDgIsSharedDg(dg), ASSERTTAG('kzjs'));     // Cannot be a shared DG
			void *pv = MsoPvAllocCore(cb, dg, szFile, li);
			#ifdef DEBUG
				if (pv != NULL)
					MsoDebugFill(pv, cb, msomfClass);
			#endif
			return pv;
			}

		__maybe_null static inline void* _cdecl operator new[](size_t cb, int dg, const CHAR* szFile, int li)
			{
			return operator new(cb, dg, szFile, li);
			}

		// Catch callers that don't define a datagroup, but do define file/line
		__maybe_null static inline void* __cdecl operator new(size_t cb, const CHAR* szFile, int li)
		{
			return operator new(cb, msodgMisc, szFile, li);
		}

		__maybe_null static inline void* __cdecl operator new[](size_t cb, const CHAR* szFile, int li)
		{
			return operator new[](cb, msodgMisc, szFile, li);
		}

		#define newDebug(dg) new(dg, __FILE__, __LINE__)

	/* Note that by #define new(dg) here, we make it hard for people to
		override the new operators ona class-by-class basis.  new
		overrides should be, in general, rare, so it's probably not a big 
		deal.  But if someone really needs to do it, they'll need to 
		#undef new or define MSO_NO_OPERATOR_NEW */

		#define new(dg) newDebug(dg)

	#else

		#define newDebug new

	#endif

#endif /* MSO_NO_OPERATOR_NEW */

#endif	/* __cplusplus && !DM96 */


/*************************************************************************
	The Plex data structure.  A plex is a low-overhead implementation
	of a varible-sized array.
*************************************************************************/

/* The generic plex structure itself */
typedef struct MSOPX
{
	/* WARNING: the following must line up with the MSOTPX template and
		the MsoPxStruct macro */
	WORD iMac, iMax;	/* used size, and total allocated size */
	unsigned cbItem:16,	/* size of each data item */
		dAlloc:15,	/* amount to grow by when reallocating larger */
		fUseCount:1;	/* if items in the plex should be use-counted */
	int dg;	/* data group to allocate out of */
	BYTE* rg;	/* the data */
} MSOPX;


/*	Handy macro for declaring a named Plex structure - must line up with
	the MSOPX structure */
#define MsoPxStruct(TYP,typ) \
		struct \
		{ \
		WORD i##typ##Mac, i##typ##Max; \
		unsigned cbItem:16, \
			dAlloc:15, \
			fUseCount:1; \
		int dg; \
		TYP *rg##typ; \
		}

/* Handy macro for enumerating over all the items in a plex ppl, using loop
	variables p and pMac */
#define FORPX(p, pMac, ppl, T) \
		for ((pMac) = ((p) = (T*)((MSOPX*)(ppl))->rg) + ((MSOPX*)(ppl))->iMac; \
			 (p) < (pMac); (p)++)

/* Handy macro for enumerating over all the items in a plex ppl backwards,
	using loop variables p and pMac */
#define FORPX2(p, pMac, ppl, T) \
		for ((p) = ((pMac) = (T*)((MSOPX*)(ppl))->rg) + ((MSOPX*)(ppl))->iMac - 1; \
			 (p) >= (pMac); (p)--)


/*************************************************************************
	Creation and destruction
*************************************************************************/

#if DEBUG
MSOAPI_(BOOL) MsoFAllocPxDebug(void** pppx, unsigned cbItem, int dAlloc, unsigned iMax, int dg, const CHAR* szFile, int li);
#define MsoFAllocPx(pppx, cbItem, dAlloc, iMax, dg) \
		MsoFAllocPxDebug(pppx, cbItem, dAlloc, iMax, dg, __FILE__, __LINE__)
#else
MSOAPI_(BOOL) MsoFAllocPx(void** pppx, unsigned cbItem, int dAlloc, unsigned iMax, int dg);
#endif


#if DEBUG
MSOAPI_(BOOL) MsoFInitPxDebug(void* pvPx, int dAlloc, int iMax, int dg, const char* szFile, int li);
#define MsoFInitPx(pvPx, dAlloc, iMax, dg) \
		MsoFInitPxDebug(pvPx, dAlloc, iMax, dg, __FILE__, __LINE__)
#else
MSOAPI_(BOOL) MsoFInitPx(void* pvPx, int dAlloc, int iMax, int dg);
#endif

MSOAPIX_(BOOL) MsoFUseCountAllocPx(void** pppx, unsigned cbItem, int dAlloc, unsigned iMax, int dg);
MSOAPI_(void) MsoFreePx(void* pvPx);

/*************************************************************************
	Lookups
*************************************************************************/

typedef int (MSOPRIVCALLTYPE* MSOPFNSGNPX)(const void*, const void*);

MSOAPIMX_(void*) MsoPLookupPx(void* pvPx, const void* pvItem, MSOPFNSGNPX pfnSgn);
MSOAPI_(BOOL) MsoFLookupPx(void* pvPx, const void* pvItem, int* pi, MSOPFNSGNPX pfnSgn);

MSOAPI_(int) MsoFLookupSortPx(const void* pvPx, const void* pvItem, int* pi, MSOPFNSGNPX pfnSgn);
MSOAPIMX_(void*) MsoPLookupSortPx(void* pvPx, const void* pvItem, MSOPFNSGNPX pfnSgn);
MSOAPIX_(BOOL) MsoFNextLookupPx(void* pvPx, int iStart, const void* pvItem, int* pi, MSOPFNSGNPX pfnSgn);

/*************************************************************************
	Adding items
*************************************************************************/

MSOAPI_(int) MsoIAppendPx(void* pvPx, const void* pv);
__inline int MsoFAppendPx(void *pvPx, const void *pv) { return (MsoIAppendPx(pvPx, pv) >= 0); }
MSOAPI_(int) MsoIAppendUniquePx(void* pvPx, const void* pv, MSOPFNSGNPX pfnSgn);
MSOAPI_(int) MsoIAppendNewPx(void** ppvPx, const void* pv, int cbItem, int dg);
MSOAPIX_(BOOL) MsoFInsertNewPx(void** ppvPx, const void* pv, int cbItem, int i, int dg);
MSOAPI_(int) MsoIInsertSortPx(void* pvPx, const void* pv, MSOPFNSGNPX pfnSgn);
MSOAPI_(int) MsoIInsertSortDupPx(void* pvPx, const void* pv, MSOPFNSGNPX pfnSgn);
MSOAPI_(BOOL) MsoFInsertPx(void* pvPx, const void* pv, int i);
MSOAPIMX_(BOOL) MsoFInsertExPx(void* pvPx, const void* pv, int i);

/*************************************************************************
	Removing items
*************************************************************************/

MSOAPI_(int) MsoFRemovePx(void* pvPx, int i, int c);
MSOAPI_(void) MsoDeletePx(void* pvPx, int i, int c);


/*************************************************************************
	Sorting
*************************************************************************/

MSOAPI_(void) MsoQuickSortPx(void* pvPx, MSOPFNSGNPX pfnSgn);


/*************************************************************************
	Miscellaneous shuffling around
*************************************************************************/

MSOAPIMX_(void) MsoMovePx(void* pvPx, int iFrom, int iTo);
MSOAPI_(BOOL) MsoFCompactPx(void* pvPx, BOOL fFull);
MSOAPI_(BOOL) MsoFResizePx(void* pvPx, int iMac, int iIns);
MSOAPI_(BOOL) MsoFGrowPx(void* pvPx, int iMac);
MSOAPIX_(void) MsoStealPx(void *pvPxSrc, void *pvPxDest);
MSOAPI_(void) MsoEmptyPx(void *pvPx);
MSOAPI_(BOOL) MsoFClonePx(void *pvPxSrc, void *pvPxDest, int dg);

/*************************************************************************
	Plex with use count items utilities
*************************************************************************/

MSOAPIX_(int) MsoIIncUsePx(void* pvPx, int i);
MSOAPIX_(int) MsoIDecUsePx(void* pvPx, int i);


/*************************************************************************
	Debug stuff
*************************************************************************/

#if DEBUG
	MSOAPI_(BOOL) MsoFValidPx(const void* pvPx);
	MSOAPI_(BOOL) MsoFWritePxBe(void* pvPx, LPARAM lParam, BOOL fSaveObj);
	MSOAPI_(BOOL) MsoFWritePxBe2(void* pvPx, LPARAM lParam, BOOL fSaveObj, 
											BOOL fAllocHasSize);
#else
	#define MsoFValidPx(pvPx) (TRUE)
#endif



/*************************************************************************
	Plex class template - this is basically a big inline class wrapper 
	around the C plex interface.
*************************************************************************/

#ifdef __cplusplus
#if DEBUG
	#define _PlexAssertInfo ,__FILE__, __LINE__
#else
	#define _PlexAssertInfo
#endif
template <class S> class MSOTPX
{
public:
	/* WARNING: the following must line up exactly with the MSOPX structure */
	WORD iMac, iMax;	/* used size, and total allocated size */
	unsigned cbItem:16,	/* size of each data item */
		dAlloc:15,	/* amount to grow by when reallocating larger */
		fUseCount:1;	/* if items in the plex should be use-counted */
	int dg;	/* data group to allocate out of */
	S* rg;	/* the data */

	/* Unexciting constructor. */
	inline MSOTPX<S>(void) 
	{ iMax = iMac = 0; AssertMsgTemplate(sizeof(S) < 0x10000, NULL); cbItem = sizeof(S); rg = NULL; }

	/* Destructor to deallocate memory */
	inline ~MSOTPX<S>(void) 
	{ if (rg) MsoFreeMem(rg, cbItem * iMax); }

	inline BOOL FValid(void) const
	{ return MsoFValidPx(this); }

	#if DEBUG
		inline BOOL FInit(int dAlloc, int iMax, int dg, const char* szFile=__FILE__, int li=__LINE__)
		{
			AssertMsgTemplate(dAlloc < 0x8000, NULL);
			return MsoFInitPxDebug(this, dAlloc, iMax, dg, szFile, li);
//			return MsoFInitPxDebug(this, dAlloc, iMax, dg, __FILE__, li);
		}
	#else
		inline BOOL FInit(int dAlloc, int iMax, int dg)
		{ return MsoFInitPx(this, dAlloc, iMax, dg); }
	#endif

	inline int IMax(void) const
	{ return iMax; }

	inline int IMac(void) const
	{ return iMac; }

	inline S* PGet(int i) const
	{
	/* AssertMsgTemplate(i >= 0 && i < iMac, NULL); */
#ifdef DEBUG
	if (i < 0 || i >= iMac)
		return (NULL);			//  Try to force crash if use bogus data
#endif
#if BCHECK
		if (rg == NULL)
			return NULL;
#endif 
		return &rg[i];
	}

	inline void Get(__no_count S* p, int i) const
	{ AssertMsgTemplate(i >= 0 && i < iMac, NULL); *p = rg[i]; }

	inline void Put(S* p, int i)
	{ AssertMsgTemplate(i >= 0 && i < iMac, NULL); rg[i] = *p; }

	// plex[i] has exactly the same semantics and performance as plex.rg[i]
	inline S& operator[](int i)
	{ AssertMsgTemplate(i >= 0 && i < iMac, NULL); return rg[i]; }

	inline const S& operator[](int i) const
	{ AssertMsgTemplate(i >= 0 && i < iMac, NULL); return rg[i]; }

	inline S* PLookup(const S* pItem, MSOPFNSGNPX pfnSgn)
	{ return (S*)MsoPLookupPx(this, pItem, pfnSgn); }

	inline S* PLookupV(const void * pItem, MSOPFNSGNPX pfnSgn)
	{ return (S*)MsoPLookupPx(this, pItem, pfnSgn); }

	inline BOOL FLookup(const S* pItem, int* pi, MSOPFNSGNPX pfnSgn)
	{ return MsoFLookupPx(this, pItem, pi, pfnSgn); }

	inline BOOL FLookupV(const void *pItem, int* pi, MSOPFNSGNPX pfnSgn)
	{ return MsoFLookupPx(this, pItem, pi, pfnSgn); }

	inline BOOL FLookupSort(const S* pItem, int* pi, MSOPFNSGNPX pfnSgn)
	{ return MsoFLookupSortPx(this, pItem, pi, pfnSgn); }

	inline BOOL FLookupSortV(const void *pItem, int* pi, MSOPFNSGNPX pfnSgn)
	{ return MsoFLookupSortPx(this, pItem, pi, pfnSgn); }

	inline S* PLookupSort(const S* pItem, MSOPFNSGNPX pfnSgn)
	{ return (S*) MsoPLookupSortPx(this, pItem, pfnSgn); }
	
	inline S* PLookupSortV(const void* pItem, MSOPFNSGNPX pfnSgn)
	{ return (S*) MsoPLookupSortPx(this, pItem, pfnSgn); }

	inline void QuickSort(MSOPFNSGNPX pfnSgn)
	{ MsoQuickSortPx(this, pfnSgn); }

	inline int FAppend(const S* p)
	{ return MsoIAppendPx(this, p) != -1; }

	inline int IAppend(const S* p)
	{ return MsoIAppendPx(this, p); }

	inline int IAppendUnique(const S* p, MSOPFNSGNPX pfnSgn)
	{ return MsoIAppendUniquePx(this, p, pfnSgn); }

	inline BOOL FInsert(const S* p, int i)
	{ return MsoFInsertPx(this, p, i); }

	inline int FInsertEx(const S* p, int i)
	{ return MsoFInsertExPx(this, p, i); }

	inline int IInsertSort(const S* p, MSOPFNSGNPX pfnSgn)
	{ return MsoIInsertSortPx(this, p, pfnSgn); }

	inline int IInsertSortDup(const S* p, MSOPFNSGNPX pfnSgn)
	{ return MsoIInsertSortDupPx(this, p, pfnSgn); }

	inline int FRemove(int i, int c)
	{ return MsoFRemovePx(this, i, c); }

	inline void Delete(int i, int c)
	{ MsoDeletePx(this, i, c); }

	inline void Move(int iFrom, int iTo)
	{ MsoMovePx(this, iFrom, iTo); }

	inline BOOL FCompact(BOOL fFull)
	{ return MsoFCompactPx(this, fFull); }

	inline BOOL FResize(int iMac, int iIns)
	{ return MsoFResizePx(this, iMac, iIns); }

	inline BOOL FReplace(const S* p, int i)
	{ 
		if (i >= iMac && !FSetIMac(i+1))
			return FALSE; 
		rg[i] = *p;
		return TRUE;
	}

	inline BOOL FSetIMac(int iMac)
	{ return MsoFResizePx(this, iMac, -1); }

	inline BOOL FSetIMax(int iMax)
	{ return MsoFGrowPx(this, iMax); }

	inline int IIncUse(int i)
	{ return MsoIIncUsePx(this, i); }

	inline int IDecUse(int i)
	{ return MsoIDecUsePx(this, i); }

	inline void Steal(void *pvPxSrc)
	{ MsoStealPx(pvPxSrc, this); }

	inline void Empty()
	{ MsoEmptyPx(this); }

	inline BOOL FClone(void *pvPxDest, int dg)
	{ return MsoFClonePx(this, pvPxDest, dg); }

	inline int BpscBulletProof(void *pmsobpcb, int cb)
	{ return MsoBpscBulletProofPx(this, pmsobpcb, cb); }

#if DEBUG
	inline BOOL FWriteBe(LPARAM lParam, BOOL fSaveObj)
	{ return MsoFWritePxBe2(this, lParam, fSaveObj, TRUE); }
#endif
};

int MSOPRIVCALLTYPE SgnLCompare(const void *pv1, const void *pv2);

#endif /* __cplusplus */

/****************************************************************************
   Shared Memory  
****************************************************************** JIMMUR **/

/* Return Value enum for Initializing a shared memory area */
enum
{
   msosmcFailed,                 // The Shared memory allocation failed
   msosmcFound,                  // The named shared memory was found
   msosmcAllocated               // The named shared memeory was created
};

/****************************************************************************
   The MSOSG structure contains all of the "Global" globals for the Office
   DLL.  To add a "Global" global simply add an item to this structure.  If
   an initial value for that item is desired, then the private Globals
   structure must also be changed.

   This whole structure is initialized in office.cpp.  If a setting is added
   or removed from this structure, that initialization will need to change,
   which means updating vrgsysmetric, vrgsysmetric2, vrgsyscolor,
   or vrgscrdevcap.

   NOTE: Unlike private shared globals, members of MSOSG are not initialized
   to 0 when this stucture is created.  If you add a member which is not
   related to the lists above, you must initialize it in MsoInitShrGlobal.
****************************************************************************/

typedef struct 
{
	// vrgsyscolor
	COLORREF crActiveCaption;	    // System Color COLOR_ACTIVECAPTION
	COLORREF crBtnFace;			    // System Color COLOR_BTNFACE
	COLORREF crBtnHighlight;	    // System Color COLOR_BTNHIGHLIGHT
	COLORREF crBtnShadow;		    // System Color COLOR_BTNSHADOW
	COLORREF crBtnText;			    // System Color COLOR_BTNTEXT
	COLORREF crCaptionText;		    // System Color COLOR_CAPTIONTEXT
	COLORREF crGrayText;			// System Color COLOR_GRAYTEXT
	COLORREF crHighlight;		    // System Color COLOR_HIGHLIGHT
	COLORREF crHighlightText;	    // System Color COLOR_HIGHLIGHTTEXT
	COLORREF crInactiveCaption;     // System Color COLOR_INACTIVECAPTION
	COLORREF crInactiveCaptionText; // System Color COLOR_INACTIVECAPTIONTEXT
	COLORREF crInfoBk;              // System Color COLOR_INFOBK
	COLORREF crInfoText;            // System Color COLOR_INFOTEXT
	COLORREF crMenu;                // System Color COLOR_MENU
	COLORREF crMenuText;            // System Color COLOR_MENUTEXT
	COLORREF crScrollbar;		    // System Color COLOR_SCROLLBAR
	COLORREF crWindow;			    // System Color COLOR_WINDOW
	COLORREF crWindowFrame;		    // System Color COLOR_WINDOWFRAME
	COLORREF crWindowText;		    // System Color COLOR_WINDOWTEXT
	COLORREF cr3DLight;			    // System Color COLOR_3DLIGHT
	COLORREF cr3DDkShadow;		    // System Color COLOR_3DDKSHADOW
	COLORREF cr3DFace;		        // System Color COLOR_3DFACE
	COLORREF cr3DShadow;		    // System Color COLOR_3DSHADOW

	COLORREF crBtnLowFreq;		    // Color used for low frequency items in menus
	COLORREF crHyperlink;		    // Color used for hyperlinks - IE's reg setting
	COLORREF crFollowedHyperlink;   // Color used for followed hyperlinks (from
	                                // IE's reg setting)
	
	// vrgsysmetric
	int smCxBorder;				// System Metric SM_CXBORDER
	int smCxDlgFrame;				// System MetrIc SM_CXDLGFRAME
	int smCxFrame;					// System MetrIc SM_CXFRAME;
	int smCxFullScreen;			// System Metric SM_CXFULLSCREEN
	int smCxIcon;					// System Metric SM_CXICON
	int smCxScreen;				// System Metric SM_CXSCREEN
	int smCxSize;					// System Metric SM_CXSIZE
	int smCxSmIcon;				// System Metric SM_CXSMICON
	int smCxVScroll;				// System Metric SM_CXVSCROLL
	int smCyBorder;				// System Metric SM_CYBORDER
	int smCyCaption;				// System Metric SM_CYCAPTION
	int smCyDlgFrame;				// System Metric SM_CYDLGFRAME
	int smCyFrame;					// System Metric SM_CYFRAME;
	int smCyFullScreen;			// System Metric SM_CYFULLSCREEN
	int smCyHScroll;				// System Metric SM_CYHSCROLL
	int smCyIcon;					// System Metric SM_CYICON
	int smCyScreen;				// System Metric SM_CYSCREEN
	int smCySize;					// System Metric SM_CYSIZE
	int smCySmIcon;				// System Metric SM_CSMICON
	int smCyVScroll;				// System Metric SM_CYVSCROLL
	int smSlowMachine;			// System Metric SM_SLOWMACHINE
	int smCyMenu;					// System Metric SM_CYMENU

	// vrgscrdevcap
	int dxvInch;					// GetDeviceCaps LOGPIXELSX
	int dyvInch;					// GetDeviceCaps LOGPIXELSY
	int dxpScreen;					// GetDeviceCaps HORZRES
	int dypScreen;					// GetDeviceCaps VERTRES
	int dxmmScreen;				// GetDeviceCaps HORZSIZE
	int dymmScreen;				// GetDeviceCaps VERTSIZE
	int ccrScreen;					// GetDeviceCaps NUMCOLORS
	int cBitsPixelScreen;		// GetDeviceCaps BITSPIXEL (adjusted for min of all monitors)
	int cPlanesScreen;			// GetDeviceCaps PLANES

	BOOL fPaletteScreen;			// GetDeviceCaps RASTERCAPS & RC_PALETTE

	COLORREF crBtnDarken;		//	Calculated color for selected Buttons

	// vrgsysmetric2
	int smCxDoubleClk;			// System Metric SM_CXDOUBLECLK
	int smCyDoubleClk;			// System Metric SM_CYDOUBLECLK
	int smCxHScroll;				// System Metric SM_CXHSCROLL
	int smCxSizeFrame;			// System Metric SM_CXSIZEFRAME
	int smCySizeFrame;			// System Metric SM_CYSIZEFRAME
	int smCxMenuSize;				// System Metric SM_CXMENUSIZE
	int smCyMenuSize;				// System Metric SM_CYMENUSIZE
	int smCxCursor;				// System Metric SM_CXCURSOR
	int smCyCursor;				// System Metric SM_CYCURSOR

	int iSmCaptionWidth;			// System Parameter non-client iSmCaptionWidth
	int iSmCaptionHeight;		// System Parameter non-client iSmCaptionHeight

	unsigned int sBitDepth;		// UNUSED!!!  Its value is always 0

	int ccrPaletteScreen;		// GetDeviceCaps SIZEPALETTE
	
	#if DEBUG
		HWND hwndMonitor;		// debug monitor window
		HWND hwndMonitored;	// window of the monitored application
		DWORD pidMonitored;	// process id we're monitoring
	#endif

	BOOL fHighContrast;    // has the user set a high contrast display setting	
	BOOL fHighContrastMenus;   // Menus drawing should assume Highcontrast mode
	BOOL fCrHyperlinkInited;   // Has crHyperlink been delay init'ed yet?
	BOOL fCrFollowedHyperlinkInited; // Has crFollowedHyperlink be inited yet?
	
	int dxMenuBtn;  // size of MDI controls on hmenu
	int dyMenuBtn;  // size of MDI controls on hmenu

	LOGFONT lfCaptionFont;

	BOOL	fAppDisabled;			// App is disabled by licensing.
	BOOL	fSystemIsUsingClearType;	// Is Clear Type font smoothing being used?

} MSOSG;

/*	Get IE's hyperlink color */
MSOAPI_(COLORREF) MsoCrGetHyperlink();

/* Get IE's followed hyperlink color */
MSOAPI_(COLORREF) MsoCrGetFollowedHyperlink();

/****************************************************************************
   Office "Public Shared Global Routines"
   These sit on top of Shared Memory Mgr routines, and are shortcuts for
   accessing the shared memory block containing data global to the office dll
   and all of its clients
   $[VSMSO]: PHarring They don't actually use shared memory any more
****************************************************************************/

/*	Initializing the Shared Globals */
MSOAPIX_(BOOL) MsoSmcInitShrGlobals();

/*	Uninitialize the Shared Globals */
MSOAPI_(void) MsoUninitShrGlobals();

/*	Lock down the "Globals" global pointer */
MSOAPI_(BOOL) MsoFLockShrGlobals(MSOSG **ppsg);

/*	Release the "Globals" global pointer */
MSOAPI_(void)  MsoUnlockShrGlobals(MSOSG **ppsg);

// For apps using Office Shared Globals structure
MSOAPI_(MSOSG *) MsoInitShrGlobal(BOOL fApp);


/****************************************************************************
   The MSOPG structure contains a handful of the per-process globals for
   the Office DLL.  We export them just because we think maybe someone
   in our process will find them useful.

   If you add or remove an item from this structure, keep it in line with
   the vrgsysbrush global used for init.
******************************************************************* RICKP **/
typedef struct 
{
	// vrgsysbrush
	HBRUSH hbrBtnface;
	HBRUSH hbrWindowframe;
	HBRUSH hbrBtnshadow;
	HBRUSH hbrBtnhighlight;
	HBRUSH hbrHighlight;
	HBRUSH hbrBtnText;
	HBRUSH hbrWindow;
	HBRUSH hbrWindowText;
	HBRUSH hbrHighlightText;
	HBRUSH hbrBtnDarken;
	HBRUSH hbrDither;
	HBRUSH hbrMenu;
	HBRUSH hbrMenuText;
	HBRUSH hbrInfoBk;
	HBRUSH hbr3DLight;
	HBRUSH hbr3DFace;
	HBRUSH hbr3DShadow;
	HBRUSH hbr3DDkShadow;

	HCURSOR hcrsArrow;
	HCURSOR hcrsHourglass;
	HCURSOR hcrsHelp;

	HBITMAP hbmpDither;
	HBITMAP hbmpDither2;
} MSOUG;

/*	Lock down the "Process" global pointer */
MSOAPIXX_(BOOL) MsoFLockGlobals(MSOUG **ppug);

/*	Release the "Process" global pointer */
MSOAPIXX_(void)  MsoUnlockGlobals(MSOUG **ppug);



/*****************************************************************************
	 Defines the IMsoShMemory interface.
	 
	 The IMsoShMemory interface is used internally by the shared memory 
	 implementation for both backward and forward compatibity for shared memory.
	 It is declared here so that future version of Office will have access to
	 this interface.
******************************************************************* JIMMUR ***/
#ifndef MSO_NO_INTERFACES
#undef  INTERFACE
#define INTERFACE   IMsoShMemory

DECLARE_INTERFACE_(IMsoShMemory, IUnknown)
{
	// IUnknown methods
	MSOMETHOD(QueryInterface)(THIS_ REFIID riid, void **ppvObj) PURE;
	MSOMETHOD_(ULONG, AddRef)(THIS) PURE;
	MSOMETHOD_(ULONG, Release)(THIS) PURE;
	
	//IMsoShMemory methods
	
	// Return the version of Shared memeory this interface supports
	MSOMETHOD_(DWORD, DwGetVersion)(THIS) PURE;
	
	// Get the version of the shared memory manager that is responcible for
	// the piece of share memory specified by the cb,dg, & dwtag.
	MSOMETHOD_(DWORD, DwGetVersionForMem)(THIS_ int cb, int dg, DWORD dwtag) PURE;
	
	//Given a pointer, return whether or not this interface manages that memory
	MSOMETHOD_(BOOL, FIsMySharedMem)(THIS_ void* pv) PURE;
	
	//PvShAlloc: Allocate into shared memory
	MSOMETHOD_(void*, PvShAlloc)(THIS_ int cb, int dg, DWORD dwtag) PURE;
	
	//ShFreePv: Delete an object from shared memory
	MSOMETHOD_(void, ShFreePv) (THIS_ void* pv) PURE;
};


#undef  INTERFACE
#define INTERFACE   IMsoMallocEx

DECLARE_INTERFACE_(IMsoMallocEx, IMalloc)
{
	// IUnknown methods
	MSOMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObj) PURE;
	MSOMETHOD_(ULONG, AddRef)(THIS) PURE;
	MSOMETHOD_(ULONG, Release)(THIS) PURE;

	// IMalloc methods
	MSOMETHOD_(void*, Alloc)(THIS_ ULONG cb);
	MSOMETHOD_(void*, Realloc)(THIS_ void* pv, ULONG cb);
	MSOMETHOD_(void,  Free)(THIS_ void* pv);
	MSOMETHOD_(SIZE_T, GetSize)(THIS_ void* pv);
	MSOMETHOD_(int,   DidAlloc)(THIS_ void* pv);
	MSOMETHOD_(void,  HeapMinimize)(THIS);

	// IMalloc exteded methods
	MSOMETHOD_(void*, Lock) (THIS_ void* pv);
	MSOMETHOD_(int,   Unlock) (THIS_ void* pv);
};

MSOAPI_(HRESULT) MsoHrGetMsoMallocEx(int fGlobalAlloc, UINT uFlags, UINT uFlagsRe, IMsoMallocEx** ppIMalloc);
#endif // MSO_NO_INTERFACES


///////////////////////////////////////////////////////////////////////////////
#if defined(DEBUG)
//  Heap Marking tracking list that is used for allocations that we are not
//  able to enumerate.  Instead, they are tracked in this list structure and
//  then we bulk mark them when doing memory checks.  This structure is
//  stored at the beginning of the memory block allocated to the routine
//  and therefore is protected by additional sentinal marker.
//  Examples of use for this are LineServices allocation callbacks and GDI+
//  allocation callbacks.
typedef struct _MSOHM {
	struct _MSOHM *phmNext;
	struct _MSOHM *phmPrev;
	DWORD  dwTest;			//  We use msomfSentinel value here
	} MSOHM;
	
#define dwTestHm 0xDeadBeefL

//  Note that we only specify the HM on alloc as the location to insert
//  the link into.  All other routines maintain the curcular link by
//  referencing off of the existing HM in the allocation passed.
#define MsoPvAllocHm(cb, dg, phm) MsoPvAllocHmCore(cb, dg, phm, __FILE__, __LINE__)
MSOAPI_(void *) MsoPvAllocHmCore(int cb, int dg, MSOHM *phm, const CHAR *szFile, int li);
MSOAPI_(void) MsoFreePvHm(void* pv);
MSOAPI_(void*) MsoPvReallocHm(void* pv, int cb);
//MSOAPI_(long) MsoCbSizePvHm(void* pv);
MSOAPI_(int) MsoFMarkPhm(MSOHM *phm, HMSOINST hInst, LPARAM lParam, int fAllocHasSize, int bt);
MSOAPI_(int) MsoFInitPhm(MSOHM *phm);

#else

#define MsoPvAllocHm(cb, dg, phm) MsoPvAllocCore(cb, dg)
#define MsoFreePvHm(pv) MsoFreePv(pv)
#define MsoPvReallocHm(pv, cb) MsoPvRealloc(pv, cb)
//#define MsoCbSizePvHm(pv) MsoCbSizPv(pv)

#endif
///////////////////////////////////////////////////////////////////////////////


#ifdef DEBUG

typedef void (MSOSTDAPICALLTYPE* MSOPFNMEMCHECK) (BOOL fEnableChecks);

///////////////////////////////////////////////////////////////////////////////
// MsoRegisterPfnMemCheck
//
// Register a global memory checking enable/disable function.  This function
// will be used to disable app-specific memory checking around function which
// are known to allocate (and not free) out of the task heap.

MSOAPI_(HRESULT) MsoRegisterPfnMemCheck(MSOPFNMEMCHECK pfnMemCheck);

// Call such a function, if registered.
MSOAPI_(HRESULT) MsoDoMemCheckHandler(BOOL fEnableChecks);

#else

#define MsoDoMemCheckHandler(f) (S_OK)

#endif


/*-----------------------------------------------------------------------------
	HGAM (Handle to Gang Alloc Memory)
-------------------------------------------------------------------- HAILIU -*/
typedef HANDLE MSOHGAM;

/*-----------------------------------------------------------------------------
	MsoHgamCreate
-------------------------------------------------------------------- HAILIU -*/
MSOAPI_(MSOHGAM) MsoHgamCreate(int cpv, int cb, int dg);

/*-----------------------------------------------------------------------------
	MsoHgamFree
-------------------------------------------------------------------- HAILIU -*/
MSOAPI_(VOID) MsoHgamFree(MSOHGAM hgam);

/*-----------------------------------------------------------------------------
	MsoPvNextOfHgam
-------------------------------------------------------------------- HAILIU -*/
MSOAPI_(PVOID) MsoPvNextOfHgam(MSOHGAM hgam);


#ifdef DECAL
#ifdef __cplusplus
template <class T>
class MSOTSTACK
{
public:
	MSOTSTACK()	{}


	// Initializes the plex
	BOOL FInit(int dAlloc, int iMax, int dg)
		{
		return m_pxt.FInit(dAlloc, iMax, dg);
		}


	// Pushes t onto the stack
	BOOL FPush(T t)
		{
		if (!FValid())
			{
			return FALSE;
			}

		return m_pxt.FAppend(&t);
		}


	// Pops top object in stack into t
	BOOL FPop(T& t)
		{
		if (!FValid() || m_pxt.IMac() == 0)
			{
			return FALSE;
			}

		int iTop = m_pxt.IMac() - 1;
		t = m_pxt[iTop];
		return m_pxt.FSetIMac(iTop);
		}


	// Copies top object in stack into t
	BOOL FPeek(T& t)
		{
		if (!FValid() || m_pxt.IMac() == 0)
			{
			return FALSE;
			}

		t = m_pxt[m_pxt.IMac() - 1];
		return TRUE;
		}


	// Count of items on the stack
	inline int CSize() { return m_pxt.IMac(); }

	// Flushes the stack
	inline void Empty() { m_pxt.Empty(); }

	// Is there anything on the stack?
	inline BOOL FEmpty() { return m_pxt.IMac() == 0; }

	// Is the stack valid? (Did FInit succeed?)
	inline BOOL FValid() { return m_pxt.FValid(); }

#ifdef DEBUG
	BOOL FWriteBe(LPARAM lParam, BOOL fSaveObj)
		{
		if (fSaveObj)
			{
			if (!MsoFSaveBe(NULL, lParam, TRUE, this, sizeof(*this), btMSOTStack))
				{
				return FALSE;
				}
			}

		m_pxt.FWriteBe(lParam, FALSE);
		}
#endif // DEBUG

private:
	// Disallow assignment the implications of copying a plex blow my mind
	MSOTSTACK(const MSOTSTACK&);
	MSOTSTACK& operator=(const MSOTSTACK&);

	MSOTPX<T> m_pxt;
};

#endif  //__cplusplus
#endif	//DECAL


#endif /* MSOALLOC_H */
