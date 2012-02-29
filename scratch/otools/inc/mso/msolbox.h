#pragma once

/*------------------------------------------------------------------------*
 * msolbox.h (previously known as lbox.h): SDM PUBLIC include file        *
 * -- LBOX definitions.                                                   *
 *                                                                        *
 * Please do not modify or check in this file without contacting MsoSdmQs.*
 *------------------------------------------------------------------------*/

#define PYRAMID_KLUDGE_MEM

//////////////////////////////////////////////////////////////////////////
// The LBOX Public Header File						//
//////////////////////////////////////////////////////////////////////////
#ifndef	LBOX_INCLUDED	// Entire file.
#define	LBOX_INCLUDED

#ifndef	SDM_TYPES_DEFINED
#define	SDM_TYPES_DEFINED
typedef	unsigned char	U8_SDM;
typedef	char		S8_SDM;
typedef	unsigned short	U16_SDM;
typedef	short		S16_SDM;

// AKadatch: all these types hold pointers
// typedef	long		S32_SDM;
// typedef	unsigned long	U32_SDM;
// typedef int INT_SDM;
// typedef unsigned int UINT_SDM;
// typedef	unsigned long	BARG_SDM;


// AKadatch: some apps use obsolete and/or modified WinNT.h
// *_PTR types should be declared manually
#include <basetsd.h>

typedef	LONG_PTR	S32_SDM;
typedef	ULONG_PTR	U32_SDM;
typedef INT_PTR		INT_SDM;
typedef UINT_PTR	UINT_SDM;
typedef	ULONG_PTR	BARG_SDM;

#endif	//SDM_TYPES_DEFINED
///////////////////////////////////////////////////////////////////////////////
//	Function storage class.

#define LBOX_NATIVE	NATIVE

#ifndef EXPORT
#define EXPORT
#endif //EXPORT

//////////////////////////////////////////////////////////////////////////
// An LBOX window							//
//////////////////////////////////////////////////////////////////////////

#ifndef	SDM_INCLUDED
typedef	RECT	LBR;
#endif	//!SDM_INCLUDED

//////////////////////////////////////////////////////////////////////////
// An LBOX								//
//////////////////////////////////////////////////////////////////////////

#ifndef	SDM_INCLUDED
typedef	struct	_LBX * * HLBX;
#endif	//!SDM_INCLUDED

#if !defined(WIN32)
#define	hlbxNil		((HLBX)NULL)
#else // !WIN32
#define	hlbxNil		(NULL)
#endif // WIN32


//////////////////////////////////////////////////////////////////////////
// LBOX Comparison Result						//
//////////////////////////////////////////////////////////////////////////

#ifndef	SDM_INCLUDED
typedef	UINT_SDM	LBC;
#endif	//!SDM_INCLUDED

// don't cast these, since we use them from MASM!!
#define	lbcEq		0
#define	lbcPrefix	1
#define	lbcLt		2
#define	lbcGt		3
#define lbcError	4	// In case of OOM during search.

//////////////////////////////////////////////////////////////////////////
// Modifier key states							//
//////////////////////////////////////////////////////////////////////////
typedef	UINT_SDM	FKEY_LBOX;

#define	fkeyNil		((FKEY_LBOX)0x0000)
#define	fkeyShift	((FKEY_LBOX)0x0001)
#define	fkeyControl 	((FKEY_LBOX)0x0002)
#define	fkeyOption 	fkeyControl
#define fkeyAlt		((FKEY_LBOX)0x0004)

//////////////////////////////////////////////////////////////////////////
// LBOX List Entry display state					//
//////////////////////////////////////////////////////////////////////////

typedef	UINT_SDM	FLBE;
// NOTE: if flbe becomes more than 16 bits stuff like dlmChangeCheckBox will break.
#define	flbeNil		((FLBE)0x0000)
#define	flbeSelected	((FLBE)0x0001)
#define	flbeEnabled	((FLBE)0x0002)
#define	flbeCursor	((FLBE)0x0004)
#define flbeCheck	((FLBE)0x0008)
#define flbeGray ((FLBE)0x0040)
// This entry has an "MRU" underline beneath it
#define flbeMruLine				((FLBE)0x0010) // Used only for Gallery controls
#define flbeDropdownDropped	((FLBE)0x0020) // Used only for Gallery controls
#define flbeOverDropdown		((FLBE)0x0080) // Used only for Gallery controls
#define flbeSmall ((FLBE)0x0100) // drawing collapsed dropdown
// This is the last entry in "MRU" list and might precede blank filler entries.
#define flbeLastBeforeMru		((FLBE)0x0200) // Used only for Gallery controls
#define flbeCursorHand			((FLBE)0x0400) // makes the cursor a hand over this entry.


//////////////////////////////////////////////////////////////////////////
// LBOX List Styles							//
//////////////////////////////////////////////////////////////////////////
typedef	UINT_SDM	FLBX;

//FUTURE: (KBrown): Rename all flbxNonClientBorder references to flbxBorder.

#define flbxNil             ((FLBX)0x00000000)
#define flbxOwnDC           ((FLBX)0x00000000) // Not used
#define flbxBorder          ((FLBX)0x00000001)
#define flbxCheckSelect     ((FLBX)0x00000002)
#define flbxScroll          ((FLBX)0x00000008)
#define flbxNotify          ((FLBX)0x00000010)
#define flbxNoSelTrack      ((FLBX)0x00000020)
#define flbxNoSelNotify     ((FLBX)0x00000040)
#define flbxNoSelFocus      ((FLBX)0x00000080)
#define flbxMulti           ((FLBX)0x00000100)
#define flbxSort            ((FLBX)0x00000200)
#define flbxNoFocus         ((FLBX)0x00000400)
#define flbxShowScroll      ((FLBX)0x00000800)
#define flbxRTL             ((FLBX)0x00001000) // Both alignment and reading-order.
#define flbxExtSelect       ((FLBX)0x00002000)
#define flbxFreeScroll      ((FLBX)0x00004000) // Leave scrolling alone during lbx operations.  Namely UpdateLbx()
#define flbxDropdown        ((FLBX)0x00008000) // Win only.
#define flbxOwnerDrag       ((FLBX)0x00010000) //
#define flbxClickCheckbox   ((FLBX)0x00040000) // only clicks inside the checkbox actually check it
                                               // for use with (flbxCheckSelect & flbxExtSelect)
#define flbxGallery         ((FLBX)0x00080000) // Draw as a Gallery control
#define flbxSmallOvrd       ((FLBX)0x00100000) // use lbg.dy default always

#define flbxNonClientBorder flbxBorder 
#define flbxExtNormal ((FLBX)0x200000) // extended listbox has normal string at least at the beginning of ppvDisp
#define FTestLbgFlbx(plbg, flbxT) ((plbg)->flbx & (flbxT))

//////////////////////////////////////////////////////////////////////////
// LBOX Calllback Prototypes						//
//////////////////////////////////////////////////////////////////////////

typedef VOID (LBOX_CALLBACK * PFNPNT)(HLBX, UINT_SDM, VOID * *, BOOL_SDM, HDC,
		LBR *, FLBE, FLBE, FLBX);

#define	pfnpntNull	((PFNPNT)NULL)

typedef ILBE_SDM (LBOX_CALLBACK * PFNCHAR)(HLBX, ILBE_SDM, UINT_SDM, BOOL_SDM *, UINT_SDM, FKEY_LBOX,
		 BOOL_SDM *);
#define	pfncharNull	((PFNCHAR)NULL)

typedef BOOL_SDM (LBOX_CALLBACK * PFNFREE)(HLBX, UINT_SDM, BOOL_SDM);
#define	pfnfreeNull	((PFNFREE)NULL)

#ifdef DEBUG
typedef BOOL_SDM (LBOX_CALLBACK * PFNFSAVEDATABE)(HLBX, UINT_SDM, LPARAM);
#define	pfnfsavedatabeNull	((PFNFSAVEDATABE)NULL)
#endif

typedef VOID (LBOX_CALLBACK * PFNFREELBE)(UINT_SDM, ILBE_SDM, VOID * *);
#define	pfnfreelbeNull	((PFNFREELBE)NULL)

#ifdef DEBUG
typedef BOOL (LBOX_CALLBACK * PFNFSAVELBEBE)(UINT_SDM, ILBE_SDM, VOID**, LPARAM);
#define pfnfsavelbebeNull ((PFNFSAVELBEBE)NULL)
#endif

typedef VOID * *  (LBOX_CALLBACK_NAT * LPFNDMND)(HLBX, ILBE_SDM, UINT_SDM);
#define	lpfndmndNull	((LPFNDMND)NULL)


//////////////////////////////////////////////////////////////////////////
// LBOX Initializer							//
//////////////////////////////////////////////////////////////////////////
typedef struct _lbi
	{
	SZ	szClass;
	FARPROC	*plpfn;
	UINT_SDM	*pcb;
	HANDLE	hinst;
	HANDLE	hinstPrev;
	HCURSOR	hcursor;
	SZ	szScrollClass;
	LPFNCOMP lpfncomp;
	}LBI;



//////////////////////////////////////////////////////////////////////////
// An LBOX List Entry							//
//////////////////////////////////////////////////////////////////////////

typedef	struct
	{
	VOID **ppvDisp; // handle to data
	FLBE   flbe;    // Only use a byte to store it.
	} LBE;

#define	ilbeNil		((ILBE_SDM)(-1))

//////////////////////////////////////////////////////////////////////////
// LBOX List Global Attributes 						//
//////////////////////////////////////////////////////////////////////////

typedef	struct _lbg
	{
	PFNPNT		pfnPaintLbx;		// default paint callback
	LPFNCOMP	lpfnLbcCompareLbx;	// compare entries callback
	PFNCHAR		pfnIlbeProcChLbx;	// handle char input callback
	LPFNDMND	lpfnPpvDemandLbx;	// demand an entry callback
	PFNFREE		pfnFreeData;		// free data callback
	PFNFREELBE	pfnFreeLbe;		// free an entry's data callback
#ifdef DEBUG
	PFNFSAVEDATABE pfnFSaveDataBe;	// debug memory accounting
	PFNFSAVELBEBE	pfnFSaveLbeBe;	// debug memory accounting
#endif
	FLBX		flbx;			// List styles
	XY_SDM		dy;			// default height of an entry

	HWND		hwndNotify;	  	// window to get notification

	ILBE_SDM	cEntryVis;	// number of entries visible in large state
				// (Dropdowns only)
	XY_SDM	dySmall;	// height of collapsed list (Dropdowns only)
	UINT_SDM	wUser;		// application defined dword
						// bad hungarian for hysterical reasons
	DWORD	clrBack;	// background color
	U32_SDM	wUser2;		// another application defined dword
	} LBG;

//////////////////////////////////////////////////////////////////////////
// LBOX Notification messages						//
//////////////////////////////////////////////////////////////////////////

typedef	UINT_SDM	LBN;

#define	lbnSelChange	((LBN)1)
#define	lbnDblClk	((LBN)2)
#define lbnCheck	((LBN)3)
#define lbnSelAgain	((LBN)4)

//////////////////////////////////////////////////////////////////////////
// LBOX Control Manager function declarations				//
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// for communicating with the routines that handle USER INPUT		//
// for LBOX								//
//////////////////////////////////////////////////////////////////////////

typedef UINT_SDM	LBOX_MSG;

#define	lbxmsgBtnDown		(LBOX_MSG) WM_LBUTTONDOWN
#define	lbxmsgBtnDblClk		(LBOX_MSG) WM_LBUTTONDBLCLK
#define	lbxmsgBtnUp		(LBOX_MSG) WM_LBUTTONUP
#define	lbxmsgRtBtnDown		(LBOX_MSG) WM_RBUTTONDOWN
#define	lbxmsgRtBtnUp			(LBOX_MSG) WM_RBUTTONUP
#define	lbxmsgRtBtnDblClk		(LBOX_MSG) WM_RBUTTONDBLCLK
#define	lbxmsgMouseMove		(LBOX_MSG) WM_MOUSEMOVE
#define lbxmsgSetCursor		(LBOX_MSG) WM_SETCURSOR
#define	lbxmsgMouseLeave		(LBOX_MSG) WM_MOUSELEAVE
#define	lbxmsgTimer		(LBOX_MSG) WM_TIMER
#define	lbxmsgSysKeyUp		(LBOX_MSG) WM_SYSKEYUP
#define	lbxmsgKeyUp		(LBOX_MSG) WM_KEYUP
#define	lbxmsgKeyDown		(LBOX_MSG) WM_KEYDOWN
#define	lbxmsgChar		(LBOX_MSG) WM_CHAR
#define	lbxmsgSbLineUp		(LBOX_MSG) SB_LINEUP
#define	lbxmsgSbLineDown	(LBOX_MSG) SB_LINEDOWN
#define	lbxmsgSbPageUp		(LBOX_MSG) SB_PAGEUP
#define	lbxmsgSbPageDown	(LBOX_MSG) SB_PAGEDOWN
#define	lbxmsgSbPosition	(LBOX_MSG) SB_THUMBPOSITION
#define	lbxmsgSbDrag		(LBOX_MSG) SB_THUMBTRACK
#define	lbxmsgSbTop		(LBOX_MSG) SB_TOP
#define	lbxmsgSbBottom		(LBOX_MSG) SB_BOTTOM
#define	lbxmsgSbEnd		(LBOX_MSG) SB_ENDSCROLL

MSOAPIX_(BOOL_SDM)	FInitLbox(LBI *);

MSOAPIX_(BOOL_SDM)	FIsPtInLbx(HLBX, PT_SDM *);

MSOAPI_(VOID)	SetCEntryVisLbx(HLBX, ILBE_SDM);
MSOAPIX_(ILBE_SDM)	CEntryVisLbx(HLBX);
MSOAPI_(U32_SDM)	WUser2Lbx(HLBX);

MSOAPI_(VOID)   UpdateLbxRespectRedraw(HLBX, ILBE_SDM, ILBE_SDM);
#define UpdateLbxTmcRespectRedraw(tmc, ilbeFirst, ilbeMin) UpdateLbxRespectRedraw(HlbxFromTmc(tmc), ilbeFirst, ilbeMin)

MSOAPI_(VOID)	UpdateLbx(HLBX, ILBE_SDM, ILBE_SDM);
MSOAPIX_(VOID) EnsureScrollEnable(HLBX);
MSOAPI_(BOOL_SDM)	FSetFirstLbx(HLBX, ILBE_SDM);
MSOAPI_(ILBE_SDM)	IlbeGetFirstLbx(HLBX);
MSOAPIMX_(UINT_SDM)	WRenderLbx(HLBX);
MSOAPIX_(VOID)	EnableLbx(HLBX, BOOL_SDM);
MSOAPI_(VOID)	ShowLbx(HLBX, BOOL_SDM);
MSOAPIX_(VOID)	SetFocusLbx(HLBX, BOOL_SDM);
MSOAPIX_(void) GetRcForLbCursor(HLBX, RECT *);
MSOAPIX_(BOOL_SDM) ProcessMouseTimerLbx(HLBX, LBOX_MSG, BARG_SDM,
				FKEY_LBOX);
MSOAPIX_(VOID)	ProcessKeybdLbx(HLBX, LBOX_MSG, UINT_SDM, FKEY_LBOX);
MSOAPIX_(VOID)	ProcessScrollBarLbx(HLBX, LBOX_MSG, UINT_SDM);
MSOAPI_(BOOL_SDM) MsoFGetLboxInfo(HDLG hdlg, TMC tmc, int *pdy, 
                   int *piFirst, int *piSel, int *piFocus, int *pcVisible);
MSOAPI_(int) MsoIGetLboxItem(HDLG hdlg, TMC tmc, POINT pt, 
                                                           RECT *prcLboxItem);
MSOAPI_(BOOL_SDM) MsoGetPrcForIlbe(HDLG hdlg, TMC tmc, int iLbe, 
                                                           RECT *prcLboxItem);                                                           
#if SDMMSAA_TROBWIS_UPDATE
MSOAPI_(BOOL) MsoFSelectLboxItem(TMC tmc, BOOL fClear, LONG lChild, HDLG hdlg);
#endif 

#define	FMouseDownLbx(hlbx)	((*hlbx)->lbd.lbt.flbt & flbtMouseDown)
//////////////////////////////////////////////////////////////////////////
// Callbacks to get the env-specific info about a list			//
//////////////////////////////////////////////////////////////////////////
VOID LBOX_CALLBACK_NAT	GetLbxLbrClient(UINT_SDM, LBR *, BOOL_SDM, BOOL_SDM);
VOID LBOX_CALLBACK_NAT	GetLbxLbrScroll(UINT_SDM, LBR *, BOOL_SDM);
BOOL_SDM LBOX_CALLBACK_NAT	FIsEnabledLbx(UINT_SDM);
BOOL_SDM LBOX_CALLBACK_NAT	FIsVisibleLbx(UINT_SDM);
HLBX LBOX_CALLBACK_NAT	HlbxGetLbx(UINT_SDM);
VOID LBOX_CALLBACK_NAT	UpdateWndLbx(UINT_SDM);

HWND LBOX_CALLBACK_NAT	HwndGetLbx(UINT_SDM);
HWND LBOX_CALLBACK_NAT	HwndGetLbxNotify(UINT_SDM);
HDC LBOX_CALLBACK_NAT	HpdcGetLbx(UINT_SDM);
VOID LBOX_CALLBACK_NAT	ReleaseLbxHpdc(UINT_SDM, HWND, HDC);


//////////////////////////////////////////////////////////////////////////
// Required Callbacks to accept certain LBOX feedback messages,	which	//
//	are needed since there is no SendMessage on the Mac.		//
//////////////////////////////////////////////////////////////////////////
//Listbox notification handler.

VOID LBOX_CALLBACK_NAT	LboxNotifyHandler(UINT_SDM, LBN);

//Listbox Timer Interface
// Timer message support.  Since the user of LBOX is handling the message
// pump, they are the ones who should pass us back timer messages.
VOID LBOX_CALLBACK_NAT	LboxSetTimerLbx(UINT_SDM,  UINT_SDM);
VOID LBOX_CALLBACK_NAT	LboxKillTimerLbx(UINT_SDM);

//////////////////////////////////////////////////////////////////////////
// LBOX List Manager function declarations				//
//////////////////////////////////////////////////////////////////////////

MSOAPI_(HLBX)	HlbxCreateLbg(LBG *, ILBE_SDM);
MSOAPI_(VOID)	SetLbxLbg(HLBX, LBG *);
MSOAPI_(VOID)   SetLbxLbeFlbe(HLBX, ILBE_SDM, FLBE);
MSOAPI_(FLBE)   FlbeLbxLbe(HLBX, ILBE_SDM);
MSOAPIX_(VOID)	DestroyLbx(HLBX);
MSOAPIX_(VOID)	DestroyScrollLbx(HLBX);
#ifdef DEBUG
MSOAPIX_(BOOL_SDM) FSaveLbxBe(HLBX hlbx, LPARAM lParam);
#endif

MSOAPI_(VOID)	GetLbxLbg(HLBX, LBG *);

// Provided for TestWizard purposes only.  Please contact MsoSdmQs if you need
// this info.
MSOAPI_(VOID) GetLbxLbd(HLBX hlbx, void *plbd);

MSOAPI_(VOID)	DeleteLbxLbe(HLBX, ILBE_SDM, ILBE_SDM);
MSOAPI_(BOOL_SDM)	FSetLbxLbe(HLBX, LBE * *, ILBE_SDM * *);
MSOAPI_(BOOL_SDM)	FGetLbxLbe(HLBX, LBE * *, ILBE_SDM * *,
				    BOOL_SDM);
MSOAPI_(ILBE_SDM)	ClbeGetLbx(HLBX);
MSOAPI_(ILBE_SDM) IlbeGetCursorLbx(HLBX);

//Direct access to the ppvDisp
MSOAPI_(VOID**) PpvLbxIlbe(HLBX, ILBE_SDM);
#define PpvLbxTmcIlbe(tmc, ilbe) PpvLbxIlbe(HlbxFromTmc(tmc), (ilbe))

///////////////////////////////////////////////////////////////////////////////
// LBOX memory allocation callbacks.					     //
///////////////////////////////////////////////////////////////////////////////
EXPORT VOID ** LBOX_CALLBACK_NAT PpvAllocLboxCb(UINT_SDM);
VOID LBOX_CALLBACK_NAT FreeLboxPpv(VOID **);
BOOL_SDM LBOX_CALLBACK_NAT FReallocLboxPpv(VOID **, UINT_SDM);
#ifdef DEBUG
EXPORT BOOL LBOX_CALLBACK_NAT FSaveLboxPpvBe(VOID **, LPARAM);
#endif

#ifdef PYRAMID_KLUDGE_MEM
///////////////////////////////////////////////////////////////////////////////
// LBOX FAR memory allocation callbacks.					    //
///////////////////////////////////////////////////////////////////////////////
typedef VOID * * PPV_LBOX;
typedef VOID * LPV_LBOX;
typedef BYTE * PB_LBOX;
PPV_LBOX	LBOX_CALLBACK_NAT	PpvAllocFarLboxCb(UINT_SDM);
VOID		LBOX_CALLBACK_NAT	FreeFarLboxPpv(PPV_LBOX);
EXPORT LPV_LBOX	LBOX_CALLBACK_NAT	LpvLockFarLboxPpv(PPV_LBOX);
EXPORT VOID		LBOX_CALLBACK_NAT	UnlockFarLboxPpv(PPV_LBOX);
EXPORT LPV_LBOX	LBOX_CALLBACK_NAT	LpvOfFarLboxPpv(PPV_LBOX);
PPV_LBOX	LBOX_CALLBACK_NAT	PpvGetFarLboxTmpCopy(PB_LBOX pb, UINT_SDM cb);
VOID		LBOX_CALLBACK_NAT	ReleaseFarLboxTmpCopy(PPV_LBOX);
#endif //PYRAMID_KLUDGE_MEM


///////////////////////////////////////////////////////////////////////////////
// LBOX Out of memory callback function declaration			     //
///////////////////////////////////////////////////////////////////////////////

typedef	UINT_SDM	LBS_LB;

#define	lbsGetHpdc	((LBS_LB)0001)
#define	lbsCreateSb	((LBS_LB)0002)
#define	lbsRegister	((LBS_LB)0003)
#define lbsAllocMem	((LBS_LB)0100)	// Lbox OOM (except for following).

// FRetryLboxError is called from both PCode and NATIVE, so must be EXPORT
BOOL_SDM	PUBLIC	FAR	PASCAL	FRetryLboxError(UINT_SDM, HLBX, LBS_LB);

//////////////////////////////////////////////////////////////////////////
// Flags for FSetSelectedLbx (future use)				//
//////////////////////////////////////////////////////////////////////////

typedef	UINT_SDM	FLBS;
#define	flbsNil		((FLBS)0x0000)
#define	flbsClear	((FLBS)0x0001)
#define	flbsUpdate	((FLBS)0x8412)

MSOAPI_(BOOL_SDM) FSetCursorLbx(HLBX, ILBE_SDM);
MSOAPI_(BOOL_SDM) FSetSelectedLbx(HLBX, ILBE_SDM * *, FLBS);
MSOAPI_(VOID)	GetSelectedLbx(HLBX, ILBE_SDM * *);
MSOAPI_(ILBE_SDM)	ClbeGetSelectedLbx(HLBX);
MSOAPI_(VOID * *) PpvGetLbxLbe(HLBX, ILBE_SDM);
MSOAPI_(VOID * *) PpvDemandLbxLbe(HLBX, ILBE_SDM);
MSOAPIX_(VOID * *) PpvDemandSdmLbx(HLBX, ILBE_SDM, UINT_SDM);

// Might need to call FreeLboxPpv() before assigning a new ppv
MSOAPI_(VOID)	SetLbxLbePpv(HLBX, ILBE_SDM, VOID * *);
MSOAPI_(BOOL_SDM)	FFreeDataLbx(HLBX);
MSOAPI_(BOOL_SDM)	FSetLbxClbe(HLBX, ILBE_SDM);
MSOAPI_(BOOL_SDM)	FAddLbxLbe(HLBX, LBE * *, ILBE_SDM, ILBE_SDM,
				    LBS_LB);
MSOAPIX_(ILBE_SDM)	IlbeSearchLbx(HLBX, ILBE_SDM, VOID * *,
				    LBC *, BOOL_SDM);

#ifdef	DEBUG
MSOAPIX_(BOOL_SDM)	FCheckLbx(HLBX);
#endif	//DEBUG



//////////////////////////////////////////////////////////////////////////
// LBOX Default Callback function declarations				//
//////////////////////////////////////////////////////////////////////////

VOID	LBOX_CALLBACK		PaintDftLbx(HLBX, UINT_SDM, VOID * *, BOOL_SDM,
				    HDC, LBR *, FLBE, FLBE, FLBX);

//	Part of the Gallery default drawing procedures.  Call before and after you draw 
//	the contents of an entry in a Gallery control.  See PaintDftLbx() as an example.
MSOAPI_(void) MsoGalleryCleanOldState(LBR *plbr, HDC hdc, FLBE flbePrev, FLBX flbx);
MSOAPI_(void) MsoGalleryDrawNewState(LBR *plbr, HDC hdc, FLBE flbeNew, FLBX flbx);

ILBE_SDM	LBOX_CALLBACK		IlbeProcChDftLbx(HLBX, ILBE_SDM, UINT_SDM, BOOL_SDM *,
				    UINT_SDM, FKEY_LBOX, BOOL_SDM *);
BOOL_SDM	LBOX_CALLBACK		FFreeDataDftLbx(HLBX, UINT_SDM, BOOL_SDM);
BOOL_SDM	LBOX_CALLBACK		FFreeDataNilLbx(HLBX, UINT_SDM, BOOL_SDM);
VOID	LBOX_CALLBACK		FreeDftLbe(UINT_SDM, ILBE_SDM, VOID * *);
VOID	LBOX_CALLBACK		FreeNilLbe(UINT_SDM, ILBE_SDM, VOID * *);
LBOX_NATIVE LBC LBOX_CALLBACK_NAT	LbcCompareDftLbx(UINT_SDM, VOID * *,
				    VOID * *);
LBOX_NATIVE VOID * * LBOX_CALLBACK_NAT PpvDemandDftLbx(HLBX, ILBE_SDM, UINT_SDM);
#ifdef DEBUG
BOOL_SDM LBOX_CALLBACK FSaveDataBeDft(HLBX, UINT_SDM, LPARAM);
BOOL_SDM LBOX_CALLBACK FSaveLbeBeDft(UINT_SDM, ILBE_SDM, VOID**, LPARAM);
BOOL_SDM LBOX_CALLBACK FSaveLbeBeNil(UINT_SDM, ILBE_SDM, VOID**, LPARAM);
#endif


//////////////////////////////////////////////////////////////////////////
// Alternate LBOX Paint Procedure							//
//
// This paint proc fits the WCHAR text (*hstz is the pointer to text) in the 
// give rectangle (*plbr). If the text is too long, "..." is put at the end.
// Toggle chars can be put in the text to make it display with a combination
// of Bold, Italic, or Underline fonts.
// If the first char of the string is the CENTER_TEXT_CHAR, the string
// will be centred in the rectangle on display.
// The paint proc also reponses to end-of-line char (\n)
//
//////////////////////////////////////////////////////////////////////////
#define BOLD_TOGGLE_CHAR (WCHAR) 1			// turn on/off BOLD
#define ITALIC_TOGGLE_CHAR (WCHAR) 2		// turn on/off ITALIC
#define UNDERLINE_TOGGLE_CHAR (WCHAR) 3		// turn on/off UNDERLINE

#define CENTER_TEXT_CHAR (WCHAR) 4			// if this char is at the beginning of the string
											// the text will be centred in the rect

void LBOX_CALLBACK PaintFormatLbx(
	HLBX		hlbx,
	UINT_SDM	wUser,
	void		**hstz,
	BOOL_SDM	fFullPaint,
	Win(HDC		hdc)
	Mac(GrafPtr	hdc),
	LBR			*plbr,
	FLBE		flbePrev,
	FLBE		flbeNew,
	FLBX		flbx);
	

BOOL LBOX_CALLBACK FRenderTextToFitRect(
	HDC hdcIn, const RECT *prect, const WCHAR *wz, 
	BOOL fRight, BOOL fGrayScale, BOOL fPaint);

//////////////////////////////////////////////////////////////////////////
// Non-windowed lbox creation routines					//
//////////////////////////////////////////////////////////////////////////
MSOAPIX_(BOOL_SDM) FCreateLboxList(HLBX);

MSOAPI_(VOID) ShowScrollLbx(HLBX, BOOL_SDM);
MSOAPIX_(VOID) RepositionScrollLbx(HLBX);
MSOAPI_(BOOL) MsoFFlyoutListBox(TMC tmc, REC rec);



#endif	//!LBOX_INCLUDED
