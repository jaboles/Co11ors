#pragma once

/*------------------------------------------------------------------------*
 * msostab.h (previously known as sdmtab.h): SDM PUBLIC include file      *
 *   -- SDM's neato tab-switcher thingy.                                  *
 *                                                                        *
 * Please do not modify or check in this file without contacting NicoleP. *
 *------------------------------------------------------------------------*/


#ifndef SDMTAB_H	//whole file
#define SDMTAB_H


typedef UINT_SDM FTSW_SDM;

/* tab switcher initialization flags */
#define ftswNil			((FTSW_SDM)0x0000)
#define ftswEnabled		((FTSW_SDM)0x0001)
#define ftswDisabled	((FTSW_SDM)0x0002)
#define ftswOnlyOne		((FTSW_SDM)0x0004)	// only allow selection, others gray
#define ftswDirectAccel	((FTSW_SDM)0x0008)	// tabs have direct accelerators
#define ftswValidate	((FTSW_SDM)0x0010)	// validate the dialog before switching
#define ftswHidSelect	((FTSW_SDM)0x0020)	// dlg's HID changes with tabswitch
#define ftswNoBorder	((FTSW_SDM)0x0040)	// Do not draw tabs/borders when only
											// one tab exists
#define ftswNewFocus	((FTSW_SDM)0x0080)	// Win 95 focus handling semantics
/* FMidEast */
#define ftswRTLReo		((FTSW_SDM)0x8000)
/* FMidEast End */

typedef struct _TSW **HTSW;
#define htswNull	((HTSW)NULL)

#define itswNil	((UINT_SDM)(-1))


#define cRowsMaxTsw 4
#define cRowsAutoTsw (-1)


MSOAPI_(BOOL_SDM)	FSetupTabSwitcher(TMC, UINT_SDM, UINT_SDM, PFN_LISTBOX, UCBK_SDM, FTSW_SDM);
MSOAPI_(VOID)	InvalidateSwitcher(TMC);
MSOAPI_(UINT_SDM)	IselGetSwitcher(TMC);
MSOAPI_(VOID)	SetSwitcherIsel(TMC, UINT_SDM);
MSOAPIMX_(UINT_SDM)	ItswDirectlyAccel(TMC, UINT_SDM);
MSOAPI_(void)	GetTmcRecItsw(TMC, UINT_SDM, REC *);
MSOAPI_(void)	GetTmcSwitcherRec(TMC, REC *);
MSOAPIX_(BOOL) MsoFPageTabEnabled(TMC tmc, UINT_SDM itsw);
MSOAPI_(void) MsoEnablePageTab(TMC tmc, UINT_SDM itsw, BOOL fEnabled);

#define ValGetTmcSwitcher(tmc)			IselGetSwitcher(tmc)
#define SetTmcValSwitcher(tmc, isel)	SetSwitcherIsel(tmc, isel)

#ifndef SDM_NOTSWCBPROTO
UCBK_SDM SDM_CALLBACK WPicTabSwitcher(TMM, SDMP *, UINT_SDM, UCBK_SDM, UCBK_SDM, TMC, UCBK_SDM);
UCBK_SDM SDM_CALLBACK WPicTabSwitcherBorder(TMM, SDMP *, UINT_SDM, UCBK_SDM, UCBK_SDM, TMC, UCBK_SDM);
#endif

#endif //SDMTAB_H	//whole file
