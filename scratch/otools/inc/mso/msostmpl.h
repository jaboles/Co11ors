#pragma once

/*------------------------------------------------------------------------*
 * msostmpl.h (previously known as sdmtmpl.h): SDM PUBLIC include file    *
 *  -- SDM Template info.                                                 *
 *      No application code should rely on the contents of this file!!    *
 *                                                                        *
 * Please do not modify or check in this file without contacting NicoleP. *
 *------------------------------------------------------------------------*/

#ifndef SDMTMPL_H	// Entire file.
#define SDMTMPL_H

/*****************************************************************************/
/* TMTs */

#define tmtEnd			((TMT)0)

#define tmtNormalMin	((TMT)1)

#define tmtStaticText	((TMT)1)	/* Items with a title */
#define tmtTooltipMin   ((TMT)2)
#define tmtPushButton	((TMT)2)
#define tmtCheckBox		((TMT)3)
#define tmtRadioButton	((TMT)4)
#define tmtGroupBox		((TMT)5)
#define tmtGroupLine	((TMT)6)
#define	tmtWithTitleMax	((TMT)7)

#define tmtEdit			 ((TMT)7)	/* Items without a title */
#define tmtFormattedText ((TMT)8)
#define tmtListBox		 ((TMT)9)
#define	tmtDropList		 ((TMT)10)
#define tmtTooltipMax     ((TMT)10)

#define tmtBitmap		  ((TMT)11)
#define tmtGeneralPicture ((TMT)12)

#define tmtScroll	((TMT)13)

#define tmtUserMin		((TMT)14)	/* User defined controls */
#define tmtUserMax		((TMT)50)

#define tmtNormalMax		((TMT)50)

#define tmtExtensionMin		((TMT)50)	/* Extensions are still special */

#define tmtConditional		((TMT)50)
#define tmtExtDummyText		((TMT)51)
#define tmtExtItemProc		((TMT)52)
#define tmtExt1				((TMT)53)	/* tmtImportTmc */
#define tmtExt2				((TMT)54)	/* tmtExtCtrlProc */
#define tmtExt3				((TMT)55)	/* reading order (for BIDI) */
#define tmtExt4				((TMT)56)

#define	tmtExtUserMin		((TMT)57)	/* Control specific extensions */
#define tmtExtUserMax		((TMT)64)

#define tmtExtensionMax		((TMT)64)

#define tmtMax			((TMT)64)	/* 6 bits */



///////////////////////////////////////////////////////////////////////////////
// ALT : Text alignment for StaticText and EditItems.
typedef WORD	ALT;

#define altLeft			((ALT)ftmsLeft)
#define altCenter		((ALT)ftmsCenter)
#define altRight		((ALT)ftmsRight)

#define FtmsOfAlt(alt)	((FTMS)(alt) & ftmsAlt)
#define	AltOfFtms(ftms)	((ALT)((ftms) & ftmsAlt))

///////////////////////////////////////////////////////////////////////////////
// FOST : StaticText options.
typedef	WORD	FOST;
#define fostNoAccel		((FOST)1)	// No acclerator.
#define fostMultiLine		((FOST)2)
#define fostLightFont		((FOST)4)


///////////////////////////////////////////////////////////////////////////////
// FOEI : EditItem options.
typedef	WORD	FOEI;
#define foeiCharValidated	((FOEI)1)
#define foeiMultiLine		((FOEI)2)
#define foeiVerticalScroll	((FOEI)4)
#define foeiSpin			((FOEI)4)

#define foei2RefEdit			((FOEI)1)


///////////////////////////////////////////////////////////////////////////////
// FOLB : ListBox options.
typedef WORD	FOLB;

#define	folbComboAtomic		((FOLB)0x0001)	// Atomic Combo ListBox.
#define	folbString		((FOLB)0x0002)	// ListBox returns string.
#define	folbSorted		((FOLB)0x0004)	// Sorted ListBox.
#define	folbMultiSelectable	((FOLB)0x0008)	// MutliSelectable ListBox.
#define	folbNoScrollBar		((FOLB)0x0010)	// No vertical scroll bar.
#define	folbDropDownSibling	((FOLB)0x0020)	// Sibling DropDown.
#define	folbDropAtomic		((FOLB)0x0040)	// Atomic Combo DropDown
						// ListBox.
#define folbSiblingAtomic	((FOLB)0x0080)	// Atomic Combo Sibling
						// DropDown ListBox.
#define folbExtended		((FOLB)0x0100)	// Send tmmCreate message.
#define folbExtSelect		((FOLB)0x0200)	// Extended-selection lbox
#define folbCheckSelect		((FOLB)0x0800)	// Checkable multi-selection lbox
#define folbGalleryControl  ((FOLB)0x1000)     // Gallery Control


///////////////////////////////////////////////////////////////////////////////
// FOSB : ScrollBar options.
typedef WORD	FOSB;

#define	fosbNoFocus			((FOSB)0x0001)	// Can't receive focus



/*****************************************************************************/
/* -- defined Dialog Item */

typedef struct _tm
	{
	BIT_SDM	tmt:6;		/* item type */
	BIT_SDM	fThinBdr:1;
	BIT_SDM	fAction:1;
	BIT_SDM	bitSlot:8;
	WORD	wSlot;		/* based pointer or wParam */
	LONG_PTR l;		/* compact rectangle or PFN */
/* FMidEast */	

/* FMidEast End */
	} TM;

/* Compact Rectangle */
typedef struct _crc
	{
	BYTE	x, y, dx, dy;		/* order critical */
	} CRC;		/* Compact Rectangle */

/*****************************************************************************/
/* BDR */
typedef	WORD BDR;

/* Note: the order and contiguity of the following 5 #defines is assumed. */
/*	Do not change.  PRB */
#define bdrNone		((BDR) 0x0000)
#define bdrThin		((BDR) 0x0001)
#define bdrThick	((BDR) 0x0002)
#define bdrCaption	((BDR) 0x0003)	/* dialog has caption */
#define bdrSysMenu	((BDR) 0x0004)	/* system menu - implies caption. */
#define bdrGoAwayBox	((BDR) 0x0004)	/* for the Mac - implies caption. */
#define bdrMask		((BDR) 0x0007)	/* actual border bits. */
#define bdrAutoPosX	((BDR) 0x0008)	/* Auto position along x-axis. */
#define bdrAutoPosY	((BDR) 0x0010)	/* Auto position along y-axis. */
#define bdrImeCtrlOn	((BDR) 0x0020)	//  FE IME option
#define bdrImeCtrlOff	((BDR) 0x0040)	//  FE IME option
#define bdrUnicode	((BDR) 0x0080)	//  Template is in Unicode form already.
#define bdrResizable ((BDR) 0x0100) // This is a resizable dialogue
#define bdrRTL        ((BDR) 0x0200) // This dialog is should look RTL
#define bdrDal ((BDR) 0x1000) // This is an auto laid out dialog


/*****************************************************************************/

typedef struct _dlgtheader
	{
	REC	rec;			/* rectangle of dialog */
	unsigned hid;		/* Help ID. */
	TMC	tmcSelInit;		/* Item with initial selection */
	PFN_DIALOG pfnDlg;	/* DialogProc.  NULL for none. */
	unsigned ctmBase;	/* # of base TMs */
	BDR	bdr;			/* Border bits. */
	WORD bpTitle;		/* based pointer to title */
	} DLGTHEADER;

/* dialog template (or template header) */
typedef struct _dlgt
	{
	REC	rec;		/* rectangle of dialog */
	unsigned hid;		/* Help ID. */
	TMC	tmcSelInit;	/* Item with initial selection */
	PFN_DIALOG pfnDlg;	/* DialogProc.  NULL for none. */
	unsigned ctmBase;	/* # of base TMs */
	BDR	bdr;		/* Border bits. */

	/* title string : based pointer */
	WORD	bpTitle;			/* based pointer to title */
	} DLGT;	/* Dialog Template */

typedef DLGT * * HDLGT;	// A near handle. 

#ifdef __cplusplus
struct DLGTW : public DLGT
{
	TM	rgtm[1];		/* array starts here */	
};
#else
typedef struct _dlgtw
	{
	union
		{
		DLGT;
		};
	TM	rgtm[1];		/* array starts here */	
	} DLGTW;

#endif // __cplusplus
typedef DLGTW * * HDLGTW;

/*****************************************************************************/
/* misc equates */

#define	bpStringFromCab		((WORD) 1)

/*****************************************************************************/

#endif	//!SDMTMPL_H	Entire file.
