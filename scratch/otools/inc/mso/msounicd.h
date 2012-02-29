#pragma once

/****************************************************************************
	msounicd.h

	Owner: NobuyaH
 	Copyright (c) 1998 Microsoft Corporation

	Unicode-related definitions and declarations.
****************************************************************************/

#ifndef MSOUNICD_H
#define MSOUNICD_H


/****************************************************************************
	Unicode SubRange (USR) definitions
*****************************************************************************/
// All updates in USR layout must be appropriately reflected in kjunicod.cpp

typedef enum UnicodeSubRanges
{
 usrBasicLatin,					//		0		// 0x0020->0x007F
 usrLatin1,						//		1		// 0x00A0->0x00FF
 usrLatinXA,					//		2		// 0x0100->0x017F
 usrLatinXB,					//		3		// 0x0180->0x024F
 usrIPAExtensions,				//		4		// 0x0250->0x02AF
 usrSpacingModLetters,			//		5		// 0x02B0->0x02FF
 usrCombDiacritical,			//		6		// 0x0300->0x036F
 usrBasicGreek,					//		7		// 0x0370->0x03CF
 usrGreekSymbolsCop,			//		8		// 0x03D0->0x03FF
 usrCyrillic,					//		9		// 0x0400->0x04FF
 usrArmenian,					//		10		// 0x0530->0x058F
 usrHebrewXA,					//		11		// 0x0590->0x05CF
 usrBasicHebrew,				//		12		// 0x05D0->0x05FF
 usrBasicArabic,				//		13		// 0x0600->0x0670
 usrArabicX,					//		14		// 0x0671->0x06FF
 usrSyriac,						//		15		// 0x0700->0x074F
 usrThaana,						//		16		// 0x0780->0x07BF
 usrDevangari,					//		17		// 0x0900->0x097F
 usrBengali,					//		18		// 0x0980->0x09FF
 usrGurmukhi,					//		19		// 0x0A00->0x0A7F
 usrGujarati,					//		20		// 0x0A80->0x0AFF
 usrOriya,						//		21		// 0x0B00->0x0B7F
 usrTamil,						//		22		// 0x0B80->0x0BFF
 usrTelugu,						//		23		// 0x0C00->0x0C7F
 usrKannada,					//		24		// 0x0C80->0x0CFF
 usrMalayalam,					//		25		// 0x0D00->0x0D7F
 usrSinhala,					//		26		// 0x0D80->0x0DFF
 usrThai,						//		27		// 0x0E00->0x0E7F
 usrLao,						//		28		// 0x0E80->0x0EFF
 usrTibetan,					//		29		// 0x0F00->0x0FFF
 usrMyanmar,					//		30		// 0x1000->0x109F
 usrBasicGeorgian,				//		31		// 0x10D0->0x10FF
 usrGeorgianExtended,			//		32		// 0x10A0->0x10CF
 usrHangulJamo,					//		33		// 0x1100->0x11FF
 usrEthiopic,					//		34		// 0x1200->0x137F
 usrCherokee,					//		35		// 0x13A0->0x13FF
 usrCanadianAbor,				//		36		// 0x1400->0x167F
 usrOgham,						//		37		// 0x1680->0x169F
 usrRunic,						//		38		// 0x16A0->0x16FF
 usrKhmer,						//		39		// 0x1780->0x17FF
 usrMongolian,					//		40		// 0x1800->0x18AF
 usrLatinExtendedAdd,			//		41		// 0x1E00->0x1EFF
 usrGreekExtended,				//		42		// 0x1F00->0x1FFF
 usrGeneralPunct,				//		43		// 0x2000->0x206F
 usrSuperAndSubscript,			//		44		// 0x2070->0x209F
 usrCurrencySymbols,			//		45		// 0x20A0->0x20CF
 usrCombDiacriticsS,			//		46		// 0x20D0->0x20FF	
 usrLetterlikeSymbols,			//		47		// 0x2100->0x214F	
 usrNumberForms,				//		48		// 0x2150->0x218F	
 usrArrows,						//		49		// 0x2190->0x21FF	
 usrMathematicalOps,			//		50		// 0x2200->0x22FF	
 usrMiscTechnical,				//		51		// 0x2300->0x23FF	
 usrControlPictures,			//		52		// 0x2400->0x243F	
 usrOpticalCharRecog,			//		53		// 0x2440->0x245F	
 usrEnclosedAlphanum,			//		54		// 0x2460->0x24FF	
 usrBoxDrawing,					//		55		// 0x2500->0x257F	
 usrBlockElements,				//		56		// 0x2580->0x259F	
 usrGeometricShapes,			//		57		// 0x25A0->0x25FF	
 usrMiscDingbats,				//		58		// 0x2600->0x26FF	
 usrDingbats,					//		59		// 0x2700->0x27BF
 usrBraillePatterns,			//		60		// 0x2800->0x28FF
 usrCJKRadicalsSupp,			//		61		// 0x2E80->0x2EFF
 usrKangxiRadicals,				//		62		// 0x2F00->0x2FDF
 usrIdeographicDesc,			//		63		// 0x2FF0->0x2FFF
 usrCJKSymAndPunct,				//		64		// 0x3000->0x303F
 usrHiragana,					//		65		// 0x3040->0x309F	
 usrKatakana,					//		66		// 0x30A0->0x30FF	
 usrBopomofo,					//		67		// 0x3100->0x312F	
 usrHangulCompatJamo,			//		68		// 0x3130->0x318F
 usrCJKMisc, 	// Kanbun		//		69		// 0x3190->0x319F
 usrExtendedBopomofo,			//		70		// 0x31A0->0x31BF
 usrEnclosedCJKLtMnth,			//		71		// 0x3200->0x32FF	
 usrCJKCompatibility,			//		72		// 0x3300->0x33FF
 usrCJKUnifiedIdeoExtA,			//		73		// 0x34E0->0x4DBF
 usrCJKUnifiedIdeo,				//		74		// 0x4E00->0x9FAF
 usrYi,							//		75		// 0xA000->0xA48F
 usrYiRadicals,					//		76		// 0xA490->0xA4CF
 usrHangul,						//		77		// 0xAC00->0xD7AF
 usrHighSurrogate,				//		78		// 0xD800->0xDBFF
 usrLowSurrogate,				//		79		// 0xDC00->0xDFFF 
 usrPrivateUseArea,				//		80		// 0xE000->0xF8FF
 usrCJKCompatibilityIdeographs,	//		81		// 0xF900->0xFA5F
 usrAlphaPresentationForms,		//		82		// 0xFB00->0xFB4F	
 usrArabicPresentationFormsA,	//		83		// 0xFB50->0xFDFF	
 usrCombiningHalfMarks,			//		84		// 0xFE20->0xFE2F	
 usrCJKCompatForms,				//		85		// 0xFE30->0xFE4F	
 usrSmallFormVariants,			//		86		// 0xFE50->0xFE6F	
 usrArabicPresentationFormsB,	//		87		// 0xFE70->0xFEFE	
 usrHFWidthForms,				//		88		// 0xFF00->0xFFEF	
 usrSpecials,					//		89		// 0xFFF0->0xFFFD	

 usrMax							//		90
} UnicodeSubRanges;

typedef enum UnicodeSubBits
{
 usbBasicLatin,
 usbLatin1,
 usbLatinXA,
 usbLatinXB,
 usbIPAExtensions,
 usbSpacingModLetters,
 usbCombDiacritical,
 usbBasicGreek,
 usbGreekSymbolsCop,
 usbCyrillic,
 usbArmenian,
 usbBasicHebrew,
 usbHebrewXA,
 usbBasicArabic,
 usbArabicX,
 usbDevangari,
 usbBengali,
 usbGurmukhi,
 usbGujarati,
 usbOriya,
 usbTamil,
 usbTelugu,
 usbKannada,
 usbMalayalam,
 usbThai,
 usbLao,
 usbBasicGeorgian,
 usbGeorgianExtended,
 usbHangulJamo,
 usbLatinExtendedAdd,
 usbGreekExtended,
 usbGeneralPunct,
 usbSuperAndSubscript,
 usbCurrencySymbols,
 usbCombDiacriticsS,	
 usbLetterlikeSymbols,	
 usbNumberForms,	
 usbArrows,	
 usbMathematicalOps,	
 usbMiscTechnical,	
 usbControlPictures,	
 usbOpticalCharRecog,	
 usbEnclosedAlphanum,	
 usbBoxDrawing,	
 usbBlockElements,	
 usbGeometricShapes,	
 usbMiscDingbats,	
 usbDingbats,	
 usbCJKSymAndPunct,
 usbHiragana,	
 usbKatakana,	
 usbBopomofo,	
 usbHangulCompatJamo,
 usbCJKMisc,	
 usbEnclosedCJKLtMnth,	
 usbCJKCompatibility,
 usbHangul,
 usbHighSurrogate,
 usbLowSurrogate,
 usbCJKUnifiedIdeo,
 usbPrivateUseArea,
 usbCJKCompatibilityIdeographs,
 usbAlphaPresentationForms,	
 usbArabicPresentationFormsA,	
 usbCombiningHalfMarks,	
 usbCJKCompatForms,	
 usbSmallFormVariants,	
 usbArabicPresentationFormsB,	
 usbHFWidthForms,	
 usbSpecials,
 usbTibetan,
 usbSyriac,
 usbThaana,
 usbSinhala,
 usbMyanmar,
 usbEthiopic,
 usbCherokee,
 usbCanadianAbor,
 usbOgham,
 usbRunic,
 usbKhmer,
 usbMongolian,
 usbBraillePatterns,
 usbYi,

 usbMax
 
} UnicodeSubBits;


/****************************************************************************
	Unicode Subrange Mask (USM) definitions

	Bit vector, one bit per Unicode subrange.
	Currently 90 bits are used, see subrange values above.

	Use routines in this header file to manipulate these.
*****************************************************************************/

#pragma pack(push, usmpack, 1)	// dense packing rules for USM.
typedef struct _USM
	{
	union
		{
		unsigned char rgb[12];
		struct
			{
			long l1;
			long l2;
			long l3;
			};
		};
	} USM;

typedef struct _USB
	{
	union
		{
		unsigned char rgb[16];
		struct
			{
			DWORD l1;
			DWORD l2;
			DWORD l3;
			DWORD l4;
			};
		};
	} USB;
		
#pragma pack(pop, usmpack)

#define cbUSM   sizeof(USM)
__inline void MsoAddUsrToUsm(int usr, USM *pusm) { \
	if (pusm)
		pusm->rgb[usr >> 3] |=  1 << (usr & 7);
	}

__inline void MsoAddUsbToFs(int usb, FONTSIGNATURE *pfs) 	{ \
	if (pfs)
		pfs->fsUsb[usb >> 5] |= 1 << (usb % (sizeof(DWORD)*8));
	}

__inline void MsoClearUsrFromUsm(int usr, USM *pusm) { \
	pusm->rgb[usr >> 3] &=  ~(1 << (usr & 7));
	}	

__inline int MsoFUsrInUsm(int usr, USM const *pusm) { \
	return (pusm->rgb[usr >> 3] & (1 << (usr & 7)));
	}	

// Perform bitwise 'and' on vector
__inline USM *MsoPusmAndUsm(USM const *pusmSrc1, USM const *pusmSrc2, USM *pusmDest) { \
	pusmDest->l1 = pusmSrc1->l1 & pusmSrc2->l1;
	pusmDest->l2 = pusmSrc1->l2 & pusmSrc2->l2;
	pusmDest->l3 = pusmSrc1->l3 & pusmSrc2->l3;
	return pusmDest;
	}	

// Perform bitwise 'or' on vector
__inline USM *MsoPusmOrUsm(USM const *pusmSrc1, USM const *pusmSrc2, USM *pusmDest) { \
	pusmDest->l1 = pusmSrc1->l1 | pusmSrc2->l1;
	pusmDest->l2 = pusmSrc1->l2 | pusmSrc2->l2;
	pusmDest->l3 = pusmSrc1->l3 | pusmSrc2->l3;
	return pusmDest;
	}	

__inline int MsoFEmptyUsm(USM const *pusm) { \
	return ((pusm->l1 | pusm->l2 | pusm->l3) == 0);
	}	

__inline void MsoClearUsm(USM *pusm) { \
	pusm->l1 = 0;
	pusm->l2 = 0;
	pusm->l3 = 0;
	}	

__inline int MsoFEqUsm(USM const *pusmSrc, USM const *pusmDest) { \
	return pusmDest->l1 == pusmSrc->l1 && pusmDest->l2 == pusmSrc->l2 && pusmDest->l3 == pusmSrc->l3;
	}

__inline void MsoInvertUsm(USM const *pusmSrc, USM *pusmDest) { \
	pusmDest->l1 = ~pusmSrc->l1;
	pusmDest->l2 = ~pusmSrc->l2;
	pusmDest->l3 = ~pusmSrc->l3;
	}	

// Clear bits that are set in mask from source and store in dest
__inline USM *MsoPusmClearUsm(const USM *pusmSrc, const USM *pusmMask, USM *pusmDest) { \
	pusmDest->l1 = pusmSrc->l1 & ~pusmMask->l1;
	pusmDest->l2 = pusmSrc->l2 & ~pusmMask->l2;
	pusmDest->l3 = pusmSrc->l3 & ~pusmMask->l3;
	return pusmDest;
	}	


/****************************************************************************
	Macros/inline functions
*****************************************************************************/

#define MsoFUsrFarEast(usr)		((usr) == usrCJKUnifiedIdeo)
#define MsoFUsrBi(usr)			((usr) == usrBasicHebrew || (usr) == usrHebrewXA || (usr) == usrBasicArabic || (usr) == usrArabicX || \
									(usr) == usrSyriac || (usr) == usrThaana)

//SOUTHASIA // schai. Used by Word
#define MsoFUsrSA(usr)	((usr) == usrThai || (usr) == usrDevangari || \
							(usr) == usrBengali || (usr) == usrGurmukhi || (usr) == usrGujarati || \
							(usr) == usrOriya || (usr) == usrTamil || (usr) ==  usrTelugu || \
							(usr) == usrKannada || (usr) == usrMalayalam || (usr) == usrSinhala || \
							(usr) == usrLao || (usr) == usrTibetan || (usr) == usrMyanmar || (usr) == usrKhmer)
//SOUTHASIA


/****************************************************************************
	Unicode Surrogate Pair support functions
*****************************************************************************/

#define wchHighSurrogateMin		0xd800
#define wchHighSurrogateMax		0xdbff
#define wchLowSurrogateMin		0xdc00
#define wchLowSurrogateMax		0xdfff

__inline BOOL MsoFWchHighSurrogate(WCHAR wch) {
	return ((wch >= wchHighSurrogateMin) && (wch <= wchHighSurrogateMax));
	}
	
__inline BOOL MsoFWchLowSurrogate(WCHAR wch) {
	return ((wch >= wchLowSurrogateMin) && (wch <= wchLowSurrogateMax));
	}
	
__inline BOOL MsoFPwchHighSurrogate( const WCHAR *pwch) {
#ifdef DEBUG
	// verify valid surrogate pair, assert if malformed string
	if (MsoFWchHighSurrogate(*pwch))
		{
		AssertInline(MsoFWchLowSurrogate(*(pwch + 1)));
		}
#endif
	return MsoFWchHighSurrogate(pwch[0]) && MsoFWchLowSurrogate(pwch[1]);
	}
	
__inline BOOL MsoFPwchLowSurrogate(const WCHAR *pwch) {
#ifdef DEBUG
	// verify valid surrogate pair, assert if malformed string
	if (MsoFWchLowSurrogate(*pwch))
		{
		AssertInline(MsoFWchHighSurrogate(*(pwch - 1)));
		}
#endif
	return MsoFWchHighSurrogate(pwch[-1]) && MsoFWchLowSurrogate(pwch[0]);
	}

__inline WCHAR *MsoPwchNext(const WCHAR *pwch) {
	if (MsoFPwchHighSurrogate(pwch))
		pwch++;
	pwch++;
	return (WCHAR *)pwch;
	}

__inline WCHAR *MsoPwchPrev(const WCHAR *pwch) {
	pwch--;
	if (MsoFPwchLowSurrogate(pwch))
		pwch--;
	return (WCHAR *)pwch;
}

/****************************************************************************
	Unicode Old Hangul suuport functions
*****************************************************************************/

#define wchOldHangulMin   0x1100
#define wchOldHangulMax   0x11FF

__inline BOOL MsoFWchOldHangul(WCHAR wch) {
	return ((wch >= wchOldHangulMin) && (wch <= wchOldHangulMax));
	}

/****************************************************************************
	Function declarations
*****************************************************************************/

MSOAPI_(int) MsoUsrFromWch(WCHAR);
MSOAPI_(int) MsoFGetUsmForChs(int, USM *);
MSOAPI_(void) MsoGetUsmForCpg(int cpg, USM *pusm);
MSOAPI_(BOOL) MsoFWchInUsr(WCHAR wch, int usr);
MSOAPI_(USM) MsoUsmFromFs(const FONTSIGNATURE* fs);
MSOAPI_(void) MsoFsFromUsm(const USM* pusm, FONTSIGNATURE* fs);
MSOAPI_(BOOL) MsoFFSAllComplex ( const FONTSIGNATURE* pfs );
MSOAPIX_(void) MsoWchFirstLastFromUsr(int usr, _Out_ WCHAR* pwchFirst, _Out_ WCHAR* pwchLast);


/****************************************************************************
	CMAP support functions
*****************************************************************************/

// these are high/low surrogate ranges just enough to cover the HK/CNS
// characters
#define wchCmapHighSurrogateMin		0xd840
#define wchCmapHighSurrogateMax		0xd869
#define wchCmapLowSurrogateMin		0xdc00
#define wchCmapLowSurrogateMax		0xdfff

#define MsoUTF16ToUCS4(wchHigh, wchLow) \
	(((int)wchHigh << 10) + (int)wchLow + 0x10000 - (0xD800 << 10) - 0xDC00)
#define ucs4CmapSurrogateMin	MsoUTF16ToUCS4(wchCmapHighSurrogateMin, wchCmapLowSurrogateMin)	
#define ucs4CmapSurrogateMax	MsoUTF16ToUCS4(wchCmapHighSurrogateMax, wchCmapLowSurrogateMax)	

// size of buffer to send to MsoFGetCharBitmapEx()
#define msocbCharBitmap		((0x10000 + ucs4CmapSurrogateMax - ucs4CmapSurrogateMin + 8) / 8)

// from kjstr.cpp
MSOAPI_(BOOL) MsoFSurrogatePairToUCS4(WCHAR w1, WCHAR w2, DWORD *plUCS4);
MSOAPI_(BOOL) MsoFUCS4ToSurrogatePair(DWORD lUCS4, _Out_ WCHAR *pw1, _Out_ WCHAR *pw2);


#if 0 //$[VSMSO]
/*---------------------------------------------------------------------------
	MsoFGetCharBitmap

	Scans the CMAP associated with the font selected into hdc and returns a
	bitmap indicating which Unicode characters have glyphs.

	pbMap parameter should point to a buffer at least 8192 bytes long.  This
	function will return a bit vector in this buffer indicating whether a
	character has valid glyph.  Use inline function MsoFWchInCmap()
	in conjunction with this bitmap.
	
	********* NOTE **********
	This function has been replaced with MsoFGetCharBitmapEx(); please do
	not use this function any more.
----------------------------------------------------------------- NobuyaH -*/
MSOAPI_(int) MsoFGetCharBitmap(BYTE *pbMap, HDC hdc);
#endif //$[VSMSO]

/*-----------------------------------------------------------------------------
	MsoFGetCharBitmapEx

	This is the enhanced version of MsoFGetCharBitmap which knows about
	surrogate pairs.
	
	pbMap parameter should point to a buffer at least msocbCharBitmap bytes
	long.
------------------------------------------------------------------- NobuyaH -*/
MSOAPI_(int) MsoFGetCharBitmapEx(BYTE *pbMap, int cbMap, HDC hdc);

/*-----------------------------------------------------------------------------
	MsoFWchInCmap

	This function returns TRUE if the indicated character exists in the supplied
	CMAP.

------------------------------------------------------------------- NobuyaH -*/
__inline int MsoFWchInCmap(const BYTE *pbMap, UINT wch)
{
	int ib = wch / 8;
	int ifl = wch % 8;

	return (pbMap[ib] & (1 << ifl));
}

/*-----------------------------------------------------------------------------
	MsoFSurrogatePairInCmap

	This function returns TRUE if the indicated surrogate pair exists in the
	supplied CMAP.
------------------------------------------------------------------- NobuyaH -*/
MSOAPI_(int) MsoFSurrogatePairInCmap(const BYTE *pbMap, WCHAR wchHigh, WCHAR wchLow);

#endif // EOF
