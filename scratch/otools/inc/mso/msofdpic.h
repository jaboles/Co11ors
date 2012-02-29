#pragma once

/******************************************************************************
	MSOFDPIC.H

	Owner: ziyiw
	Copyright (c) 1994 Microsoft Corporation

	12/23/97	Port Date and time core code from WORD9

********************************************************************* ziyiw **/
#ifndef MSOFDPIC_H
#define MSOFDPIC_H
#if 0 //$[VSMSO]
#pragma warning( push )
#pragma warning( disable:4201 ) // nonstandard extension used : nameless struct/union

/* field picture/format argument option characters */
#define chFldPicReqDigit    	'0'
#define chFldPicOptDigit    	'#'
#define chFldPicTruncDigit  	'x'
#define chFldPicNegSign     	'-'
#define chFldPicSign        	'+'
#define chFldPicExpSeparate 	';'
#define chFldPicQuote       	'\''  // use FChFldPicQuote(ch)
#define chFldPicSequence    	'`'
#define chFldPicAMPMSep     	'/'

// AMEQUALSYR macro -- for some languages, "a" as in AM/PM is
// 	also "a" as in year (ano, annee, etc.)
#define FAmEqualsYrLid(lid) \
				((lid) == lidFrench 			|| \
				 (lid) == lidFrenchCanadian 	|| \
				 (lid) == lidSpanish 			|| \
				 (lid) == lidPortBrazil 		|| \
				 (lid) == lidPortIberian 		|| \
				 (lid) == lidItalian 			|| \
				 (lid) == lidCatalan)

//DefineIds("dyhmms¡;0x#-+\'‘’\`.,,RC\0genaw", idsFldPicChars, sttFieldpic)

// items from ichFldPicDay - ichFldPicOf are localized
#define ichFldPicMin          	0
#define ichFldPicNull			0
#define ichFldPicDay			1
#define ichFldPicYear			2
#define ichFldPicHour			3
#define ichFldPicMonth			4
#define ichFldPicMinute			5
#define ichFldPicSecond       	6
#define ichFldPicOf				7		/* second letter of "of" -- a hack for
												Spanish and Portuguese.  Must NOT
								   			match any other chFldPic letter!
								   			For English this is a random (seldom
								   			used) character */
#define ichFldPicMax          	8

// Warning!  Make sure items below match with their
//  chFldPic counterparts
#define ichFldPicExpSeparate	8	
#define ichFldPicReqDigit		9	
#define ichFldPicTruncDigit	    10
#define ichFldPicOptDigit		11
#define ichFldPicNegSign		12
#define ichFldPicSign			13
#define ichFldPicQuote1         14
#define ichFldPicQuote2         15
#define ichFldPicQuote3         16
#define ichFldPicSequence		17
#define ichFldPicDecimalSep     18  // used for field translation purposes only
#define ichFldPicThousandSep    19  // used for field translation purposes only
#define ichFldPicListSep        20  // used for field translation purposes only
#define ichFldPicRow            21  // used for field translation purposes only
#define ichFldPicCol            22  // used for field translation purposes only
#define ichFldPicEof			23
#define ichFldPicEra            24	//  g
#define ichFldPicEraNZ          25	//  e
#define ichFldPicEraNZCompat    26	//  n
#define ichFldPicDayFE          27	//  a
#define ichFldPicDayFECompat    28	//  w
#define ichFldPicTpeMonthD	    29	//  o
#define ichFldPicTpeHourD12  	30	//  r
#define ichFldPicTpeMinuteD  	31	//  i
#define ichFldPicTpeSecondD  	32	//  c


//SOUTHASIA	//a-sagraw Changes for thai 97 compatibility

#define ichFldPicBuddhistYear	33
#define ichFldPicThaiDay		34 
#define ichFldPicThaiMonth		35
#define ichFldPicThaiYear		36
#define ichFldPicThaiHour		37
#define ichFldPicThaiMinute		38
#define ichFldPicThaiSecond		39
#define ichFldPicThaiHour24		40	// v-sakvin - 10/12/98 : 97Compatibility
#define ichFldPicEofFE			41
#define cchMaxFldPicChars		42
//SOUTHASIA

// Date/Time codes
#define msoiszDttmMin        0
#define msoiszDateShort      0  //  S1 iszDateMin
#define msoiszDateLongDay    1  //  L1   
#define msoiszDateLong       2  //  L2
#define msoiszDateShortAlt   3  //  S2
#define msoiszDateISO        4  //  S3
#define msoiszDateShortMon   5  //  S4
#define msoiszDateShortSlash 6  //  S5
#define msoiszDateShortAbb   7  //  L3
#define msoiszDateEnglish    8  //  L4
#define msoiszDateMonthYr    9  //  L5 
#define msoiszDateMon_Yr    10  //  L7 iszDateMac
#define msoiszTimeDatePM    11  //  T1 iszTimeMin
#define msoiszTimeDateSecPM 12  //  T2
#define msoiszTimePM        13  //  T3
#define msoiszTimeSecPM     14  //  T4
#define msoiszTime24        15  //  T5
#define msoiszTimeSec24     16  //  T6 
#define msoiszFEExtra1		17	// 	
#define msoiszFEExtra2		18	// 
#define msoiszFEExtra3		19	// 	
#define msoiszFEExtra4		20	// 	
#define msoiszFEExtra5		21	// 	
#define msoiszDttmMax	    21

#define msoiszDateMac 		msoiszDateMon_Yr					// Obsolete define, don't use them anymore
#define msoiszTimeMac 		msoiszTimeSec24						// Obsolete define, don't use them anymore
#define msoiszDateMin 		msoiszDateShort						// Obsolete define, don't use them anymore
#define msoiszTimeMin 		msoiszTimeDatePM					// Obsolete define, don't use them anymore
#define cmsoiszDateNum 		(msoiszDateMac - msoiszDateMin + 1)	// Obsolete define, don't use them anymore
#define cmsoiszTimeNum 		(msoiszTimeMac - msoiszTimeMin + 1)	// Obsolete define, don't use them anymore
#define cmsoiszDttmNum		(msoiszDttmMax - msoiszDttmMin + 1)

// Error codes, stord in wArg, with fptError
#define msoistErrUnmatchedQuote			17
#define msoistErrUnknownPicCh			18

// Special mode for generating DTTM pictures
#define msofDttmNone					0x0000	// Nothing set
#define msofDttmHebrewNumbering			0x0001	// BIDI, use Hebrew numbering
#define msofDttmBiDiUseShort			0x0002	// BIDI, use Short format
#define msofDttmHideROC					0x0004  // CHN, hide ROC
#define msofDttmBidiCalendar			0x0008  // Hijri or Lunar Calendar.
//SOUTHASIA
#define msofDttmBuddhistCalendar		0x0010  //a-sagraw Date/Time for Thai/Vietnamese (Budhhist Calendar)
#define msofDttmSakaEra					0X0020  //a-sagraw Date/Time for Hindi calendar
//SOUTHASIA

//SOUTHASIA
#define BUDDHIST_CAL_OFFSET		543
//SOUTHASIA

#ifdef NEVER
/* Custom data in Picture Parse Block */
typedef struct _PPBCUST
	{
	CP 			cpArg;
	CR 			*rgcr;     	/* should be ccrMaxPic */
	}PPBCUST;
#endif //NEVER

#endif //$[VSMSO]

/* F I E L D  P I C T U R E  T O K E N S */
	/* common tokens */  
#define fptCommonMin    20
#define fptEof          20
#define fptError        21
#define fptQuote        22
#define fptSequence     23
#define fptCommonMax    24
	/* numeric tokens */
#define fptReqDigit     10
#define fptOptDigit     11
#define fptTruncDigit   12
#define fptDecimal      13
#define fptThousand     14
#define fptNegSign      15
#define fptReqSign      16
	/* date/time tokens */
#define fptMonth1       30  /* 1-12 */
#define fptMonth2       31  /* 01-12 */
#define fptMonth3       32  /* Jan-Dec */
#define fptMonth4       33  /* January-December */
#define fptDay1         34  /* 1-31 */
#define fptDay2         35  /* 01-31 */
#define fptDay3         36  /* Sun-Sat */
#define fptDay4         37  /* Sunday-Saturday */
#define fptYear1        38  /* 00-99 */
#define fptYear2        39  /* 1900-2156 */
#define fptHour1        40  /* 1-12 */
#define fptHour2        41  /* 01-12 */
#define fptHour3        42  /* 0-23 */
#define fptHour4        43  /* 00-23 */
#define fptMinute1      44  /* 0-59 */
#define fptMinute2      45  /* 00-59 */
#define fptAMPMLC       46  /* am/pm, a/p */
#define fptAMPMUC       47  /* AM/PM, A/P */
#define fptSecond1      48  /* 0-59 */
#define fptSecond2      49  /* 00-59 */

//  date/time tokens for Japanese
#define fptOfsJapanese  50                      /* Japanese offset */
#define fptJpnMonth     (fptOfsJapanese + 0)    /* Mutsuki-Shiwasu */
#define fptJpnDay       (fptOfsJapanese + 1)    /* Nichi-Do */
#define fptJpnEraSbc    (fptOfsJapanese + 2)    /* Era name in SBC */
#define fptJpnEraDbc1   (fptOfsJapanese + 3)    /* Era name in DBC (1 char) */
#define fptJpnEraDbc2   (fptOfsJapanese + 4)    /* Era name in DBC */
#define fptJpnYear1     (fptOfsJapanese + 5)    /* 1-99 */
#define fptJpnYear2     (fptOfsJapanese + 6)    /* 01-99 */
#define fptJpnYearD1    (fptOfsJapanese + 7)	/* DBC emperor year name */
#define fptJpnYearD2    (fptOfsJapanese + 8)	/* DBC yyyy year name */
#define fptJpnAMPM      (fptOfsJapanese + 9)	/* Gozen/Gogo */

//	data/time tokens for Chinese
#define fptOfsChinese		60                      /* Chinese offset */
#define fptTpeYearS1   		(fptOfsChinese)    		/* e */
#define fptTpeYearS2   		(fptOfsChinese + 1)    	/* ee */
#define fptTpeYearS3   		(fptOfsChinese + 2)    	/* eee */
#define fptTpeYearD1   		(fptOfsChinese + 3)    	/* E */
#define fptTpeYearD2   		(fptOfsChinese + 4)    	/* EE */
#define fptTpeYearD3   		(fptOfsChinese + 5)    	/* EEE */
#define fptTpeYearD4   		(fptOfsChinese + 6)    	/* EEEE */
#define fptFEMonthD    		(fptOfsChinese + 7)		/* O */ // generalize it to all FE
#define fptFEDayD      		(fptOfsChinese + 8)		/* A */ // generalize it to all FE
#define fptTpeWeekD      	(fptOfsChinese + 9)		/* W */
#define fptTpeHourD12   	(fptOfsChinese + 10)    /* r */
#define fptTpeHourAMPMD12 	(fptOfsChinese + 11)    /* rr */
#define fptTpeHourD24   	(fptOfsChinese + 12)    /* R */
#define fptTpeHourAMPMS12 	(fptOfsChinese + 13)	/* RR */
#define fptTpeMinuteD   	(fptOfsChinese + 14)    /* I */
#define fptTpeSecondD   	(fptOfsChinese + 15)    /* C */
#define fptChnAMPM			(fptOfsChinese + 16)	/* AMPM */

//	data/time tokens for Korean                     
#define fptOfsKorean		77                  /* Korean offset */
#define fptKorDay       	(fptOfsKorean + 0)  	  
#define fptKorAMPM      	(fptOfsKorean + 1)	/* Ojen/ Ohu */
#define	fptKorEraYear		(fptOfsKorean + 2)	/* Dangi */

//	data/time tokens for Taiwan
#define	fptTpeEraYear		(fptOfsKorean + 3)	/* gg */

#define fptFEAMPM	fptJpnAMPM	// quick & dirty fix


//SOUTHASIA
//date/time tokens for Thai    //a-sagraw
#define fptOfsThai       	81
#define fptThaiYear1     	(fptOfsThai + 0)
#define fptThaiYear2     	(fptOfsThai + 1)
#define fptThaiMonth     	(fptOfsThai + 2)
#define fptThaiDay1			(fptOfsThai + 3)
#define fptThaiDay2      	(fptOfsThai + 4)
#define fptThaiHour1      	(fptOfsThai + 5)
#define fptThaiHour2      	(fptOfsThai + 6)
#define fptThaiMinute    	(fptOfsThai + 7)
#define fptThaiSecond    	(fptOfsThai + 8) 
#define	fptThaiHour241		(fptOfsThai + 9)
#define	fptThaiHour242		(fptOfsThai + 10)

//date/time tokens for Hindi
#define fptOfsHindi       92
#define fptHindiYear1     (fptOfsHindi + 0)
#define fptHindiYear2     (fptOfsHindi + 1)
#define fptHindiMonth     (fptOfsHindi + 2)
#define fptHindiDay1      (fptOfsHindi + 3)
#define fptHindiDay2      (fptOfsHindi + 4)
#define fptHindiHour      (fptOfsHindi + 5)
#define fptHindiMinute    (fptOfsHindi + 6)
#define fptHindiSecond    (fptOfsHindi + 7) 
//SOUTHASIA

#undef cchMaxPic
#define cchMaxPic			64

#undef cbMaxCurrency
#define cbMaxCurrency		7

typedef unsigned short BF;
typedef unsigned char BFC;

#if 0 //$[VSMSO]
/* Date/Time Structure */
/* A smaller version of a DTR, or an expanded version of a DTTM! */
typedef struct _msodts
	{
	BF mon  :4;  /* 1-12 */
	BF mint :6;  /* 0-59 */
	BF sec  :6;  /* 0-59 */
	BF hr   :5;  /* 0-23 */
	BF wdy  :3;  /* 0(Sun)-6(Sat) */
	BF dom  :5;  /* 1-31 */
	BF yr   :12; /* 1900-4095 */
	} MSODTS;

// BIDI Date info block
typedef struct _msobdi {
    LID lid;
    int	dtDay;		
    int	dtMonth;		
    int	dtYear;		
    int	dtWDay;		
    int	dtWStart;	
    union {
    	struct {
    		WCHAR dtXszHebWDay[30];		
    		WCHAR dtXszHebMonth[30];		
    		WCHAR dtXszHebYear[10];		
    		WCHAR dtXszHebDay[10];		
    		WCHAR dtXszHebHoliday[40];	
    		};
    	struct {
    		WCHAR dtXszAraWDay[30];
    		WCHAR dtXszAraMonth[30];
    		};
    	struct {
    		WCHAR dtXszPerWDay[30];
    		WCHAR dtXszPerMonth[30];
    		};
    	struct {
    		WCHAR dtXszWDay[30];
    		WCHAR dtXszMonth[30];
    		WCHAR dtXszWDayShort[10];
    		WCHAR dtXszMonthShort[10];
    		};
    	};
    }MSOBDI;


/* CORE Picture Parse Block */
typedef struct _ppb
	{
	LID				lid;				/* language id associated with this picture str */
	LID				lidFE;				/* FE language id */
	WCHAR 			*pxchNext;			/* next char position */
	int 			wArg;				/* error id, or the size of arg */
	WCHAR 			*pxchArg;			/* pointer to the arg, if any */
	int 			pos;				/* for numbering use, only */
	int 			posBlw;				/* for numbering use, only */
	WCHAR 			*rgxch; 			/* should be cchMaxPic+1 */
	WCHAR			xchThousand;		/* get from locale */
	WCHAR			xchDecimal; 		/* get from locale */
	WCHAR			xszCurrency[cbMaxCurrency];
	void			*pCustom;			/* any custom data structure */

	union 
		{
		BFC	fTempFlags : 8;
		struct
			{
			BFC 	fThouSep 	: 1;
			BFC 	fImpSign 	: 1;
			BFC		fGenetive 	: 1;
			BFC		fEraYear 	: 1;
			BFC		fNoLength	: 1;
			BFC		fUnused 	: 3;
			};
		};

	} MSOPPB;

#define	msocaltWestern				0
#define msocaltArabicHijri			1
#define msocaltHebrewLunar			2
#define msocaltChineseNational	3
#define msocaltJapaneseEmperor	4
#define msocaltThaiBuddhist		5
#define msocaltKoreanDanki			6
//SOUTHASIA	// v-sakvin - 10/7/98 : Calendar Supports
#define msocaltSakaEra				7
//SOUTHASIA
//a-tzipb Arabic Transliterate calendars
#define msocaltTranslitEnglish   8  // Gregorian Transliterated English calendar
#define msocaltTranslitFrench    9  // Gregorian Transliterated French calendar


#define msocaltMin					msocaltWestern
//SOUTHASIA	// v-sakvin - 10/07/98 : Calendar Supports
//#define msocaltMax					msocaltSakaEra
#define msocaltMax               msocaltTranslitFrench     
//SOUTHASIA

// Format types for UI support
#define msoDttmOptDate	0x01	// use date formats
#define msoDttmOptTime	0x02	// use time formats

// Public MSO Dttm APIs
MSOAPI_(int) MsoGetDttmFormatCore(int iszDateCode, LID lid, WCHAR *xszDate, int *pcch);
MSOAPI_(int) MsoFGetDefTimeFormat(LID lid, WCHAR *wzTime, int *pcchMax);
MSOAPI_(int) MsoGetDttmPictureCore(MSODTS *pdts, const WCHAR *xszPic, WCHAR *xszDest, int *pcchMaxDest, LID lid, LID lidFE, MSOBDI *pbidi, UCHAR fMode);
MSOAPI_(int) MsoFDttmFormatLidFECalCompat(const WCHAR *xszFormat, LID lid, int cal);
MSOAPI_(int) MsoGetCalNameXsz(int cal, WCHAR *xsz, int cch);
MSOAPI_(int) MsoFHideROC(void);

#define MsoGetDttmPicture(pdts, xszsrc, xszdest, pcchMaxDest, lid)				\
				MsoGetDttmPictureCore((pdts), 									\
									  (xszsrc), 								\
									  (xszdest), 								\
									  pcchMaxDest,								\
									  (FLidFarEast(lid) ? lidAmerican : (lid)), \
									  (FLidFarEast(lid) ? (lid) : lidAmerican), \
									  NULL, 									\
									  (UCHAR) msofDttmNone)

// Public MSO Dttm APIs - extension
MSOAPI_(int) MsoGetIntlDttmPictureCore ( MSODTS *pdts, const WCHAR *xszPic, WCHAR *xszDest, int *pcchMaxDest, LID lid, LID lidFE, int calType, UCHAR fMode);
MSOAPI_(int) MsoGetIntlCalList (LID lid, _Out_cap_(pListSize) int *pcaltList, int *pListSize);
MSOAPI_(int) MsoGetDttmIszList (LID lid, _Out_cap_(pListSize) int *piszList, int *pListSize);

// Public MSO Dttm APIs - language specific preferences
MSOAPI_(int) MsoFGetDttmPrefLast(LID* lid, int *pcalt, int *piszDate, int *piszTime, int *pOpt, WCHAR* wszDatePic, int cchDate, WCHAR* wszTimePic, int cchTime);
MSOAPI_(int) MsoFGetDttmPrefLid(LID lid, int *pcalt, int *piszDate, int *piszTime, int *pOpt, WCHAR* wszDatePic, int cchDate, WCHAR* wszTimePic, int cchTime);
MSOAPI_(int) MsoFSetDttmPrefLid(LID lid, int calt, int iszDate, int iszTime, int iOpt, const WCHAR* wszDatePic, const WCHAR* wszTimePic);

// PPB parse helper functions
MSOAPI_(WCHAR*) MsoPxchInPppb(WCHAR xch, const WCHAR *pxch);
MSOAPI_(int) 	MsoCchDtsFptToRgxchCore(MSODTS *pdts, int fpt, WCHAR *pxch, int cch, LID lid, LID lidFE, MSOBDI *pbidi, UCHAR fMode, MSOPPB *pppb);
MSOAPI_(int) 	MsoFptNextPppbCore(MSOPPB *pppb, int fNumeric);
MSOAPI_(int) 	MsoInitPppbCore(int fNumeric, int wSign, const WCHAR *rgxch, LID lid, LID lidFE, MSOPPB *ppb, void *pCustom);
MSOAPI_(int) 	MsoGetFldPicChars(LID lid, WCHAR *xszFldPic, int *pcch);

#endif //$[VSMSO]
// Other useful MSO APIs
MSOAPI_(int) 	MsoGetDefWeekdayNameLid(LID lid, int iDay, int fAbbrev, _Out_z_cap_(cchMax) WCHAR *xsz, DWORD cchMax);
MSOAPI_(int) 	MsoGetDefMonthNameLid(LID lid, int iMon, int fAbbrev, _Out_z_cap_(cchMax) WCHAR *xsz, DWORD cchMax);
#if 0 //$[VSMSO]
MSOAPI_(void) 	MsoConvertPictureSwitch(WCHAR *xszPic, int cchPic);
MSOAPI_(WCHAR*) MsoXszDttmException(LID lid, WCHAR *xszString, int ixchMax, int iszFormat, const WCHAR *xszFldPicChars);
#endif //$[VSMSO]
MSOAPI_(int) 	MsoCchGetLocaleInfo(LID lid, LCTYPE LCType, _Out_z_cap_(cchMax) WCHAR *xsz, int cchMax);
MSOAPIX_(int) 	MsoSzToWzLid(const CHAR* sz, _Out_z_cap_(ixchMax) WCHAR* wz, int ixchMax, LID lid);
MSOAPIX_(WCHAR*) MsoWzStripSpaces(_Inout_z_ WCHAR *pwch);
#if 0 //$[VSMSO]
MSOAPI_(int) MsoFEnumLangStrings(int ilst, WCHAR *wtz, int cchMax, LID *plid);
MSOAPI_(int) MsoGetCalList (LID lid, _Out_cap_(pListSize) int *pcaltList, int *pListSize); 
MSOAPI_(int) MsoGetIntlCalType ( LID lid, int msoCalType );
MSOAPI_(int) MsoGetSystemHijriAdvance();
MSOAPI_(int) MsoFCvtBidiDate(const MSODTS *pdts, MSOBDI *pbidiDate, LID lid, int cal);

//SOUTHASIA //a-sagraw
#define HINDI_NUM_ZERO	0x0966
#define	THAI_NUM_ZERO	0x0E50
MSOAPI_(int) MsoCchUnsToPpxchThaiHindiNum(unsigned int n, WCHAR **ppxch, const WCHAR *pxchMac, int cDigits, WCHAR xchNumZero);
void MsoGetSakaDt(MSODTS, MSODTS*);
//SOUTHASIA

#pragma warning( pop )

#endif //$[VSMSO]
#endif //MSOFDPIC_H
