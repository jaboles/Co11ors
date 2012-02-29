#pragma once

/*************************************************************************
 	msostr.h

	Owner: rickp
 	Copyright (c) 1994 Microsoft Corporation

	Standard string utilities for Office.  We have fairly a complete
	set of operations on the SZ (null-terminated single-byte string),
	WZ (null-terminated unicode string), and WTZ (null-terminated
	unicode string with length word) types, and a few random other
	string representations thrown in.
*************************************************************************/

#if !defined(MSOSTR_H)
#define MSOSTR_H

#if !defined(MSO_CONST_CLEAN_STR_FUNCTIONS)
#define MSO_CONST_CLEAN_STR_FUNCTIONS 1
#endif

#if !defined(MSOSTD_H)
#include <msostd.h>
#endif
#include <msointl.h>

#if !defined(MSODEBUG_H)
#include <msodebug.h>
#endif

#if !defined(_INC_MINMAX)
#include <minmax.h>
#endif

#if defined(__cplusplus)
extern "C" {
#endif

/*************************************************************************
	Unicode character utilities
*************************************************************************/

/*
Special Turkish sort order defined by Office - different from system's
default Turkish sort order in that undotted i's sort before dotted i's -
in the system, they both sort equal. To use, or in the mask with the
LCID passed to the MsoCompareString* routines.
*/

#define MSO_TURKISH_SORT 0xf
#define MSO_TURKISH_SORT_MASK 0x000f0000

MSOAPI_(BOOL) MsoFPuncWch(WCHAR wch);

MSOAPI_(BOOL) MsoFAlphaWch(WCHAR wch);

MSOAPI_(BOOL) MsoFSpecChWch(WCHAR wch);

MSOAPI_(BOOL) MsoFComplexMarkWch(WCHAR wch);

MSOAPI_(BOOL) MsoFSpaceWch(WCHAR wch);

MSOAPIXX_(BOOL) MsoFSpaceCh(CHAR ch);

MSOAPI_(BOOL) MsoFStripLeadingAndEndingWSWz(const WCHAR* wzIn, _Deref_out_z_ WCHAR** wzOut);

MSOAPI_(int) MsoCompareStringA(LCID Locale, DWORD dwCmpFlags,
		LPCSTR lpString1, int cchCount1, LPCSTR lpString2, int cchCount2);

MSOAPI_(int) MsoCompareStringW(LCID Locale, DWORD dwCmpFlags,
		LPCWSTR lpString1, int cchCount1, LPCWSTR lpString2, int cchCount2);

MSOAPI_(int) MsoCompareStringLegacyW(LCID Locale, DWORD dwCmpFlags,
		LPCWSTR lpString1, int cchCount1, LPCWSTR lpString2, int cchCount2);

/*	Returns fTrue if and only if wch is a digit. */
MSOAPI_(BOOL) MsoFDigitWch(WCHAR wch);

/*	Returns TRUE if and only if ch is a digit. */
MSOAPI_(BOOL) MsoFDigitCh(int ch);

/* Returns TRUE if and only if wch is a hex digit. */
MSOAPIX_(BOOL) MsoFHexDigitWch(WCHAR wch);

/*	Returns fTrue if and only if wch is an alphanumeric character. */
MSOAPIX_(BOOL) MsoFAlphaNumWch(WCHAR wch);

/*	Returns TRUE if and only if wch is a digit or an alpha with upper/lowercase. */
MSOAPIX_(BOOL) MsoFUpperLowerDigitWch(WCHAR wch);

/* Converts wch to lower case and returns the result. */
MSOAPIMX_(WCHAR) MsoWchLower(WCHAR wch);

/*  Returns FS_* font signature flags indicating code pages which support
    the supplied Unicode character. */
MSOAPI_(int) MsoFsFromWch(WCHAR wch);

/*	Returns fTrue if and only if wch is a FE char. */
MSOAPI_(BOOL) MsoFFEWch(WCHAR wch);

/* Returns an FS_* font signature flag corresponding to a particular code
	page, these two routines can be used to work out spans of characters
	which fit in to a give code page, hence can be represented by a single
	Win95 Charset. */
MSOAPI_(int) MsoFsCpgFromCpg(int cpg);

/* Return the charset that corresponds to a given code page (i.e. that
	contains the glyphs for the characters in that code page). */
MSOAPI_(BYTE) MsoChsFromCpg(int cpg);

/* Return the code page from a font signature bit array.  Based on the Word
	code but this handles a full array (but the code page returned is fairly
	random) - i.e. multiple bits may be set in fsCpg. */
MSOAPI_(CPG) MsoCpgFromFsCpg(int fsCpg);

MSOAPI_(int) MsoFToggleCharCode(LPCWSTR pwtSrc, _Out_cap_(cchDest) LPWSTR pwtDest, int cchDest, int fExact);


//SOUTHASIA

BOOL MsoIsWchVietToneMark(WCHAR wch);
/*	Returns fTrue if and only if wch is a Indic vowel. */
MSOAPI_(BOOL) MsoIsWchIndicVowel(WCHAR wch);

BOOL MsoFIsThaiText(LPCWSTR lpwch, LPCWSTR lpwchEnd);
BOOL MsoFIsThaiChar(WCHAR wch);

//SOUTHASIA


/*************************************************************************
	Unicode Japanese character normalization routines from Word.
    (Fuzzy Normalization)
*************************************************************************/

/* Fuzzy Option Flags Structure - for Fuzzy Normalization.*/
typedef struct _MSOGRPFFUZ
	{
	union
		{
		unsigned int grpfFuzExp;
		struct
			{
			ULONG		fEqCase : 1;				// ( 0) upper/lower case
			ULONG		fEqByte : 1;				// ( 1) single/double byte
			ULONG		fEqMinus : 1;				// ( 2) minus/long vowel
			ULONG		fEqHira_Kata : 1;			// ( 3) hiragana/katakana
			ULONG		fEqSmallKana : 1;			// ( 4) small kana
			ULONG		fEqRepSymbol : 1;			// ( 5) repeated symbols
			ULONG		fEqOldKana : 1;				// ( 6) old kana
			ULONG		fEqLongVowel : 1;			// ( 7) several exp with long vowel
			ULONG		fEqD_Z : 1;					// ( 8) DI/JI DU/ZU
			ULONG		fEqB_V : 1;					// ( 9) BA/VA
			ULONG		fEqT_C : 1;					// (10) TSI/THI/CHI
			ULONG		fEqDhi_Zi : 1;				// (11) DHI/ZI
			ULONG		fEqY_A : 1;					// (12) YA/A
			ULONG		fEqS_SH : 1;				// (13) SHE/SE JE/ZE
			ULONG		fEqKI_KU : 1;				// (14) KI/KU
			ULONG		fEqF_H : 1;					// (15) FU/HU
			ULONG		fEqKanji : 1;				// (16) kanji ("itaiji")
			ULONG		fEqSoundsLike : 1;			// (17) sounds like

			ULONG		fIgnorePunct : 1;			// (18) ignore punctuation char
			ULONG		fIgnoreSpace : 1;			// (19) ignore space char

			ULONG		fDontIgnore : 1;			// (20) ignore nothing
			ULONG		fDontIgnoreLast : 1;		// (21) ignore nothing at the end of str

			ULONG		: 10;
			};
		};
	} MSOGRPFFUZ;

/* 	%%Function: MsoCwchGetAltChar
	%%Contact: katsun,kander

	Get alternate characters of wchOrig.
	Result characters are put in pwchResult which includes the original
	character.
	If return value is zero, there is no alternate characters. */
MSOAPI_(int) MsoCwchGetAltChar(WCHAR wchOrig, WCHAR *pwchResult, int cwchResultMax);

/*	%%Function: MsoCwchNormFuzzy
	%%Contact: katsun,kander

	Normalize fuzzy expression in rgwchSrc into rgwchDest.
	rgwchDest must be big enough to hold the result string.
	It does not ignore anything if fDontIgnore is fTrue.
	Return cwchDest if succeed. */
MSOAPI_(int) MsoCwchNormFuzzy(const WCHAR *rgwchSrc, int cwchSrc, WCHAR *rgwchDest, int cwchDestMax, MSOGRPFFUZ *pgrpffuz);



/*************************************************************************
	Simple string utilities
*************************************************************************/

#define MsoWzCchCopy MsoWzCopy
#define MsoSzCchCopy MsoSzCopy
#define MsoWzCchAppend MsoWzAppend

/*	Returns the length, in bytes, of the SBCS/DBCS string sz. */
MSOAPI_(int) MsoCchSzLen(__rzcount(return) const CHAR* sz);

/*	Returns the length of the Unicode string wz. */
MSOAPI_(int) MsoCchWzLen(__rzcount(return) const WCHAR* wz);

/*	Returns the length of the Unicode string wtz. */
__inline WORD MsoCchWtzLen(__rtzcount(return) const WCHAR* wtz) { return (WORD)wtz[0]; }

/*	Returns the length of the Unicode string wt. */
__inline WORD MsoCchWtLen(__rtcount(return) const WCHAR* wt) { return (WORD)wt[0]; }

/* Copies the lesser of len(wzFrom) or n characters (n includes the
	null terminator) from wzFrom to wzTo. */
MSOAPI_(WCHAR *) MsoWzCopy(_In_z_ const WCHAR* wzFrom, _Out_z_cap_(cch) WCHAR* wzTo, int cch );

/*	Copies the lesser of len(szFrom) or n characters (n includes the
	null terminator) from szFrom to szTo. */
MSOAPI_(char *) MsoSzCopy(_In_z_ const CHAR* szFrom, _Out_z_cap_(cch) CHAR* szTo, int cch);

MSOAPI_(BOOL) MsoWtCopy(const WCHAR* wtFrom, _Out_cap_(cch) WCHAR* wtTo, int cch);
MSOAPI_(BOOL) MsoWtRgwchCchCopy(const WCHAR* rgwchFrom, int cchFrom, _Out_cap_(cchMax) WCHAR* wtTo, int cchMax);
MSOAPI_(BOOL) MsoWtzCopy(const WCHAR *wtzFrom, _Out_cap_(cch) WCHAR *wtzTo, int cch);
MSOAPI_(BOOL) MsoWtzRgwchCchCopy(const WCHAR* rgwchFrom, int cchFrom, _Out_cap_(cchMax) WCHAR* wtzTo, int cchMax);

__inline void MsoWtzCopyNull(const WCHAR *wtzFrom, _Out_z_cap_(cchTo) WCHAR *wtzTo, int cchTo)
{
	if (wtzFrom == NULL)
		wtzTo[0] = wtzTo[1] = 0;
	else
		MsoWtzCopy(wtzFrom, wtzTo, cchTo);
}

/*  A simple string copy, creating a new copy in the mark buffer */
MSOAPIX_(WCHAR *) MsoWzCopyMark(const WCHAR *wzFrom);

/*	Returns pointer to first occurrence of ch in rgch, NULL if not found. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const CHAR*) MsoRgchIndex(const CHAR *rgch, int cch, int ch);
#else
MSOAPI_(CHAR*) MsoRgchIndex(const CHAR *rgch, int cch, int ch);
#endif
__inline CHAR* MsoRgchIndexNonConst(_In_z_ CHAR *rgch, int cch, int ch)
{
	OACR_USE_PTR(rgch); return (CHAR*)MsoRgchIndex(rgch, cch, ch );
}

/*	Returns pointer to first occurrence of wch in rgwch, NULL if not found. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoRgwchIndex(const WCHAR *rgwch, int cwch, int wch);
#else
MSOAPI_(WCHAR*) MsoRgwchIndex(const WCHAR *rgwch, int cwch, int wch);
#endif
__inline WCHAR* MsoRgwchIndexNonConst(_In_z_ WCHAR *rgwch, int cwch, int ch)
{
	OACR_USE_PTR(rgwch); return (WCHAR*)MsoRgwchIndex(rgwch, cwch, ch );
}

/*	Returns pointer to last occurrence of wch in rgwch, NULL if not found. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoRgwchIndexRight(const WCHAR *rgwch, int cwch, WCHAR wch);
#else
MSOAPI_(WCHAR*) MsoRgwchIndexRight(const WCHAR *rgwch, int cwch, WCHAR wch);
#endif
__inline WCHAR* MsoRgwchIndexRightNonConst(_In_z_ WCHAR *rgwch, int cwch, int ch)
{
	OACR_USE_PTR(rgwch); return (WCHAR*)MsoRgwchIndexRight(rgwch, cwch, ch );
}

/*	Returns the length, in bytes, of the sz derived from wz.  -1 on error */
MSOAPI_(int) MsoCpCchSzLenFromWz(UINT CodePage, const WCHAR *wz);

/*	Append the WTZ string wtzFrom onto the end of wtzTo. */
MSOAPI_(void) MsoWtzAppend(const WCHAR* wtzFrom, _Out_z_cap_(cch) WCHAR* wtzTo, int cch);

/* Appends cch-strlen(wzTo) chars from wzFrom onto the end of wzTo, and
	returns a pointer to the end of the destination string. */
MSOAPI_(WCHAR*) MsoWzAppend(const WCHAR* wzFrom, _Out_z_cap_(cchTo) WCHAR* wzTo, int cchTo);

/* Appends cch-strlen(szTo) chars from szFrom onto the end of szTo, and
	returns a pointer to the end of the destination string. */
MSOAPI_(CHAR*) MsoSzAppend(const char *szFrom, _Out_z_cap_(cchTo) char *szTo, int cchTo);

/* Copy and append routines when the source string is not null terminated.
 * cchDest is always the total size of the destination, and the destination
 * string will always be null terminated after the call.
 */
MSOAPI_(CHAR *) MsoRgchCopy(const CHAR *rgchSrc, int cchSrc, _Out_z_cap_(cchDest) CHAR *szDest, int cchDest);
MSOAPI_(WCHAR *) MsoRgwchCopy(const WCHAR *rgwchSrc, int cchSrc, _Out_z_cap_(cchDest) WCHAR *wzDest, int cchDest);
MSOAPI_(CHAR *) MsoRgchAppend(const CHAR *rgchSrc, int cchSrc, _Out_z_cap_(cchDest) CHAR *szDest, int cchDest);
MSOAPI_(WCHAR *) MsoRgwchAppend(const WCHAR *rgwchSrc, int cchSrc, _Out_z_cap_(cchDest) WCHAR *wzDest, int cchDest);
MSOAPI_(WCHAR *) MsoRgwchWtzAppend(const WCHAR *rgwchSrc, int cchSrc, _Out_z_cap_(cchDest) WCHAR *wtzDest, int cchDest);
MSOAPI_(WCHAR *) MsoRgwchWtAppend(const WCHAR *rgwchSrc, int cchSrc, _Out_z_cap_(cchDest) WCHAR *wtDest, int cchDest);

#if (OACR)
__inline WCHAR * MsoWzRgwchCchCopy(const WCHAR *rgwchFrom, __zcount(cch) WCHAR *wzTo, int cch)
{
	return MsoRgwchCopy(rgwchFrom, cch, wzTo, (cch) + 1);
}
#else
// WARNING! this is unsafe.  There is no target buffer
#define MsoWzRgwchCchCopy(rgwchFrom, wzTo, cch) MsoRgwchCopy(rgwchFrom, cch, wzTo, (cch) + 1)
#endif


/* Allocate a copy of the string on the heap.  Free using MsoPvFree */
#ifdef DEBUG
#define MsoSzClone(sz, msodg) MsoSzCloneCore(sz, msodg, __FILE__, __LINE__)
#define MsoWzClone(wz, msodg) MsoWzCloneCore(wz, msodg, __FILE__, __LINE__)
MSOAPI_(CHAR *) MsoSzCloneCore(const CHAR *sz, int msodg, const char *szFile, int iLine);
MSOAPI_(WCHAR *) MsoWzCloneCore(const WCHAR *wz, int msodg, const char *szFile, int iLine);
#else
MSOAPI_(CHAR *) MsoSzClone(const CHAR *sz, int msodg);
MSOAPI_(WCHAR *) MsoWzClone(const WCHAR *wz, int msodg);
#endif

/* get text part of length-preceeded strings */

/* Returns a pointer to the text part of the string wtz */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
#if defined __cplusplus
extern "C++"
{
__inline const WCHAR* MsoWzFromWtz(     const WCHAR* wtz) { return wtz+1; }
__inline       WCHAR* MsoWzFromWtz(__no_count WCHAR* wtz) { return wtz+1; }
}
#else
__inline const WCHAR* MsoWzFromWtz(const WCHAR* wtz) { return wtz+1; }
#endif
__inline WCHAR* MsoWzFromWtzNonConst(__no_count WCHAR* wtz) { return wtz+1; }
#else
__inline WCHAR* MsoWzFromWtz(const WCHAR* wtz) { return (WCHAR*)(wtz+1); }
#endif

/* Returns a pointer to the text part of the string wt */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
#if defined __cplusplus
extern "C++"
{
__inline const WCHAR* MsoRgwchFromWt(     const WCHAR* wt) { return wt+1; }
__inline       WCHAR* MsoRgwchFromWt(__no_count WCHAR* wt) { return wt+1; }
}
#else
__inline const WCHAR* MsoRgwchFromWt(const WCHAR* wt) { return wt+1; }
#endif
__inline WCHAR* MsoRgwchFromWtNonConst(__no_count WCHAR* wt) { return wt+1; }
#else
__inline WCHAR* MsoRgwchFromWt(const WCHAR* wt) { return (WCHAR*)(wt+1); }
#endif

/* Returns the length of a wt string */
__inline WCHAR MsoCwchFromWt(const WCHAR *wt) { return wt[0]; }

/* Returns the length of a wtz string */
#define MsoCwchFromWtz(wtz) MsoCwchFromWt(wtz)

/* Checks if wtz is a valid string */
#if DEBUG
	MSOAPIXX_(BOOL) MsoFValidWtz(const WCHAR* wtz);
#else
	#define MsoFValidWtz(wtz) (1)
#endif


/*************************************************************************
	String comparisons
*************************************************************************/

enum
{
	msocsExact = 0x4,	/* exact string matches required */

	msocsCase = 0x4,	/* case-sensitive comparisons */
	msocsIgnoreCase = 0x1,	/* case-insensitive comparisons */

	msocsDiacriticals = 0x4,	/* diacritical differences compare different */
	msocsIgnoreDiacriticals = 0x2,	/* diacriticals ignored */
	msocsIgnoreKana = 0x8,
	msocsIgnoreWidth = 0x10,
	msocsIgnoreNonSpace = 0x20
};

/* Test if the null-terminated strings sz1 and sz2 are the same. */
MSOAPI_(BOOL) MsoFSzEqual(const CHAR* sz1, const CHAR* sz2, int cs);

MSOAPI_(BOOL) MsoFWzEqual(const WCHAR* wz1, const WCHAR* wz2, int cs);

MSOAPI_(BOOL) MsoFWtzEqual(const WCHAR* wtz1, const WCHAR* wtz2, int cs);

/* Test if the two strings rgch1 (length cch1) and rgch2 (length cch2) are
	the same. */
MSOAPI_(BOOL) MsoFRgchEqual(const CHAR* rgch1, int cch1, const CHAR* rgch2, int cch2, int cs);

/*	Test if the two Unicode strings rgwch1 (length cch1) and rgwch2
	(length cch2) are the same. */
MSOAPI_(BOOL) MsoFRgwchEqual(const WCHAR* rgch1, int cch1, const WCHAR* rgch2, int cch2, int cs);

/* Compares the sort order of the two strings sz1 and sz2.  Returns
	-1 if sz1 < sz2, 0 if they sort equally, and +1 if sz1 > sz2. */
MSOAPI_(int) MsoSgnSzCompare(const CHAR* sz1, const CHAR* sz2, int cs);

MSOAPI_(int) MsoSgnWzCompare(const WCHAR* wz1, const WCHAR* wz2, int cs);

MSOAPI_(int) MsoSgnWtzCompare(const WCHAR* wtz1, const WCHAR* wtz2, int cs);

/* Compares the sort order of the two strings rgch1 (length cch1)
   and rgch2 (length cch2).  Returns -1 if rgch1 < rgch2, 0 if they
   sort equally, and +1 if rgch1 > rgch2. */
MSOAPI_(int) MsoSgnRgchCompare(const CHAR* rgch1, int cch1, const CHAR* rgch2, int cch2, int cs);

MSOAPI_(int) MsoSgnRgwchCompare(const WCHAR* rgch1, int cch1, const WCHAR* rgch2, int cch2, int cs);

/*	Compares the sort order of the two unicode strings rgwch1
	(# of chars cch1) and rgwch2 (# of chars cch2).  Returns -1 if
	rgwch1 < rgwch2, 0 if they sort equally, and +1 if rgwch1 > rgwch2.
	The string compare is based on locale - identified by the sort
	and language IDs. */
MSOAPIMX_(int) MsoSgnRgwchCompareLoc(const WCHAR* rgch1, int cch1, const WCHAR* rgch2, int cch2, int cs, WORD wLangId, WORD wSortId);


/*************************************************************************
	Substrings and searches
*************************************************************************/

/*	Finds the first occurance of the character ch in the null-terminated
	string sz, and returns a pointer to it.  Returns NULL if the character
	was not in the string. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const CHAR*) MsoSzIndex(const CHAR* sz, int ch);
#else
MSOAPI_(CHAR*) MsoSzIndex(const CHAR* sz, int ch);
#endif
__inline CHAR* MsoSzIndexNonConst(__no_count CHAR* sz, int ch)
{
	OACR_USE_PTR(sz); return (CHAR*)MsoSzIndex(sz, ch);
}

/*	Finds the last (rightmost) occurance of the character ch in the null-terminated
	string sz, and returns a pointer to it.  Returns NULL if the character
	was not in the string. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const CHAR*) MsoSzIndexRight(const CHAR* sz, int ch);
#else
MSOAPI_(CHAR*) MsoSzIndexRight(const CHAR* sz, int ch);
#endif
__inline CHAR* MsoSzIndexRightNonConst(__no_count CHAR* sz, int ch)
{
	OACR_USE_PTR(sz); return (CHAR*)MsoSzIndexRight(sz, ch);
}

/*  Finds the first occurance of one of szCharSet in the null-terminated
	string sz, and returns a pointer to it.  Returns NULL if none of the characters
	was not in the string.

	Does not work on DBCS */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPIX_(const CHAR*) MsoSzIndexOneOf(const CHAR* sz, const CHAR *szCharSet);
#else
MSOAPIX_(CHAR*) MsoSzIndexOneOf(const CHAR* sz, const CHAR *szCharSet);
#endif
__inline CHAR* MsoSzIndexOneOfNonConst(__no_count CHAR* sz, const CHAR *szCharSet)
{
	OACR_USE_PTR(sz); return (CHAR*)MsoSzIndexOneOf(sz, szCharSet);
}

#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoWzIndexOneOf(const WCHAR* wz, const WCHAR *szCharSet);
#else
MSOAPI_(WCHAR*) MsoWzIndexOneOf(const WCHAR* wz, const WCHAR *szCharSet);
#endif
__inline WCHAR* MsoWzIndexOneOfNonConst(__no_count WCHAR* wz, const WCHAR* wzCharSet)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoWzIndexOneOf(wz, wzCharSet);
}

/*	Finds the last (rightmost) occurance of the character wch in the null-terminated
	UNINCODE  string cwz, and returns a pointer to it.  Returns NULL if the character
	was not in the string. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoWzIndexRight(const WCHAR* wz, int wch);
#else
MSOAPI_(WCHAR*) MsoWzIndexRight(const WCHAR* wz, int wch);
#endif
__inline WCHAR* MsoWzIndexRightNonConst(__no_count WCHAR* wz, int wch)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoWzIndexRight(wz, wch);
}

/*	Finds the last (rightmost) occurance of the character wch in the null-terminated
	UNINCODE  string wz starting from char position ich, and returns a pointer to it.
	Returns NULL if the character was not in the string. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoCchWzIndexRight(const WCHAR* wz, INT_PTR ich, int wch);
#else
MSOAPI_(WCHAR*) MsoCchWzIndexRight(const WCHAR* wz, INT_PTR ich, int wch);
#endif
__inline WCHAR* MsoCchWzIndexRightNonConst(__no_count WCHAR* wz, INT_PTR ich, int wch)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoCchWzIndexRight(wz, ich, wch);
}

/* Replaces all occurances of wchOut with wchIn. */
MSOAPI_(void) MsoReplaceAllOfWchWithWch (__no_ecount WCHAR *wz, WCHAR wchOut, WCHAR wchIn);
/* Replaces all occurances of pwzOut with pwzIn, and return the new allocated string*/
MSOAPI_(WCHAR *) MsoReplaceAllOfWzWithWz(const WCHAR *wz, const WCHAR *pwzOut, const WCHAR *pwzIn, int cs);

/*	Just like MsoSzIndex, except the string is guaranteed to be full of
	single-byte ANSI characters (i.e., no DBCS characters) */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPIMX_(const CHAR*) MsoSzIndex1(const CHAR* sz, int ch);
#else
MSOAPIMX_(CHAR*) MsoSzIndex1(const CHAR* sz, int ch);
#endif
__inline CHAR* MsoSzIndex1NonConst(__no_count CHAR* sz, int ch)
{
	OACR_USE_PTR(sz); return (CHAR*)MsoSzIndex1(sz, ch);
}

/*	Finds the first occurance of the character ch in the null-terminated
	string wz, and returns a pointer to it.  Returns NULL if the character
	was not in the string. */
MSOAPI_(const WCHAR*) MsoWzIndex(const WCHAR* wz, int ch);
__inline WCHAR* MsoWzIndexNonConst(__no_count WCHAR* wz, int ch)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoWzIndex(wz,ch);
}

/* Finds the occurance of the substring szFind inside the string sz, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const CHAR*) MsoSzStrStr(const CHAR* sz, const CHAR* szFind);
#else
MSOAPI_(CHAR*) MsoSzStrStr(const CHAR* sz, const CHAR* szFind);
#endif
__inline CHAR* MsoSzStrStrNonConst(__no_count CHAR* sz, const CHAR* szFind)
{
	OACR_USE_PTR(sz); return (CHAR*)MsoSzStrStr(sz,szFind);
}

/*	Finds the occurance of the substring szFind inside the string sz, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found

	This version pays attention to cs and does not require NULL termination. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const CHAR*) MsoPchStrStr(const CHAR* pch, int cch, const CHAR* szFind, int cs);
#else
MSOAPI_(CHAR*) MsoPchStrStr(const CHAR* pch, int cch, const CHAR* szFind, int cs);
#endif
__inline CHAR* MsoPchStrStrNonConst(_In_z_ CHAR* pch, int cch, const CHAR* szFind, int cs)
{
	OACR_USE_PTR(pch); return (CHAR*)MsoPchStrStr(pch, cch, szFind, cs);
}

/*	Finds the occurance of the substring pwchFind inside the string pwch, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found. This version pays attention to cs and does not require NULL
	termination. */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPIX_(const WCHAR *) MsoPwchStrStr(const WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind, int cs);
#else
MSOAPIX_(WCHAR *) MsoPwchStrStr(const WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind, int cs);
#endif
__inline WCHAR* MsoPwchStrStrNonConst(_In_z_ WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind, int cs)
{
   OACR_USE_PTR(pwch); return (WCHAR*)MsoPwchStrStr(pwch, cwch, pwchFind, cwchFind, cs);
}

/*	Finds the occurance of the substring pwchFind inside the string pwch, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoPwchStrStrFast(const WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind);
#else
MSOAPI_(WCHAR*) MsoPwchStrStrFast(const WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind);
#endif
__inline WCHAR* MsoPwchStrStrFastNonConst(_In_z_ WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind)
{
   OACR_USE_PTR(pwch); return (WCHAR*)MsoPwchStrStrFast(pwch, cwch, pwchFind, cwchFind);
}
/*	Finds the last (rightmost) occurance of the substring pwchFind inside the
	string pwch, and returns a pointer to the beginning of it.
	Returns NULL if the substring is not found */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoPwchStrStrRightFast(const WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind);
#else
MSOAPI_(WCHAR*) MsoPwchStrStrRightFast(const WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind);
#endif
__inline WCHAR* MsoPwchStrStrRightFastNonConst(_In_z_ WCHAR *pwch, int cwch, const WCHAR *pwchFind, int cwchFind)
{
   OACR_USE_PTR(pwch); return (WCHAR*)MsoPwchStrStrRightFast(pwch, cwch, pwchFind, cwchFind);
}

/*	Finds the occurance of the substring wzFind inside the string wz, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found  */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoWzStrStr(const WCHAR* wz, const WCHAR* wzFind);
#else
MSOAPI_(WCHAR*) MsoWzStrStr(const WCHAR* wz, const WCHAR* wzFind);
#endif
__inline WCHAR* MsoWzStrStrNonConst(__no_count WCHAR* wz, const WCHAR* wzFind)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoWzStrStr(wz, wzFind);
}

/* Ex version lets you specify case sensitive vs. insensitive etc */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPI_(const WCHAR*) MsoWzStrStrEx(const WCHAR* wz, const WCHAR* wzFind, int cs);
#else
MSOAPI_(WCHAR*) MsoWzStrStrEx(const WCHAR* wz, const WCHAR* wzFind, int cs);
#endif
__inline WCHAR* MsoWzStrStrExNonConst(__no_count WCHAR* wz, const WCHAR* wzFind, int cs)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoWzStrStrEx(wz, wzFind, cs);
}

/* Ex version lets you specify cch in wzFind to match against, and case sensitive vs. insensitive etc */
#if ( MSO_CONST_CLEAN_STR_FUNCTIONS )
MSOAPIX_(const WCHAR*) MsoWzStrStrEx2(const WCHAR* wz, const WCHAR* wzFind, int cchFind, int cs);
#else
MSOAPIX_(WCHAR*) MsoWzStrStrEx2(const WCHAR* wz, const WCHAR* wzFind, int cchFind, int cs);
#endif
__inline WCHAR* MsoWzStrStrEx2NonConst(__no_count WCHAR* wz, const WCHAR* wzFind, int cchFind, int cs)
{
	OACR_USE_PTR(wz); return (WCHAR*)MsoWzStrStrEx2(wz, wzFind, cchFind, cs);
}

/*	Given a pattern string with special token symbols in it, inserts
	other strings in place of the token, based on a variable argument
	list.  The token "|0" is replaced by the 0th variable arg, "|1" is
	replaced by the 1st variable arg, "|2" by the 2nd variable arg, etc.
	Especially useful for international error messages, where the order
	of inserted strings may change depending on the language.

	If the "|" character is not followed by a number, inserts the 0th
	variable arg.

	The pattern string is passed in in szPat, the result is stored in
	the buffer szOut, which is of length cchOut. */

#define __stdcall __cdecl
// MSOAPI_(int) MsoInsertSz(CHAR* szOut, int cchOut, const CHAR* szPat, const CHAR* sz0, ...);
// MSOMACPUB int __cdecl MsoInsertWtz(WCHAR* wtzOut, int cchOut, const WCHAR* wtzPat, const WCHAR* wtz0, ...);
// MSOAPI_(int) MsoInsertWz(WCHAR* wzOut, int cchOut, const WCHAR* wzPat, const WCHAR* wz0, ...);
// MSOAPIX_(int) MsoInsertIdsWtz(WCHAR* wtzOut, int cchOut, HINSTANCE hinst, int idsPat, const WCHAR* wtz0, ...);

#if defined(FRONTPAGE_BUILD) || defined(OFFICE_BUILD) || defined(LIME_BUILD)
MSOAPI_(int) MsoCchInsertSz    ( _Out_z_cap_(cchOut) CHAR* szOut , int cchOut, const  CHAR* szPat , int iCount, ...);
#endif
MSOAPIX_(int) MsoCchInsertWtz   (_Out_z_cap_(cchOut) WCHAR* wtzOut, int cchOut, const WCHAR* wtzPat, int iCount, ...);
MSOAPI_(int) MsoCchInsertWz    (_Out_z_cap_(cchOut) WCHAR* wzOut , int cchOut, const WCHAR* wzPat , int iCount, ...);
MSOAPIX_(int) MsoCchInsertIdsWtz(_Out_z_cap_(cchOut) WCHAR* wtzOut, int cchOut, HINSTANCE hinst, int idsPat, int iCount, ...);
MSOAPI_(int) MsoCchInsertIdsWz (_Out_z_cap_(cchOut) WCHAR* wtzOut, int cchOut, HINSTANCE hinst, int idsPat, int iCount, ...);

#undef __stdcall


/*************************************************************************
	Unicode and character set string conversions
*************************************************************************/

/* converts a wide unicode string szFrom into the system's native
	default character set in szTo */
MSOAPI_(int) MsoWzToSz(const WCHAR* wzFrom, _Out_z_cap_(cchMax) CHAR* szTo, int cchMax);
MSOAPIX_(void) MsoWzToSzSimple(const WCHAR* wzFrom, _Out_z_cap_(cchMax) CHAR* szTo, int cchMax);
MSOAPI_(int) MsoWtzToSz(const WCHAR* wtzFrom, _Out_z_cap_(cchMax) CHAR* szTo, int cchMax);

/* Any particular reason we're missing MsoWtzToStz? */

/* converts a system's native character set string szFrom into a
	wide unicode string in szTo */
MSOAPI_(int) MsoSzToWz(const CHAR *sz, _Out_z_cap_(cch) WCHAR *wz, int cch);
#define MsoSzToWzCch(szFrom, wzTo, cchTo) (MsoSzToWz(szFrom, wzTo, cchTo)+1)

MSOAPIX_(void) MsoSzToWzSimple(const CHAR* sz, _Out_z_cap_(cchMax) WCHAR* wz, int cchMax);
MSOAPI_(int) MsoCpSzToWz(UINT cp, const CHAR* sz, _Out_z_cap_(cchMax) WCHAR* wz, int cchMax);
MSOAPIX_(WCHAR) MsoWchFromCpCh(UINT cp, const CHAR ch);

/*	Converts the unicode string wtz into the Ansi */
MSOAPI_(int) MsoSzToWtz(const CHAR* sz, _Out_z_cap_(cchMax) WCHAR* wtz, int cchMax);
MSOAPI_(int) MsoWtzToWz(const WCHAR* wtz, _Out_z_cap_(cchMax) WCHAR* wz, int cchMax);
MSOAPI_(int) MsoWzToWtz(const WCHAR* wz, _Out_z_cap_(cchMax) WCHAR* wtz, int cchMax);

/*	Converts the Unicode string in rgwchFrom into an Ansi string at
	rgchTo.  This API will work in-place.
	Returns FALSE if the target buffer is not large enough for the
	conversion.  Returns TRUE on success. */
MSOAPI_(BOOL) MsoFRgwchToRgch(const WCHAR* rgwchFrom, _Out_z_cap_(cbTo) CHAR* rgchTo, int cbTo);

/* converts a wide unicode string in rgchFrom (length cchFrom) into the
	system's native default character set in rgchTo */
MSOAPI_(int) MsoRgwchToRgch(const WCHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) CHAR* rgchTo, int cchTo);

/*	Converts a unicode string rgchFrom (with length cchFrom) into the
	equivalent string in the specified character set. If rgchTo
	is non-NULL, the converted string is stored there.  The length in
	bytes of the converted string is returned, even if rgchTo is NULL. */
MSOAPI_(int) MsoRgwchToCpRgch(UINT CodePage, const WCHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) CHAR* rgchTo, int cchTo);
MSOAPI_(int) MsoRgwchToCpRgchEx(UINT CodePage, const WCHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) CHAR* rgchTo, int cchTo, BOOL *fDefault);

/*	Converts the Ansi string in rgchFrom with length cchFrom into a
	Unicode string at rgchTo.  The rgchTo buffer is assumed to be cchTo
	character slong.  Returns the length of the converted string.  This
	API will work in-place. */
MSOAPI_(int) MsoRgchToRgwch(const CHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) WCHAR* rgchTo, int cchTo);

/*	Converts the MultiByte string of code page CodePage in rgchFrom with
	length cchFrom into a Unicode string at rgwchTo.  The rgwchTo buffer
	is assumed to be cchTo characters long.  Returns the length of the
	converted string.  This API will work in-place. */
MSOAPI_(int) MsoCpRgchToRgwch(UINT CodePage, const CHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) WCHAR* rgwchTo, int cchTo);

/*	Converts the Ansi string in rgchFrom with length cchFrom into a
	Unicode string at rgwchTo.  This API will work in-place.
	Returns FALSE if the target buffer is not large enough for the
	conversion.  Returns TRUE on success. */
MSOAPI_(BOOL) MsoFRgchToRgwch(const CHAR* rgchFrom, _Out_z_cap_(cchTo) WCHAR* rgwchTo, int cchTo);

/* converts a system's native character set string rgchFrom (with
	length cchFrom) into a wide unicode string in rgchTo */
#ifdef DEBUG
#define MsoSzMarkWtz(wtz) MsoSzMarkWtzCore(wtz, __FILE__, __LINE__)
MSOAPIX_(CHAR*) MsoSzMarkWtzCore(const WCHAR* wtz, const CHAR* szFile, int line);
#else
#define MsoSzMarkWtz(wtz) MsoSzMarkWtzCore(wtz)
MSOAPIX_(CHAR*) MsoSzMarkWtzCore(const WCHAR* wtz);
#endif

/*	Converts the unicode string wz into the Ansi, allocated from the
	mark memory manager.  The string should be freed using MsoReleaseSz. */
#ifdef DEBUG
#define MsoSzMarkWz(wz) MsoSzMarkWzCore(wz, __FILE__, __LINE__)
MSOAPI_(CHAR*) MsoSzMarkWzCore(const WCHAR* wz, const CHAR* szFile, int line);
#else
#define MsoSzMarkWz(wz) MsoSzMarkWzCore(wz)
MSOAPI_(CHAR*) MsoSzMarkWzCore(const WCHAR* wz);
#endif

/*	Releases the string sz allocated by MsoSzMarkWtz, MsoSzMarkWz, et.al. */
#define MsoReleaseSz(sz) MsoReleaseMem(sz)

/*	Translates an ANSI (single/double byte character set) character string
	sz in the system default codepage into a Unicode string.  The unicode
	string is allocated out of the mark/release memory heap. */
#ifdef DEBUG
#define MsoWzMarkSz(sz) MsoWzMarkSzCore(sz, __FILE__, __LINE__)
MSOAPI_(WCHAR*) MsoWzMarkSzCore(const CHAR* sz, const CHAR* szFile, int line);
#else
#define MsoWzMarkSz(sz) MsoWzMarkSzCore(sz)
MSOAPI_(WCHAR*) MsoWzMarkSzCore(const CHAR* sz);
#endif

/*	Translates an ANSI (single byte character set, all chars < 128) character
	string sz in the system default codepage into a Unicode string.  The
	unicode string is allocated out of the mark/release memory heap.
					THIS IS NOT FOR DBCS!!!!     */
#ifdef DEBUG
#define MsoWzMarkSzSimple(sz) MsoWzMarkSzSimpleCore(sz, __FILE__, __LINE__)
MSOAPIX_(WCHAR*) MsoWzMarkSzSimpleCore(const CHAR* sz, const CHAR* szFile, int line);
#else
#define MsoWzMarkSzSimple(sz) MsoWzMarkSzSimpleCore(sz)
MSOAPIX_(WCHAR*) MsoWzMarkSzSimpleCore(const CHAR* sz);
#endif


/*	Releases a Unicode string allocated by MsoWzMarkSz or MsoWtzMarkSz. */
#define MsoReleaseWz(wz) MsoReleaseMem(wz)

/*	conversion routines for OLECHAR to WCHAR - note that if an OLECHAR and
	a WCHAR are the same thing, these routines to nothing */
	#define MsoOzMarkWz(wz) ((OLECHAR*)(wz))
	#define MsoReleaseOzWz(oz)
	#define MsoWzMarkOz(oz) ((WCHAR*)(oz))
	#define MsoReleaseWzOz(wz)


/*************************************************************************
	Miscellaneous string utilities
*************************************************************************/

/*	Decodes the unsigned integer u into ASCII text in base wBase.  The
	string is stored in the rgch buffer, which is assumed to be large
	enough to hold the number's text and a null terminator.  Returns
	the length of the text decoded. */
MSOAPI_(int) MsoSzDecodeUint(_Out_z_cap_(cch) CHAR* rgch, int cch, unsigned u, int wBase);

/*	Decodes the unsigned u into Unicode text in base wBase.  The string is
	stored in the rgwch buffer, which is assumed to be large enough to
	hold the number's text and a null terminator.  Returns the length
	of the text decoded. */
MSOAPI_(int) MsoWzDecodeUint(_Out_z_cap_(cch) WCHAR* rgwch, int cch, unsigned u, int wBase);

/* Decodes the integer w into Unicode text in base wBase, filling any
   unused space with leading zeros.  Returns number of significant digits. */
MSOAPI_(int) MsoWzDecodeUIntFill(_Out_z_cap_(cch) WCHAR* rgwch, int cch, unsigned u, int wBase);

/*	Decodes the integer w into Unicode text in base wBase.  The string is
	stored in the rgwch buffer, which is assumed to be large enough to
	hold the number's text and a null terminator.  Returns the length
	of the text decoded. */
MSOAPI_(int) MsoWzDecodeInt(_Out_z_cap_(cch) WCHAR* rgwch, int cch, int w, int wBase);

/*	Decodes the signed integer w into ASCII text in base wBase.  The
	string is stored in the rgch buffer, which is assumed to be large
	enough to hold the number's text and a null terminator.  Returns
	the length of the text decoded. */
MSOAPI_(int) MsoSzDecodeInt(_Out_z_cap_(cch) CHAR* rgch, int cch, int w, int wBase);

/* Parses the Unicode text at rgwch into *pw. Returns the count of
   characters considered part of the number. Continues reading characters
   and encoding them into *pw until it encounters a non-digit.
   Handles overflow by returning zero.*/
//#if defined (ZENSTAT_LIB_DEF) || !defined (STATIC_LIB_DEF)
#if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
MSOAPI_(int) MsoParseInt64Wz(const WCHAR* rgwch, __int64 *pw);
MSOAPI_(int) MsoParseUInt64Wz(const WCHAR* rgwch, unsigned __int64 *pw);
#endif

/* Parses the Unicode text at rgwch into *pw. Returns the count of
   characters considered part of the number.
   Returns 0 on overflow */
MSOAPI_(int) MsoParseIntWz(const WCHAR* rgwch, int *pw);
// Temporary name mapping -- we'll remove this API soon
#define MsoIntEncodeWz MsoParseIntWz

/* Parses the Unicode text at rgwch into *pw. Returns the count of
   characters considered part of the number.
   Returns 0 on overflow */
MSOAPI_(int) MsoParseUIntWz(const WCHAR* rgwch, unsigned *pw);

/*	Parses the Ansi text at rgch into *pw. Returns the count of
   characters considered part of the number. Continues reading
   characters and encoding them into *pw until it encounters a
   non-digit.  Returns 0 on overflow */
MSOAPI_(int) MsoParseIntSz(const CHAR* rgch, int *pw);

/*	Parses the Ansi text at rgch into *pw. Returns the count of
   characters considered part of the number. Continues reading
   characters and encoding them into *pw until it encounters a
   non-digit.  Returns 0 on overflow */
MSOAPI_(int) MsoParseUIntSz(const CHAR* rgch, unsigned *pw);

// Temporary name mapping -- we'll remove this API soon
#define MsoIntEncodeSz MsoParseIntSz

MSOAPI_(int) MsoParseHexIntWz(const WCHAR *wz, int *pw);
MSOAPIX_(int) MsoParseHexIntSz(const CHAR  *sz, int *pw);

/* Parses the Unicode text at rgwch into *pw. Returns the count of
   characters considered part of the number.
   Returns 0 on overflow */
MSOAPI_(int) MsoParseHexUIntWz(const WCHAR *wz, unsigned *pw);

/*	Expands MSO-environment strings and replaces them with their defined
	values.  Subsequently calls the system API to expand and replace any
	remaining environment-variable strings. */
MSOAPI_(DWORD) MsoExpandEnvironmentStrings(LPCWSTR lpSrc, _Out_cap_(cchDest) LPWSTR lpDest, DWORD cchDest);

/*  Allocates a new wz which is a copy of the given one. */
MSOAPI_(WCHAR*) MsoWzCreateFromWz(const WCHAR *wz);

/*************************************************************************
	String table resources
*************************************************************************/

/* String identifiers and their associated pieces */

#define MsoIdsFromSttIdsl(stt,idsl) ((long)MAKELONG(idsl,stt))
#define MsoIdslFromIds(ids) LOWORD(ids)
#define MsoSttFromIds(ids) HIWORD(ids)

#define msoidslNil 0xFFFFU
#define msosttNil  0xFFFFU
#define msoidsNil MsoIdsFromSttIdsl(msosttNil,msoidslNil)


/* The string table resource format */

enum
{
	msotstNil = 0,	/* invalid string table type */
	msotstAllocated = 0,	/* allocated index table type */
	msotstFixed = 1,	/* fixed index table type */
	msotstNoCompress = 2,	/* no compression */
	msotstSlow = 4,	/* slow reverse lookups */
	msotstFast = 8,	/* fast O(ln n) reverse lookups */
	msotstCaseSensitive = 16,	/* case sensitive reverse lookups */
	msotstSimple = 32,	/* simple (no index) tables */
	msotstCompress = 64, /* simple compress if SBCS charset, no 2nd byte */
	msotstHuffCompress = 128 /* huffman and substring compression */
};

typedef struct MSOFLOM
{
	WORD idsl;	/* string index */
	WORD bwt;	/* offset to beginning of string */
} MSOFLOM;

typedef struct MSOSTT
{
	WORD tst;	/* type of string table */
	WORD lang;	/* NLS language ID */
	WORD sort;	/* NLS sorting ID */
	union
		{
		struct	/* tstAllocated */
			{
			WORD idslMac;	/* number of strings */
			WORD mpidslbwt[1];	/* idslMac entries */
			};
		struct	/* tstFixed */
			{
			WORD cflom;	/* number of strings */
			MSOFLOM rgflom[1];	/* cflom entries */
			};
		};
	/* for tstFast, there is a rgflomSorted array here */
	/* the actual grwt data follows */
} MSOSTT;


// Huffman Table Resource format
typedef struct MSODSTE
{   // Decode State Table Entry
	WORD mpbidste[2];
	WCHAR rgwchKey[4];
} MSODSTE;

#define msobwtNil 0xFFFFU

#define __stdcall __cdecl

/* Loads a non-OSTRMAN string with id ids from the resource file
	specified by hinst.  result is left in sz, which is a buffer big
	enough to hold of cch Ansi characters.  Returns # of chars
	if successful. */
MSOAPIX_(int) MsoCchStdLoadSz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) char *sz, int cch);

/* Loads a non-OSTRMAN string with id ids from the resource file
	specified by hinst.  result is left in wz, which is a buffer big
	enough to hold of cch Unicode characters.  Returns # of chars
	if successful. */
MSOAPIX_(int) MsoCchStdLoadWz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) WCHAR* wz, int cch);

#undef __stdcall

/*	Loads a string with id ids from the resource file specified by hinst.
	result is left in wtz, which is a buffer big enough to hold of cch
	Unicode characters.  Returns TRUE if successful */
MSOAPI_(BOOL) MsoFLoadWtz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) WCHAR* wtz, int cch);

/*	Loads a string with id ids from the resource file specified by hinst.
	result is left in wz, which is a buffer big enough to hold of cch
	Unicode characters.  Returns TRUE if successful */
MSOAPI_(BOOL) MsoFLoadWz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) WCHAR* wz, int cch);

/*	Loads a string with id ids from the resource file specified by hinst.
	result is left in wz, which is a buffer big enough to hold of cch
	Unicode characters.  Returns # of chars if successful */
MSOAPI_(int) MsoCchLoadWz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) WCHAR* wz, int cch);

/*	Loads a string with id ids from the resource file specified by hinst.
	result is left in rgch, which is a buffer of length cch.  Returns length
	of the string if found, FALSE if not.  The returned string is
	guaranteed to be null-terminated */
MSOAPI_(int) MsoCchLoadSz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) CHAR* rgch, int cch);

/*	Loads a string with id ids from the resource file specified by hinst.
	result is left in rgch, which is a buffer of length cch.  Returns the
	actual length of the string.  The returned string is guaranteed to be
	null-terminated */
MSOAPI_(BOOL) MsoFLoadSz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) CHAR* sz, int cch);


/* Expand Ids Matching Struct.  Refer to the comments for MsoFExpandIdsWz for details. */
typedef struct _MSOEIMS
{
	const WCHAR* wzPlaceholder;
	int cchPlaceholder;
	const WCHAR* wzTextToSubst;
	int cchTextToSubst;
} MSOEIMS;

#define WZ_DEFAULT_PLACEHOLDER		L"%1"
#define CCH_DEFAULT_PLACEHOLDER		2

/* this is similar to MsoFExpandIdsWz, but it can be used on arbitrary strings.  Populate wzResult with
     the string which contains the substitution codes. */
MSOAPI_(void) MsoExpandWz(_Out_cap_(cchWzMax) WCHAR* wzResult, int cchWzMax, MSOEIMS* rgeims, int ceims, int msocs);
/* this is similar to MsoFLoadWz, but it should be used for ids values which have
   "substitution codes."  See the comments on the actual function in the MSO code for
   details. */
MSOAPI_(BOOL) MsoFExpandIdsWz(HINSTANCE hinst, int ids, _Out_cap_(cchWzMax) WCHAR* wzResult, int cchWzMax, MSOEIMS* rgeims, int ceims);


MSOAPIX_(CHAR*) MsoMarkLoadStz(HINSTANCE hinst, int ids);
MSOAPIX_(WCHAR*) MsoMarkLoadWtz(HINSTANCE hinst, int ids);

#define msoStrip	0x0001
#define msoCaps	0x0002
/*	This function performs upper case conversion for a given
	Unicode character given its language id.
	Returns the input Unicode character if no conversion is found.
	Currently works only for the 5 Win and 5 Mac code pages needed for the
	Pan European version of the apps.
	wflags has the following bits that are currently used -
	msoStrip if you want to strip all the accents */
MSOAPI_(WCHAR) MsoWchToUpperLid(WCHAR wch, LID lid, WORD wflags);

/*	This function performs upper case conversion for a given
	Unicode character. It is invariant with respect to lcids and
	must be kept this way because word needs an invariant upper casing. */
MSOAPI_(WCHAR) MsoWchToUpperInvariant(WCHAR wch);

/* This function performs upper case conversion for a given
	Unicode character. Returns the input Unicode character if no
	conversion is found. Currently works only for the 10 code pages
	needed for Pan European version of the apps. */
MSOAPI_(WCHAR) MsoWchToUpper(WCHAR wch);

/*	This function strips the accent of WinLatin1 and Greek characters.
	Please note that European Latin, Extended Latin, and Cyrillic
	characters are unaffected. */
MSOAPI_(WCHAR) MsoWchStripAccent(WCHAR wch);

/*	This function performs lower case conversion for a given
	Unicode character given its language id.
	Currently works only for the 5 Win and 5 Mac code pages needed for the
	Pan European version of the apps.
	The wflags param is unused currently. */
MSOAPI_(WCHAR) MsoWchToLowerLid(WCHAR wch, LID lid, WORD wflags);

/* This function performs lower case conversion for a given
	Unicode character. Returns the input Unicode character if no
	conversion is found. Currently works only for the 10 code pages
	needed for Pan European version of the apps. */
MSOAPI_(WCHAR) MsoWchToLower(WCHAR wch);

/*	This function performs upper case conversion for a given wz.
	Note:  If the input string size exceeds cchMaxSz then it is truncated*/
MSOAPI_(void) MsoWzUpper(__no_ecount WCHAR * wz);
MSOAPI_ (void) MsoPwchUpper(_Out_cap_(cch) WCHAR *pwch, int  cch);

/*	This function performs lower case conversion for a given wz.
	Note:  If the input string size exceeds cchMaxSz then it is truncated*/
MSOAPI_(void) MsoWzLower(__no_ecount WCHAR *wz);
MSOAPI_ (void) MsoPwchLower(_Out_cap_(cch) WCHAR *pwch, int  cch);

#ifdef OFFICE_BUILD
/*	This function performs upper case conversion for a given sz.
	Note:  If the input string size exceeds cchMaxSz then it is truncated */
MSOAPIX_(void) MsoSzUpper(__no_ecount CHAR *sz);
#endif

/* This function performs lower case conversion for a given sz.
	Note:  If the input string size exceeds cchMaxSz then it is truncated */
MSOAPI_(void) MsoSzLower(__no_ecount CHAR *sz);

/* These functions convert between Unicode and the UTF8 codepage, including
	surrogate support. See implementation for more detailed comments. */
MSOAPI_(int) UTF8ToUnicode(LPCSTR lpSrcStr, int *cchSrc, _Out_cap_(cchDest) LPWSTR lpDestStr,
	int cchDest);
MSOAPI_(int) UnicodeToUTF8(LPCWSTR lpSrcStr, int cchSrc, _Out_cap_(cchDest) LPSTR lpDestStr,
	int cchDest);
MSOAPIX_(int) UnicodeToUTF8Core(LPCWSTR lpSrcStr, int cchSrc, int *pcchSrcLeft,
	_Out_cap_(cchDest) LPSTR lpDestStr, int cchDest);

/*	This function provides our own layer to MultiByteToWideChar() call.  We
	provide our own lookup table if it's not available from NLS subsystem.
	Currently works for 20 SBCS Win and 5 Mac code pages. */
MSOAPI_(int) MsoMultiByteToWideChar(UINT CodePage, DWORD dwFlags,
		LPCSTR lpMultiByteStr, INT_PTR cchMultiByte, _Out_opt_cap_(cchWideChar) LPWSTR lpWideCharStr,
		int cchWideChar);

/*	This function provides our own layer to MultiByteToWideChar() call.  We
	provide our own lookup table if it's not available from NLS subsystem.
	Currently works for 20 SBCS Win and 5 Mac code pages. */
MSOAPI_(int) MsoWideCharToMultiByte(UINT CodePage,
									DWORD   dwFlags,
									LPCWSTR lpWideChar,
									int     cchWideChar,
									_Out_cap_(cchMultiByte) LPSTR   lpMultiByte,
									INT_PTR cchMultiByte,
									LPCSTR  lpDefaultChar,
									LPBOOL  lpUsedDefaultChar);

/*	This function provides our own layer to GetStringTypeExW() call. */
MSOAPI_(BOOL) MsoGetStringTypeExW(LCID Locale,
								  DWORD dwInfoType,
								  LPCWSTR lpSrcStr,
								  int cchSrc,
								  LPWORD lpCharType);

#ifdef OFFICE_BUILD
// Returns whether a given Unicode string can be losslessly converted into the given codepage.
MSOAPIX_(BOOL) MsoFExpressibleWzInCodePage (const WCHAR *wz, int cpg);
#endif

/* The next 2 apis are from Office 95.  The latter is only used by Word. */
/* Checks if ch is a DBCS lead byte. */

MSOAPIX_(BOOL) FDbcsFirstByte(unsigned char ch);

// All ctypes of all acceptable languages should
// be defined here.
#define ctypeNone 0

// This must be the same as in langspec.h !!!
#define ctypeSpace 1
#define ctypeAlnum 2
#define ctypeSBKatakana 3
#define ctypeSymbol 4
#define ctypeDBSpace 5
#define ctypeDBHiragana 6
#define ctypeDBKatakana 7
#define ctypeDBEKana 8
#define ctypeDBAlnum 9
#define ctypeDBKanji 10
#define ctypeDBSymbol 11
#define ctypeDBHangeul 12
#define ctypeDBHangeulJamo 13
#define ctypeDBHanja 14
#define ctypeDBEUDC 15
#define ctypeDBGreek 16
#define ctypeDBJapanese 17
#define ctypeDBRussian 18
#define ctypeDBHebrew 19
#define ctypeDBArabic 20
#define ctypeDBDevanagari 21
#define ctypeDBOriya 22
#define ctypeDBTamil 23
#define ctypeDBThai 24

// Classes of characters
#define classNone 0
#define classSpace 1
#define classTerminal 2
#define classNT 3
#define classTNT 4

// Possible trigger values
#define itrigger0 0
#define itrigger1 1
#define itrigger2 2
#define itrigger3 3
#define itrigger4 4
#define itrigger5 5
#define itrigger6 6
#define itrigger7 7

MSOAPI_(int) MsoIOFCTriggerFromXchXch(int xch1, int xch2);

#define wchUnicodeBig    0xFFFE
#define wchUnicodeLittle 0xFEFF
#define lUTF8BOM         0xBFBBEFL
#define lUTF8BOMMask     0xFFFFFFL

/* Office's (actually Word's and maintained by them) implementation of DrawTextW */
MSOAPI_(int) MsoDrawTextW(HDC hdc, LPCWSTR lpwchText, int cchText, RECT FAR *lprc, UINT format);
MSOAPI_(int) MsoDrawTextExW(HDC hdc, LPCWSTR lpwchText, int cchText,
							LPRECT lprc, DWORD dwDTformat,
							LPDRAWTEXTPARAMS lpDTparams);

/*	Compares the sort order of the two unicode strings rgwch1
	(# of chars cch1) and rgwch2 (# of chars cch2).  Returns negative if
	rgwch1 < rgwch2, 0 if they sort equally, and positive if rgwch1 > rgwch2.
	The string compare is based on Unicode order.  MsoWchToUpper is used
	as an invariant ToUpper.  It must behave the same no matter which language
	we run on. */
MSOAPI_(int) MsoSgnRgwchCompareInvariant(const WCHAR *rgwch1, int cch1, const WCHAR *rgwch2, int cch2, int cs);

/* Compares two 7-bit ASCII only strings case insensitively.  Does so
   without doing a unicode conversion and without requiring the language
   Dll.
 */
MSOAPIX_(BOOL) MsoFRgchIEqualFast(const char *rgch1, const char *rgch2, int cch);
/*-----------------------------------------------------------------------------
	MsoFRgwchIEqualFast

	cheap rgch compare, wz1 must be a unicode string of lower case
	7-bit chars (0-127)
-------------------------------------------------------------------- HAILIU -*/
MSOAPIX_(BOOL) MsoFRgwchIEqualFast(LPCWSTR rgwch1, LPCWSTR rgwch2, int cch);
MSOAPI_(BOOL) MsoFSzIEqualFast(const char *rgch1, const char *rgch2);


#if defined(__cplusplus)
}
#endif

MSOAPI_(BOOL) MsoFMarkWzToWtz(const WCHAR *wz, _Inout_ WCHAR **pwtz);

// tokenize the string
MSOAPI_(PWSTR) MsoWzToken(PWSTR * pwzToken, PCWSTR wzDelimiterList);

#if defined(OFFICE_BUILD) || defined(WORD_BUILD) || defined(EXCEL_BUILD) || defined(ACCESS_BUILD) || defined(PPT_BUILD) || defined(OIS_BUILD)
/*these scan an array IN REVERSE ORDER
	 - put the most common cases last in the array*/
#define countof(arr)		(sizeof(arr) / sizeof(arr[0]))
#define ArrHasB(arr, item)	arrScanB(arr, item, countof(arr))
#define ArrHasW(arr, item)	arrScanW(arr, item, countof(arr))
#define ArrHasD(arr, item)	arrScanD(arr, item, countof(arr))

__inline int __cdecl arrScanB(const char *arr, char item, int cnt)
{	//ret: 0 = item not in array, non-zero = found
	do ; while (--cnt >= 0 && arr[cnt] != item);
	return ++cnt;
}

__inline int __cdecl arrScanW(const WORD *arr, WORD item, int cnt)
{	//ret: 0 = item not in array, non-zero = found
	do ; while (--cnt >= 0 && arr[cnt] != item);
	return ++cnt;
}

__inline int __cdecl arrScanD(const DWORD *arr, DWORD item, int cnt)
{	//ret: 0 = item not in array, non-zero = found
	do ; while (--cnt >= 0 && arr[cnt] != item);
	return ++cnt;
}


#undef isspace
__inline int __cdecl isspace(int c)
{
    return MsoFSpaceCh((CHAR)c);
}

#ifndef ZENSTAT_LIB_DEF
#undef isdigit
#define isdigit(c) _mso_isdigit(c)
__inline int __cdecl _mso_isdigit(int c)
{
    return MsoFDigitCh((CHAR)c);
}
#endif //ZENSTAT_LIB_DEF

#undef iswdigit
__inline int __cdecl iswdigit(wint_t c)
{
    return MsoFDigitWch((WCHAR)c);
}
#define _WCTYPE_INLINE_DEFINED		//  NOTE:  To keep the crt headers from giving us this again.

#if !defined(ACCESS_BUILD) && !defined(OFFICE_BUILD)
#undef alloca
#define alloca    Do not use
#endif
#endif // if defined(*BUILD)

// 2nd definition will not compile if passed a pointer (rather than an array)
// Arrays match 1st template type (so sizeof(Verify) is 1).
// Pointers match 2nd template type (so sizeof(Verify) is a compile error).
// Note that cElements() is still a _compile-time constant_, no code is ever actually run
#ifndef cElements
#ifndef __cplusplus
#define cElements(ary) (sizeof(ary) / sizeof(ary[0]))
#elif OACR
#define cElements(ary) (sizeof(ary) / sizeof(ary[0]))
#elif defined _PREFAST_
// Avoid PreFAST warning 260: Sizeof*sizeof is almost always wrong
#define cElements(ary) (sizeof(ary) / sizeof(ary[0]))
#else
// C++ template magic to ensure that the cElements macro is used only on fixed-sized array types.
// If the second one matches, then we get sizeof(void) which is a compile-time error
template<typename T> static char cElementsVerify(void const *, T) throw() { return 0; }
template<typename T> static void cElementsVerify(T *const, T *const *) throw() {};
// Note: PreFAST balks on this line with warning 260 because of the sizeof*sizeof
// PHarring: Consider: Use division instead, e.g.: sizeof(arr)/sizeof(*(arr))/sizeof(cElementsVerify(arr,&(arr)))
#define cElements(arr) (sizeof(cElementsVerify(arr,&(arr))) * sizeof(arr)/sizeof(*(arr)))
#endif
#endif

// Use for things like MsoWzCopy(wzSrc, RgC(wzDst)) --> MsoWzCopy(wzSrc, wzDst, cElements(wzDst))
#ifndef RgC
#define RgC(ary) (ary),cElements(ary)
#endif

#define RgCb(ary) (ary),cElements(ary)*sizeof(ary[0])
#define MsoEnsureZeroTermination( wz ) wz[cElements(wz)-1] = 0

//
// Special & Misc.
//

/*************************************************************************

	end of OString section

*************************************************************************/

int IDigitValueOfWch (WCHAR wch);

//======================================================================
// macro to shut down WCHAR* to BSTR conversion warnings
// use as last resort

#if( OACR )
BSTR MsoTreatAsBSTR(_In_z_ WCHAR*);
#define MSO_TREAT_AS_BSTR( wz ) MsoTreatAsBSTR(wz)
#else
#define MSO_TREAT_AS_BSTR( wz ) (wz)
#endif

#endif // MSOSTR_H
