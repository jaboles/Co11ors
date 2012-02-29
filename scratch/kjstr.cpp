/*************************************************************************
	kjstr.cpp

	Owner: rickp
	Copyright (c) 1994 Microsoft Corporation

	Tons and tons of those cool office string utility routines,
	including copy, conversion, comparison, searches and stuff like
	that.
*************************************************************************/

#include "offpch.h"
#include "msocrit.hpp"
#include "kjstr.h"

#pragma OPT_SPEED

#define EXCPT_TRUE 0x0001
#define EXCPT_CHANGEDLOCALE 0x0004

#define minInline(a, b)         ((a) <= (b) ? a : b)
#define FIsLowerFast(ch)        ((unsigned)((ch)-'a') <= 'z'-'a')

__inline void GetDwordWeight(WCHAR wch, DWORD *pWeight);
__inline void ChangeSMWeightKorean(SORTKEY *pSortkey, WCHAR wch);

BOOL FCmpTextEq(const WCHAR* rgwch1, int cch1, const WCHAR* rgwch2, int cch2, int cs);
MSOAPI_(LPVOID) MsoLoadPres(HINSTANCE hinst, int resType, int resId);

#define ASCII_TOUPPER(x) (FBetween(x,'a','z') ? (x) - ('a' - 'A') : (x))

/*------------------------------------------------------------------------
	WzTruncCopy

	Copies the null-terminated string wzSrc to wzDst, when cchDst is
	the size of the buffer at wzDst (in chars).  Truncates the string if it
	is too big to fit.  Resulting string is always null-terminated.

	Returns a point to the end of the string (i.e. the NIL)
--------------------------------------------------------------- RGIESEN -*/
WCHAR *MSOAPICALLTYPE WzTruncCopy(_Out_z_cap_(cchDst) WCHAR *wzDst, const WCHAR *wzSrc, unsigned cchDst)
{
	MsoAssertTag(cchDst > 0, ASSERTTAG('eqrh'));

    unsigned cchSrc = MsoCchWzLen(wzSrc) + 1;
    if (cchDst >= cchSrc)
    {
        memmove(wzDst, wzSrc, cchSrc*2);

        if (MsoFGetDebugCheck(msodcMemoryFillStrings))
        {
            MsoDebugFill(wzDst+cchSrc,(cchDst-cchSrc)*sizeof(WCHAR), msomfBuffer);
        }

        return wzDst + cchSrc - 1;
    }
    else
    {
        cchSrc = cchDst - 1;
        memmove(wzDst, wzSrc, cchSrc*2);
        wzDst[cchSrc] = 0;

        return wzDst+cchSrc;
    }
}

/*------------------------------------------------------------------------
	SzTruncCopy

	Copies the null-terminated string szSrc to szDst, when cbDst is
	the size of the buffer at szDst.   Truncates the string if it is
	too big to fit.  Resulting string is always null-terminated.

	Returns a point to the end of the string (i.e. the NIL)
---------------------------------------------------------------- RICKP -*/
char *MSOAPICALLTYPE SzTruncCopy(_Out_z_cap_(cchDst) CHAR *szDst, const CHAR *szSrc, unsigned cchDst)
{
	MsoAssertTag(cchDst > 0, ASSERTTAG('eqrj'));

	unsigned cchSrc = MsoCchSzLen(szSrc) + 1;
    if (cchDst >= cchSrc)
    {
        memmove(szDst, szSrc, cchSrc);

        if (MsoFGetDebugCheck(msodcMemoryFillStrings))
        {
            MsoDebugFill(szDst+cchSrc,cchDst-cchSrc, msomfBuffer);
        }

        return szDst + cchSrc - 1;
    }
    else
    {
        cchSrc = cchDst - 1;
        memmove(szDst, szSrc, cchSrc);
        szDst[cchSrc] = 0;
        
        return szDst+cchSrc;
    }
}

/*------------------------------------------------------------------------
FCmpTextEq

Taken from the Excel routine of the same name.  FCmpTextEq is supposed
to be fast while still being able to handle international issues.
---------------------------------------------------------------- DHSU -*/
BOOL FCmpTextEq(const WCHAR* rgwch1, int cch1, const WCHAR* rgwch2, int cch2, int cs)
{
    int cch;
    int ch1, ch2;
    const WCHAR * pch1;
    const WCHAR * pch2;

    MsoAssertTag(cch1 >= 0, ASSERTTAG('jrvb'));
    MsoAssertTag(cch2 >= 0, ASSERTTAG('jrvc'));

    // Set the flags for MsoCompareStringW
    UINT flags = 0;
    if (!(cs & msocsCase))
    {
        flags |= NORM_IGNORECASE;
    }

    if (cs & msocsIgnoreKana)
    {
        flags |= NORM_IGNOREKANATYPE;
    }

    if (cs & msocsIgnoreWidth)
    {
        flags |= NORM_IGNOREWIDTH;
    }

    if (cs & msocsIgnoreNonSpace)
    {
        flags |= NORM_IGNORENONSPACE;
    }

    LCID lcidT = GetUserDefaultLCID();
    if (PRIMARYLANGID(LANGIDFROMLCID(lcidT)) == LANG_TURKISH)
    {
        lcidT |= MSO_TURKISH_SORT_MASK;
    }

    // Pass off to MsoCompareStringW if we are not doing a simple case
    // insensitive compare (NORM_IGNORECASE = 1)
    if (flags > NORM_IGNORECASE)
    {
        return MsoCompareStringW(lcidT, flags, rgwch1, cch1, rgwch2, cch2) == 2;
    }

    pch1 = rgwch1;
    pch2 = rgwch2;
    for (cch = minInline(cch1, cch2); cch != 0; cch--)
    {
        // Get the characters and optimistically compare.
        ch1 = *pch1++;
        ch2 = *pch2++;
        if (ch1 == ch2)
        {
            continue;
        }

        // They may still be equal.
        if (FIsLowerFast(ch1))
        {
            ch1 -= 'a'-'A';
        }

        if (FIsLowerFast(ch2))
        {
            ch2 -= 'a'-'A';
        }

        if (ch1 == ch2)
        {
            continue;
        }

        // They may yet be equal if they're international, but this will be slow!
        if ((ch1|ch2) >= 0x0080)
        {
            return MsoCompareStringW(lcidT, flags, rgwch1, cch1, rgwch2, cch2) == 2;
        }

        // Not equal.
        return fFalse;
    }

    // If the lengths are equal then we are done
    if (cch1 == cch2)
    {
        return fTrue;
    }

    // If we got here then the first cch characters were equal but the lengths are unequal.
    // It is possible for two int'l strings to have different length and still be equal! (LevL)
    ch1 = ch2 = '\0';
    if(cch1 > cch2)
    {
        ch1 = *pch1;
    }
    else
    {
        ch2 = *pch2;
    }

    if (ch1 >= 0x0080 || ch2 >= 0x0080)
    {
        return MsoCompareStringW(lcidT, flags, rgwch1, cch1, rgwch2, cch2) == 2;
    }

    return fFalse;
}


//#if !STATIC_LIB_DEF
/*------------------------------------------------------------------------
	MsoWzCopy

	Copies the lesser of len(wzFrom) or n characters (n includes the
	null terminator) from wzFrom to wzTo.
-------------------------------------------------------------- NICOLEP -*/
MSOAPI_(WCHAR *) MsoWzCchCopy(_In_z_ const WCHAR* wzFrom, _Out_z_cap_(cch) WCHAR* wzTo, int cch)
{
    if (cch <= 0)
    {
        MsoAssertTag(cch == 0, ASSERTTAG('pxpn'));
        return wzTo;
    }

	return WzTruncCopy(wzTo, wzFrom, cch);
}

/*------------------------------------------------------------------------
	MsoSzCopy

	Copies the lesser of len(szFrom) or n characters (n includes the
	null terminator) from szFrom to szTo.
-------------------------------------------------------------- NICOLEP -*/
MSOAPI_(CHAR *) MsoSzCchCopy(_In_z_ const CHAR* szFrom, _Out_z_cap_(cch) CHAR* szTo, int cch)
{
    if (cch <= 0)
    {
        MsoAssertTag(cch == 0, ASSERTTAG('pxpo'));
        return 0;
    }

    return SzTruncCopy(szTo, szFrom, cch);
}

/*------------------------------------------------------------------------
	MsoWzCopyMark

	Makes a copy of a wz on the mark buffer
-------------------------------------------------------------- MICHMARC -*/
MSOAPIX_(WCHAR *) MsoWzCopyMark(const WCHAR *wz)
{
    DWORD cb = (MsoCchWzLen(wz)+1)*sizeof(WCHAR);
    WCHAR *wzOut;

    if (!MsoFMarkMem((void **)&wzOut, cb))
    {
        return NULL;
    }

    memcpy(wzOut, wz, cb);
    return wzOut;
}

/*------------------------------------------------------------------------
	MsoCchSzLen

	Returns the length, in bytes, of the SBCS/DBCS string sz.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoCchSzLen(__rzcount(return) const CHAR* sz)
{
	return static_cast<int>(strlen(sz));
}

/*------------------------------------------------------------------------
	MsoCchWzLen

	Returns the length of the Unicode string wz.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoCchWzLen(__rzcount(return) const WCHAR* wz)
{
    return static_cast<int>(wcslen(wz));
}

/*------------------------------------------------------------------------
	MsoSzAppend

	Append null-terminated string wzFrom onto the end of wzTo, such that the
	resulting string is <cch and
	returns a pointer to the end of the destination string.

-------------------------------------------------------------- chrismcb -*/
MSOAPI_(CHAR*) MsoSzAppend(const CHAR* szFrom, _Out_z_cap_(cchTo) CHAR* szTo, int cchTo)
{
    if (!szFrom || !szTo)
    {
        MsoAssertTag(FALSE, ASSERTTAG('ozkq'));
        return szTo;
    }

	MsoAssertSzTag(MsoCchSzLen(szFrom) + MsoCchSzLen(szTo) < cchTo, "the buffer is too short; string will be truncated", ASSERTTAG('uvge'));

    while (*szTo != '\0')
    {
        szTo++;
        cchTo--;
    }

    if (cchTo <= 0)
    {
        MsoNotReachedSzTag("Cch passed in is less than buffer size of szTo", ASSERTTAG('waqh'));
        return szTo;
    }

	return SzTruncCopy(szTo, szFrom, cchTo);
}

/*------------------------------------------------------------------------
	MsoFSzEqual

	Test if the null-terminated strings sz1 and sz2 are the same, taking
	case sensitivity cs into effect.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(BOOL) MsoFSzEqual(const CHAR* sz1, const CHAR* sz2, int cs)
{
	return MsoFRgchEqual(sz1, MsoCchSzLen(sz1), sz2, MsoCchSzLen(sz2), cs);
}

/*------------------------------------------------------------------------
	MsoFWzEqual

	Test if the null-terminated strings wz1 and wz2 are the same, taking
	case sensitivity cs into effect.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(BOOL) MsoFWzEqual(const WCHAR* wz1, const WCHAR* wz2, int cs)
{
	return MsoFRgwchEqual(wz1, MsoCchWzLen(wz1), wz2, MsoCchWzLen(wz2), cs);
}

/*------------------------------------------------------------------------
	MsoFRgchEqual

	Test if the two strings rgch1 (length cch1) and rgch2 (length cch2)
	are the same.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(BOOL) MsoFRgchEqual(const CHAR* rgch1, int cch1, const CHAR* rgch2, int cch2, int cs)
{
	if (cs == msocsExact)
    {
		return cch1 == cch2 && memcmp(rgch1, rgch2, cch1) == 0;
    }

	LCID lcid = GetUserDefaultLCID();

    // Exlcude Turkey because of O10/154031, the four turkish I's
    if (((cs == msocsIgnoreCase) || (cs == 0)) && (PRIMARYLANGID(LANGIDFROMLCID(lcid)) != LANG_TURKISH))
    {
        bool fEnglish = (((lcid & 0x03FF) == 0x0009) && ((lcid & 0xF400) <= 0x3400));
        // The LCID compare allows:
        //		0009 English, default		0409 English, US
        //		0809 English, UK			0C09 English, Australia
        //		1009 English, Canada		1409 English, NZ
        //		1809 English, Ireland		1C09 English, South Africa
        //		2009 English, Jamaca		2409 English, Caribean
        //		2809 English, Belize		2C09 English, Trinidad
        //		3009 English, Zimbabwe		3409 English, Philippines
        // In these locals, _both_ chars need to have the high-bit set for something
        // odd to happen.  Other locales merely need _one_ high-bit set.

        #ifdef DEBUG
        // Make sure we still act like MsoSgnRchCompare
        BOOL fRetReal = (MsoSgnRgchCompare(rgch1, cch1, rgch2, cch2, msocsIgnoreCase) == 0);
        #endif

        while(cch1)
        {
            if (cch2 == 0)
            {
                MsoAssertTag(fRetReal == FALSE, ASSERTTAG('mvjm'));
                return FALSE;
            }

            // Funny characters, call real routine.
            if (((*rgch1 & 0x80) && (*rgch2 & 0x80)) ||     // Both are ext chars, must do real search
                (!fEnglish && ((*rgch1 | *rgch2) & 0x80)))  // One ext char, non english, must to real search
            {
                MsoAssertTag((MsoSgnRgchCompare(rgch1, cch1, rgch2, cch2, msocsIgnoreCase) == 0) == fRetReal, ASSERTTAG('mvjn'));

                return (MsoSgnRgchCompare(rgch1, cch1, rgch2, cch2, msocsIgnoreCase) == 0);
            }

            if (ASCII_TOUPPER(*rgch1) != ASCII_TOUPPER(*rgch2))
            {
                MsoAssertTag(fRetReal == FALSE, ASSERTTAG('mvjo'));
                return FALSE;
            }

            rgch1++;
            cch2--;
            rgch2++;
            cch1--;
        }

		MsoAssertTag(fRetReal == (cch2 == 0), ASSERTTAG('mvjp'));
		return cch2 == 0;
    }

    return (MsoSgnRgchCompare(rgch1, cch1, rgch2, cch2, cs) == 0);
}

/*------------------------------------------------------------------------
	MsoFRgwchEqual

	Test if the two Unicode strings rgwch1 (length cch1) and rgwch2
	(length cch2) are the same.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(BOOL) MsoFRgwchEqual(const WCHAR* rgwch1, int cch1, const WCHAR* rgwch2, int cch2, int cs)
{
	if (cs == msocsExact)
    {
		return cch1 == cch2 && memcmp(rgwch1, rgwch2, cch1 * sizeof(WCHAR)) == 0;
    }
	else
    {	
        return FCmpTextEq(rgwch1, cch1, rgwch2, cch2, cs);
    }
}

/*------------------------------------------------------------------------
	MsoSgnRgchCompare

	Compares the sort order of the two strings rgch1 (length cch1)
	and rgch2 (length cch2).  Returns -1 if rgch1 < rgch2, 0 if they
	sort equally, and +1 if rgch1 > rgch2.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoSgnRgchCompare(const CHAR* rgch1, int cch1, const CHAR* rgch2, int cch2, int cs)
{
	MsoAssertTag(msocsCase == 0x4 && msocsIgnoreCase != 0, ASSERTTAG('eqsj'));

	UINT flags = 0;
	if (!(cs & msocsCase))
    {
		flags |= NORM_IGNORECASE;
    }

	if (cs & msocsIgnoreKana)
    {
		flags |= NORM_IGNOREKANATYPE;
    }

	if (cs & msocsIgnoreWidth)
    {
		flags |= NORM_IGNOREWIDTH;
    }

	if (cs & msocsIgnoreNonSpace)
    {
		flags |= NORM_IGNORENONSPACE;
    }

	WCHAR *rgwch1 = NULL, *rgwch2 = NULL;
	int cwch1, cwch2;
	if (cch1 < 32768)
    {
		MsoFMarkMem((void**)&rgwch1, cch1*sizeof(WCHAR));
    }

	else
    {
		rgwch1 = (WCHAR *)MsoPvAlloc(cch1*sizeof(WCHAR), msodgMisc);
    }

	if (!rgwch1)
    {
		return FALSE;
    }

	cwch1 = MsoMultiByteToWideChar(CP_ACP, 0, rgch1, cch1, rgwch1, cch1);
	if (cch2 < 32768)
    {
		MsoFMarkMem(reinterpret_cast<void**>(&rgwch2), cch2 * sizeof(WCHAR));
    }
	else
    {
		rgwch2 = reinterpret_cast<WCHAR *>(MsoPvAlloc(cch2 * sizeof(WCHAR), msodgMisc));
    }

    if (!rgwch2)
    {
        if (cch1 < 32768)
        {
            MsoReleaseMem(rgwch1);
        }
        else
        {
            MsoFreePv(rgwch1);
        }

        return FALSE;
    }

	cwch2 = MsoMultiByteToWideChar(CP_ACP, 0, rgch2, cch2, rgwch2, cch2);

	LCID lcidT = GetUserDefaultLCID();
	if (PRIMARYLANGID(LANGIDFROMLCID(lcidT)) == LANG_TURKISH)
    {
		lcidT |= MSO_TURKISH_SORT_MASK;
    }

	int iRet = MsoCompareStringW(lcidT, flags, rgwch1, cwch1, rgwch2, cwch2) - 2;
	if (cch2 < 32768)
    {
		MsoReleaseMem(rgwch2);
    }
	else
    {
		MsoFreePv(rgwch2);
    }

	if (cch1 < 32768)
    {
		MsoReleaseMem(rgwch1);
    }
	else
    {
		MsoFreePv(rgwch1);
    }

	return iRet;
}

MSOAPI_(int) MsoSgnRgwchCompare(const WCHAR* rgwch1, int cch1, const WCHAR* rgwch2, int cch2, int cs)
{
	UINT flags = 0;
	if (!(cs & msocsCase))
    {
		flags |= NORM_IGNORECASE;
    }

	if (cs & msocsIgnoreKana)
    {
		flags |= NORM_IGNOREKANATYPE;
    }

	if (cs & msocsIgnoreWidth)
    {
		flags |= NORM_IGNOREWIDTH;
    }

	if (cs & msocsIgnoreNonSpace)
    {
		flags |= NORM_IGNORENONSPACE;
    }

	LCID lcidT = GetUserDefaultLCID();
	if (PRIMARYLANGID(LANGIDFROMLCID(lcidT)) == LANG_TURKISH)
    {
        lcidT |= MSO_TURKISH_SORT_MASK;
    }

	return MsoCompareStringW(lcidT, flags, rgwch1, cch1, rgwch2, cch2) - 2;
}

/*------------------------------------------------------------------------
	MsoSzIndex

	Finds the first occurance of the character ch in the null-terminated
	string sz, and returns a pointer to it.  Returns NULL if the character
	was not in the string.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(const CHAR*) MsoSzIndex(const CHAR* sz, int ch)
{
    for (const CHAR* pch = sz; *pch; )
    {
        int chT = *(reinterpret_cast<unsigned char*>(const_cast<CHAR*>(pch)));
        if (IsDBCSLeadByte(chT))
        {
            chT = (chT << 8) | *(reinterpret_cast<unsigned char*>(const_cast<CHAR*>(pch + 1)));

            if (chT == ch)
            {
                return pch;
            }

            pch += 2;
        }
        else
        {
            if (chT == ch)
            {
                return pch;
            }

            pch++;
        }
    }

	return NULL;
}

/*--------------------------------------------------------------------------
	MsoSzIndexRight

	Finds the last (rightmost) occurance of the character ch in the
	null-terminated ANSI string csz, and returns a pointer to it.
	Returns NULL if the character was not in the string.
---------------------------------------------------------------- SHAMIKB -*/
MSOAPI_(const CHAR*) MsoSzIndexRight(const CHAR* csz, int ch)
{
	const CHAR* pchResult = MsoSzIndex(csz,ch);
	if (pchResult == NULL)
    {
		return NULL;
    }

	UINT cbChar = IsDBCSLeadByte(*pchResult) ? 2 : 1;
	const CHAR* pchNext = pchResult + cbChar;
	const CHAR* pchOld = pchResult;

    while ((pchResult = MsoSzIndex(pchNext, ch)) != 0)
    {
        pchNext = pchResult + cbChar;
        pchOld = pchResult;
    }

	return pchOld;
}

/*--------------------------------------------------------------------------
	Finds the last (rightmost) occurance of the character wch in the
	null-terminated UNINCODE  string cwz, and returns a pointer to it.
	Returns NULL if the character was not in the string.
---------------------------------------------------------------- SHAMIKB -*/
MSOAPI_(const WCHAR*) MsoWzIndexRight(const WCHAR* cwz, int wch)
{
	for (const WCHAR* pwch = cwz + MsoCchWzLen(cwz) - 1; pwch >= cwz; pwch--)
    {
		if (*pwch == wch)
        {
			return pwch;
        }
    }

	return NULL;
}

/*------------------------------------------------------------------------
	MsoWzStrStr

	Finds the occurance of the substring wzFind inside the string wz, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found
---------------------------------------------------------------- RICKP -*/
MSOAPI_(const WCHAR*) MsoWzStrStr(const WCHAR* wz, const WCHAR* wzFind)
{
	const int cch = MsoCchWzLen(wz);
	const int cchFind = MsoCchWzLen(wzFind);

	for (int ich = 0; ich+cchFind <= cch; ich++)
    {
		if (MsoFRgwchEqual(&wz[ich], cchFind, wzFind, cchFind, msocsExact))
        {
			return &wz[ich];
        }
    }

	return NULL;
}

/*------------------------------------------------------------------------
	MsoWzStrStrEx

	Finds the occurance of the substring wzFind inside the string wz, and
	returns a pointer to the beginning of it. Returns NULL if the substring
	is not found
	Ex version lets you specify case sensitive vs. insensitive etc
--------------------------------------------------------------- SHAMIKB -*/
MSOAPI_(const WCHAR*) MsoWzStrStrEx(const WCHAR* wz, const WCHAR* wzFind, int cs)
{
    const int cch = MsoCchWzLen(wz);
    const int cchFind = MsoCchWzLen(wzFind);

    for (int ich = 0; ich+cchFind <= cch; ich++)
    {
        if (MsoFRgwchEqual(&wz[ich], cchFind, wzFind, cchFind, cs))
        {
            return &wz[ich];
        }
    }

    return NULL;
}

/*------------------------------------------------------------------------
	MsoFRgwchToRgch

	Converts the Unicode string in rgwchFrom into an Ansi string at
	rgchTo.  This API will work in-place.
	Returns FALSE if the target buffer is not large enough for the
	conversion.  Returns TRUE on success.
---------------------------------------------------------------- DGray -*/
MSOAPI_(BOOL) MsoFRgwchToRgch(const WCHAR* rgwchFrom, _Out_z_cap_(cbTo) CHAR* rgchTo, int cbTo)
{
	DWORD dw;
	MsoAssertTag(cbTo > 0, ASSERTTAG('wnco'));

    if (cbTo <= 0)
    {
        return FALSE;
    }

	const int cchFrom = MsoCchWzLen(rgwchFrom) + 1;

	if ((cbTo > (dw = MsoRgwchToRgch(rgwchFrom, cchFrom, rgchTo, cbTo))) ||
	    ((dw == cbTo) && (rgchTo[cbTo-1] == '\0') ) )
	{
		return TRUE;
	}

	rgchTo[cbTo-1] = '\0';	// assure that the buffer terminates
	
    return FALSE;
}

/*------------------------------------------------------------------------
	MsoRgwchToRgch

	Converts a unicode string rgchFrom (with length cchFrom) into the
	equivalent string in the system's native character set. If rgchTo
	is non-NULL, the converted string is stored there.  The length in
	bytes of the converted string is returned, even if rgchTo is NULL.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoRgwchToRgch(const WCHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) CHAR* rgchTo, int cchTo)
{
   return MsoRgwchToCpRgch(CP_ACP, rgchFrom, cchFrom, rgchTo, cchTo);
}

/*------------------------------------------------------------------------
	MsoRgwchToCpRgch

	Calls MsoRgwchToCpRgch with the BOOL fDefault set to NULL
--------------------------------------------------------------   RICKP -*/
MSOAPI_(int) MsoRgwchToCpRgch(UINT CodePage, const WCHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) CHAR* rgchTo, int cchTo)
{
	return MsoRgwchToCpRgchEx(CodePage, rgchFrom, cchFrom, rgchTo, cchTo, NULL);
}

/*------------------------------------------------------------------------
	MsoRgwchToCpRgch

	Converts a unicode string rgchFrom (with length cchFrom) into the
	equivalent string in the specified character set. If rgchTo
	is non-NULL, the converted string is stored there.  The length in
	bytes of the converted string is returned, even if rgchTo is NULL.
	Supports CP_ACP and CP_OEMCP defines
--------------------------------------------------------------   RICKP -*/
MSOAPI_(int) MsoRgwchToCpRgchEx(UINT CodePage, const WCHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) CHAR* rgchTo, int cchTo, BOOL *pfDefault)
{
    #if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
	DWORD dwFlags = WC_NO_BEST_FIT_CHARS;
    #else
	DWORD dwFlags = 0;
    #endif // STATIC_LIB_DEF || ZENSTAT_LIB_DEF

	if ( CodePage == CP_UTF8 || CodePage == CP_CHINAGB18030 )
    {
		dwFlags = 0;
    }

	if (pfDefault)
    {
		*pfDefault = fFalse;
    }

	if (static_cast<void*>(const_cast<WCHAR*>(rgchFrom)) == static_cast<void*>(rgchTo))
	{
		WCHAR* rgwchT;
		BOOL fMarkMem;

		MsoAssertTag(cchFrom > 0, ASSERTTAG('eqso'));

        if (cchFrom >= 32768 || !MsoFMarkMem(reinterpret_cast<void**>(&rgwchT), cchFrom * sizeof(WCHAR)))
        {
            rgwchT = reinterpret_cast<WCHAR *>(MsoPvAlloc(cchFrom * sizeof(WCHAR), msodgMisc));
            fMarkMem = fFalse;
        }
        else
        {
            fMarkMem = fTrue;
        }

        if (!rgwchT)
        {
            return 0;
        }

        memcpy(rgwchT, rgchFrom, cchFrom * sizeof(WCHAR));

        const int cchRet = MsoWideCharToMultiByte(CodePage, dwFlags, rgwchT, cchFrom, rgchTo, cchTo, NULL, pfDefault);

        if (fMarkMem)
        {
            MsoReleaseMem(rgwchT);
        }
        else
        {
            MsoFreePv(rgwchT);
        }

        return cchRet;
	}

	int ich = 0;
    if (rgchTo == NULL)
    {
        for ( ; cchFrom > 0; ich++, cchFrom--, rgchFrom++)
        {
            if ((*rgchFrom & 0xff80) != 0)
            {
                return MsoWideCharToMultiByte(CodePage, dwFlags, rgchFrom, cchFrom, NULL, 0, NULL, pfDefault) + ich;
            }
        }
    }
    else
    {
        for ( ; cchFrom > 0 && ich < cchTo; ich++, cchFrom--, rgchFrom++, rgchTo++)
        {
            if ((*rgchFrom & 0xff80) != 0)
            {
                return MsoWideCharToMultiByte(CodePage, dwFlags, rgchFrom, cchFrom, rgchTo, cchTo-ich, NULL, pfDefault) + ich;
            }

            *rgchTo = static_cast<CHAR>(*rgchFrom);
        }
    }

	return ich;
}

/*------------------------------------------------------------------------
	MsoRgchToRgwch

	Converts the Ansi string in rgchFrom with length cchFrom into a
	Unicode string at rgwchTo.  The rgwchTo buffer is assumed to be cchTo
	characters long.  Returns the length of the converted string.  This
	API will work in-place.
---------------------------------------------------------------- RICKP -*/
//  REVIEW:  PETERO:  #define this one
MSOAPI_(int) MsoRgchToRgwch(const CHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) WCHAR* rgwchTo, int cchTo)
{
   return MsoCpRgchToRgwch(CP_ACP, rgchFrom, cchFrom, rgwchTo, cchTo);
}

/*------------------------------------------------------------------------
	MsoCpRgchToRgwch

	Converts the MultiByte string of code page CodePage in rgchFrom with
	length cchFrom into a Unicode string at rgwchTo.  The rgwchTo buffer
	is assumed to be cchTo characters long.  Returns the length of the
	converted string.  This API will work in-place.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoCpRgchToRgwch(UINT CodePage, const CHAR* rgchFrom, int cchFrom, _Out_z_cap_(cchTo) WCHAR* rgwchTo, int cchTo)
{
	if (cchFrom == 0)
    {
		return 0;
    }

	if (static_cast<void*>(const_cast<CHAR*>(rgchFrom)) == static_cast<void*>(rgwchTo))
	{
		BOOL fHeap = FALSE;
		CHAR* rgchT = NULL;
		MsoAssertExpTag(cchFrom >= 0, ASSERTTAG('eqsp'));   // -1 to calculate length not supported

        if (cchFrom < 0)
        {
            return 0;
        }

        if ((cchFrom > 4096) || !MsoFMarkMem(reinterpret_cast<void**>(&rgchT), cchFrom))
        {
            rgchT = reinterpret_cast<CHAR *>(MsoPvAlloc(cchFrom, msodgMisc));
            fHeap = TRUE;
        }

		if (!rgchT)
        {
			return 0;
        }

		memcpy(rgchT, rgchFrom, cchFrom);
		int cchRet =  MsoMultiByteToWideChar(CodePage, 0, rgchT, cchFrom, rgwchTo, cchTo);

		if (fHeap)
        {
            MsoFreePv(rgchT);
        }
		else
        {
			MsoReleaseMem(rgchT);
        }

		return cchRet;
	}

	if (cchFrom > cchTo || cchFrom == (-1))
    {
		return MsoMultiByteToWideChar(CodePage, 0, rgchFrom, cchFrom, rgwchTo, cchTo);
    }

	const CHAR* pchFrom = rgchFrom;
	const CHAR* pchMax = pchFrom + cchFrom;
    while (*pchFrom < 128)
    {
        *rgwchTo++ = *pchFrom++;

        if (pchFrom >= pchMax)
        {
            return cchFrom;
        }
    }

	int iRet = MsoMultiByteToWideChar(CodePage, 0, pchFrom, static_cast<int>(pchMax - pchFrom), rgwchTo, 
                                      static_cast<int>(cchTo - (pchFrom - rgchFrom)));

	if (!iRet)
    {
		return 0;
    }

	return static_cast<int>(pchFrom - rgchFrom + iRet);
}

#ifdef DEBUG
int FAllSimpleSz(const CHAR *sz)
{
	int cch = MsoCchSzLen(sz);

	while (cch--)
    {
		if (*sz++ & 0x80)
        {
			return FALSE;
        }
    }

	return TRUE;
}

int FAllSimpleWz(const WCHAR *wz)
{
	int cch = MsoCchWzLen(wz);

	while (cch--)
    {
		if (*wz++ & 0xFF80)
        {
			return FALSE;
        }
    }

	return TRUE;
}
#endif

MSOAPI_(int) MsoWtzToSz(const WCHAR* wtz, _Out_z_cap_(cchMax) CHAR* sz, int cchMax)
{
	int cch = wtz[0] + 1;

	return (MsoRgwchToRgch(&wtz[1], cch, sz, cchMax) - 1);
}

// If sz == null, then ret size of (eventual) conversion.
MSOAPI_(int) MsoWzToSz(const WCHAR* wz, _Out_z_cap_(cchMax) CHAR* sz, int cchMax)
{
	MsoAssertTag(wz, ASSERTTAG('eqsq'));
//	Assert(!FAllSimpleWz(wz));

	int cch = MsoCchWzLen(wz) + 1;

	return (MsoRgwchToRgch(wz, cch, sz, cchMax) - 1);
}

// Returns the unicode length of the string.
// If this length is < cch, then it also
// converts the string storing it into wz (otherwise, wz is unchanged)
MSOAPIX_(int) MsoSzToWz(const CHAR *sz, _Out_z_cap_(cch) WCHAR *wz, int cch)
{
	MsoAssertTag(sz != NULL, ASSERTTAG('lbcw'));
	MsoAssertTag((wz != NULL) || (cch == 0), ASSERTTAG('lbcx'));
	MsoAssertTag(cch != -1, ASSERTTAG('lbcy'));

	const DWORD cchNeed = MsoMultiByteToWideChar(CP_ACP, 0, sz, -1, NULL, 0);

    if (cchNeed <= cch)
    {
        MsoMultiByteToWideChar(CP_ACP, 0, sz, -1, wz, cch);
        MsoDebugFill(wz + cchNeed, (cch - cchNeed)*sizeof(WCHAR), msomfBuffer);
        MsoAssertTag(MsoCchWzLen(wz)+1 == cchNeed, ASSERTTAG('ozku'));
    }

	return (cchNeed - 1);
}

/*-----------------------------------------------------------------------------
	MsoSzToWzSimple

	Convert to Unicode assuming the sz contains only characters < 128

-------------------------------------------------------------------- PeterO -*/
MSOAPIX_(void) MsoSzToWzSimple(const CHAR* sz, _Out_z_cap_(cchMax) WCHAR* wz, int cchMax)
{
	MsoAssertTag(FAllSimpleSz(sz), ASSERTTAG('eqsu'));

    while (*sz && cchMax > 1)
    {
        MsoAssertTag(!(*sz & 0x80), ASSERTTAG('eqsv'));
        *wz++ = static_cast<WCHAR>(*sz++);
        cchMax--;
    }
	if (cchMax > 0)
    {
		*wz = static_cast<WCHAR>(NULL);			// copy the NULL
    }
}

/*------------------------------------------------------------------------
	MsoWzMarkSzSimpleCore

	Convert to Unicode assuming the sz contains only characters < 128
	The unicode string is allocated out of the mark/release memory heap.
---------------------------------------------------------------- DGray -*/
#ifdef DEBUG
MSOAPIX_(WCHAR*) MsoWzMarkSzSimpleCore(const CHAR* sz, const CHAR* szFile, int line)
#else
MSOAPIX_(WCHAR*) MsoWzMarkSzSimpleCore(const CHAR* sz)
#endif
{
	MsoAssertTag(sz, ASSERTTAG('eqsw'));
	MsoAssertTag(FAllSimpleSz(sz), ASSERTTAG('eqsx'));

	WCHAR* wz;

	const int cch = MsoCchSzLen(sz);
	const int cwch = cch * 2;				// This is never DBCS, just *2
	MsoAssertTag(cwch + 1 > 0, ASSERTTAG('eqsy'));

    #ifdef DEBUG
	if (!MsoFMarkMemCore(reinterpret_cast<void**>(&wz), (cwch + 1) * 2, szFile, line))
    #else
	if (!MsoFMarkMemCore(reinterpret_cast<void**>(&wz), (cwch + 1) * 2))
    #endif
    {
		return NULL;
    }

	MsoSzToWzSimple(sz, wz, (cch + 1) * 2);

	return wz;
}

/*------------------------------------------------------------------------
	MsoFLoadWtz

	Loads a string with id ids from the resource file specified by hinst.
	result is left in wtz, which is a buffer big enough to hold of cch
	Unicode characters.  Returns TRUE if successful
---------------------------------------------------------------- RICKP -*/
MSOAPI_(BOOL) MsoFLoadWtz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) WCHAR* wtz, int cch)
{
	const int cchRead = MsoCchStdLoadWz(hinst, ids, (wtz + 1), (cch - 1));

    if (cchRead < 1)
    {
        if (wtz)
        {
            wtz[0] = wtz[1] = 0;
        }

        return FALSE;
    }

	*wtz = static_cast<WCHAR>(cchRead);

	return TRUE;
}

/*-----------------------------------------------------------------------------
	_MsoCchInsertWz

	Does the real work for MsoCchInsertWz(), etc

	pchMax points to character after last valid character
-------------------------------------------------------------------- KBROWN -*/
int _MsoCchInsertWz(__endptr(pchMax) WCHAR *pchOut, const WCHAR *pchMax, const WCHAR *pchPat, const WCHAR **rgwz)
{
	const WCHAR *wzOut = pchOut;
	MsoDebugFill(pchOut, sizeof(WCHAR)*(pchMax - pchOut), msomfBuffer);
	MsoAssertTag(pchOut<=pchMax, ASSERTTAG('wuef'));

    for (const WCHAR *pchScan = pchPat; ; )
    {
        if (*pchScan != 0 && *pchScan != L'|')
        {
            pchScan++;
        }
        else
        {
            // copy over pattern not copied yet
            int cch = min(pchScan - pchPat, (pchMax - pchOut) - 1);
            memcpy(pchOut, pchPat, cch * sizeof(WCHAR));

            pchOut += cch;
            if (*pchScan == 0)
            {
                break;
            }

            // pull the token number in the pattern and get the specified
            // string to insert
            pchScan++;
            CONST WCHAR *wz;
            CONST WCHAR wzBar[] = L"|";
            if (*pchScan >= L'0' && *pchScan <= L'9')
            {
                wz = rgwz[*pchScan - L'0'];
				if (wz == NULL)
                {
					break;
                }

                pchScan++;
                goto LInsert;
            }
            else if (*pchScan == L'|')
            {
                // Allow | to be in the message if specified as ||.
                wz = wzBar;
                pchScan++;

            LInsert:
                // insert the string into the output buffer
                pchPat = pchScan;
                MsoAssertTag(wz != NULL, ASSERTTAG('eqww'));

                if (wz != NULL)
                {
                    // the -1 ensures that we don't copy over the last spot which is 
                    // for the null terminator.
                    cch = min(MsoCchWzLen(wz), (pchMax - pchOut) - 1); // PERF (leehu): pull MsoCchWzLen() out of min macro
                    memcpy(pchOut, wz, cch * sizeof(WCHAR));
                    pchOut += cch;
                }

                if (pchOut >= pchMax)
                {
                    // NOTE (rolandr): It's ok to cast here since pchOut<pchMax afterwards
                    pchOut = (WCHAR*)pchMax - 1;  // leave room for NIL, or return a -1 if wzOut == pchMax (an error)
                    break;
                }
            }
            else
            {
                wz = rgwz[0];
                goto LInsert;
            }
        }
    }

    MsoAssertSzTag(pchOut < pchMax, "String too long.", ASSERTTAG('eqwx'));

	//REVIEW: (kbrown) If the string is too long, then wzOut won't be null terminiated.
	//AssertSz(pchOut < pchMax, wzOut);

	return static_cast<int>(pchOut - wzOut);
}

/*------------------------------------------------------------------------
	MsoCchInsertWz

	Acts like MsoInsertSz().

---------------------------------------------------------------- MIKEKELL -*/
MSOAPI_(int) MsoCchInsertWz(_Out_z_cap_(cchOut) WCHAR* wzOut, int cchOut, const WCHAR* wzPat, int iCount, ...)
{
	WCHAR* wzPatT = NULL;
    if (wzOut == wzPat)
    {
        if (!MsoFMarkMem(reinterpret_cast<void**>(&wzPatT), (MsoCchWzLen(wzPat) + 1) * sizeof(WCHAR)))
        {
            return 0;
        }

        memcpy(wzPatT, wzPat, (MsoCchWzLen(wzPat) + 1) * sizeof(WCHAR));
        wzPat = wzPatT;
    }

	//Build a list of arguments.  We have a maximum of 10 (0-9).
	MsoAssertMsgTag(iCount < 11, "Maximum of 10 arguments supported", ASSERTTAG('eqwz'));
    MsoAssertMsgTag(iCount >= 0, "Minimum of  0 arguments supported", ASSERTTAG('eqwz'));

	const WCHAR *rgwz[10] = {0};
	va_list va;

	va_start(va, iCount);
	int i = 0;

    while (i < iCount)
    {
        rgwz[i++] = va_arg(va, const WCHAR *);
    }

    va_end(va);

	//Call the helper funciton.
	WCHAR *pchOut = wzOut + _MsoCchInsertWz(wzOut, wzOut + cchOut, wzPat, rgwz);
	*pchOut = 0;

	if (wzPatT != NULL)
    {
		MsoReleaseMem(wzPatT);
    }

	return static_cast<int>(pchOut - wzOut);
}

/*------------------------------------------------------------------------
	MsoCchStdLoadWz

	Loads a non-OSTRMAN string with id ids from the resource file
	specified by hinst.  result is left in wz, which is a buffer big
	enough to hold of cch Unicode characters.  Returns # of chars
	if successful.

	Note:  Derivative of MsoCchLoadWz to support retrieval of non-Office-
	type strings by resource ID.
-------------------------------------------------------------- JOELDOW -*/

//NOTE: They do this horribleness in msostr.h, where MsoCchStdLoadWz is declared, thus if I don't do it here I get an error about
//a redecleration with different type modifiers :(
#define __stdcall __cdecl

MSOAPIX_(int) MsoCchStdLoadWz(HINSTANCE hinst, int ids, _Out_z_cap_(cch) WCHAR* wz, int cch)
{
	MsoAssertTag(cch > 0 && wz, ASSERTTAG('eqxc'));

	int cchRead = 0;
    if (wz && (cch > 0) && (ids >= 0))
    {
        cchRead = LoadStringW(hinst, ids, wz, cch);
        if (cchRead < 1)
        {
            if (wz)
            {
                *wz = 0;
            }

            return 0;
        }
    }

	return cchRead;
}

#undef __stdcall

/*------------------------------------------------------------------------
	MsoWzDecodeInt

	Decodes the signed integer w into ASCII text in base wBase. The string
	is stored in the rgch buffer, which is assumed to be large enough to hold
	the number's text and a null terminator. Returns the length of the text
	decoded.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoWzDecodeInt(_Out_z_cap_(cch) WCHAR* rgwch, int cch, int w, int wBase)
{
    if (cch <= 0)
    {
        MsoAssertTag(cch == 0, ASSERTTAG('pxpp'));
        return 0;
    }

    if (w < 0)
    {
        *rgwch = '-';
        return MsoWzDecodeUint(rgwch+1, cch-1, -w, wBase) + 1;
    }

	return MsoWzDecodeUint(rgwch, cch, w, wBase);
}

/*------------------------------------------------------------------------
	MsoWzDecodeUInt

	Decodes the integer w into Unicode text in base wBase.
	Returns the length of the text decoded.
---------------------------------------------------------------- RICKP -*/
const char rgchHex[] = "0123456789ABCDEF";
MSOAPI_(int) MsoWzDecodeUint(_Out_z_cap_(cch) WCHAR* rgwch, int cch, unsigned u, int wBase)
{
	MsoAssertTag(wBase >= 2 && wBase <= 16, ASSERTTAG('eqxf'));
	MsoDebugFill(rgwch, cch * sizeof(WCHAR), msomfBuffer);

	if (cch == 1)
    {
		*rgwch = 0;
    }

	if (cch <= 1)
    {
		return 0;
    }

    if (u == 0)
    {
        rgwch[0] = L'0';
        rgwch[1] = 0;
        return 1;
    }

	int cDigits = 0;
	unsigned uT = u;

    while(uT)
    {
        cDigits++;
        uT /= wBase;
    }

	if (cDigits >= cch || cDigits < 0)
    {
		return 0;
    }

	rgwch += cDigits;
	*rgwch-- = 0;
	uT = u;

    while(uT)
    {
        *rgwch-- = rgchHex[uT % wBase];
        uT /= wBase;
    }

	return cDigits;
}

/*------------------------------------------------------------------------
	MsoSzDecodeInt

	Decodes the signed integer w into ASCII text in base wBase. The string
	is stored in the rgch buffer, which is assumed to be large enough to hold
	the number's text and a null terminator. Returns the length of the text
	decoded.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoSzDecodeInt(_Out_z_cap_(cch) CHAR* rgch, int cch, int w, int wBase)
{
    if (cch <= 0)
    {
        MsoAssertTag(cch == 0, ASSERTTAG('pxpq'));
        return 0;
    }

    if (w < 0)
    {
        *rgch = '-';
        return MsoSzDecodeUint(rgch+1, cch-1, -w, wBase) + 1;
    }

    return MsoSzDecodeUint(rgch, cch, w, wBase);
}

/*------------------------------------------------------------------------
	MsoSzDecodeUint

	Decodes the unsigned integer u into ASCII text in base wBase. The string
	is stored in the rgch buffer, which is assumed to be large enough to hold
	the number's text and a null terminator. Returns the length of the text
	decoded.
---------------------------------------------------------------- RICKP -*/
MSOAPI_(int) MsoSzDecodeUint(_Out_z_cap_(cch) CHAR* rgch, int cch, unsigned u, int wBase)
{
    if (!(rgch && (cch > 0)))
    {
        MsoAssertTag(FALSE, ASSERTTAG('ozkva'));
        return 0;
    }

	MsoAssertTag(wBase >= 2 && wBase <= 16, ASSERTTAG('ozkv'));
	MsoDebugFill(rgch, cch, msomfBuffer);

	if (cch == 1)
    {
		*rgch = 0;
    }

	if (cch <= 1)
    {
		return 0;
    }

    if (u == 0)
    {
        rgch[0] = '0';
        rgch[1] = 0;
        return 1;
    }

	int cDigits = 0;
	unsigned uT = u;

    while(uT)
    {
        cDigits++;
        uT /= wBase;
    }

	if (cDigits >= cch || cDigits < 0)
    {
		return 0;
    }

    rgch += cDigits;
    *rgch-- = 0;
    uT = u;

    while(uT)
    {
        *rgch-- = rgchHex[uT % wBase];
        uT /= wBase;
    }

	return cDigits;
}

/*-----------------------------------------------------------------------------
	IDigitValueOfWch

	Digit value of a character.
------------------------------------------------------------------ IlyaV ----*/
int IDigitValueOfWch (WCHAR wch)
{
    // ASCII digits, Fullwidth digits.
    // Arabic-Indic, Extended Arabic-Indic.
    // Superscript 0, 4..9.  Subscript 0..9.
    if (FBetween(wch,0x30,0x39) || FBetween(wch,0xFF10,0xFF19) ||
        FBetween(wch,0x0660,0x0669) || FBetween(wch,0x06F0,0x06F9) ||
        FBetween(wch,0x2070,0x2079) || FBetween(wch,0x2080,0x2089))
    {
        return wch & 0x000F;
    }
    // Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada,
    // Malayalam.
    else if (FBetween(wch,0x0966,0x096F) || FBetween(wch,0x09E6,0x09EF) ||
             FBetween(wch,0x0A66,0x0A6F) || FBetween(wch,0x0AE6,0x0AEF) ||
             FBetween(wch,0x0B66,0x0B6F) || FBetween(wch,0x0BE7,0x0BEF) ||
             FBetween(wch,0x0C66,0x0C6F) || FBetween(wch,0x0CE6,0x0CEF) ||
             FBetween(wch,0x0D66,0x0D6F))
    {
        return (wch & 0x000F) - 6;
    }
    // Thai, Lao, Tibetan.
    else if (FBetween(wch,0x0E50,0x0E59) || FBetween(wch,0x0ED0,0x0ED9) || FBetween(wch,0x0F20,0x0F29))
    {
        return wch & 0x000F;
    }
    // Superscript 2, 3, 1.
    else if (wch==0xB2 || wch==0xB3 || wch==0xB9)
    {
        return wch & 0x0007;
    }
    // Tamil 10, 100, 1000;
    else if (FBetween(wch,0x0BF0,0x0BF2))
    {
        return wch==0x0BF0 ? 10 : wch==0x0BF1 ? 100 : 1000;
    }
    else
    {
        MsoAssertTag (!MsoFDigitWch(wch), ASSERTTAG('eqxj'));
        return 0;
    }
}

/*-----------------------------------------------------------------------------
ParseIntWz

This function does the actual conversion of a wz to an int (signed/unsigned)

------------------------------------------------------------------ MARTINTH -*/
int ParseIntWz(const WCHAR *rgwch, int *pdw, BOOL fSigned)
{
    unsigned long dw;
    unsigned long maxval;
    unsigned digval;
    BOOL fNeg = fFalse;
    const WCHAR *pwch;

    dw = 0;
    pwch = rgwch;

    // Skip white space
    while (MsoFSpaceWch(*pwch))
    {
        ++pwch;
    }

    if (*pwch == L'-')
    {
        // We didn't want signed numbers
        if (!fSigned)
        {
            return(0);
        }

        fNeg = fTrue;
        pwch++;
    }
    else if (*pwch == L'+')		// Skip the sign
    {
        ++pwch;
    }

    maxval = ULONG_MAX / 10;
    while (MsoFDigitWch(*pwch))
    {
        digval = IDigitValueOfWch(*pwch);

        // Make sure we don't overflow
        if (dw < maxval || (dw == maxval && static_cast<unsigned long>(digval) <= ULONG_MAX % 10))
        {
            dw = 10*dw + digval;
        }
        else
        {
            return 0;  // overflow
        }

        pwch++;
    }

    // We could still have overflowed, let's make sure...
    if ((fSigned) && ((fNeg && (dw > -LONG_MIN)) || (!fNeg && (dw > LONG_MAX))))
    {
        return(0);
    }

    *pdw = dw;
    if (fSigned && fNeg)
    {
        *pdw = -(static_cast<long>(dw));
    }

    return static_cast<int>(pwch - rgwch);
}

/*----------------------------------------------------------------------------
MsoParseIntWz

Parses the Unicode text at rgwch into *pw. Returns the count of
characters considered part of the number. Continues reading characters
and encoding them into *pw until it encounters a non-digit.
Handles overflow by returning zero.
------------------------------------------------------------------ MartinTh --*/
MSOAPI_(int) MsoParseIntWz(const WCHAR* rgwch, int *pw)
{
    return ParseIntWz(rgwch, pw, TRUE);
}

/*----------------------------------------------------------------------------
MsoParseUIntWz

Parses the Unicode text at rgwch into *pw. Returns the count of
characters considered part of the number. Continues reading characters
and encoding them into *pw until it encounters a non-digit.
Handles overflow by returning zero.
------------------------------------------------------------------ MartinTh --*/
MSOAPI_(int) MsoParseUIntWz(const WCHAR* rgwch, unsigned *pw)
{
    return ParseIntWz(rgwch, (int*)pw, FALSE);
}

/*-----------------------------------------------------------------------------
ParseIntSz

This function does the actual conversion of an sz to an int (signed/unsigned)

------------------------------------------------------------------ MARTINTH -*/
int ParseIntSz(const CHAR *rgch, int *pdw, BOOL fSigned)
{
    unsigned long dw;
    BOOL fNeg = fFalse;
    unsigned long maxval;
    unsigned digval;
    const CHAR *pch;

    dw = 0;
    pch = rgch;

    // Skip white space
    while (MsoFSpaceCh(*pch))
    {
        ++pch;
    }

    if (*pch == '-')
    {
        // We didn't want signed numbers
        if (!fSigned)
        {
            return(0);
        }

        fNeg = TRUE;
        pch++;
    }
    else if (*pch == '+')
    {
        ++pch;
    }

    maxval = ULONG_MAX / 10;
    while (MsoFDigitCh(*pch))
    {
        digval = *pch - '0';

        // Make sure we don't overflow
        if (dw < maxval || (dw == maxval && static_cast<unsigned long>(digval) <= ULONG_MAX % 10))
        {
            dw = 10*dw + digval;
        }
        else
        {
            return 0;	// overflow
        }

        pch++;
    }

    // We could still have overflowed, let's make sure...
    if ((fSigned) && ((fNeg && (dw > -LONG_MIN)) || (!fNeg && (dw > LONG_MAX))))
    {
        return(0);
    }

    *pdw = dw;
    if (fSigned && fNeg)
    {
        *pdw =  -(static_cast<long>(dw));
    }

    return static_cast<int>(pch - rgch);
}

/*--------------------------------------------------------------------------
   MsoParseIntSz

   Parses the Ansi text at rgch into *pw. Returns the count of
   characters considered part of the number. Continues reading
   characters and encoding them into *pw until it encounters a
   non-digit.
	Handles overflow by returning zero.
------------------------------------------------------------------ RICKP -*/
MSOAPI_(int) MsoParseIntSz(const CHAR* rgch, int *pdw)
{
	return ParseIntSz(rgch, pdw, fTrue);
}

/*--------------------------------------------------------------------------
   MsoParseIntSz

   Parses the Ansi text at rgch into *pw. Returns the count of
   characters considered part of the number. Continues reading
   characters and encoding them into *pw until it encounters a
   non-digit.
	Handles overflow by returning zero.
------------------------------------------------------------------ RICKP -*/
MSOAPI_(int) MsoParseUIntSz(const CHAR* rgch, unsigned *pdw)
{
	return ParseIntSz(rgch, (int*)pdw, FALSE);
}

// multibyte to unicode conversion tables

//SOUTHASIA //schai
const WCHAR mpwchConv874[] = // Thai ANSI Codepage (874) conversion
	{
	0x20AC, 0x0081, 0x0082, 0x0083, 0x0084, 0x2026, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x0E01, 0x0E02, 0x0E03, 0x0E04, 0x0E05, 0x0E06, 0x0E07,
	0x0E08, 0x0E09, 0x0E0A, 0x0E0B, 0x0E0C, 0x0E0D, 0x0E0E, 0x0E0F,
	0x0E10, 0x0E11, 0x0E12, 0x0E13, 0x0E14, 0x0E15, 0x0E16, 0x0E17,
	0x0E18, 0x0E19, 0x0E1A, 0x0E1B, 0x0E1C, 0x0E1D, 0x0E1E, 0x0E1F,
	0x0E20, 0x0E21, 0x0E22, 0x0E23, 0x0E24, 0x0E25, 0x0E26, 0x0E27,
	0x0E28, 0x0E29, 0x0E2A, 0x0E2B, 0x0E2C, 0x0E2D, 0x0E2E, 0x0E2F,
	0x0E30, 0x0E31, 0x0E32, 0x0E33, 0x0E34, 0x0E35, 0x0E36, 0x0E37,
	0x0E38, 0x0E39, 0x0E3A, 0xF8C1, 0xF8C2, 0xF8C3, 0xF8C4, 0x0E3F,
	0x0E40, 0x0E41, 0x0E42, 0x0E43, 0x0E44, 0x0E45, 0x0E46, 0x0E47,
	0x0E48, 0x0E49, 0x0E4A, 0x0E4B, 0x0E4C, 0x0E4D, 0x0E4E, 0x0E4F,
	0x0E50, 0x0E51, 0x0E52, 0x0E53, 0x0E54, 0x0E55, 0x0E56, 0x0E57,
	0x0E58, 0x0E59, 0x0E5A, 0x0E5B, 0xF8C5, 0xF8C6, 0xF8C7, 0xF8C8
	};
//SOUTHASIA

const WCHAR mpwchConv1250[] = // Eastern Roman ANSI Codepage (1250) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0083, 0x201E, 0x2026, 0x2020, 0x2021,
	0x0088, 0x2030, 0x0160, 0x2039, 0x015A, 0x0164, 0x017D, 0x0179,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x0098, 0x2122, 0x0161, 0x203A, 0x015B, 0x0165, 0x017E, 0x017A,
	0x00A0, 0x02C7, 0x02D8, 0x0141, 0x00A4, 0x0104, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x015E, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x017B,
	0x00B0, 0x00B1, 0x02DB, 0x0142, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x0105, 0x015F, 0x00BB, 0x013D, 0x02DD, 0x013E, 0x017C,
	0x0154, 0x00C1, 0x00C2, 0x0102, 0x00C4, 0x0139, 0x0106, 0x00C7,
	0x010C, 0x00C9, 0x0118, 0x00CB, 0x011A, 0x00CD, 0x00CE, 0x010E,
	0x0110, 0x0143, 0x0147, 0x00D3, 0x00D4, 0x0150, 0x00D6, 0x00D7,
	0x0158, 0x016E, 0x00DA, 0x0170, 0x00DC, 0x00DD, 0x0162, 0x00DF,
	0x0155, 0x00E1, 0x00E2, 0x0103, 0x00E4, 0x013A, 0x0107, 0x00E7,
	0x010D, 0x00E9, 0x0119, 0x00EB, 0x011B, 0x00ED, 0x00EE, 0x010F,
	0x0111, 0x0144, 0x0148, 0x00F3, 0x00F4, 0x0151, 0x00F6, 0x00F7,
	0x0159, 0x016F, 0x00FA, 0x0171, 0x00FC, 0x00FD, 0x0163, 0x02D9
	};

const WCHAR mpwchConv1251[] = // Eastern Cyrillic ANSI Codepage (1251) conversion
	{
	0x0402, 0x0403, 0x201A, 0x0453, 0x201E, 0x2026, 0x2020, 0x2021,
	0x20AC, 0x2030, 0x0409, 0x2039, 0x040A, 0x040C, 0x040B, 0x040F,
	0x0452, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x0098, 0x2122, 0x0459, 0x203A, 0x045A, 0x045C, 0x045B, 0x045F,
	0x00A0, 0x040E, 0x045E, 0x0408, 0x00A4, 0x0490, 0x00A6, 0x00A7,
	0x0401, 0x00A9, 0x0404, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x0407,
	0x00B0, 0x00B1, 0x0406, 0x0456, 0x0491, 0x00B5, 0x00B6, 0x00B7,
	0x0451, 0x2116, 0x0454, 0x00BB, 0x0458, 0x0405, 0x0455, 0x0457,
	0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417,
	0x0418, 0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E, 0x041F,
	0x0420, 0x0421, 0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427,
	0x0428, 0x0429, 0x042A, 0x042B, 0x042C, 0x042D, 0x042E, 0x042F,
	0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437,
	0x0438, 0x0439, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F,
	0x0440, 0x0441, 0x0442, 0x0443, 0x0444, 0x0445, 0x0446, 0x0447,
	0x0448, 0x0449, 0x044A, 0x044B, 0x044C, 0x044D, 0x044E, 0x044F
	};

const WCHAR mpwchConv1252[] = // Western ANSI Codepage (1252) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x00D0, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x00DD, 0x00DE, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x00F0, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x00FD, 0x00FE, 0x00FF
	};

const WCHAR mpwchConv1253[] = // Greek ANSI Codepage (1253) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x0088, 0x2030, 0x008A, 0x2039, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x0098, 0x2122, 0x009A, 0x203A, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x0385, 0x0386, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0xF8F9, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x2015,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x0384, 0x00B5, 0x00B6, 0x00B7,
	0x0388, 0x0389, 0x038A, 0x00BB, 0x038C, 0x00BD, 0x038E, 0x038F,
	0x0390, 0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0396, 0x0397,
	0x0398, 0x0399, 0x039A, 0x039B, 0x039C, 0x039D, 0x039E, 0x039F,
	0x03A0, 0x03A1, 0xF8FA, 0x03A3, 0x03A4, 0x03A5, 0x03A6, 0x03A7,
	0x03A8, 0x03A9, 0x03AA, 0x03AB, 0x03AC, 0x03AD, 0x03AE, 0x03AF,
	0x03B0, 0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B6, 0x03B7,
	0x03B8, 0x03B9, 0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BE, 0x03BF,
	0x03C0, 0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C5, 0x03C6, 0x03C7,
	0x03C8, 0x03C9, 0x03CA, 0x03CB, 0x03CC, 0x03CD, 0x03CE, 0xF8FB
	};

const WCHAR mpwchConv1254[] = // Turkish ANSI Codepage (1254) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x008E, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x009E, 0x0178,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x011E, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x0130, 0x015E, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x011F, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x0131, 0x015F, 0x00FF
	};

const WCHAR mpwchConv1255[] = // Hebrew ANSI Codepage (1255) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x008A, 0x2039, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x02DC, 0x2122, 0x009A, 0x203A, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x20AA, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00D7, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00F7, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x05B0, 0x05B1, 0x05B2, 0x05B3, 0x05B4, 0x05B5, 0x05B6, 0x05B7,
	0x05B8, 0x05B9, 0x05BA, 0x05BB, 0x05BC, 0x05BD, 0x05BE, 0x05BF,
	0x05C0, 0x05C1, 0x05C2, 0x05C3, 0x05F0, 0x05F1, 0x05F2, 0x05F3,
	0x05F4, 0xF88D, 0xF88E, 0xF88F, 0xF890, 0xF891, 0xF892, 0xF893,
	0x05D0, 0x05D1, 0x05D2, 0x05D3, 0x05D4, 0x05D5, 0x05D6, 0x05D7,
	0x05D8, 0x05D9, 0x05DA, 0x05DB, 0x05DC, 0x05DD, 0x05DE, 0x05DF,
	0x05E0, 0x05E1, 0x05E2, 0x05E3, 0x05E4, 0x05E5, 0x05E6, 0x05E7,
	0x05E8, 0x05E9, 0x05EA, 0xF894, 0xF895, 0x200E, 0x200F, 0xF896
	};

const WCHAR mpwchConv1256[] = // Arabic ANSI Codepage (1256) conversion
	{
	0x20AC, 0x067E, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x0679, 0x2039, 0x0152, 0x0686, 0x0698, 0x0688,
	0x06AF, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x06A9, 0x2122, 0x0691, 0x203A, 0x0153, 0x200C, 0x200D, 0x06BA,
	0x00A0, 0x060C, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x06BE, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x061B, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x061F,
	0x06C1, 0x0621, 0x0622, 0x0623, 0x0624, 0x0625, 0x0626, 0x0627,
	0x0628, 0x0629, 0x062A, 0x062B, 0x062C, 0x062D, 0x062E, 0x062F,
	0x0630, 0x0631, 0x0632, 0x0633, 0x0634, 0x0635, 0x0636, 0x00d7,
	0x0637, 0x0638, 0x0639, 0x063A, 0x0640, 0x0641, 0x0642, 0x0643,
	0x00E0, 0x0644, 0x00E2, 0x0645, 0x0646, 0x0647, 0x0648, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x0649, 0x064A, 0x00EE, 0x00EF,
	0x064B, 0x064C, 0x064D, 0x064E, 0x00F4, 0x064F, 0x0650, 0x00F7,
	0x0651, 0x00F9, 0x0652, 0x00FB, 0x00FC, 0x200E, 0x200F, 0x06D2
	};

const WCHAR mpwchConv1257[] = // Baltic ANSI Codepage (1257) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0083, 0x201E, 0x2026, 0x2020, 0x2021,
	0x0088, 0x2030, 0x008A, 0x2039, 0x008C, 0x00A8, 0x02C7, 0x00B8,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x0098, 0x2122, 0x009A, 0x203A, 0x009C, 0x00AF, 0x02DB, 0x009F,
	0x00A0, 0xF8FC, 0x00A2, 0x00A3, 0x00A4, 0xF8FD, 0x00A6, 0x00A7,
	0x00D8, 0x00A9, 0x0156, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00C6,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00F8, 0x00B9, 0x0157, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00E6,
	0x0104, 0x012E, 0x0100, 0x0106, 0x00C4, 0x00C5, 0x0118, 0x0112,
	0x010C, 0x00C9, 0x0179, 0x0116, 0x0122, 0x0136, 0x012A, 0x013B,
	0x0160, 0x0143, 0x0145, 0x00D3, 0x014C, 0x00D5, 0x00D6, 0x00D7,
	0x0172, 0x0141, 0x015A, 0x016A, 0x00DC, 0x017B, 0x017D, 0x00DF,
	0x0105, 0x012F, 0x0101, 0x0107, 0x00E4, 0x00E5, 0x0119, 0x0113,
	0x010D, 0x00E9, 0x017A, 0x0117, 0x0123, 0x0137, 0x012B, 0x013C,
	0x0161, 0x0144, 0x0146, 0x00F3, 0x014D, 0x00F5, 0x00F6, 0x00F7,
	0x0173, 0x0142, 0x015B, 0x016B, 0x00FC, 0x017C, 0x017E, 0x02D9
	};

//SOUTHASIA // schai
const WCHAR mpwchConv1258[] = // Vietnamese ANSI Codepage (1258) conversion
	{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x008A, 0x2039, 0x0152, 0x008D, 0x008E, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x02DC, 0x2122, 0x009A, 0x203A, 0x0153, 0x009D, 0x009E, 0x0178,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x0102, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x0300, 0x00CD, 0x00CE, 0x00CF,
	0x0110, 0x00D1, 0x0309, 0x00D3, 0x00D4, 0x01A0, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x01AF, 0x0303, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x0103, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x0301, 0x00ED, 0x00EE, 0x00EF,
	0x0111, 0x00F1, 0x0323, 0x00F3, 0x00F4, 0x01A1, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x01B0, 0x20AB, 0x00FF
	};
//SOUTHASIA

const WCHAR mpwchConv20866[] = // KOI8-R (Cyrillic) (20866) conversion (0x80-0xff)
	{
	0x2500, 0x2502, 0x250C, 0x2510, 0x2514, 0x2518, 0x251C, 0x2524,
	0x252C, 0x2534, 0x253C, 0x2580, 0x2584, 0x2588, 0x258C, 0x2590,
	0x2591, 0x2592, 0x2593, 0x2320, 0x25A0, 0x2022, 0x221A, 0x2248,
	0x2264, 0x2265, 0x00A0, 0x2321, 0x00B0, 0x00B2, 0x00B7, 0x00F7,
	0x2550, 0x2551, 0x2552, 0x0451, 0x2553, 0x2554, 0x2555, 0x2556,
	0x2557, 0x2558, 0x2559, 0x255A, 0x255B, 0x255C, 0x255D, 0x255E,
	0x255F, 0x2560, 0x2561, 0x0401, 0x2562, 0x2563, 0x2564, 0x2565,
	0x2566, 0x2567, 0x2568, 0x2569, 0x256A, 0x256B, 0x256C, 0x00A9,
	0x044E, 0x0430, 0x0431, 0x0446, 0x0434, 0x0435, 0x0444, 0x0433,
	0x0445, 0x0438, 0x0439, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E,
	0x043F, 0x044F, 0x0440, 0x0441, 0x0442, 0x0443, 0x0436, 0x0432,
	0x044C, 0x044B, 0x0437, 0x0448, 0x044D, 0x0449, 0x0447, 0x044A,
	0x042E, 0x0410, 0x0411, 0x0426, 0x0414, 0x0415, 0x0424, 0x0413,
	0x0425, 0x0418, 0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E,
	0x041F, 0x042F, 0x0420, 0x0421, 0x0422, 0x0423, 0x0416, 0x0412,
	0x042C, 0x042B, 0x0417, 0x0428, 0x042D, 0x0429, 0x0427, 0x042A
	};

const WCHAR mpwchConv28591[] = // ISO-8859-1 (Latin1, West European) (28591) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x00D0, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x00DD, 0x00DE, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x00F0, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x00FD, 0x00FE, 0x00FF
	};

const WCHAR mpwchConv28592[] = // ISO-8859-2 (Latin2, East European) (28592) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x0104, 0x02D8, 0x0141, 0x00A4, 0x013D, 0x015A, 0x00A7,
	0x00A8, 0x0160, 0x015E, 0x0164, 0x0179, 0x00AD, 0x017D, 0x017B,
	0x00B0, 0x0105, 0x02DB, 0x0142, 0x00B4, 0x013E, 0x015B, 0x02C7,
	0x00B8, 0x0161, 0x015F, 0x0165, 0x017A, 0x02DD, 0x017E, 0x017C,
	0x0154, 0x00C1, 0x00C2, 0x0102, 0x00C4, 0x0139, 0x0106, 0x00C7,
	0x010C, 0x00C9, 0x0118, 0x00CB, 0x011A, 0x00CD, 0x00CE, 0x010E,
	0x0110, 0x0143, 0x0147, 0x00D3, 0x00D4, 0x0150, 0x00D6, 0x00D7,
	0x0158, 0x016E, 0x00DA, 0x0170, 0x00DC, 0x00DD, 0x0162, 0x00DF,
	0x0155, 0x00E1, 0x00E2, 0x0103, 0x00E4, 0x013A, 0x0107, 0x00E7,
	0x010D, 0x00E9, 0x0119, 0x00EB, 0x011B, 0x00ED, 0x00EE, 0x010F,
	0x0111, 0x0144, 0x0148, 0x00F3, 0x00F4, 0x0151, 0x00F6, 0x00F7,
	0x0159, 0x016F, 0x00FA, 0x0171, 0x00FC, 0x00FD, 0x0163, 0x02D9
	};

const WCHAR mpwchConv28593[] = // ISO-8859-3 (Latin3, South European) (28593) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x0126, 0x02D8, 0x00A3, 0x00A4, 0x00A5, 0x0124, 0x00A7,
	0x00A8, 0x0130, 0x015E, 0x011E, 0x0134, 0x00AD, 0x00AE, 0x017B,
	0x00B0, 0x0127, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x0125, 0x00B7,
	0x00B8, 0x0131, 0x015F, 0x011F, 0x0135, 0x00BD, 0x00BE, 0x017C,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x010A, 0x0108, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x00D0, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x0120, 0x00D6, 0x00D7,
	0x011C, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x016C, 0x015C, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x010B, 0x0109, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x00F0, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x0121, 0x00F6, 0x00F7,
	0x011D, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x016D, 0x015D, 0x02D9
	};

const WCHAR mpwchConv28594[] = // ISO-8859-4 (Latin4, North European) (28594) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x0104, 0x0138, 0x0156, 0x00A4, 0x0128, 0x013B, 0x00A7,
	0x00A8, 0x0160, 0x0112, 0x0122, 0x0166, 0x00AD, 0x017D, 0x00AF,
	0x00B0, 0x0105, 0x02DB, 0x0157, 0x00B4, 0x0129, 0x013C, 0x02C7,
	0x00B8, 0x0161, 0x0113, 0x0123, 0x0167, 0x014A, 0x017E, 0x014B,
	0x0100, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x012E,
	0x010C, 0x00C9, 0x0118, 0x00CB, 0x0116, 0x00CD, 0x00CE, 0x012A,
	0x0110, 0x0145, 0x014C, 0x0136, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x0172, 0x00DA, 0x00DB, 0x00DC, 0x0168, 0x016A, 0x00DF,
	0x0101, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x012F,
	0x010D, 0x00E9, 0x0119, 0x00EB, 0x0117, 0x00ED, 0x00EE, 0x012B,
	0x0111, 0x0146, 0x014D, 0x0137, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x0173, 0x00FA, 0x00FB, 0x00FC, 0x0169, 0x016B, 0x02D9
	};

const WCHAR mpwchConv28595[] = // ISO-8859-5 (Cyrillic) (28595) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x0401, 0x0402, 0x0403, 0x0404, 0x0405, 0x0406, 0x0407,
	0x0408, 0x0409, 0x040A, 0x040B, 0x040C, 0x00AD, 0x040E, 0x040F,
	0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417,
	0x0418, 0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E, 0x041F,
	0x0420, 0x0421, 0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427,
	0x0428, 0x0429, 0x042A, 0x042B, 0x042C, 0x042D, 0x042E, 0x042F,
	0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437,
	0x0438, 0x0439, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F,
	0x0440, 0x0441, 0x0442, 0x0443, 0x0444, 0x0445, 0x0446, 0x0447,
	0x0448, 0x0449, 0x044A, 0x044B, 0x044C, 0x044D, 0x044E, 0x044F,
	0x2116, 0x0451, 0x0452, 0x0453, 0x0454, 0x0455, 0x0456, 0x0457,
	0x0458, 0x0459, 0x045A, 0x045B, 0x045C, 0x00A7, 0x045E, 0x045F
	};

const WCHAR mpwchConv28596[] = // ISO-8859-6 (Arabic) (28596) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x060C, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x061B, 0x00BC, 0x00BD, 0x00BE, 0x061F,
	0x00C0, 0x0621, 0x0622, 0x0623, 0x0624, 0x0625, 0x0626, 0x0627,
	0x0628, 0x0629, 0x062A, 0x062B, 0x062C, 0x062D, 0x062E, 0x062F,
	0x0630, 0x0631, 0x0632, 0x0633, 0x0634, 0x0635, 0x0636, 0x0637,
	0x0638, 0x0639, 0x063A, 0x00DB, 0x00DC, 0x00DD, 0x00DE, 0x00DF,
	0x0640, 0x0641, 0x0642, 0x0643, 0x0644, 0x0645, 0x0646, 0x0647,
	0x0648, 0x0649, 0x064A, 0x064B, 0x064C, 0x064D, 0x064E, 0x064F,
	0x0650, 0x0651, 0x0652, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x00FD, 0x00FE, 0x00FF
	};

const WCHAR mpwchConv28597[] = // ISO-8859-7 (Greek) (28597) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x02BD, 0x02BC, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x2015,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x0384, 0x0385, 0x0386, 0x00B7,
	0x0388, 0x0389, 0x038A, 0x00BB, 0x038C, 0x00BD, 0x038E, 0x038F,
	0x0390, 0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0396, 0x0397,
	0x0398, 0x0399, 0x039A, 0x039B, 0x039C, 0x039D, 0x039E, 0x039F,
	0x03A0, 0x03A1, 0x00D2, 0x03A3, 0x03A4, 0x03A5, 0x03A6, 0x03A7,
	0x03A8, 0x03A9, 0x03AA, 0x03AB, 0x03AC, 0x03AD, 0x03AE, 0x03AF,
	0x03B0, 0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B6, 0x03B7,
	0x03B8, 0x03B9, 0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BE, 0x03BF,
	0x03C0, 0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C5, 0x03C6, 0x03C7,
	0x03C8, 0x03C9, 0x03CA, 0x03CB, 0x03CC, 0x03CD, 0x03CE, 0x00FF
	};

const WCHAR mpwchConv28598[] = // ISO-8859-8 (Hebrew) (28598) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00D7, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x203E,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00F7, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x00D0, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x00DD, 0x00DE, 0x2017,
	0x05D0, 0x05D1, 0x05D2, 0x05D3, 0x05D4, 0x05D5, 0x05D6, 0x05D7,
	0x05D8, 0x05D9, 0x05DA, 0x05DB, 0x05DC, 0x05DD, 0x05DE, 0x05DF,
	0x05E0, 0x05E1, 0x05E2, 0x05E3, 0x05E4, 0x05E5, 0x05E6, 0x05E7,
	0x05E8, 0x05E9, 0x05EA, 0x00FB, 0x00FC, 0x00FD, 0x00FE, 0x00FF
	};

const WCHAR mpwchConv28599[] = // ISO-8859-9 (Latin5, Turkish) (28599) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x011E, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x0130, 0x015E, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x011F, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x0131, 0x015F, 0x00FF
	};

const WCHAR mpwchConv28605[] = // ISO-8859-15 (Latin9) (28605) conversion (0x80-0xff)
	{
	0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
	0x0088, 0x0089, 0x008A, 0x008B, 0x008C, 0x008D, 0x008E, 0x008F,
	0x0090, 0x0091, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097,
	0x0098, 0x0099, 0x009A, 0x009B, 0x009C, 0x009D, 0x009E, 0x009F,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x20AC, 0x00A5, 0x0160, 0x00A7,
	0x0161, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x017D, 0x00B5, 0x00B6, 0x00B7,
	0x017E, 0x00B9, 0x00BA, 0x00BB, 0x0152, 0x0153, 0x0178, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x00D0, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x00DD, 0x00DE, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x00F0, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x00FD, 0x00FE, 0x00FF
	};

const WCHAR mpwchConv10000[] = // Roman Mac Codepage (10000) conversion (0x80-0xff)
	{
	0x00c4, 0x00c5, 0x00c7, 0x00c9, 0x00d1, 0x00d6, 0x00dc, 0x00e1,
	0x00e0, 0x00e2, 0x00e4, 0x00e3, 0x00e5, 0x00e7, 0x00e9, 0x00e8,
	0x00ea, 0x00eb, 0x00ed, 0x00ec, 0x00ee, 0x00ef, 0x00f1, 0x00f3,
	0x00f2, 0x00f4, 0x00f6, 0x00f5, 0x00fa, 0x00f9, 0x00fb, 0x00fc,
	0x2020, 0x00b0, 0x00a2, 0x00a3, 0x00a7, 0x2022, 0x00b6, 0x00df,
	0x00ae, 0x00a9, 0x2122, 0x00b4, 0x00a8, 0x2260, 0x00c6, 0x00d8,
	0x221e, 0x00b1, 0x2264, 0x2265, 0x00a5, 0x00b5, 0x2202, 0x2211,
	0x220f, 0x03c0, 0x222b, 0x00aa, 0x00ba, 0x2126, 0x00e6, 0x00f8,
	0x00bf, 0x00a1, 0x00ac, 0x221a, 0x0192, 0x2248, 0x2206, 0x00ab,
	0x00bb, 0x2026, 0x00a0, 0x00c0, 0x00c3, 0x00d5, 0x0152, 0x0153,
	0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0x00f7, 0x25ca,
	0x00ff, 0x0178, 0x2044, 0x20AC, 0x2039, 0x203a, 0xfb01, 0xfb02,
	0x2021, 0x00b7, 0x201a, 0x201e, 0x2030, 0x00c2, 0x00ca, 0x00c1,
	0x00cb, 0x00c8, 0x00cd, 0x00ce, 0x00cf, 0x00cc, 0x00d3, 0x00d4,
	0xf8ff, 0x00d2, 0x00da, 0x00db, 0x00d9, 0x0131, 0x02c6, 0x02dc,
	0x00af, 0x02d8, 0x02d9, 0x02da, 0x00b8, 0x02dd, 0x02db, 0x02c7
	};

const WCHAR mpwchConv10006[] = // Greek Mac Codepage (10006) conversion (0x80-0xff)
	{
	0x00c4, 0x00b9, 0x00b2, 0x00c9, 0x00b3, 0x00d6, 0x00dc, 0x0385,
	0x00e0, 0x00e2, 0x00e4, 0x0384, 0x00a8, 0x00e7, 0x00e9, 0x00e8,
	0x00ea, 0x00eb, 0x00a3, 0x2122, 0x00ee, 0x00ef, 0x2022, 0x00bd,
	0x2030, 0x00f4, 0x00f6, 0x00a6, 0x00ad, 0x00f9, 0x00fb, 0x00fc,
	0x2020, 0x0393, 0x0394, 0x0398, 0x039b, 0x039e, 0x03a0, 0x00df,
	0x00ae, 0x00a9, 0x03a3, 0x03aa, 0x00a7, 0x2260, 0x00b0, 0x0387,
	0x0391, 0x00b1, 0x2264, 0x2265, 0x00a5, 0x0392, 0x0395, 0x0396,
	0x0397, 0x0399, 0x039a, 0x039c, 0x03a6, 0x03ab, 0x03a8, 0x03a9,
	0x03ac, 0x039d, 0x00ac, 0x039f, 0x03a1, 0x2248, 0x03a4, 0x00ab,
	0x00bb, 0x2026, 0x00a0, 0x03a5, 0x03a7, 0x0386, 0x0388, 0x0153,
	0x2013, 0x2015, 0x201c, 0x201d, 0x2018, 0x2019, 0x00f7, 0x0389,
	0x038a, 0x038c, 0x038e, 0x03ad, 0x03ae, 0x03af, 0x03cc, 0x038f,
	0x03cd, 0x03b1, 0x03b2, 0x03c8, 0x03b4, 0x03b5, 0x03c6, 0x03b3,
	0x03b7, 0x03b9, 0x03be, 0x03ba, 0x03bb, 0x03bc, 0x03bd, 0x03bf,
	0x03c0, 0x03ce, 0x03c1, 0x03c3, 0x03c4, 0x03b8, 0x03c9, 0x03c2,
	0x03c7, 0x03c5, 0x03b6, 0x03ca, 0x03cb, 0x0390, 0x03b0, 0xf8a0
	};

const WCHAR mpwchConv10007[] = // Cyrillic Mac Codepage (10007) conversion (0x80-0xff)
	{
	0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417,
	0x0418, 0x0419, 0x041a, 0x041b, 0x041c, 0x041d, 0x041e, 0x041f,
	0x0420, 0x0421, 0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427,
	0x0428, 0x0429, 0x042a, 0x042b, 0x042c, 0x042d, 0x042e, 0x042f,
	0x2020, 0x00b0, 0x00a2, 0x00a3, 0x00a7, 0x2022, 0x00b6, 0x0406,
	0x00ae, 0x00a9, 0x2122, 0x0402, 0x0452, 0x2260, 0x0403, 0x0453,
	0x221e, 0x00b1, 0x2264, 0x2265, 0x0456, 0x00b5, 0x2202, 0x0408,
	0x0404, 0x0454, 0x0407, 0x0457, 0x0409, 0x0459, 0x040a, 0x045a,
	0x0458, 0x0405, 0x00ac, 0x221a, 0x0192, 0x2248, 0x2206, 0x00ab,
	0x00bb, 0x2026, 0x00a0, 0x040b, 0x045b, 0x040c, 0x045c, 0x0455,
	0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0x00f7, 0x201e,
	0x040e, 0x045e, 0x040f, 0x045f, 0x2116, 0x0401, 0x0451, 0x044f,
	0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437,
	0x0438, 0x0439, 0x043a, 0x043b, 0x043c, 0x043d, 0x043e, 0x043f,
	0x0440, 0x0441, 0x0442, 0x0443, 0x0444, 0x0445, 0x0446, 0x0447,
	0x0448, 0x0449, 0x044a, 0x044b, 0x044c, 0x044d, 0x044e, 0x00a4
	};

const WCHAR mpwchConv10029[] = // Latin 2 Mac Codepage (100029) conversion (0x80-0xff)
	{
	0x00c4, 0x0100, 0x0101, 0x00c9, 0x0104, 0x00d6, 0x00dc, 0x00e1,
	0x0105, 0x010c, 0x00e4, 0x010d, 0x0106, 0x0107, 0x00e9, 0x0179,
	0x017a, 0x010e, 0x00ed, 0x010f, 0x0112, 0x0113, 0x0116, 0x00f3,
	0x0117, 0x00f4, 0x00f6, 0x00f5, 0x00fa, 0x011a, 0x011b, 0x00fc,
	0x2020, 0x00b0, 0x0118, 0x00a3, 0x00a7, 0x2022, 0x00b6, 0x00df,
	0x00ae, 0x00a9, 0x2122, 0x0119, 0x00a8, 0x2260, 0x0123, 0x012e,
	0x012f, 0x012a, 0x2264, 0x2265, 0x012b, 0x0136, 0x2202, 0x2211,
	0x0142, 0x013b, 0x013c, 0x013d, 0x013e, 0x0139, 0x013a, 0x0145,
	0x0146, 0x0143, 0x00ac, 0x221a, 0x0144, 0x0147, 0x2206, 0x00ab,
	0x00bb, 0x2026, 0x00a0, 0x0148, 0x0150, 0x00d5, 0x0151, 0x014c,
	0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0x00f7, 0x25ca,
	0x014d, 0x0154, 0x0155, 0x0158, 0x2039, 0x203a, 0x0159, 0x0156,
	0x0157, 0x0160, 0x201a, 0x201e, 0x0161, 0x015a, 0x015b, 0x00c1,
	0x0164, 0x0165, 0x00cd, 0x017d, 0x017e, 0x016a, 0x00d3, 0x00d4,
	0x016b, 0x016e, 0x00da, 0x016f, 0x0170, 0x0171, 0x0172, 0x0173,
	0x00dd, 0x00fd, 0x0137, 0x017b, 0x0141, 0x017c, 0x0122, 0x02c7
	};

const WCHAR mpwchConv10081[] = // Turkish Mac Codepage (10081) conversion (0x80-0xff)
	{
	0x00c4, 0x00c5, 0x00c7, 0x00c9, 0x00d1, 0x00d6, 0x00dc, 0x00e1,
	0x00e0, 0x00e2, 0x00e4, 0x00e3, 0x00e5, 0x00e7, 0x00e9, 0x00e8,
	0x00ea, 0x00eb, 0x00ed, 0x00ec, 0x00ee, 0x00ef, 0x00f1, 0x00f3,
	0x00f2, 0x00f4, 0x00f6, 0x00f5, 0x00fa, 0x00f9, 0x00fb, 0x00fc,
	0x2020, 0x00b0, 0x00a2, 0x00a3, 0x00a7, 0x2022, 0x00b6, 0x00df,
	0x00ae, 0x00a9, 0x2122, 0x00b4, 0x00a8, 0x2260, 0x00c6, 0x00d8,
	0x221e, 0x00b1, 0x2264, 0x2265, 0x00a5, 0x00b5, 0x2202, 0x2211,
	0x220f, 0x03c0, 0x222b, 0x00aa, 0x00ba, 0x2126, 0x00e6, 0x00f8,
	0x00bf, 0x00a1, 0x00ac, 0x221a, 0x0192, 0x2248, 0x2206, 0x00ab,
	0x00bb, 0x2026, 0x00a0, 0x00c0, 0x00c3, 0x00d5, 0x0152, 0x0153,
	0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0x00f7, 0x25ca,
	0x00ff, 0x0178, 0x011e, 0x011f, 0x0130, 0x0131, 0x015e, 0x015f,
	0x2021, 0x00b7, 0x201a, 0x201e, 0x2030, 0x00c2, 0x00ca, 0x00c1,
	0x00cb, 0x00c8, 0x00cd, 0x00ce, 0x00cf, 0x00cc, 0x00d3, 0x00d4,
	0xf8ff, 0x00d2, 0x00da, 0x00db, 0x00d9, 0xf8a0, 0x02c6, 0x02dc,
	0x00af, 0x02d8, 0x02d9, 0x02da, 0x00b8, 0x02dd, 0x02db, 0x02c7
	};

const WCHAR mpccToUpper[] =
	{	// 247..495 (0x00F7..0x01EF)
/*															0x0178,
	0x0100, 0x0100, 0x0102, 0x0102, 0x0104, 0x0104, 0x0106, 0x0106,
	0x0108, 0x0108, 0x010a, 0x010a, 0x010c, 0x010c, 0x010e, 0x010e,
	0x0110, 0x0110, 0x0112, 0x0112, 0x0114, 0x0114, 0x0116, 0x0116,
	0x0118, 0x0118, 0x011a, 0x011a, 0x011c, 0x011c, 0x011e, 0x011e,
	0x0120, 0x0120, 0x0122, 0x0122, 0x0124, 0x0124, 0x0126, 0x0126,
	0x0128, 0x0128, 0x012a, 0x012a, 0x012c, 0x012c, 0x012e, 0x012e,
	0x0130, 0x0049, 0x0132, 0x0132, 0x0134, 0x0134, 0x0136, 0x0136,
	0x0138, 0x0139, 0x0139, 0x013b, 0x013b, 0x013d, 0x013d, 0x013f,
	0x013f, 0x0141, 0x0141, 0x0143, 0x0143, 0x0145, 0x0145, 0x0147,
	0x0147, 0x0149, */
    #define UPPERofs1	0
	/*0131..0131*/	0x0049,	//this one gets a table entry to bypass Turkish check
    #define UPPERofs2	(0x0131 - 0x0131 + 1 + UPPERofs1)
	/*017A..01DD*/	0x0179, 0x017b, 0x017b, 0x017d, 0x017d, 0x017f,
	0x0180, 0x0181, 0x0182, 0x0182, 0x0184, 0x0184, 0x0186, 0x0187,
	0x0187, 0x0189, 0x018a, 0x018b, 0x018b, 0x018d, 0x018e, 0x018f,
	0x0190, 0x0191, 0x0192, 0x0193, 0x0194, 0x0195, 0x0196, 0x0197,
	0x0198, 0x0198, 0x019a, 0x019b, 0x019c, 0x019d, 0x019e, 0x019f,
	0x01a0, 0x01a0, 0x01a2, 0x01a2, 0x01a4, 0x01a4, 0x01a6, 0x01a7,
	0x01a7, 0x01a9, 0x01aa, 0x01ab, 0x01ac, 0x01ac, 0x01ae, 0x01af,
	0x01af, 0x01b1, 0x01b2, 0x01b3, 0x01b3, 0x01b5, 0x01b5, 0x01b7,
	0x01b8, 0x01b8, 0x01ba, 0x01bb, 0x01bc, 0x01bc, 0x01be, 0x01bf,
	0x01c0, 0x01c1, 0x01c2, 0x01c3, 0x01c4, 0x01c4, 0x01c4, 0x01c7,
	0x01c7, 0x01c7, 0x01ca, 0x01ca, 0x01ca, 0x01cd, 0x01cd, 0x01cf,
	0x01cf, 0x01d1, 0x01d1, 0x01d3, 0x01d3, 0x01d5, 0x01d5, 0x01d7,
	0x01d7, 0x01d9, 0x01d9, 0x01db, 0x01db, 0x018e,
    #define UPPERofs3	(0x01DD - 0x017A + 1 + UPPERofs2)
	//01F0..0252 no change
	/*0253..0292*/			0x0181, 0x0186, 0x0255, 0x0189, 0x018a,
	0x0258, 0x018f, 0x025a, 0x0190, 0x025c, 0x025d, 0x025e, 0x025f,
	0x0193, 0x0261, 0x0262, 0x0194, 0x0264, 0x0265, 0x0266, 0x0267,
	0x0197, 0x0196, 0x026a, 0x026b, 0x026c, 0x026d, 0x026e, 0x019c,
	0x0270, 0x0271, 0x019d, 0x0273, 0x0274, 0x019f, 0x0276, 0x0277,
	0x0278, 0x0279, 0x027a, 0x027b, 0x027c, 0x027d, 0x027e, 0x027f,
	0x0280, 0x0281, 0x0282, 0x01a9, 0x0284, 0x0285, 0x0286, 0x0287,
	0x01ae, 0x0289, 0x01b1, 0x01b2, 0x028c, 0x028d, 0x028e, 0x028f,
	0x0290, 0x0291, 0x01b7,
    #define UPPERofs4	(0x0292 - 0x0253 + 1 + UPPERofs3)
	//0293..03AB no change except 0390->03AA
	/*03AC..03F1*/					0x0386, 0x0388, 0x0389, 0x038a,
	0x03ab, 0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0396, 0x0397,
	0x0398, 0x0399, 0x039a, 0x039b, 0x039c, 0x039d, 0x039e, 0x039f,
	0x03a0, 0x03a1, 0x03a3, 0x03a3,	0x03a4, 0x03a5, 0x03a6, 0x03a7,
	0x03a8, 0x03a9, 0x03aa, 0x03ab,	0x038c, 0x038e, 0x038f, 0x03cf,
	0x0392, 0x0398, 0x03d2, 0x03d3,	0x03d4, 0x03a6, 0x03a0, 0x03d7,
	0x03d8, 0x03d9, 0x03da, 0x03da,	0x03dc, 0x03dc, 0x03de, 0x03de,
	0x03e0, 0x03e0, 0x03e2, 0x03e2,	0x03e4, 0x03e4, 0x03e6, 0x03e6,
	0x03e8, 0x03e8, 0x03ea, 0x03ea,	0x03ec, 0x03ec, 0x03ee, 0x03ee,
	0x039a, 0x03a1,
	}; // mpccToUpper

const WCHAR adjUpper[] = {
//	 a		 b		 c	-- if <= b && >= a then -= c
	0x0061, 0x007A,	0x20,		// if char is between the first 2, add the third
	0x00E0,	0x00F6,	0x20,		// must arrange in ascending order
	0x00F8,	0x00FE,	0x20,

	0x00FF,	0x00FF,	0x00FF-0x0178,
	0x0100,	0x0130,	0,				// 0 = clear bottom bit
	0x0131,	0x0131,	0x0131 - UPPERofs1,
	0x0132,	0x0138,	0,				// 0 = clear bottom bit
	0x0139,	0x0149,	1,				// 1 = dec and set bottom bit
	0x014A,	0x0178,	0,				// etc.
	0x017A,	0x01DD,	0x017A - UPPERofs2,		// offset into mpccToUpper
	0x01DF, 0x01EF, 0,
	0x0253,	0x0292,	0x0253 - UPPERofs3,		//	"
	0x0390,	0x0390,	0x0390-0x03AA,
	0x03AC,	0x03F1,	0x03AC - UPPERofs4,		//	"
	0x0430,	0x044F,	0x20,
	0x0451,	0x045C,	0x50,
	0x045E,	0x045F,	0x50,
	0x0460,	0x0481,	0,
	0x0491,	0x04C0,	0,
	0x04C1,	0x04CC,	1,
	0x04D9,	0x04D9,	0,
	0x04E9,	0x04E9,	0,
	0x0561,	0x0586,	0x30,
	0x24D0,	0x24E9,	0x1A,
	0xFF41,	0xFF5A,	0x20,
	0xFFFF,	0xFFFF,	1, };		// catch all

///* Unicode String Compare */

#pragma MSO_BOOTDATA

//SUKHABUT
#define UNISORT_HIBYTE
#include "unisort.inc"
#undef UNISORT_HIBYTE
//SUKHABUT

#pragma MSO_ENDBOOTDATA

//SUKHABUT
#define UNISORT_EXCEPTIONS
#include "unisort.inc"
#undef UNISORT_EXCEPTIONS
//SUKHABUT

MSOCONSTFIXUP(TBL_PTRS) vtblptrs =
    {
        CREVDIACRITIC,
        CDBLCOMPRESS,
        CIDEOLCID,
        CEXPANSION,
        CCOMPRESS,
        CEXCEPT,
        CMULTIWT,
        rgRevDiacritic,
        rgDblCompress,
        rgIdeoLcid,
        rgExpand,
        rgCompressHdr,
        rgCompress,
        rgExceptHdr,
        rgExcept,
        rgMultiWt
    };

#pragma MSO_BOOTDATA
const TBL_PTRS *pTblPtrs = &vtblptrs;
#pragma MSO_ENDBOOTDATA

//SUKHABUT
#define UNISORT_LOBYTE
#include "unisort.inc"
#undef UNISORT_LOBYTE
//SUKHABUT

/*
 *  State Table.
 */
#define STATE_DW                  1    /* normal diacritic weight state */
#define STATE_REVERSE_DW          2    /* reverse diacritic weight state */
#define STATE_CW                  4    /* case weight state */


///*
// *  Bit Mask Values.
// *
// *  NOTE: Due to intel byte reversal, the DWORD value is backwards:
// *                CW   DW   SM   AW
// *
// *  Case Weight (CW) - 8 bits:
// *    bit 0   => width
// *    bit 4   => case
// *    bit 5   => kana
// *    bit 6,7 => compression
// */
#define CMP_MASKOFF_NONE          0xffffffff
#define CMP_MASKOFF_DW            0xff00ffff
#define CMP_MASKOFF_CW            0xe7ffffff
#define CMP_MASKOFF_DW_CW         0xe700ffff
#define CMP_MASKOFF_KANA          0xdfffffff
#define CMP_MASKOFF_WIDTH         0xfeffffff
#define CMP_MASKOFF_KANA_WIDTH    0xdeffffff

///*
// *  Invalid weight value.
// */
#define CMP_INVALID_WEIGHT        0xffffffff
#define CMP_INVALID_FAREAST       0xffff0000
#define CMP_INVALID_UW            0xffff

///*
// *  Forward Declarations.
// */
int LongCompareStringW(PLOC_HASH pHashN, DWORD dwCmpFlags, LPCWSTR lpString1, int cchCount1, LPCWSTR lpString2, int cchCount2, WORD wExcept);

/*-------------------------------------------------------------------------*\
 *                           INTERNAL MACROS                               *
\*-------------------------------------------------------------------------*/

/*------------------------------------------------------------------------
* NOT_END_STRING
*
* Checks to see if the search has reached the end of the string.
* It returns TRUE if the counter is not at zero (counting backwards) and
* the null termination has not been reached (if -1 was passed in the count
* parameter.
*
-------------------------------------------------------------- SHAMIKB -*/
#define NOT_END_STRING( ct, ptr, cchIn )                                    \
	( (ct != 0) && (!((*(ptr) == 0) && (cchIn == -2))) )

/*------------------------------------------------------------------------
* AT_STRING_END
*
* Checks to see if the pointer is at the end of the string.
* It returns TRUE if the counter is zero or if the null termination
* has been reached (if -2 was passed in the count parameter).
*
-------------------------------------------------------------- SHAMIKB -*/
#define AT_STRING_END( ct, ptr, cchIn )                                     \
	( (ct == 0) || ((*(ptr) == 0) && (cchIn == -2)) )

/*------------------------------------------------------------------------
* REMOVE_STATE
*
* Removes the current state from the state table.  This should only be
* called when the current state should not be entered for the remainder
* of the comparison.  It decrements the counter going through the state
* table and decrements the number of states in the table.
*
-------------------------------------------------------------- SHAMIKB -*/
#define REMOVE_STATE( value )      ( State &= ~value )

/*------------------------------------------------------------------------
* POINTER_FIXUP
*
* Fixup the string pointers if expansion characters were found.
* Then, advance the string pointers and decrement the string counters.
*
-------------------------------------------------------------- SHAMIKB -*/
#define POINTER_FIXUP()                            \
{                                                  \
    /*                                             \
     *  Fixup the pointers (if necessary).         \
     */                                            \
    if (pSave1 && (--cExpChar1 == 0))              \
    {                                              \
        /*                                         \
         *  Done using expansion temporary buffer. \
         */                                        \
        pString1 = pSave1;                         \
        pSave1 = NULL;                             \
     }                                             \
                                                   \
    if (pSave2 && (--cExpChar2 == 0))              \
    {                                              \
        /*                                         \
         *  Done using expansion temporary buffer. \
         */                                        \
         pString2 = pSave2;                        \
         pSave2 = NULL;                            \
    }                                              \
                                                   \
    /*                                             \
     *  Advance the string pointers.               \
     */                                            \
    pString1++;                                    \
    pString2++;                                    \
}

/*------------------------------------------------------------------------
* SCAN_LONGER_STRING
*
* Scans the longer string for diacritic, case, and special weights.
*
-------------------------------------------------------------- SHAMIKB -*/
__inline int SCAN_LONGER_STRING(int ct, LPCWSTR ptr, int cchIn, BOOL ret, DWORD *pWeight1,
                                BOOL fIgnoreDiacritic, WORD wExcept, BOOL fIgnoreSymbol,
                                int WhichCase, int WhichExtra, int WhichPunct1, 
                                int WhichPunct2, int WhichDiacritic)
{
    /*
    *  Search through the rest of the longer string to make sure
    *  all characters are not to be ignored.  If find a character that
    *  should not be ignored, return the given return value immediately.
    *
    *  The only exception to this is when a nonspace mark is found.  If
    *  another DW difference has been found earlier, then use that.
    */
    while (NOT_END_STRING( ct, ptr, cchIn ))
    {
        GetDwordWeight( *ptr, pWeight1 );
        switch ( GET_SCRIPT_MEMBER( pWeight1 ) )
        {
        case ( UNSORTABLE ):
            {
                break;
            }
        case ( NONSPACE_MARK ):
            {
                if ((!fIgnoreDiacritic) && (!WhichDiacritic))
                {
                    return ( ret );
                }
                break;
            }
        case ( NLS_PUNCTUATION ) :
        case ( SYMBOL_1 ) :
        case ( SYMBOL_2 ) :
        case ( SYMBOL_3 ) :
        case ( SYMBOL_4 ) :
        case ( SYMBOL_5 ) :
            {
                if (!fIgnoreSymbol)
                {
                    return ( ret );
                }
                break;
            }
        case ( EXPANSION ) :
        case ( FAREAST_SPECIAL ) :
        default :
            {
                return ( ret );
            }
        case ( RESERVED_2 ) :
        case ( RESERVED_3 ) :
            {
                /*
                *  Fill out the case statement so the compiler
                *  will use a jump table.
                */
                break;
            }
        }

        /*
        *  Advance pointer and decrement counter.
        */
        ptr++;
        ct--;
    }

    /*
    *  Need to check diacritic, case, extra, and special weights for
    *  final return value.  Still could be equal if the longer part of
    *  the string contained only characters to be ignored.
    *
    *  NOTE:  The following checks MUST REMAIN IN THIS ORDER:
    *            Diacritic, Case, Extra, Punctuation.
    */
    if (WhichDiacritic)
    {
        return ( WhichDiacritic );
    }
    if (WhichCase)
    {
        return ( WhichCase );
    }
    if (WhichExtra)
    {
        if (!fIgnoreDiacritic)
        {
            if (GET_WT_FOUR( &WhichExtra ))
            {
                return ( GET_WT_FOUR( &WhichExtra ) );
            }
            if (GET_WT_FIVE( &WhichExtra ))
            {
                return ( GET_WT_FIVE( &WhichExtra ) );
            }
        }
        if (GET_WT_SIX( &WhichExtra ))
        {
            return ( GET_WT_SIX( &WhichExtra ) );
        }
        if (GET_WT_SEVEN( &WhichExtra ))
        {
            return ( GET_WT_SEVEN( &WhichExtra ) );
        }
    }
    if (WhichPunct1)
    {
        return ( WhichPunct1 );
    }
    if (WhichPunct2)
    {
        return ( WhichPunct2 );
    }

    return ( 2 );
}

/*------------------------------------------------------------------------
	Get the sorting information for the Unicode char wch.
-------------------------------------------------------------- SHAMIKB -*/

// TODO shamikb - split back into GetSortkeyOther for perf reasons
void GetSortkeyOther(BYTE bLow, const BYTE *pByte, SORTKEY *pSortkey);

void GetSortkey(WCHAR wch, SORTKEY *pSortkey)
{
	BYTE bHigh = HIBYTE(wch);
	BYTE bLow = LOBYTE(wch);

	const BYTE* pByte = rgUnicodeHiByte[bHigh];
	pSortkey->UW.SM_AW.Alpha = pSortkey->UW.SM_AW.Script = pSortkey->Diacritic = pSortkey->Case = 0;

    if (!pByte)
    {
        return;
    }
    switch (*pByte)
    {
    case 1:
        {
            const ENTRYTYPE1 *pEntry = reinterpret_cast<const ENTRYTYPE1 *>(pByte);

            pSortkey->UW.SM_AW.Script = rgb1b3b4[pEntry->woffb1b3b4[bLow]].b1;
            pSortkey->Diacritic = rgb1b3b4[pEntry->woffb1b3b4[bLow]].b3;
            pSortkey->Case = rgb1b3b4[pEntry->woffb1b3b4[bLow]].b4;
            pSortkey->UW.SM_AW.Alpha = pEntry->rgb2[bLow];

            break;
        }
    case 2:
        {
            const ENTRYTYPE2 *pEntry = reinterpret_cast<const ENTRYTYPE2 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;
            pSortkey->UW.SM_AW.Script = pEntry->rgb1[bLow];
            pSortkey->UW.SM_AW.Alpha = pEntry->rgb2[bLow];

            break;
        }
    case 3:
        {
            const ENTRYTYPE3 *pEntry = reinterpret_cast<const ENTRYTYPE3 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;

            if (bLow < pEntry->bCutoff)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[0] + bLow);
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[1] + bLow - pEntry->bCutoff);
            }
            break;
        }
    case 4:
        {
            const ENTRYTYPE4 *pEntry = reinterpret_cast<const ENTRYTYPE4 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;

            if (bLow < pEntry->bCutoff1)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[0] + bLow);
            }
            else if (bLow < pEntry->bCutoff2)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[1] + bLow - pEntry->bCutoff1);
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[2];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[2] + bLow - pEntry->bCutoff2);
            }
            break;
        }
    case 5:
        {
            const ENTRYTYPE5 *pEntry = reinterpret_cast<const ENTRYTYPE5 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;
            
            if (bLow < pEntry->bCutoff1)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[0] + bLow);
            }
            else if (bLow < pEntry->bCutoff2)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[1] + bLow - pEntry->bCutoff1);
            }
            else if (bLow < pEntry->bCutoff3)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[2];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[2] + bLow - pEntry->bCutoff2);
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[3];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[3] + bLow - pEntry->bCutoff3);
            }
            break;
        }
    case 6:
        {
            const ENTRYTYPE6 *pEntry = reinterpret_cast<const ENTRYTYPE6 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;

            if (bLow < pEntry->bCutoff1)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[0] + bLow);
            }
            else if (bLow < pEntry->bCutoff2)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[1] + bLow - pEntry->bCutoff1);
            }
            else if (bLow < pEntry->bCutoff3)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[2];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[2] + bLow - pEntry->bCutoff2);
            }
            else if (bLow < pEntry->bCutoff4)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[3];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[3] + bLow - pEntry->bCutoff3);
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[4];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[4] + bLow - pEntry->bCutoff4);
            }
            break;
        }
    case 7:
        {
            const ENTRYTYPE7 *pEntry = reinterpret_cast<const ENTRYTYPE7 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;

            if (bLow < pEntry->bCutoff1)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[0] + bLow);
            }
            else if (bLow < pEntry->bCutoff2)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[1] + bLow - pEntry->bCutoff1);
            }
            else if (bLow < pEntry->bCutoff3)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[2];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[2] + bLow - pEntry->bCutoff2);
            }
            else if (bLow < pEntry->bCutoff4)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[3];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[3] + bLow - pEntry->bCutoff3);
            }
            else if (bLow < pEntry->bCutoff5)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[4];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[4] + bLow - pEntry->bCutoff4);
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[5];
                pSortkey->UW.SM_AW.Alpha = static_cast<BYTE>(pEntry->rgb2[5] + bLow - pEntry->bCutoff5);
            }
            break;
        }
    case 8:
        {
            const ENTRYTYPE8 *pEntry = reinterpret_cast<const ENTRYTYPE8 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;

            if (bLow < pEntry->bCutoff)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
            }

            pSortkey->UW.SM_AW.Alpha = pEntry->rgb2[bLow];
            break;
        }
    case 9:
        {
            const ENTRYTYPE9 *pEntry = reinterpret_cast<const ENTRYTYPE9 *>(pByte);

            pSortkey->Diacritic = pEntry->b3;
            pSortkey->Case = pEntry->b4;

            if (bLow < pEntry->bCutoff1)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[0];
            }
            else if (bLow < pEntry->bCutoff2)
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[1];
            }
            else
            {
                pSortkey->UW.SM_AW.Script = pEntry->rgb1[2];
            }

            pSortkey->UW.SM_AW.Alpha = pEntry->rgb2[bLow];
            break;
        }
    case 10:
        {
            const ENTRYTYPE10 *pEntry = reinterpret_cast<const ENTRYTYPE10 *>(pByte);

            pSortkey->UW.SM_AW.Alpha = pEntry->rgByte4[bLow].b2;
            pSortkey->UW.SM_AW.Script = pEntry->rgByte4[bLow].b1;
            pSortkey->Diacritic = pEntry->rgByte4[bLow].b3;
            pSortkey->Case = pEntry->rgByte4[bLow].b4;
            break;
        }
    default:
        MsoAssertTag(0, ASSERTTAG('eqxv'));
    }
    return;
}

/*------------------------------------------------------------------------
* QuickScanLongerString
*
* Scans the longer string for diacritic, case, and special weights.
* Assumes that both strings are null-terminated.
*
-------------------------------------------------------------- SHAMIKB -*/
__inline int QuickScanLongerString (const WCHAR *ptr, int ret, int WhichDiacritic, 
                                    int WhichCase, int WhichExtra, int WhichPunct1, 
                                    int WhichPunct2, WORD wExcept)
{
	SORTKEY SortkeyT;

	/*
	*  Search through the rest of the longer string to make sure
	*  all characters are not to be ignored.  If find a character that
	*  should not be ignored, return the given return value immediately.
	*
	*  The only exception to this is when a nonspace mark is found.  If
	*  another DW difference has been found earlier, then use that.
	*/
    while (*ptr != 0)
    {
        GetSortkey(*ptr, &SortkeyT);
        switch ( GET_SCRIPT_MEMBER( &SortkeyT ) )
        {
        case ( UNSORTABLE ):
            {
                break;
            }
        case ( NONSPACE_MARK ):
            {
                if (!WhichDiacritic)
                {
                    return ( ret );
                }
                break;
            }
        default :
            {
                return ( ret );
            }
        }

        /*
        *  Advance pointer.
        */
        ptr++;
    }

    /*
    *  Need to check diacritic, case, extra, and special weights for
    *  final return value.  Still could be equal if the longer part of
    *  the string contained only unsortable characters.
    *
    *  NOTE:  The following checks MUST REMAIN IN THIS ORDER:
    *            Diacritic, Case, Extra, Punctuation.
    */
    if (WhichDiacritic)
    {
        return ( WhichDiacritic );
    }
    if (WhichCase)
    {
        return ( WhichCase );
    }
    if (WhichExtra)
    {
        if (GET_WT_FOUR( &WhichExtra ))
        {
            return ( GET_WT_FOUR( &WhichExtra ) );
        }
        if (GET_WT_FIVE( &WhichExtra ))
        {
            return ( GET_WT_FIVE( &WhichExtra ) );
        }
        if (GET_WT_SIX( &WhichExtra ))
        {
            return ( GET_WT_SIX( &WhichExtra ) );
        }
        if (GET_WT_SEVEN( &WhichExtra ))
        {
            return ( GET_WT_SEVEN( &WhichExtra ) );
        }
    }
    if (WhichPunct1)
    {
        return ( WhichPunct1 );
    }
    if (WhichPunct2)
    {
        return ( WhichPunct2 );
    }

    return ( 2 );
}

#define FAR_EAST_SWAP(dwT, wt) dwT |= ((wt)&0xffff0000)

/*------------------------------------------------------------------------
* GET_FAREAST_WEIGHT
*
* Returns the weight for the far east special case in "wt".  This currently
* includes the Cho-on, the Repeat, and the Kana characters.
*
-------------------------------------------------------------- SHAMIKB -*/
#define GET_FAREAST_WEIGHT( wt,                                                             \
                            uw,                                                             \
                            mask,                                                           \
                            pBegin,                                                         \
                            pCur,                                                           \
                            ExtraWt )                                                       \
{                                                                                           \
    int ct;       /* loop counter */								                        \
    BYTE PrevSM;  /* previous script member value */			                            \
    BYTE PrevAW;  /* previous alphanumeric value */			                                \
    BYTE PrevCW;  /* previous case value */						                            \
    BYTE AW;      /* alphanumeric value */						                            \
    BYTE CW;      /* case value */									                        \
    DWORD PrevWt; /* previous weight */							                            \
    DWORD dwT;    /* temp var for compoler bug workaround*/                                 \
                                                                                            \
                                                                                            \
    /*                                                                                      \
     *  Get the alphanumeric weight and the case weight of the                              \
     *  current code point.                                                                 \
     */                                                                                     \
    AW = GET_ALPHA_NUMERIC( &wt );                                                          \
    CW = GET_CASE( &wt );                                                                   \
    ExtraWt = static_cast<DWORD>(0);                                                        \
    PrevWt = CMP_INVALID_FAREAST;                                                           \
                                                                                            \
    /*                                                                                      \
     *  Special case Repeat and Cho-On.                                                     \
     *    AW = 0  =>  Repeat                                                                \
     *    AW = 1  =>  Cho-On                                                                \
     *    AW = 2+ =>  Kana                                                                  \
     */                                                                                     \
    if (AW <= MAX_SPECIAL_AW)                                                               \
    {                                                                                       \
        /*                                                                                  \
         *  If the script member of the previous character is                               \
         *  invalid, then give the special character an                                     \
         *  invalid weight (highest possible weight) so that it                             \
         *  will sort AFTER everything else.                                                \
         */                                                                                 \
        ct = 1;                                                                             \
        while ((pCur - ct) >= pBegin)                                                       \
        {                                                                                   \
            GetDwordWeight( *(pCur - ct), &PrevWt );                                        \
            PrevWt &= mask;                                                                 \
            PrevSM = GET_SCRIPT_MEMBER( &PrevWt );                                          \
            if (PrevSM < FAREAST_SPECIAL)                                                   \
            {                                                                               \
                if (PrevSM == EXPANSION)                                                    \
                {                                                                           \
                    PrevWt = CMP_INVALID_FAREAST;                                           \
                }                                                                           \
                else                                                                        \
                {                                                                           \
                    /*                                                                      \
                     *  UNSORTABLE or NONSPACE_MARK.                                        \
                     *                                                                      \
                     *  Just ignore these, since we only care about the                     \
                     *  previous UW value.                                                  \
                     */                                                                     \
                     PrevWt = CMP_INVALID_FAREAST;                                          \
                     ct++;                                                                  \
                     continue;                                                              \
                 }                                                                          \
            }                                                                               \
            else if (PrevSM == FAREAST_SPECIAL)                                             \
            {                                                                               \
                PrevAW = GET_ALPHA_NUMERIC( &PrevWt );                                      \
                if (PrevAW <= MAX_SPECIAL_AW)                                               \
                {                                                                           \
                    /*                                                                      \
                     *  Handle case where two special chars follow                          \
                     *  each other.  Keep going back in the string.                         \
                     */                                                                     \
                     PrevWt = CMP_INVALID_FAREAST;                                          \
                     ct++;                                                                  \
                    continue;                                                               \
                }                                                                           \
                                                                                            \
                UNICODE_WT( &PrevWt ) =	MAKE_UNICODE_WT( KANA, PrevAW );                    \
                                                                                            \
                /*                                                                          \
                 *  Only build weights 4, 5, 6, and 7 if the                                \
                 *  previous character is KANA.                                             \
                 *                                                                          \
                 *  Always:                                                                 \
                 *    4W = previous CW  &  ISOLATE_SMALL                                    \
                 *    6W = previous CW  &  ISOLATE_KANA                                     \
                 *                                                                          \
                 */                                                                         \
                 PrevCW = GET_CASE( &PrevWt );                                              \
                 GET_WT_FOUR( &ExtraWt ) = static_cast<BYTE>(PrevCW & ISOLATE_SMALL);       \
                 GET_WT_SIX( &ExtraWt )  = static_cast<BYTE>(PrevCW & ISOLATE_KANA);        \
                                                                                            \
                 if (AW == AW_REPEAT)                                                       \
                 {                                                                          \
                     /*                                                                     \
                      *  Repeat:                                                            \
                      *    UW = previous UW                                                 \
                      *    5W = WT_FIVE_REPEAT                                              \
                      *    7W = previous CW  &  ISOLATE_WIDTH                               \
                      */                                                                    \
                      uw = UNICODE_WT( &PrevWt );                                           \
                      uw = WSwapIfMac(uw);                                                  \
                      GET_WT_FIVE( &ExtraWt )  = WT_FIVE_REPEAT;                            \
                      GET_WT_SEVEN( &ExtraWt ) = static_cast<BYTE>(PrevCW & ISOLATE_WIDTH); \
                }                                                                           \
                else                                                                        \
                {                                                                           \
                    /*                                                                      \
                     *  Cho-On:                                                             \
                     *    UW = previous UW  &  CHO_ON_UW_MASK                               \
                     *    5W = WT_FIVE_CHO_ON                                               \
                     *    7W = current  CW  &  ISOLATE_WIDTH                                \
                     */                                                                     \
                     uw = UNICODE_WT( &PrevWt );                                            \
                     uw = WSwapIfMac(uw);                                                   \
                     uw &= CHO_ON_UW_MASK;                                                  \
                     GET_WT_FIVE( &ExtraWt )  = WT_FIVE_CHO_ON;                             \
                     GET_WT_SEVEN( &ExtraWt ) = static_cast<BYTE>(CW & ISOLATE_WIDTH);      \
                }                                                                           \
            }                                                                               \
            else                                                                            \
            {                                                                               \
                uw = GET_UNICODE( &PrevWt );                                                \
                uw = WSwapIfMac(uw);                                                        \
            }                                                                               \
                                                                                            \
            break;                                                                          \
        }                                                                                   \
    }                                                                                       \
    else                                                                                    \
    {                                                                                       \
        /*                                                                                  \
         *  Kana:                                                                           \
         *    SM = KANA                                                                     \
         *    AW = current AW                                                               \
         *    4W = current CW  &  ISOLATE_SMALL                                             \
         *    5W = WT_FIVE_KANA                                                             \
         *    6W = current CW  &  ISOLATE_KANA                                              \
         *    7W = current CW  &  ISOLATE_WIDTH                                             \
         */                                                                                 \
         uw = MAKE_WIN_UNICODE_WT( KANA, AW );                                              \
         GET_WT_FOUR( &ExtraWt )  = static_cast<BYTE>(CW & ISOLATE_SMALL);                  \
         GET_WT_FIVE( &ExtraWt )  = WT_FIVE_KANA;                                           \
         GET_WT_SIX( &ExtraWt )   = static_cast<BYTE>(CW & ISOLATE_KANA);                   \
         GET_WT_SEVEN( &ExtraWt ) = static_cast<BYTE>(CW & ISOLATE_WIDTH);                  \
    }                                                                                       \
                                                                                            \
    /*                                                                                      \
     *  Get the weight for the far east special case and store it in wt.                    \
     */                                                                                     \
    if ((AW > MAX_SPECIAL_AW) || (PrevWt != CMP_INVALID_FAREAST))                           \
    {                                                                                       \
        /*                                                                                  \
         *  Always:                                                                         \
         *    DW = current DW                                                               \
         *    CW = minimum CW                                                               \
         */                                                                                 \
         dwT = WSwapIfMac(uw);                                                              \
         FAR_EAST_SWAP(dwT, wt);                                                            \
         wt = dwT;                                                                          \
         CASE_WT( &wt ) = MIN_CW;                                                           \
     }                                                                                      \
     else                                                                                   \
     {                                                                                      \
        uw = CMP_INVALID_UW;                                                                \
        wt = CMP_INVALID_FAREAST;                                                           \
        ExtraWt = 0;                                                                        \
    }                                                                                       \
}

/*------------------------------------------------------------------------
	Get the weight for the Unicode character wch
-------------------------------------------------------------- SHAMIKB -*/
__inline void GetDwordWeight(WCHAR wch, DWORD *pWeight)
{
	SORTKEY SortkeyT;

	GetSortkey(wch, &SortkeyT);

	if (PRIMARYLANGID(LANGIDFROMLCID(vHashN.Locale)) == LANG_KOREAN)
    {
		ChangeSMWeightKorean(&SortkeyT, wch);
    }

	*pWeight = MAKE_PSORTKEY_DWORD(&SortkeyT);
}

/*------------------------------------------------------------------------
	Get the weight for the Unicode character wch in the Korean locale.
	Korean sorting is different from everything else because Korea wants
	Korean chars to sort before Latin and all other ANSI chars.
-------------------------------------------------------------- SHAMIKB -*/
__inline void ChangeSMWeightKorean(SORTKEY *pSortkey, WCHAR wch)
{
    if (pSortkey->UW.SM_AW.Script == 3)		// Japanese Kana
    {
        if ((pSortkey->UW.SM_AW.Alpha != 0) && (pSortkey->UW.SM_AW.Alpha != 1))
        {
            pSortkey->UW.SM_AW.Script = 148;		// 34 + 114
        }
        // else no change (iteration mark or Cho-On)
    }
    else if (pSortkey->UW.SM_AW.Script < 14)
    {
        // no change for digits, symbols and punctuation
    }
    else if (wch >= 0xE000  && wch <= 0xF8FF)   // Private Use Area
    {
        // no change
    }
    else if (pSortkey->UW.SM_AW.Script < 128)
    {
        pSortkey->UW.SM_AW.Script += 114;		// move down non-Korean chars
    }
    else if (pSortkey->UW.SM_AW.Script >=128)
    {
        pSortkey->UW.SM_AW.Script -= 114;		// move up Korean Hangul/Hanja
    }
}

/*------------------------------------------------------------------------
	Fill up the locale info structure
-------------------------------------------------------------- SHAMIKB -*/
__inline void FillpHashN(DWORD Locale, LOC_HASH *pHashN)
{
    /* for FillpHashN */
    DWORD ctr;					/* loop counter */
    REVERSE_DW const *pRevDW;		/* ptr to reverse diacritic table */
    DBL_COMPRESS const *pDblComp;	/* ptr to double compression table */
    COMPRESS_HDR const *pCompHdr;	/* ptr to compression header */

    pHashN->Locale = Locale;

    /*
    *  Check for Reverse Diacritic Locale.
    */
    pRevDW = pTblPtrs->pReverseDW;
    for (ctr = pTblPtrs->NumReverseDW; ctr > 0; ctr--, pRevDW++)
    {
        if (*pRevDW == static_cast<DWORD>(Locale))
        {
            pHashN->IfReverseDW = TRUE;
            break;
        }
    }

    /*
    *  Check for Compression.
    */
    pCompHdr = pTblPtrs->pCompressHdr;
    for (ctr = pTblPtrs->NumCompression; ctr > 0; ctr--, pCompHdr++)
    {
        if (pCompHdr->Locale == static_cast<DWORD>(Locale))
        {
            pHashN->IfCompression = TRUE;
            pHashN->pCompHdr = pCompHdr;

            if (pCompHdr->Num2 > 0)
            {
                pHashN->pCompress2 = reinterpret_cast<PCOMPRESS_2>((reinterpret_cast<LPWORD>(const_cast<COMPRESS*>(pTblPtrs->pCompression))) + 
                                                                   (pCompHdr->Offset));
            }
            if (pCompHdr->Num3 > 0)
            {
                pHashN->pCompress3 = reinterpret_cast<PCOMPRESS_3>((reinterpret_cast<LPWORD>(const_cast<COMPRESS*>(pTblPtrs->pCompression))) +
                                                                   (pCompHdr->Offset) + (pCompHdr->Num2 * (sizeof(COMPRESS_2) / sizeof(WORD))));
            }

            break;
        }
    }

    /*
    *  Check for Double Compression.
    */
    if (pHashN->IfCompression)
    {
        pDblComp = pTblPtrs->pDblCompression;

        for (ctr = pTblPtrs->NumDblCompression; ctr > 0; ctr--, pDblComp++)
        {
            if (*pDblComp == (DWORD)Locale)
            {
                pHashN->IfDblCompression = TRUE;
                break;
            }
        }
    }
}

typedef struct
{
	BYTE bHigh;
	BYTE *pEntry;
} DEFEXCEPT;

int cDefExcept;
#define CDEFEXCEPTMAX 200
DEFEXCEPT rgDefExcept[CDEFEXCEPTMAX];

#ifndef ZENSTAT_LIB_DEF
__inline void RestoreDefaultTables()
#else
void RestoreDefaultTables()
#endif //ZENSTAT_LIB_DEF
{
    for (int i = 0; i < cDefExcept; i++)
    {
        MsoFreePv(rgUnicodeHiByte[rgDefExcept[i].bHigh]);
        rgUnicodeHiByte[rgDefExcept[i].bHigh] = rgDefExcept[i].pEntry;
    }

    cDefExcept = 0;
}

#define MASK_SORT_OFF 0xfff0ffff

__inline BOOL FAddExceptions(LCID Locale)
{
    #if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
	int cEntries;
	ENTRYTYPE10 *pEntry10;
	const EXCEPT *pExcept;
	SORTKEY SortkeyT;
	WCHAR wch;
	const WORD *pData;
	const IDEO_INFO *pKSCScript = NULL, *pKSCAlpha = NULL; 
	const IDEO_CANTCOMPRESS *pKSCCantCompress = NULL;
	BOOL fMsoTurkSort = FALSE;

	#if !STATIC_LIB_DEF
	if (!FEnsureMsoIntlDll())
    {
		return FALSE;
    }
	#endif

	if (Locale == lidKazakh)
    {
		return FALSE;
    }

	if (LANGIDFROMLCID(Locale) == lidTaiwan && SORTIDFROMLCID(Locale) == 0x03)
    {
		return FALSE;
    }

    if (PRIMARYLANGID(LANGIDFROMLCID(Locale)) == LANG_TURKISH
        && SORTIDFROMLCID(Locale) == MSO_TURKISH_SORT)
    {
        Locale &= MASK_SORT_OFF;
        fMsoTurkSort = TRUE;
    }

    for (int i = 0; i < CEXCEPT; i++)
    {
        if (rgExceptHdr[i].Locale == Locale)
        {
            pExcept = &rgExcept[rgExceptHdr[i].Offset/3];
            cEntries = rgExceptHdr[i].NumEntries;
            MsoAssertTag(cDefExcept == 0, ASSERTTAG('eqxx'));
            cDefExcept = 0;

            for (int j = 0; j < cEntries; j++)
            {
                if (*(rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)]) == 10)
                {
                    pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)]);
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b1 = HIBYTE(pExcept[j].Unicode);
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b2 = LOBYTE(pExcept[j].Unicode);
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b3 = pExcept[j].Diacritic;
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b4 = pExcept[j].Case;
                    continue;
                }

				if ((pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(MsoPvAlloc(sizeof(ENTRYTYPE10), msodgMisc))) == NULL)
                {
					return FALSE;
                }

				rgDefExcept[cDefExcept].bHigh = 	HIBYTE(pExcept[j].UCP);
				rgDefExcept[cDefExcept++].pEntry = rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)];

				MsoAssertTag(cDefExcept < CDEFEXCEPTMAX, ASSERTTAG('eqxy'));
				pEntry10->bID = 10;
                for (WORD k = 0; k < 256; k++)
                {
                    wch = static_cast<WCHAR>((pExcept[j].UCP & 0xff00) | k);
                    GetSortkey(wch, &SortkeyT);
                    pEntry10->rgByte4[k].b1 = SortkeyT.UW.SM_AW.Script;
                    pEntry10->rgByte4[k].b2 = SortkeyT.UW.SM_AW.Alpha;
                    pEntry10->rgByte4[k].b3 = SortkeyT.Diacritic;
                    pEntry10->rgByte4[k].b4 = SortkeyT.Case;
                }

				pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b1 = HIBYTE(pExcept[j].Unicode);
				pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b2 = LOBYTE(pExcept[j].Unicode);
				pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b3 = pExcept[j].Diacritic;
				pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b4 = pExcept[j].Case;

				rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)] = (BYTE *)pEntry10;
            }
            break;
        }
    }

    int iScript = 1, iAlpha = 1;
    for (int i = 0; i < CIDEOLCID; i++)
    {
        if (rgIdeoLcid[i].Locale == Locale)
        {
            pData = reinterpret_cast<WORD *>(MsoLoadPres(vhinstMsoIntl, MSORTCompStr, rgIdeoLcid[i].id));
            if (!pData)
            {
                return FALSE;
            }

            if (Locale == 0x412)		// special case korean
            {
                pKSCScript = reinterpret_cast<IDEO_INFO *>(MsoLoadPres(vhinstMsoIntl, MSORTCompStr, MSORIKSCSCRIPT));
                pKSCAlpha = reinterpret_cast<IDEO_INFO *>(MsoLoadPres(vhinstMsoIntl, MSORTCompStr, MSORIKSCALPHA));
                pKSCCantCompress = reinterpret_cast<IDEO_CANTCOMPRESS *>(MsoLoadPres(vhinstMsoIntl, MSORTCompStr, MSORIKSCCANTCOMPRESS));

                if (!pKSCScript || !pKSCAlpha || !pKSCCantCompress)
                {
                    return FALSE;
                }
            }
            for (int j = 0; j < rgIdeoLcid[i].cData; j++)
            {
                if (!pData[j])
                {
                    continue;
                }

                if (*(rgUnicodeHiByte[HIBYTE(pData[j])]) == 10)
                {
                    pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[HIBYTE(pData[j])]);
                    if (Locale != 0x412)		// special case korean
                    {
                        pEntry10->rgByte4[LOBYTE(pData[j])].b1 = static_cast<BYTE>(128 + j/254);
                        pEntry10->rgByte4[LOBYTE(pData[j])].b2 = static_cast<BYTE>(2 + j%254);
                        pEntry10->rgByte4[LOBYTE(pData[j])].b3 = 2;
                    }
                    else
                    {
                        if (iScript < CKSCSCRIPT && j >= pKSCScript[iScript].Offset)
                        {
                            iScript++;
                        }

                        pEntry10->rgByte4[LOBYTE(pData[j])].b1 = pKSCScript[iScript-1].Value;

                        if (iAlpha < CKSCALPHA && j >= pKSCAlpha[iAlpha].Offset)
                        {
                            iAlpha++;
                        }

                        pEntry10->rgByte4[LOBYTE(pData[j])].b2 = pKSCAlpha[iAlpha-1].Value;
                        pEntry10->rgByte4[LOBYTE(pData[j])].b3 = static_cast<BYTE>(65 + (j-pKSCAlpha[iAlpha-1].Offset)*2);
                    }

                    pEntry10->rgByte4[LOBYTE(pData[j])].b4 = 2;
                    continue;
                }

                if ((pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(MsoPvAlloc(sizeof(ENTRYTYPE10), msodgMisc))) == NULL)
                {
                    return FALSE;
                }

                rgDefExcept[cDefExcept].bHigh = 	HIBYTE(pData[j]);
                rgDefExcept[cDefExcept++].pEntry = rgUnicodeHiByte[HIBYTE(pData[j])];
                MsoAssertTag(cDefExcept < CDEFEXCEPTMAX, ASSERTTAG('eqxz'));

                pEntry10->bID = 10;
                for (WORD k = 0; k < 256; k++)
                {
                    wch = static_cast<WCHAR>((pData[j] & 0xff00) | k);
                    GetSortkey(wch, &SortkeyT);
                    pEntry10->rgByte4[k].b1 = SortkeyT.UW.SM_AW.Script;
                    pEntry10->rgByte4[k].b2 = SortkeyT.UW.SM_AW.Alpha;
                    pEntry10->rgByte4[k].b3 = SortkeyT.Diacritic;
                    pEntry10->rgByte4[k].b4 = SortkeyT.Case;
                }
                if (Locale != 0x412)		// special case korean
                {
                    pEntry10->rgByte4[LOBYTE(pData[j])].b1 = static_cast<BYTE>(128 + j/254);
                    pEntry10->rgByte4[LOBYTE(pData[j])].b2 = static_cast<BYTE>(2 + j%254);
                    pEntry10->rgByte4[LOBYTE(pData[j])].b3 = 2;
                }
                else
                {
                    if (iScript < CKSCSCRIPT && j >= pKSCScript[iScript].Offset)
                    {
                        iScript++;
                    }

                    pEntry10->rgByte4[LOBYTE(pData[j])].b1 = pKSCScript[iScript-1].Value;

                    if (iAlpha < CKSCALPHA && j >= pKSCAlpha[iAlpha].Offset)
                    {
                        iAlpha++;
                    }

                    pEntry10->rgByte4[LOBYTE(pData[j])].b2 = pKSCAlpha[iAlpha-1].Value;
                    pEntry10->rgByte4[LOBYTE(pData[j])].b3 = static_cast<BYTE>(65 + (j-pKSCAlpha[iAlpha-1].Offset)*2);
                }

                rgUnicodeHiByte[HIBYTE(pData[j])] = (BYTE *)pEntry10;
            }

            if (Locale == 0x412)		// special case korean
            {
                for (int j = 0; j < CKSCCANTCOMPRESS; j++)
                {
                    if (*(rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)]) == 10)
                    {
                        pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)]);

                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b1 = pKSCCantCompress[j].b1;
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b2 = pKSCCantCompress[j].b2;
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b3 = pKSCCantCompress[j].b3;
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b4 = pKSCCantCompress[j].b4;
                        continue;
                    }

                    if ((pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(MsoPvAlloc(sizeof(ENTRYTYPE10), msodgMisc))) == NULL)
                    {
                        return FALSE;
                    }

                    rgDefExcept[cDefExcept].bHigh = 	HIBYTE(pKSCCantCompress[j].wUnicode);
                    rgDefExcept[cDefExcept++].pEntry = rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)];
                    MsoAssertTag(cDefExcept < CDEFEXCEPTMAX, ASSERTTAG('eqya'));

                    pEntry10->bID = 10;
                    for (WORD k = 0; k < 256; k++)
                    {
                        wch = static_cast<WCHAR>((pKSCCantCompress[j].wUnicode & 0xff00) | k);
                        GetSortkey(wch, &SortkeyT);
                        pEntry10->rgByte4[k].b1 = SortkeyT.UW.SM_AW.Script;
                        pEntry10->rgByte4[k].b2 = SortkeyT.UW.SM_AW.Alpha;
                        pEntry10->rgByte4[k].b3 = SortkeyT.Diacritic;
                        pEntry10->rgByte4[k].b4 = SortkeyT.Case;
                    }

                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b1 = pKSCCantCompress[j].b1;
                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b2 = pKSCCantCompress[j].b2;
                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b3 = pKSCCantCompress[j].b3;
                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b4 = pKSCCantCompress[j].b4;
                    rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)] = reinterpret_cast<BYTE *>(pEntry10);
                }
            }

            return TRUE;
        }
    }

	/*
	Special Turkish sort order defined by Office - different from system's
	default Turkish sort order in that undotted i's sort before dotted i's -
	in the system, they both sort equal. To use, or in the mask with the
	LCID passed to the MsoCompareString* routines.
	*/
    if (fMsoTurkSort)
    {
        // 0x0131 undotted lower case i
        MsoAssertTag(*(rgUnicodeHiByte[0x01]) == 10, ASSERTTAG('eqyd'));
        pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[0x01]);
        pEntry10->rgByte4[0x31].b2 = 0x31;


        // 0x0049 undotted upper case I
        MsoAssertTag(*(rgUnicodeHiByte[0x00]) == 10, ASSERTTAG('eqye'));
        pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[0x00]);
        pEntry10->rgByte4[0x49].b2 = 0x31;
    }
	return TRUE;
    #else
	// now we are in the STATIC LIB case!
	// make ostrman work -- MRuhlen

	int cEntries;
	ENTRYTYPE10 *pEntry10;
	const EXCEPT *pExcept;
	SORTKEY SortkeyT;
	WCHAR wch;
	const WORD *pData;
	const IDEO_INFO *pKSCScript = NULL, *pKSCAlpha = NULL; // *pKSCUniptScript;
	// const WCHAR *pKSCUnipt;
	const IDEO_CANTCOMPRESS *pKSCCantCompress = NULL;
	BOOL fMsoTurkSort = FALSE;

    if (PRIMARYLANGID(LANGIDFROMLCID(Locale)) == LANG_TURKISH && SORTIDFROMLCID(Locale) == MSO_TURKISH_SORT)
    {
        Locale &= MASK_SORT_OFF;
        fMsoTurkSort = TRUE;
    }
    for (int i = 0; i < CEXCEPT; i++)
    {
        if (rgExceptHdr[i].Locale == Locale)
        {
            pExcept = &rgExcept[rgExceptHdr[i].Offset/3];
            cEntries = rgExceptHdr[i].NumEntries;
            MsoAssertTag(cDefExcept == 0, ASSERTTAG('eqyf'));
            
            cDefExcept = 0;
            for (int j = 0; j < cEntries; j++)
            {
                if (*(rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)]) == 10)
                {
                    pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)]);
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b1 = HIBYTE(pExcept[j].Unicode);
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b2 = LOBYTE(pExcept[j].Unicode);
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b3 = pExcept[j].Diacritic;
                    pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b4 = pExcept[j].Case;
                    continue;
                }

                if ((pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(MsoPvAlloc(sizeof(ENTRYTYPE10), msodgMisc))) == NULL)
                {
                    return FALSE;
                }

                rgDefExcept[cDefExcept].bHigh = 	HIBYTE(pExcept[j].UCP);
                rgDefExcept[cDefExcept++].pEntry = rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)];
                MsoAssertTag(cDefExcept < CDEFEXCEPTMAX, ASSERTTAG('eqyg'));

                pEntry10->bID = 10;
                for (WORD k = 0; k < 256; k++)
                {
                    wch = static_cast<WCHAR>((pExcept[j].UCP & 0xff00) | k);
                    GetSortkey(wch, &SortkeyT);
                    pEntry10->rgByte4[k].b1 = SortkeyT.UW.SM_AW.Script;
                    pEntry10->rgByte4[k].b2 = SortkeyT.UW.SM_AW.Alpha;
                    pEntry10->rgByte4[k].b3 = SortkeyT.Diacritic;
                    pEntry10->rgByte4[k].b4 = SortkeyT.Case;
                }

                pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b1 = HIBYTE(pExcept[j].Unicode);
                pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b2 = LOBYTE(pExcept[j].Unicode);
                pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b3 = pExcept[j].Diacritic;
                pEntry10->rgByte4[LOBYTE(pExcept[j].UCP)].b4 = pExcept[j].Case;

                rgUnicodeHiByte[HIBYTE(pExcept[j].UCP)] = reinterpret_cast<BYTE *>(pEntry10);
            }
            break;
        }
    }

    int iScript = 1, iAlpha = 1;
    for (int i = 0; i < CIDEOLCID; i++)
    {
        if (rgIdeoLcidStatic[i].Locale == Locale)
        {
            pData = reinterpret_cast<WORD *>(rgIdeoLcidStatic[i].id);
            if (!pData)
            {
                return FALSE;
            }

            if (Locale == 0x412)		// special case korean
            {
                pKSCScript = reinterpret_cast<IDEO_INFO *>(rgKSCScript);
                pKSCAlpha = reinterpret_cast<IDEO_INFO *>(rgKSCAlpha);
                pKSCCantCompress = reinterpret_cast<IDEO_CANTCOMPRESS *>(rgKSCCantCompress);

                if (!pKSCScript || !pKSCAlpha || !pKSCCantCompress)
                {
                    return FALSE;
                }
            }

            for (int j = 0; j < rgIdeoLcidStatic[i].cData; j++)
            {
                if (!pData[j])
                {
                    continue;
                }

                if (*(rgUnicodeHiByte[HIBYTE(pData[j])]) == 10)
                {
                    pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[HIBYTE(pData[j])]);

                    if (Locale != 0x412)		// special case korean
                    {
                        pEntry10->rgByte4[LOBYTE(pData[j])].b1 = static_cast<BYTE>(128 + j/254);
                        pEntry10->rgByte4[LOBYTE(pData[j])].b2 = static_cast<BYTE>(2 + j%254);
                        pEntry10->rgByte4[LOBYTE(pData[j])].b3 = 2;
                    }
                    else
                    {
                        if (iScript < CKSCSCRIPT && j >= pKSCScript[iScript].Offset)
                        {
                            iScript++;
                        }

                        pEntry10->rgByte4[LOBYTE(pData[j])].b1 = pKSCScript[iScript-1].Value;

                        if (iAlpha < CKSCALPHA && j >= pKSCAlpha[iAlpha].Offset)
                        {
                            iAlpha++;
                        }

                        pEntry10->rgByte4[LOBYTE(pData[j])].b2 = pKSCAlpha[iAlpha-1].Value;
                        pEntry10->rgByte4[LOBYTE(pData[j])].b3 = static_cast<BYTE>(65 + (j-pKSCAlpha[iAlpha-1].Offset)*2);
                    }

                    pEntry10->rgByte4[LOBYTE(pData[j])].b4 = 2;
                    continue;
                }

                if ((pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(MsoPvAlloc(sizeof(ENTRYTYPE10), msodgMisc))) == NULL)
                {
                    return FALSE;
                }

                rgDefExcept[cDefExcept].bHigh = 	HIBYTE(pData[j]);
                rgDefExcept[cDefExcept++].pEntry = rgUnicodeHiByte[HIBYTE(pData[j])];
                MsoAssertTag(cDefExcept < CDEFEXCEPTMAX, ASSERTTAG('eqyh'));

                pEntry10->bID = 10;
                for (WORD k = 0; k < 256; k++)
                {
                    wch = static_cast<WCHAR>((pData[j] & 0xff00) | k);
                    GetSortkey(wch, &SortkeyT);
                    pEntry10->rgByte4[k].b1 = SortkeyT.UW.SM_AW.Script;
                    pEntry10->rgByte4[k].b2 = SortkeyT.UW.SM_AW.Alpha;
                    pEntry10->rgByte4[k].b3 = SortkeyT.Diacritic;
                    pEntry10->rgByte4[k].b4 = SortkeyT.Case;
                }

                if (Locale != 0x412)		// special case korean
                {
                    pEntry10->rgByte4[LOBYTE(pData[j])].b1 = static_cast<BYTE>(128 + j/254);
                    pEntry10->rgByte4[LOBYTE(pData[j])].b2 = static_cast<BYTE>(2 + j%254);
                    pEntry10->rgByte4[LOBYTE(pData[j])].b3 = 2;
                }
                else
                {
                    if (iScript < CKSCSCRIPT && j >= pKSCScript[iScript].Offset)
                    {
                        iScript++;
                    }

                    pEntry10->rgByte4[LOBYTE(pData[j])].b1 = pKSCScript[iScript-1].Value;

                    if (iAlpha < CKSCALPHA && j >= pKSCAlpha[iAlpha].Offset)
                    {
                        iAlpha++;
                    }

                    pEntry10->rgByte4[LOBYTE(pData[j])].b2 = pKSCAlpha[iAlpha-1].Value;
                    pEntry10->rgByte4[LOBYTE(pData[j])].b3 = static_cast<BYTE>(65 + (j-pKSCAlpha[iAlpha-1].Offset)*2);
                }
                rgUnicodeHiByte[HIBYTE(pData[j])] = reinterpret_cast<BYTE *>pEntry10;
            }

            if (Locale == 0x412)		// special case korean
            {
                for (int j = 0; j < CKSCCANTCOMPRESS; j++)
                {
                    if (*(rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)]) == 10)
                    {
                        pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)]);
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b1 = pKSCCantCompress[j].b1;
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b2 = pKSCCantCompress[j].b2;
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b3 = pKSCCantCompress[j].b3;
                        pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b4 = pKSCCantCompress[j].b4;
                        continue;
                    }

                    if ((pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(MsoPvAlloc(sizeof(ENTRYTYPE10), msodgMisc))) == NULL)
                    {
                        return FALSE;
                    }

                    rgDefExcept[cDefExcept].bHigh = 	HIBYTE(pKSCCantCompress[j].wUnicode);
                    rgDefExcept[cDefExcept++].pEntry = rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)];
                    MsoAssertTag(cDefExcept < CDEFEXCEPTMAX, ASSERTTAG('eqyi'));
                    
                    pEntry10->bID = 10;
                    for (WORD k = 0; k < 256; k++)
                    {
                        wch = static_cast<WCHAR>((pKSCCantCompress[j].wUnicode & 0xff00) | k);
                        GetSortkey(wch, &SortkeyT);
                        pEntry10->rgByte4[k].b1 = SortkeyT.UW.SM_AW.Script;
                        pEntry10->rgByte4[k].b2 = SortkeyT.UW.SM_AW.Alpha;
                        pEntry10->rgByte4[k].b3 = SortkeyT.Diacritic;
                        pEntry10->rgByte4[k].b4 = SortkeyT.Case;
                    }

                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b1 = pKSCCantCompress[j].b1;
                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b2 = pKSCCantCompress[j].b2;
                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b3 = pKSCCantCompress[j].b3;
                    pEntry10->rgByte4[LOBYTE(pKSCCantCompress[j].wUnicode)].b4 = pKSCCantCompress[j].b4;
                    rgUnicodeHiByte[HIBYTE(pKSCCantCompress[j].wUnicode)] = reinterpret_cast<BYTE *>(pEntry10);
                }
            }
            return TRUE;
        }
    }

    /*
    Special Turkish sort order defined by Office - different from system's
    default Turkish sort order in that undotted i's sort before dotted i's -
    in the system, they both sort equal. To use, or in the mask with the
    LCID passed to the MsoCompareString* routines.
    */
    if (fMsoTurkSort)
    {
        // 0x0131 undotted lower case i
        MsoAssertTag(*(rgUnicodeHiByte[0x01]) == 10, ASSERTTAG('eqyl'));
        pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[0x01]);

        // pEntry10->rgByte4[0x31].b3 = 1;
        pEntry10->rgByte4[0x31].b2 = 0x31;

        // 0x0049 undotted upper case I
        MsoAssertTag(*(rgUnicodeHiByte[0x00]) == 10, ASSERTTAG('eqym'));

        pEntry10 = reinterpret_cast<ENTRYTYPE10 *>(rgUnicodeHiByte[0x00]);
        pEntry10->rgByte4[0x49].b2 = 0x31;
    }
    return TRUE;
    #endif // STATIC_LIB_DEF || ZENSTAT_LIB_DEF
}

/*------------------------------------------------------------------------
* MsoCompareStringLegacyW
*
* Compares two wide character strings of the same locale according to the
* supplied locale handle.
*
* Note: Different from system's CompareStringW in following ways. If tables
* ever updated be sure to maintain these special cases. Both these will
* be fixed in NT SUR:
* 1. Turkish undotted i sorts before all the other i's in case sensitive
* and insensitive cases.
* 2. Greek final small sigma sorts same as small sigma in case insens. and
* after small sigma in case sens.
*
-------------------------------------------------------------- SHAMIKB -*/
MSOAPI_(int) MsoCompareStringLegacyW(LCID Locale, DWORD dwCmpFlags, LPCWSTR lpString1, int cchCount1, LPCWSTR lpString2, int cchCount2)
{
	LPWSTR pString1;	/* ptr to go thru string 1 */
	LPWSTR pString2;	/* ptr to go thru string 2 */
	PLOC_HASH pHashN = &vHashN;			/* ptr to LOC hash node */
	BOOL fIgnorePunct;			/* flag to ignore punctuation (not symbol) */
	DWORD State;				/* state table */
	DWORD Mask;				/* mask for weights */
	DWORD Weight1;				/* full weight of char - string 1 */
	DWORD Weight2;				/* full weight of char - string 2 */
	int WhichDiacritic;		/* DW => 1 = str1 smaller, 3 = str2 smaller */
	int WhichCase;				/* CW => 1 = str1 smaller, 3 = str2 smaller */
	int WhichPunct1;			/* SW => 1 = str1 smaller, 3 = str2 smaller */
	int WhichPunct2;			/* SW => 1 = str1 smaller, 3 = str2 smaller */
	LPWSTR pSave1;				/* ptr to saved pString1 */
	LPWSTR pSave2;				/* ptr to saved pString2 */
	int cExpChar1=0, cExpChar2=0;	/* ct of expansions in tmp */

	DWORD ExtraWt1, ExtraWt2;	/* extra weight values (for far east) */
	DWORD WhichExtra;			/* XW => wts 4, 5, 6, 7 (for far east) */
	WORD wExcept;

    // ok, now we inline FLoadExceptions by hand for better exception handling
	LCID LocaleT;
	wExcept = 0;

    LocaleT = Locale;
    if (LANGIDFROMLCID(Locale) == LANG_USER_DEFAULT)
    {
        LocaleT = GetUserDefaultLCID();
    }
    if (LANGIDFROMLCID(Locale) == LANG_SYSTEM_DEFAULT)
    {
        LocaleT = GetSystemDefaultLCID();
    }
    if (LocaleT == vTableLocale && vfErrNoExceptions)
    {
        return CompareStringW(Locale, dwCmpFlags, lpString1, cchCount1, lpString2, cchCount2);
    }

    // ok NOW we block
	#if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
	OFFICE_CRITICAL_SECTION ocs;
	#endif

    if (LocaleT == vTableLocale)
    {
        wExcept = EXCPT_TRUE;
    }
    else
    {
        memset(&vHashN, 0, sizeof(LOC_HASH));
        FillpHashN(LocaleT, &vHashN);
        RestoreDefaultTables();
        vTableLocale = LocaleT;

        if (!FAddExceptions(LocaleT))
        {
            vfErrNoExceptions = TRUE;

            return CompareStringW(Locale, dwCmpFlags, lpString1, cchCount1, lpString2, cchCount2);
        }

        vfErrNoExceptions = FALSE;
        wExcept |= EXCPT_CHANGEDLOCALE;
    }

	/*
	*  Call longer compare string if any of the following is true:
	*	- locale is invalid
	*	- compression locale
	*	- either count is not -1
	*	- dwCmpFlags is not 0 or ignore case   (see NOTE below)
	*
	*  NOTE:  If the value of NORM_IGNORECASE ever changes, this
	*		code should check for:
	*			( (dwCmpFlags != 0)  &&  (dwCmpFlags != NORM_IGNORECASE) )
	*		Since NORM_IGNORECASE is equal to 1, we can optimize this
	*		by checking for > 1.
	*/
	dwCmpFlags &= (~LOCALE_USE_CP_ACP);

    if ( (pHashN == NULL) || (pHashN->IfCompression) || (cchCount1 > -1) || (cchCount2 > -1) || (dwCmpFlags > NORM_IGNORECASE) )
    {
        return ( LongCompareStringW( pHashN, dwCmpFlags, lpString1, ((cchCount1 <= -1) ? -2 : cchCount1), lpString2, 
                                     ((cchCount2 <= -1) ? -2 : cchCount2), wExcept) );
    }

    if (IS_INVALID_LOCALE(Locale))
    {
        SetLastError( ERROR_INVALID_PARAMETER );
        return ( 0 );
    }

	/*
	*  Initialize string pointers.
	*/
	pString1 = const_cast<LPWSTR>(lpString1);
	pString2 = const_cast<LPWSTR>(lpString2);

	/*
	*  Invalid Parameter Check:
	*	- NULL string pointers
	*/
    if ((pString1 == NULL) || (pString2 == NULL))
    {
        SetLastError( ERROR_INVALID_PARAMETER );
        return ( 0 );
    }

	/*
	*  Do a wchar by wchar compare.
	*/
    while (TRUE)
    {
        /*
        *  See if characters are equal.
        *  If characters are equal, increment pointers and continue
        *  string compare.
        *
        *  NOTE: Loop is unrolled 8 times for performance.
        */
        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;

        if ((*pString1 != *pString2) || (*pString1 == 0))
        {
            break;
        }

        pString1++;
        pString2++;
    }

	/*
	*  If strings are both at null terminators, return equal.
	*/
    if (*pString1 == *pString2)
    {
        MsoAssertTag(*pString1 == 0, ASSERTTAG('eqyp'));
        return ( 2 );
    }

	/*
	*  Initialize flags, pointers, and counters.
	*/
	fIgnorePunct = FALSE;
	WhichDiacritic = 0;
	WhichCase = 0;
	WhichPunct1 = 0;
	WhichPunct2 = 0;
	pSave1 = NULL;
	pSave2 = NULL;
	ExtraWt1 = (DWORD)0;
	WhichExtra = (DWORD)0;

	/*
	*  Switch on the different flag options.  This will speed up
	*  the comparisons of two strings that are different.
	*
	*  The only two possibilities in this optimized section are
	*  no flags and the ignore case flag.
	*/
    if (dwCmpFlags == 0)
    {
        Mask = CMP_MASKOFF_NONE;
    }
    else
    {
        Mask = CMP_MASKOFF_CW;
    }

    State = (pHashN->IfReverseDW) ? STATE_REVERSE_DW : STATE_DW;
    State |= STATE_CW;

	/*
	*  Compare each character's sortkey weight in the two strings.
	*/
    while ( (*pString1 != 0) && (*pString2 != 0) )
    {
        GetDwordWeight( *pString1, &Weight1 );
        GetDwordWeight( *pString2, &Weight2 );
        Weight1 &= Mask;
        Weight2 &= Mask;

        if (Weight1 != Weight2)
        {
            BYTE sm1 = GET_SCRIPT_MEMBER( &Weight1 );   /* script member 1 */
            BYTE sm2 = GET_SCRIPT_MEMBER( &Weight2 );   /* script member 2 */
            WORD uw1 = GET_UNICODE_SM( &Weight1, sm1 ); /* unicode weight 1 */
            uw1 = WSwapIfMac(uw1);
            WORD uw2 = GET_UNICODE_SM( &Weight2, sm2 ); /* unicode weight 2 */
            uw2 = WSwapIfMac(uw2);
            BYTE dw1 = 0;								/* diacritic weight 1 */
            BYTE dw2 = 0;								/* diacritic weight 2 */
            BOOL fContinue = FALSE;						/* flag to continue loop */
            DWORD Wt = 0;								/* temp weight holder */
            WCHAR pTmpBuf1[MAX_TBL_EXPANSION] = L"";	/* temp buffer for exp 1 */
            WCHAR pTmpBuf2[MAX_TBL_EXPANSION] = L"";	/* temp buffer for exp 2 */


            /*
            *  If Unicode Weights are different and no special cases,
            *  then we're done.  Otherwise, we need to do extra checking.
            *
            *  Must check ENTIRE string for any possibility of Unicode Weight
            *  differences.  As soon as a Unicode Weight difference is found,
            *  then we're done.  If no UW difference is found, then the
            *  first Diacritic Weight difference is used.  If no DW difference
            *  is found, then use the first Case Difference.  If no CW
            *  difference is found, then use the first Extra Weight
            *  difference.  If no EW difference is found, then use the first
            *  Special Weight difference.
            *  difference.
            */
            if ((uw1 != uw2) || (sm1 == FAREAST_SPECIAL))
            {
                /*
                *  Initialize the continue flag.
                */
                fContinue = FALSE;

                /*
                *  Check for Unsortable characters and skip them.
                *  This needs to be outside the switch statement.  If EITHER
                *  character is unsortable, must skip it and start over.
                */
                if (sm1 == UNSORTABLE)
                {
                    pString1++;
                    fContinue = TRUE;
                }
                if (sm2 == UNSORTABLE)
                {
                    pString2++;
                    fContinue = TRUE;
                }
                if (fContinue)
                {
                    continue;
                }

                /*
                *  Switch on the script member of string 1 and take care
                *  of any special cases.
                */
                switch (sm1)
                {
                case ( NONSPACE_MARK ) :
                    {
                        /*
                        *  Nonspace only - look at diacritic weight only.
                        */
                        if ( (WhichDiacritic == 0) || (State & STATE_REVERSE_DW) )
                        {
                            WhichDiacritic = 3;

                            /*
                            *  Remove state from state machine.
                            */
                            REMOVE_STATE( STATE_DW );
                        }

                        /*
                        *  Adjust pointer and set flags.
                        */
                        pString1++;
                        fContinue = TRUE;

                        break;
                    }
                case ( NLS_PUNCTUATION ) :
                    {
                        /*
                        *  If the ignore punctuation flag is set, then skip
                        *  over the punctuation.
                        */
                        if (fIgnorePunct)
                        {
                            pString1++;
                            fContinue = TRUE;
                        }
                        else if (sm2 != NLS_PUNCTUATION)
                        {
                            /*
                            *  The character in the second string is
                            *  NOT punctuation.
                            */
                            if (WhichPunct2)
                            {
                                /*
                                *  Set WP 2 to show that string 2 is smaller,
                                *  since a punctuation char had already been
                                *  found at an earlier position in string 2.
                                *
                                *  Set the Ignore Punctuation flag so we just
                                *  skip over any other punctuation chars in
                                *  the string.
                                */
                                WhichPunct2 = 3;
                                fIgnorePunct = TRUE;
                            }
                            else
                            {
                                /*
                                *  Set WP 1 to show that string 2 is smaller,
                                *  and that string 1 has had a punctuation
                                *  char - since no punctuation chars have
                                *  been found in string 2.
                                */
                                WhichPunct1 = 3;
                            }

                            /*
                            *  Advance pointer 1, and set flag to true.
                            */
                            pString1++;
                            fContinue = TRUE;
                        }

                        /*
                        *  Do NOT want to advance the pointer in string 1 if
                        *  string 2 is also a punctuation char.  This will
                        *  be done later.
                        */

                        break;
                    }
                case ( EXPANSION ) :
                    {
                        /*
                        *  Save pointer in pString1 so that it can be
                        *  restored.
                        */
                        pSave1 = pString1;
                        pString1 = pTmpBuf1;

                        /*
                        *  Expand character into temporary buffer.
                        */
                        pTmpBuf1[0] = GET_EXPANSION_1( &Weight1 );
                        pTmpBuf1[1] = GET_EXPANSION_2( &Weight1 );

                        /*
                        *  Set cExpChar1 to the number of expansion characters
                        *  stored.
                        */
                        cExpChar1 = MAX_TBL_EXPANSION;

                        fContinue = TRUE;
                        break;
                    }
                case ( FAREAST_SPECIAL ) :
                    {
                        /*
                        *  Get the weight for the far east special case
                        *  and store it in Weight1.
                        */
                        GET_FAREAST_WEIGHT( Weight1, uw1, Mask, lpString1, pString1, ExtraWt1 );

                        if (sm2 != FAREAST_SPECIAL)
                        {
                            /*
                            *  The character in the second string is
                            *  NOT a fareast special char.
                            *
                            *  Set each of weights 4, 5, 6, and 7 to show
                            *  that string 2 is smaller (if not already set).
                            */
                            if ( (GET_WT_FOUR( &WhichExtra ) == 0) && (GET_WT_FOUR( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_FOUR( &WhichExtra ) = 3;
                            }

                            if ( (GET_WT_FIVE( &WhichExtra ) == 0) && (GET_WT_FIVE( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_FIVE( &WhichExtra ) = 3;
                            }

                            if ( (GET_WT_SIX( &WhichExtra ) == 0) && (GET_WT_SIX( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_SIX( &WhichExtra ) = 3;
                            }

                            if ( (GET_WT_SEVEN( &WhichExtra ) == 0) && (GET_WT_SEVEN( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_SEVEN( &WhichExtra ) = 3;
                            }
                        }

                        break;
                    }
                case ( RESERVED_2 ) :
                case ( RESERVED_3 ) :
                case ( UNSORTABLE ) :
                    {
                        /*
                        *  Fill out the case statement so the compiler
                        *  will use a jump table.
                        */
                        break;
                    }
                }

                /*
                *  Switch on the script member of string 2 and take care
                *  of any special cases.
                */
                switch (sm2)
                {
                case ( NONSPACE_MARK ) :
                    {
                        /*
                        *  Nonspace only - look at diacritic weight only.
                        */
                        if ( (WhichDiacritic == 0) || (State & STATE_REVERSE_DW) )
                        {
                            WhichDiacritic = 1;

                            /*
                            *  Remove state from state machine.
                            */
                            REMOVE_STATE( STATE_DW );
                        }

                        /*
                        *  Adjust pointer and set flags.
                        */
                        pString2++;
                        fContinue = TRUE;

                        break;
                    }
                case ( NLS_PUNCTUATION ) :
                    {
                        /*
                        *  If the ignore punctuation flag is set, then skip
                        *  over the punctuation.
                        */
                        if (fIgnorePunct)
                        {
                            /*
                            *  Pointer 2 will be advanced after if-else
                            *  statement.
                            */
                            ;
                        }
                        else if (sm1 != NLS_PUNCTUATION)
                        {
                            /*
                            *  The character in the first string is
                            *  NOT punctuation.
                            */
                            if (WhichPunct1)
                            {
                                /*
                                *  Set WP 1 to show that string 1 is smaller,
                                *  since a punctuation char had already
                                *  been found at an earlier position in
                                *  string 1.
                                *
                                *  Set the Ignore Punctuation flag so we just
                                *  skip over any other punctuation in the
                                *  string.
                                */
                                WhichPunct1 = 1;
                                fIgnorePunct = TRUE;
                            }
                            else
                            {
                                /*
                                *  Set WP 2 to show that string 1 is smaller,
                                *  and that string 2 has had a punctuation
                                *  char - since no punctuation chars have
                                *  been found in string 1.
                                */
                                WhichPunct2 = 1;
                            }

                            /*
                            *  Pointer 2 will be advanced after if-else
                            *  statement.
                            */
                        }
                        else
                        {
                            /*
                            *  Both code points are punctuation.
                            *
                            *  See if either of the strings has encountered
                            *  punctuation chars previous to this.
                            */
                            if (WhichPunct1)
                            {
                                /*
                                *  String 1 has had a punctuation char, so
                                *  it should be the smaller string (since
                                *  both have punctuation chars).
                                */
                                WhichPunct1 = 1;
                            }
                            else if (WhichPunct2)
                            {
                                /*
                                *  String 2 has had a punctuation char, so
                                *  it should be the smaller string (since
                                *  both have punctuation chars).
                                */
                                WhichPunct2 = 3;
                            }
                            else
                            {
                                /*
                                *  Position is the same, so compare the
                                *  special weights.  Set WhichPunct1 to
                                *  the smaller special weight.
                                */
                                WhichPunct1 = ( ( (GET_ALPHA_NUMERIC( &Weight1 ) < GET_ALPHA_NUMERIC( &Weight2 )) ) ? 1 : 3 );
                            }

                            /*
                            *  Set the Ignore Punctuation flag so we just
                            *  skip over any other punctuation in the string.
                            */
                            fIgnorePunct = TRUE;

                            /*
                            *  Advance pointer 1.  Pointer 2 will be
                            *  advanced after if-else statement.
                            */
                            pString1++;
                        }

                        /*
                        *  Advance pointer 2 and set flag to true.
                        */
                        pString2++;
                        fContinue = TRUE;

                        break;
                    }
                case ( EXPANSION ) :
                    {
                        /*
                        *  Save pointer in pString1 so that it can be
                        *  restored.
                        */
                        pSave2 = pString2;
                        pString2 = pTmpBuf2;

                        /*
                        *  Expand character into temporary buffer.
                        */
                        pTmpBuf2[0] = GET_EXPANSION_1( &Weight2 );
                        pTmpBuf2[1] = GET_EXPANSION_2( &Weight2 );

                        /*
                        *  Set cExpChar2 to the number of expansion characters
                        *  stored.
                        */
                        cExpChar2 = MAX_TBL_EXPANSION;

                        fContinue = TRUE;
                        break;
                    }
                case ( FAREAST_SPECIAL ) :
                    {
                        /*
                        *  Get the weight for the far east special case
                        *  and store it in Weight2.
                        */
                        GET_FAREAST_WEIGHT( Weight2, uw2, Mask, lpString2, pString2, ExtraWt2 );

                        if (sm1 != FAREAST_SPECIAL)
                        {
                            /*
                            *  The character in the first string is
                            *  NOT a fareast special char.
                            *
                            *  Set each of weights 4, 5, 6, and 7 to show
                            *  that string 1 is smaller (if not already set).
                            */
                            if ( (GET_WT_FOUR( &WhichExtra ) == 0) && (GET_WT_FOUR( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_FOUR( &WhichExtra ) = 1;
                            }
                            
                            if ( (GET_WT_FIVE( &WhichExtra ) == 0) && (GET_WT_FIVE( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_FIVE( &WhichExtra ) = 1;
                            }

                            if ( (GET_WT_SIX( &WhichExtra ) == 0) && (GET_WT_SIX( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_SIX( &WhichExtra ) = 1;
                            }
                            
                            if ( (GET_WT_SEVEN( &WhichExtra ) == 0) && (GET_WT_SEVEN( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_SEVEN( &WhichExtra ) = 1;
                            }
                        }
                        else
                        {
                            /*
                            *  Characters in both strings are fareast
                            *  special chars.
                            *
                            *  Set each of weights 4, 5, 6, and 7
                            *  appropriately (if not already set).
                            */
                            if ( (GET_WT_FOUR( &WhichExtra ) == 0) && ( GET_WT_FOUR( &ExtraWt1 ) != GET_WT_FOUR( &ExtraWt2 ) ) )
                            {
                                GET_WT_FOUR( &WhichExtra ) = static_cast<BYTE>(( GET_WT_FOUR( &ExtraWt1 ) < GET_WT_FOUR( &ExtraWt2 ) ) ? 1 : 3);
                            }

                            if ( (GET_WT_FIVE( &WhichExtra ) == 0) && ( GET_WT_FIVE( &ExtraWt1 ) != GET_WT_FIVE( &ExtraWt2 ) ) )
                            {
                                GET_WT_FIVE( &WhichExtra ) = static_cast<BYTE>(( GET_WT_FIVE( &ExtraWt1 ) < GET_WT_FIVE( &ExtraWt2 ) ) ? 1 : 3);
                            }
                            if ( (GET_WT_SIX( &WhichExtra ) == 0) && ( GET_WT_SIX( &ExtraWt1 ) != GET_WT_SIX( &ExtraWt2 ) ) )
                            {
                                GET_WT_SIX( &WhichExtra ) = static_cast<BYTE>(( GET_WT_SIX( &ExtraWt1 ) < GET_WT_SIX( &ExtraWt2 ) ) ? 1 : 3);
                            }

                            if ( (GET_WT_SEVEN( &WhichExtra ) == 0) && ( GET_WT_SEVEN( &ExtraWt1 ) != GET_WT_SEVEN( &ExtraWt2 ) ) )
                            {
                                GET_WT_SEVEN( &WhichExtra ) = static_cast<BYTE>(( GET_WT_SEVEN( &ExtraWt1 ) < GET_WT_SEVEN( &ExtraWt2 ) ) ? 1 : 3);
                            }
                        }

                        break;
                    }
                case ( RESERVED_2 ) :
                case ( RESERVED_3 ) :
                case ( UNSORTABLE ) :
                    {
                        /*
                        *  Fill out the case statement so the compiler
                        *  will use a jump table.
                        */
                        break;
                    }
                }

                /*
                *  See if the comparison should start again.
                */
                if (fContinue)
                {
                    continue;
                }

                /*
                *  We're not supposed to drop down into the state table if
                *  unicode weights are different, so stop comparison and
                *  return result of unicode weight comparison.
                */
                if (uw1 != uw2)
                {
                    return ( (uw1 < uw2) ? 1 : 3 );
                }
            }

            /*
            *  For each state in the state table, do the appropriate
            *  comparisons.	(UW1 == UW2)
            */
            if (State & (STATE_DW | STATE_REVERSE_DW))
            {
                /*
                *  Get the diacritic weights.
                */
                dw1 = GET_DIACRITIC( &Weight1 );
                dw2 = GET_DIACRITIC( &Weight2 );

                if (dw1 != dw2)
                {
                    /*
                    *  Look ahead to see if diacritic follows a
                    *  minimum diacritic weight.  If so, get the
                    *  diacritic weight of the nonspace mark.
                    */
                    while (*(pString1 + 1) != 0)
                    {
                        GetDwordWeight( *(pString1 + 1), &Wt );
                        if (GET_SCRIPT_MEMBER( &Wt ) == NONSPACE_MARK)
                        {
                            dw1 = (BYTE)(dw1 + GET_DIACRITIC( &Wt ));
                            pString1++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    while (*(pString2 + 1) != 0)
                    {
                        GetDwordWeight( *(pString2 + 1), &Wt );
                        if (GET_SCRIPT_MEMBER( &Wt ) == NONSPACE_MARK)
                        {
                            dw2 = (BYTE)(dw2 + GET_DIACRITIC( &Wt ));
                            pString2++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    /*
                    *  Save which string has the smaller diacritic
                    *  weight if the diacritic weights are still
                    *  different.
                    */
                    if (dw1 != dw2)
                    {
                        WhichDiacritic = (dw1 < dw2) ? 1 : 3;

                        /*
                        *  Remove state from state machine.
                        */
                        REMOVE_STATE( STATE_DW );
                    }
                }
            }
            if (State & STATE_CW)
            {
                /*
                *  Get the case weights.
                */
                if (GET_CASE( &Weight1 ) != GET_CASE( &Weight2 ))
                {
                    /*
                    *  Save which string has the smaller case weight.
                    */
                    WhichCase = (GET_CASE( &Weight1 ) < GET_CASE( &Weight2 )) ? 1 : 3;

                    /*
                    *  Remove state from state machine.
                    */
                    REMOVE_STATE( STATE_CW );
                }
            }
        }

        /*
        *  Fixup the pointers.
        */
        POINTER_FIXUP();
    }

    /*
    *  If the end of BOTH strings has been reached, then the unicode
    *  weights match exactly.  Check the diacritic, case and special
    *  weights.  If all are zero, then return success.  Otherwise,
    *  return the result of the weight difference.
    *
    *  NOTE:  The following checks MUST REMAIN IN THIS ORDER:
    *			Diacritic, Case, Punctuation.
    */
    if (*pString1 == 0)
    {
        if (*pString2 == 0)
        {
            if (WhichDiacritic)
            {
                return ( WhichDiacritic );
            }
            if (WhichCase)
            {
                return ( WhichCase );
            }
            if (WhichExtra)
            {
                if (GET_WT_FOUR( &WhichExtra ))
                {
                    return ( GET_WT_FOUR( &WhichExtra ) );
                }

                if (GET_WT_FIVE( &WhichExtra ))
                {
                    return ( GET_WT_FIVE( &WhichExtra ) );
                }

                if (GET_WT_SIX( &WhichExtra ))
                {
                    return ( GET_WT_SIX( &WhichExtra ) );
                }

                if (GET_WT_SEVEN( &WhichExtra ))
                {
                    return ( GET_WT_SEVEN( &WhichExtra ) );
                }
            }
            if (WhichPunct1)
            {
                return ( WhichPunct1 );
            }
            if (WhichPunct2)
            {
                return ( WhichPunct2 );
            }

            return ( 2 );
        }
        else
        {
            /*
            *  String 2 is longer.
            */
            pString1 = pString2;
        }
    }

    /*
    *  Scan to the end of the longer string.
    */
    return QuickScanLongerString( pString1, ((*pString2 == 0) ? 3 : 1), WhichDiacritic, WhichCase, WhichExtra, 
                                   WhichPunct1, WhichPunct2, wExcept);
}

/*-----------------------------------------------------------------------------
	MsoIsValidLocale

	This function is equivalent to OS API call IsValidLocale(), except this
	version provides caching mechanism.
------------------------------------------------------------------- NOBUYAH -*/
MSOAPI_(BOOL) MsoIsValidLocale(LCID lcid, DWORD dwFlags)
{
	static LCID lcidCache = 0;
	static DWORD dwFlagsCache = 0;
	static BOOL fValidCache;

    // check to see if cached
    if (lcid == lcidCache && dwFlags == dwFlagsCache)
    {
        MsoAssertSzTag(IsValidLocale(lcid, dwFlags) == fValidCache, "MsoIsValidLocale() cache is incorrect", ASSERTTAG('qvgl'));

        return fValidCache;
    }

    // not in cache, call OS and cache results
    fValidCache = IsValidLocale(lcid, dwFlags);
    lcidCache = lcid;
    dwFlagsCache = dwFlags;

    return fValidCache;
}

MSOAPI_(int) MsoCompareStringW(LCID Locale, DWORD dwCmpFlags, LPCWSTR lpString1, int cchCount1, LPCWSTR lpString2, int cchCount2)
{
	LCID LocaleT;

    LocaleT = Locale;
    if (LANGIDFROMLCID(Locale) == LANG_USER_DEFAULT)
    {
        LocaleT = GetUserDefaultLCID();
    }
    if (LANGIDFROMLCID(Locale) == LANG_SYSTEM_DEFAULT)
    {
        LocaleT = LOCALE_USER_DEFAULT;
    }

    // starting with Office.NET, we will now use system's sort table unless the
    // locale is not installed, in which case we fall back on our own built-in
    // table
    if (MsoIsValidLocale(LocaleT, LCID_INSTALLED) && PRIMARYLANGID(LANGIDFROMLCID(Locale)) != LANG_TURKISH)
    {
        return CompareStringW(LocaleT, dwCmpFlags, lpString1, cchCount1, lpString2, cchCount2);
    }

    return MsoCompareStringLegacyW(Locale, dwCmpFlags, lpString1, cchCount1, lpString2, cchCount2);
}

/*-------------------------------------------------------------------------*\
 *						INTERNAL ROUTINES							*
\*-------------------------------------------------------------------------*/

/*------------------------------------------------------------------------
* LongCompareStringW
*
* Compares two wide character strings of the same locale according to the
* supplied locale handle.
*
-------------------------------------------------------------- SHAMIKB -*/
int LongCompareStringW(PLOC_HASH pHashN, DWORD dwCmpFlags, LPCWSTR lpString1, int cchCount1, 
                       LPCWSTR lpString2, int cchCount2, WORD wExcept)
{
	int ctr1 = cchCount1;		/* loop counter for string 1 */
	int ctr2 = cchCount2;		/* loop counter for string 2 */
	register LPWSTR pString1;	/* ptr to go thru string 1 */
	register LPWSTR pString2;	/* ptr to go thru string 2 */
	BOOL IfCompress;			/* if compression in locale */
	BOOL IfDblCompress1;		/* if double compression in string 1 */
	BOOL IfDblCompress2;		/* if double compression in string 2 */
	BOOL fEnd1;				/* if at end of string 1 */
	BOOL fIgnorePunct;			/* flag to ignore punctuation (not symbol) */
	BOOL fIgnoreDiacritic;		/* flag to ignore diacritics */
	BOOL fIgnoreSymbol;		/* flag to ignore symbols */
	BOOL fStringSort;			/* flag to use string sort */
	DWORD State;				/* state table */
	DWORD Mask;				/* mask for weights */
	DWORD Weight1;				/* full weight of char - string 1 */
	DWORD Weight2;				/* full weight of char - string 2 */
	int WhichDiacritic;		/* DW => 1 = str1 smaller, 3 = str2 smaller */
	int WhichCase;				/* CW => 1 = str1 smaller, 3 = str2 smaller */
	int WhichPunct1;			/* SW => 1 = str1 smaller, 3 = str2 smaller */
	int WhichPunct2;			/* SW => 1 = str1 smaller, 3 = str2 smaller */
	LPWSTR pSave1;				/* ptr to saved pString1 */
	LPWSTR pSave2;				/* ptr to saved pString2 */
	int cExpChar1 = 0, cExpChar2 = 0;	/* ct of expansions in tmp */

	DWORD ExtraWt1, ExtraWt2;	/* extra weight values (for far east) */
	DWORD WhichExtra;			/* XW => wts 4, 5, 6, 7 (for far east) */

	/*
	*  Initialize string pointers.
	*/
	pString1 = const_cast<LPWSTR>(lpString1);
	pString2 = const_cast<LPWSTR>(lpString2);

	/*
	*  Invalid Parameter Check:
	*	- invalid locale (hash node)
	*	- either string is null
	*/
    if ( (pHashN == NULL) || (pString1 == NULL) || (pString2 == NULL) )
    {
        SetLastError( ERROR_INVALID_PARAMETER );
        return ( 0 );
    }

	/*
	*  Invalid Flags Check:
	*	- invalid flags
	*/
    if (dwCmpFlags & CS_INVALID_FLAG)
    {
        SetLastError( ERROR_INVALID_FLAGS );
        return ( 0 );
    }

    /*
    *  Check if compression in the given locale.  If not, then
    *  try a wchar by wchar compare.  If strings are equal, this
    *  will be quick.
    */
    if ((IfCompress = pHashN->IfCompression) == FALSE)
    {
        /*
        *  Compare each wide character in the two strings.
        */
        while ( NOT_END_STRING( ctr1, pString1, cchCount1 ) && NOT_END_STRING( ctr2, pString2, cchCount2 ) )
        {
            /*
            *  See if characters are equal.
            */
            if (*pString1 == *pString2)
            {
                /*
                *  Characters are equal, so increment pointers,
                *  decrement counters, and continue string compare.
                */
                pString1++;
                pString2++;
                ctr1--;
                ctr2--;
            }
            else
            {
                /*
                *  Difference was found.  Fall into the sortkey
                *  check below.
                */
                break;
            }
        }

        /*
        *  If the end of BOTH strings has been reached, then the strings
        *  match exactly.  Return success.
        */
        if ( AT_STRING_END( ctr1, pString1, cchCount1 ) && AT_STRING_END( ctr2, pString2, cchCount2 ) )
        {
            return ( 2 );
        }
    }

	/*
	*  Initialize flags, pointers, and counters.
	*/
	fIgnorePunct = dwCmpFlags & NORM_IGNORESYMBOLS;
	fIgnoreDiacritic = dwCmpFlags & NORM_IGNORENONSPACE;
	fIgnoreSymbol = fIgnorePunct;
	fStringSort = dwCmpFlags & SORT_STRINGSORT;
	WhichDiacritic = 0;
	WhichCase = 0;
	WhichPunct1 = 0;
	WhichPunct2 = 0;
	pSave1 = NULL;
	pSave2 = NULL;
	ExtraWt1 = (DWORD)0;
	WhichExtra = (DWORD)0;

	/*
	*  Set the weights to be invalid.  This flags whether or not to
	*  recompute the weights next time through the loop.  It also flags
	*  whether or not to start over (continue) in the loop.
	*/
	Weight1 = CMP_INVALID_WEIGHT;
	Weight2 = CMP_INVALID_WEIGHT;

	/*
	*  Switch on the different flag options.  This will speed up
	*  the comparisons of two strings that are different.
	*/
	State = STATE_CW;
	Mask = CMP_MASKOFF_NONE;	//  NOTE:  PETERO:  This is to make the compiler happy

	//  as it isn't able to realize that all 4 possible cases are handled in this switch
	//  and there is no way for Mask to not be initialized.
    switch (dwCmpFlags & (NORM_IGNORECASE | NORM_IGNORENONSPACE))
    {
    case ( 0 ) :
        {
            Mask = CMP_MASKOFF_NONE;
            State |= (pHashN->IfReverseDW) ? STATE_REVERSE_DW : STATE_DW;

            break;
        }

    case ( NORM_IGNORECASE ) :
        {
            Mask = CMP_MASKOFF_CW;
            State |= (pHashN->IfReverseDW) ? STATE_REVERSE_DW : STATE_DW;

            break;
        }

    case ( NORM_IGNORENONSPACE ) :
        {
            Mask = CMP_MASKOFF_DW;

            break;
        }

    case ( NORM_IGNORECASE | NORM_IGNORENONSPACE ) :
        {
            Mask = CMP_MASKOFF_DW_CW;

            break;
        }
    }

    switch (dwCmpFlags & (NORM_IGNOREKANATYPE | NORM_IGNOREWIDTH))
    {
    case ( 0 ) :
        {
            break;
        }

    case ( NORM_IGNOREKANATYPE ) :
        {
            Mask &= CMP_MASKOFF_KANA;

            break;
        }

    case ( NORM_IGNOREWIDTH ) :
        {
            Mask &= CMP_MASKOFF_WIDTH;

            if (dwCmpFlags & NORM_IGNORECASE)
            {
                REMOVE_STATE( STATE_CW );
            }

            break;
        }

    case ( NORM_IGNOREKANATYPE | NORM_IGNOREWIDTH ) :
        {
            Mask &= CMP_MASKOFF_KANA_WIDTH;

            if (dwCmpFlags & NORM_IGNORECASE)
            {
                REMOVE_STATE( STATE_CW );
            }

            break;
        }
    }

    /*
    *  Compare each character's sortkey weight in the two strings.
    */
    while ( NOT_END_STRING(ctr1, pString1, cchCount1) && NOT_END_STRING(ctr2, pString2, cchCount2) )
    {
        if (Weight1 == CMP_INVALID_WEIGHT)
        {
            GetDwordWeight( *pString1, &Weight1 );
            Weight1 &= Mask;
        }
        if (Weight2 == CMP_INVALID_WEIGHT)
        {
            GetDwordWeight( *pString2, &Weight2 );
            Weight2 &= Mask;
        }

        /*
        *  If compression locale, then need to check for compression
        *  characters even if the weights are equal.  If it's not a
        *  compression locale, then we don't need to check anything
        *  if the weights are equal.
        */
        if ( (IfCompress) && (GET_COMPRESSION( &Weight1 ) || GET_COMPRESSION( &Weight2 )) )
        {
            int ctr;				/* loop counter */
            PCOMPRESS_3 pComp3;		/* ptr to compress 3 table */
            PCOMPRESS_2 pComp2;		/* ptr to compress 2 table */
            int If1;				/* if compression found in string 1 */
            int If2;				/* if compression found in string 2 */
            int CompVal;			/* compression value */
            int IfEnd1;				/* if exists 1 more char in string 1 */
            int IfEnd2;				/* if exists 1 more char in string 2 */


            /*
            *  Check for compression in the weights.
            */
            If1 = GET_COMPRESSION( &Weight1 );
            If2 = GET_COMPRESSION( &Weight2 );
            CompVal = ( (If1 > If2) ? If1 : If2 );

            if (ctr1 >= 1)
            {
                IfEnd1 = AT_STRING_END( ctr1 - 1, pString1 + 1, cchCount1 );
            }
            else
            {
                IfEnd1 = true;
            }

            if (ctr2 >= 1)
            {
                IfEnd2 = AT_STRING_END( ctr2 - 1, pString2 + 1, cchCount2 );
            }
            else
            {
                IfEnd2 = true;
            }

            if (pHashN->IfDblCompression == FALSE)
            {
                /*
                *  NO double compression, so don't check for it.
                */
                switch ( CompVal )
                {
                    /*
                    *  Check for 3 characters compressing to 1.
                    */
                case ( COMPRESS_3_MASK ) :
                    {
                        /*
                        *  Check character in string 1 and string 2.
                        */
                        if ( ((If1) && (!IfEnd1) && !AT_STRING_END( ctr1 - 2, pString1 + 2, cchCount1 )) ||
                            ((If2) && (!IfEnd2) && !AT_STRING_END( ctr2 - 2, pString2 + 2, cchCount2 )) )
                        {
                            ctr = pHashN->pCompHdr->Num3;
                            pComp3 = pHashN->pCompress3;
                            for (; ctr > 0; ctr--, pComp3++)
                            {
                                /*
                                *  Check character in string 1.
                                */
                                if ( (If1) && (pComp3->UCP1 == *pString1) &&
                                     (pComp3->UCP2 == *(pString1 + 1)) && (pComp3->UCP3 == *(pString1 + 2)) )
                                {
                                    /*
                                    *  Found compression for string 1.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight1 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp3->Weights ));
                                    Weight1 &= Mask;
                                    pString1 += 2;
                                    ctr1 -= 2;

                                    /*
                                    *  Set boolean for string 1 - search is
                                    *  complete.
                                    */
                                    If1 = 0;

                                    /*
                                    *  Break out of loop if both searches are
                                    *  done.
                                    */
                                    if (If2 == 0)
                                    {
                                        break;
                                    }
                                }

                                /*
                                *  Check character in string 2.
                                */
                                if ( (If2) && (pComp3->UCP1 == *pString2) &&
                                     (pComp3->UCP2 == *(pString2 + 1)) && (pComp3->UCP3 == *(pString2 + 2)) )
                                {
                                    /*
                                    *  Found compression for string 2.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight2 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp3->Weights ));
                                    Weight2 &= Mask;
                                    pString2 += 2;
                                    ctr2 -= 2;

                                    /*
                                    *  Set boolean for string 2 - search is
                                    *  complete.
                                    */
                                    If2 = 0;

                                    /*
                                    *  Break out of loop if both searches are
                                    *  done.
                                    */
                                    if (If1 == 0)
                                    {
                                        break;
                                    }
                                }
                            }
                            if (ctr > 0)
                            {
                                break;
                            }
                        }

                        /*
                        *  Fall through if not found.
                        */
                        __fallthrough;
                    }

                    /*
                    *  Check for 2 characters compressing to 1.
                    */
                case ( COMPRESS_2_MASK ) :
                    {
                        /*
                        *  Check character in string 1 and string 2.
                        */
                        if ( ((If1) && (!IfEnd1)) || ((If2) && (!IfEnd2)) )
                        {
                            ctr = pHashN->pCompHdr->Num2;
                            pComp2 = pHashN->pCompress2;
                            for (; ((ctr > 0) && (If1 || If2)); ctr--, pComp2++)
                            {
                                /*
                                *  Check character in string 1.
                                */
                                if ( (If1) && (pComp2->UCP1 == *pString1) && (pComp2->UCP2 == *(pString1 + 1)) )
                                {
                                    /*
                                    *  Found compression for string 1.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight1 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp2->Weights ));
                                    Weight1 &= Mask;
                                    pString1++;
                                    ctr1--;

                                    /*
                                    *  Set boolean for string 1 - search is
                                    *  complete.
                                    */
                                    If1 = 0;

                                    /*
                                    *  Break out of loop if both searches are
                                    *  done.
                                    */
                                    if (If2 == 0)
                                    {
                                        break;
                                    }
                                }

                                /*
                                *  Check character in string 2.
                                */
                                if ( (If2) && (pComp2->UCP1 == *pString2) && (pComp2->UCP2 == *(pString2 + 1)) )
                                {
                                    /*
                                    *  Found compression for string 2.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight2 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp2->Weights ));
                                    Weight2 &= Mask;
                                    pString2++;
                                    ctr2--;

                                    /*
                                    *  Set boolean for string 2 - search is
                                    *  complete.
                                    */
                                    If2 = 0;

                                    /*
                                    *  Break out of loop if both searches are
                                    *  done.
                                    */
                                    if (If1 == 0)
                                    {
                                        break;
                                    }
                                }
                            }
                            if (ctr > 0)
                            {
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                /*
                *  Double Compression exists, so must check for it.
                */
                IfDblCompress1 = 0;
                if ( !IfEnd1 && ((IfDblCompress1 = (*pString1 == *(pString1 + 1))) != 0) )
                {
                    /*
                    *  Advance past the first code point to get to the
                    *  compression character.
                    */
                    pString1++;
                    ctr1--;
                }

                IfDblCompress2 = 0;
                if ( !IfEnd2 && ((IfDblCompress2 = (*pString2 == *(pString2 + 1))) != 0) )
                {
                    /*
                    *  Advance past the first code point to get to the
                    *  compression character.
                    */
                    pString2++;
                    ctr2--;
                }

                switch ( CompVal )
                {
                    /*
                    *  Check for 3 characters compressing to 1.
                    */
                case ( COMPRESS_3_MASK ) :
                    {
                        /*
                        *  Check character in string 1.
                        */
                        if ( (If1) && (!IfEnd1) && !AT_STRING_END( ctr1 - 2, pString1 + 2, cchCount1 ) )
                        {
                            ctr = pHashN->pCompHdr->Num3;
                            pComp3 = pHashN->pCompress3;
                            for (; ctr > 0; ctr--, pComp3++)
                            {
                                /*
                                *  Check character in string 1.
                                */
                                if ( (pComp3->UCP1 == *pString1) && (pComp3->UCP2 == *(pString1 + 1)) && (pComp3->UCP3 == *(pString1 + 2)) )
                                {
                                    /*
                                    *  Found compression for string 1.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight1 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp3->Weights ));
                                    Weight1 &= Mask;
                                    if (!IfDblCompress1)
                                    {
                                        pString1 += 2;
                                        ctr1 -= 2;
                                    }

                                    /*
                                    *  Set boolean for string 1 - search is
                                    *  complete.
                                    */
                                    If1 = 0;
                                    break;
                                }
                            }
                        }

                        /*
                        *  Check character in string 2.
                        */
                        if ( (If2) && (!IfEnd2) && !AT_STRING_END( ctr2 - 2, pString2 + 2, cchCount2 ) )
                        {
                            ctr = pHashN->pCompHdr->Num3;
                            pComp3 = pHashN->pCompress3;
                            for (; ctr > 0; ctr--, pComp3++)
                            {
                                /*
                                *  Check character in string 2.
                                */
                                if ( (pComp3->UCP1 == *pString2) && (pComp3->UCP2 == *(pString2 + 1)) && (pComp3->UCP3 == *(pString2 + 2)) )
                                {
                                    /*
                                    *  Found compression for string 2.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight2 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp3->Weights ));
                                    Weight2 &= Mask;
                                    if (!IfDblCompress2)
                                    {
                                        pString2 += 2;
                                        ctr2 -= 2;
                                    }

                                    /*
                                    *  Set boolean for string 2 - search is
                                    *  complete.
                                    */
                                    If2 = 0;
                                    break;
                                }
                            }
                        }

                        /*
                        *  Fall through if not found.
                        */
                        if ((If1 == 0) && (If2 == 0))
                        {
                            break;
                        }

                        __fallthrough;
                    }

                    /*
                    *  Check for 2 characters compressing to 1.
                    */
                case ( COMPRESS_2_MASK ) :
                    {
                        /*
                        *  Check character in string 1.
                        */
                        if ( (If1) && (!IfEnd1) )
                        {
                            ctr = pHashN->pCompHdr->Num2;
                            pComp2 = pHashN->pCompress2;
                            for (; ctr > 0; ctr--, pComp2++)
                            {
                                /*
                                *  Check character in string 1.
                                */
                                if ( (pComp2->UCP1 == *pString1) && (pComp2->UCP2 == *(pString1 + 1)) )
                                {
                                    /*
                                    *  Found compression for string 1.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight1 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp2->Weights ));
                                    Weight1 &= Mask;
                                    if (!IfDblCompress1)
                                    {
                                        pString1++;
                                        ctr1--;
                                    }

                                    /*
                                    *  Set boolean for string 1 - search is
                                    *  complete.
                                    */
                                    If1 = 0;
                                    break;
                                }
                            }
                        }

                        /*
                        *  Check character in string 2.
                        */
                        if ( (If2) && (!IfEnd2) )
                        {
                            ctr = pHashN->pCompHdr->Num2;
                            pComp2 = pHashN->pCompress2;
                            for (; ctr > 0; ctr--, pComp2++)
                            {
                                /*
                                *  Check character in string 2.
                                */
                                if ( (pComp2->UCP1 == *pString2) && (pComp2->UCP2 == *(pString2 + 1)) )
                                {
                                    /*
                                    *  Found compression for string 2.
                                    *  Get new weight and mask it.
                                    *  Increment pointer and decrement counter.
                                    */
                                    Weight2 = HiWordSwapIfMac(MAKE_SORTKEY_DWORD( pComp2->Weights ));
                                    Weight2 &= Mask;
                                    if (!IfDblCompress2)
                                    {
                                        pString2++;
                                        ctr2--;
                                    }

                                    /*
                                    *  Set boolean for string 2 - search is
                                    *  complete.
                                    */
                                    If2 = 0;
                                    break;
                                }
                            }
                        }
                    }
                }

                /*
                *  Reset the pointer back to the beginning of the double
                *  compression.  Pointer fixup at the end will advance
                *  them correctly.
                *
                *  If double compression, we advanced the pointer at
                *  the beginning of the switch statement.  If double
                *  compression character was actually found, the pointer
                *  was NOT advanced.  We now want to decrement the pointer
                *  to put it back to where it was.
                *
                *  The next time through, the pointer will be pointing to
                *  the regular compression part of the string.
                */
                if (IfDblCompress1)
                {
                    pString1--;
                    ctr1++;
                }
                if (IfDblCompress2)
                {
                    pString2--;
                    ctr2++;
                }
            }
        }

        /*
        *  Check the weights again.
        */
        if (Weight1 != Weight2)
        {
            /*
            *  Weights are still not equal, even after compression
            *  check, so compare the different weights.
            */
            BYTE sm1 = GET_SCRIPT_MEMBER( &Weight1 );   /* script member 1 */
            BYTE sm2 = GET_SCRIPT_MEMBER( &Weight2 );   /* script member 2 */
            WORD uw1 = GET_UNICODE_SM( &Weight1, sm1 ); /* unicode weight 1 */
            uw1 = WSwapIfMac(uw1);
            WORD uw2 = GET_UNICODE_SM( &Weight2, sm2 ); /* unicode weight 2 */
            uw2 = WSwapIfMac(uw2);
            BYTE dw1 = 0;								/* diacritic weight 1 */
            BYTE dw2 = 0;								/* diacritic weight 2 */
            DWORD Wt = 0;								/* temp weight holder */
            WCHAR pTmpBuf1[MAX_TBL_EXPANSION] = L"";	/* temp buffer for exp 1 */
            WCHAR pTmpBuf2[MAX_TBL_EXPANSION] = L"";	/* temp buffer for exp 2 */

            /*
            *  If Unicode Weights are different and no special cases,
            *  then we're done.  Otherwise, we need to do extra checking.
            *
            *  Must check ENTIRE string for any possibility of Unicode Weight
            *  differences.  As soon as a Unicode Weight difference is found,
            *  then we're done.  If no UW difference is found, then the
            *  first Diacritic Weight difference is used.  If no DW difference
            *  is found, then use the first Case Difference.  If no CW
            *  difference is found, then use the first Extra Weight
            *  difference.  If no EW difference is found, then use the first
            *  Special Weight difference.
            */
            if ( (uw1 != uw2) || ((sm1 <= SYMBOL_5) && (sm1 >= FAREAST_SPECIAL)) )
            {
                /*
                *  Check for Unsortable characters and skip them.
                *  This needs to be outside the switch statement.  If EITHER
                *  character is unsortable, must skip it and start over.
                */
                if (sm1 == UNSORTABLE)
                {
                    pString1++;
                    ctr1--;
                    Weight1 = CMP_INVALID_WEIGHT;
                }
                if (sm2 == UNSORTABLE)
                {
                    pString2++;
                    ctr2--;
                    Weight2 = CMP_INVALID_WEIGHT;
                }

                /*
                *  Check for Ignore Nonspace and Ignore Symbol.  If
                *  Ignore Nonspace is set and either character is a
                *  nonspace mark only, then we need to advance the
                *  pointer to skip over the character and continue.
                *  If Ignore Symbol is set and either character is a
                *  punctuation char, then we need to advance the
                *  pointer to skip over the character and continue.
                *
                *  This step is necessary so that a string with a
                *  nonspace mark and a punctuation char following one
                *  another are properly ignored when one or both of
                *  the ignore flags is set.
                */
                if (fIgnoreDiacritic)
                {
                    if (sm1 == NONSPACE_MARK)
                    {
                        pString1++;
                        ctr1--;
                        Weight1 = CMP_INVALID_WEIGHT;
                    }
                    if (sm2 == NONSPACE_MARK)
                    {
                        pString2++;
                        ctr2--;
                        Weight2 = CMP_INVALID_WEIGHT;
                    }
                }
                if (fIgnoreSymbol)
                {
                    if (sm1 == NLS_PUNCTUATION)
                    {
                        pString1++;
                        ctr1--;
                        Weight1 = CMP_INVALID_WEIGHT;
                    }
                    if (sm2 == NLS_PUNCTUATION)
                    {
                        pString2++;
                        ctr2--;
                        Weight2 = CMP_INVALID_WEIGHT;
                    }
                }
                if ((Weight1 == CMP_INVALID_WEIGHT) || (Weight2 == CMP_INVALID_WEIGHT))
                {
                    continue;
                }

                /*
                *  Switch on the script member of string 1 and take care
                *  of any special cases.
                */
                switch (sm1)
                {
                case ( NONSPACE_MARK ) :
                    {
                        /*
                        *  Nonspace only - look at diacritic weight only.
                        */
                        if ( !fIgnoreDiacritic )
                        {
                            if ( (WhichDiacritic == 0) ||
                                (State & STATE_REVERSE_DW) )
                            {
                                WhichDiacritic = 3;

                                /*
                                *  Remove state from state machine.
                                */
                                REMOVE_STATE( STATE_DW );
                            }
                        }

                        /*
                        *  Adjust pointer and counter and set flags.
                        */
                        pString1++;
                        ctr1--;
                        Weight1 = CMP_INVALID_WEIGHT;

                        break;
                    }
                case ( SYMBOL_1 ) :
                case ( SYMBOL_2 ) :
                case ( SYMBOL_3 ) :
                case ( SYMBOL_4 ) :
                case ( SYMBOL_5 ) :
                    {
                        /*
                        *  If the ignore symbol flag is set, then skip over
                        *  the symbol.
                        */
                        if (fIgnoreSymbol)
                        {
                            pString1++;
                            ctr1--;
                            Weight1 = CMP_INVALID_WEIGHT;
                        }

                        break;
                    }
                case ( NLS_PUNCTUATION ) :
                    {
                        /*
                        *  If the ignore punctuation flag is set, then skip
                        *  over the punctuation char.
                        */
                        if (fIgnorePunct)
                        {
                            pString1++;
                            ctr1--;
                            Weight1 = CMP_INVALID_WEIGHT;
                        }
                        else if (!fStringSort)
                        {
                            /*
                            *  Use WORD sort method.
                            */
                            if (sm2 != NLS_PUNCTUATION)
                            {
                                /*
                                *  The character in the second string is
                                *  NOT punctuation.
                                */
                                if (WhichPunct2)
                                {
                                    /*
                                    *  Set WP 2 to show that string 2 is
                                    *  smaller, since a punctuation char had
                                    *  already been found at an earlier
                                    *  position in string 2.
                                    *
                                    *  Set the Ignore Punctuation flag so we
                                    *  just skip over any other punctuation
                                    *  chars in the string.
                                    */
                                    WhichPunct2 = 3;
                                    fIgnorePunct = TRUE;
                                }
                                else
                                {
                                    /*
                                    *  Set WP 1 to show that string 2 is
                                    *  smaller, and that string 1 has had
                                    *  a punctuation char - since no
                                    *  punctuation chars have been found
                                    *  in string 2.
                                    */
                                    WhichPunct1 = 3;
                                }

                                /*
                                *  Advance pointer 1 and decrement counter 1.
                                */
                                pString1++;
                                ctr1--;
                                Weight1 = CMP_INVALID_WEIGHT;
                            }

                            /*
                            *  Do NOT want to advance the pointer in string 1
                            *  if string 2 is also a punctuation char.  This
                            *  will be done later.
                            */
                        }

                        break;
                    }
                case ( EXPANSION ) :
                    {
                        /*
                        *  Save pointer in pString1 so that it can be
                        *  restored.
                        */
                        pSave1 = pString1;
                        pString1 = pTmpBuf1;

                        /*
                        *  Add one to counter so that subtraction doesn't end
                        *  comparison prematurely.
                        */
                        ctr1++;

                        /*
                        *  Expand character into temporary buffer.
                        */
                        pTmpBuf1[0] = GET_EXPANSION_1( &Weight1 );
                        pTmpBuf1[1] = GET_EXPANSION_2( &Weight1 );

                        /*
                        *  Set cExpChar1 to the number of expansion characters
                        *  stored.
                        */
                        cExpChar1 = MAX_TBL_EXPANSION;

                        Weight1 = CMP_INVALID_WEIGHT;
                        break;
                    }
                case ( FAREAST_SPECIAL ) :
                    {
                        /*
                        *  Get the weight for the far east special case
                        *  and store it in Weight1.
                        */
                        GET_FAREAST_WEIGHT( Weight1, uw1, Mask, lpString1, pString1, ExtraWt1 );

                        if (sm2 != FAREAST_SPECIAL)
                        {
                            /*
                            *  The character in the second string is
                            *  NOT a fareast special char.
                            *
                            *  Set each of weights 4, 5, 6, and 7 to show
                            *  that string 2 is smaller (if not already set).
                            */
                            if ( (GET_WT_FOUR( &WhichExtra ) == 0) && (GET_WT_FOUR( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_FOUR( &WhichExtra ) = 3;
                            }

                            if ( (GET_WT_FIVE( &WhichExtra ) == 0) && (GET_WT_FIVE( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_FIVE( &WhichExtra ) = 3;
                            }

                            if ( (GET_WT_SIX( &WhichExtra ) == 0) && (GET_WT_SIX( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_SIX( &WhichExtra ) = 3;
                            }

                            if ( (GET_WT_SEVEN( &WhichExtra ) == 0) && (GET_WT_SEVEN( &ExtraWt1 ) != 0) )
                            {
                                GET_WT_SEVEN( &WhichExtra ) = 3;
                            }
                        }

                        break;
                    }
                case ( RESERVED_2 ) :
                case ( RESERVED_3 ) :
                case ( UNSORTABLE ) :
                    {
                        /*
                        *  Fill out the case statement so the compiler
                        *  will use a jump table.
                        */
                        break;
                    }
                }

                /*
                *  Switch on the script member of string 2 and take care
                *  of any special cases.
                */
                switch (sm2)
                {
                case ( NONSPACE_MARK ) :
                    {
                        /*
                        *  Nonspace only - look at diacritic weight only.
                        */
                        if ( !fIgnoreDiacritic )
                        {
                            if ( (WhichDiacritic == 0) || (State & STATE_REVERSE_DW) )

                            {
                                WhichDiacritic = 1;

                                /*
                                *  Remove state from state machine.
                                */
                                REMOVE_STATE( STATE_DW );
                            }
                        }

                        /*
                        *  Adjust pointer and counter and set flags.
                        */
                        pString2++;
                        ctr2--;
                        Weight2 = CMP_INVALID_WEIGHT;

                        break;
                    }
                case ( SYMBOL_1 ) :
                case ( SYMBOL_2 ) :
                case ( SYMBOL_3 ) :
                case ( SYMBOL_4 ) :
                case ( SYMBOL_5 ) :
                    {
                        /*
                        *  If the ignore symbol flag is set, then skip over
                        *  the symbol.
                        */
                        if (fIgnoreSymbol)
                        {
                            pString2++;
                            ctr2--;
                            Weight2 = CMP_INVALID_WEIGHT;
                        }

                        break;
                    }
                case ( NLS_PUNCTUATION ) :
                    {
                        /*
                        *  If the ignore punctuation flag is set, then
                        *  skip over the punctuation char.
                        */
                        if (fIgnorePunct)
                        {
                            /*
                            *  Advance pointer 2 and decrement counter 2.
                            */
                            pString2++;
                            ctr2--;
                            Weight2 = CMP_INVALID_WEIGHT;
                        }
                        else if (!fStringSort)
                        {
                            /*
                            *  Use WORD sort method.
                            */
                            if (sm1 != NLS_PUNCTUATION)
                            {
                                /*
                                *  The character in the first string is
                                *  NOT punctuation.
                                */
                                if (WhichPunct1)
                                {
                                    /*
                                    *  Set WP 1 to show that string 1 is
                                    *  smaller, since a punctuation char had
                                    *  already been found at an earlier
                                    *  position in string 1.
                                    *
                                    *  Set the Ignore Punctuation flag so we
                                    *  just skip over any other punctuation
                                    *  chars in the string.
                                    */
                                    WhichPunct1 = 1;
                                    fIgnorePunct = TRUE;
                                }
                                else
                                {
                                    /*
                                    *  Set WP 2 to show that string 1 is
                                    *  smaller, and that string 2 has had
                                    *  a punctuation char - since no
                                    *  punctuation chars have been found
                                    *  in string 1.
                                    */
                                    WhichPunct2 = 1;
                                }

                                /*
                                *  Pointer 2 and counter 2 will be updated
                                *  after if-else statement.
                                */
                            }
                            else
                            {
                                /*
                                *  Both code points are punctuation chars.
                                *
                                *  See if either of the strings has encountered
                                *  punctuation chars previous to this.
                                */
                                if (WhichPunct1)
                                {
                                    /*
                                    *  String 1 has had a punctuation char, so
                                    *  it should be the smaller string (since
                                    *  both have punctuation chars).
                                    */
                                    WhichPunct1 = 1;
                                }
                                else if (WhichPunct2)
                                {
                                    /*
                                    *  String 2 has had a punctuation char, so
                                    *  it should be the smaller string (since
                                    *  both have punctuation chars).
                                    */
                                    WhichPunct2 = 3;
                                }
                                else
                                {
                                    /*
                                    *  Position is the same, so compare the
                                    *  special weights.   Set WhichPunct1 to
                                    *  the smaller special weight.
                                    */
                                    WhichPunct1 = ( ( (GET_ALPHA_NUMERIC( &Weight1 ) < GET_ALPHA_NUMERIC( &Weight2 )) ) ? 1 : 3 );
                                }

                                /*
                                *  Set the Ignore Punctuation flag.
                                */
                                fIgnorePunct = TRUE;

                                /*
                                *  Advance pointer 1 and decrement counter 1.
                                *  Pointer 2 and counter 2 will be updated
                                *  after if-else statement.
                                */
                                pString1++;
                                ctr1--;
                                Weight1 = CMP_INVALID_WEIGHT;
                            }

                            /*
                            *  Advance pointer 2 and decrement counter 2.
                            */
                            pString2++;
                            ctr2--;
                            Weight2 = CMP_INVALID_WEIGHT;
                        }

                        break;
                    }
                case ( EXPANSION ) :
                    {
                        /*
                        *  Save pointer in pString1 so that it can be restored.
                        */
                        pSave2 = pString2;
                        pString2 = pTmpBuf2;

                        /*
                        *  Add one to counter so that subtraction doesn't end
                        *  comparison prematurely.
                        */
                        ctr2++;

                        /*
                        *  Expand character into temporary buffer.
                        */
                        pTmpBuf2[0] = GET_EXPANSION_1( &Weight2 );
                        pTmpBuf2[1] = GET_EXPANSION_2( &Weight2 );

                        /*
                        *  Set cExpChar2 to the number of expansion characters
                        *  stored.
                        */
                        cExpChar2 = MAX_TBL_EXPANSION;

                        Weight2 = CMP_INVALID_WEIGHT;
                        break;
                    }
                case ( FAREAST_SPECIAL ) :
                    {
                        /*
                        *  Get the weight for the far east special case
                        *  and store it in Weight2.
                        */
                        GET_FAREAST_WEIGHT( Weight2, uw2, Mask, lpString2, pString2, ExtraWt2 );

                        if (sm1 != FAREAST_SPECIAL)
                        {
                            /*
                            *  The character in the first string is
                            *  NOT a fareast special char.
                            *
                            *  Set each of weights 4, 5, 6, and 7 to show
                            *  that string 1 is smaller (if not already set).
                            */
                            if ( (GET_WT_FOUR( &WhichExtra ) == 0) && (GET_WT_FOUR( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_FOUR( &WhichExtra ) = 1;
                            }

                            if ( (GET_WT_FIVE( &WhichExtra ) == 0) && (GET_WT_FIVE( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_FIVE( &WhichExtra ) = 1;
                            }

                            if ( (GET_WT_SIX( &WhichExtra ) == 0) && (GET_WT_SIX( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_SIX( &WhichExtra ) = 1;
                            }

                            if ( (GET_WT_SEVEN( &WhichExtra ) == 0) && (GET_WT_SEVEN( &ExtraWt2 ) != 0) )
                            {
                                GET_WT_SEVEN( &WhichExtra ) = 1;
                            }
                        }
                        else
                        {
                            /*
                            *  Characters in both strings are fareast
                            *  special chars.
                            *
                            *  Set each of weights 4, 5, 6, and 7
                            *  appropriately (if not already set).
                            */
                            if ( (GET_WT_FOUR( &WhichExtra ) == 0) && ( GET_WT_FOUR( &ExtraWt1 ) != GET_WT_FOUR( &ExtraWt2 ) ) )
                            {
                                GET_WT_FOUR( &WhichExtra ) = static_cast<BYTE>(( GET_WT_FOUR( &ExtraWt1 ) < GET_WT_FOUR( &ExtraWt2 ) ) ? 1 : 3);
                            }

                            if ( (GET_WT_FIVE( &WhichExtra ) == 0) && ( GET_WT_FIVE( &ExtraWt1 ) != GET_WT_FIVE( &ExtraWt2 ) ) )
                            {
                                GET_WT_FIVE( &WhichExtra ) = static_cast<BYTE>(( GET_WT_FIVE( &ExtraWt1 ) < GET_WT_FIVE( &ExtraWt2 ) ) ? 1 : 3);
                            }

                            if ( (GET_WT_SIX( &WhichExtra ) == 0) && ( GET_WT_SIX( &ExtraWt1 ) != GET_WT_SIX( &ExtraWt2 ) ) )
                            {
                                GET_WT_SIX( &WhichExtra ) = static_cast<BYTE>(( GET_WT_SIX( &ExtraWt1 ) < GET_WT_SIX( &ExtraWt2 ) ) ? 1 : 3);
                            }

                            if ( (GET_WT_SEVEN( &WhichExtra ) == 0) && ( GET_WT_SEVEN( &ExtraWt1 ) != GET_WT_SEVEN( &ExtraWt2 ) ) )
                            {
                                GET_WT_SEVEN( &WhichExtra ) = static_cast<BYTE>(( GET_WT_SEVEN( &ExtraWt1 ) < GET_WT_SEVEN( &ExtraWt2 ) ) ? 1 : 3);
                            }
                        }

                        break;
                    }
                case ( RESERVED_2 ) :
                case ( RESERVED_3 ) :
                case ( UNSORTABLE ) :
                    {
                        /*
                        *  Fill out the case statement so the compiler
                        *  will use a jump table.
                        */
                        break;
                    }
                }

                /*
                *  See if the comparison should start again.
                */
                if ((Weight1 == CMP_INVALID_WEIGHT) || (Weight2 == CMP_INVALID_WEIGHT))
                {
                    continue;
                }

                /*
                *  We're not supposed to drop down into the state table if
                *  the unicode weights are different, so stop comparison
                *  and return result of unicode weight comparison.
                */
                if (uw1 != uw2)
                {
                    return ( (uw1 < uw2) ? 1 : 3 );
                }
            }

            /*
            *  For each state in the state table, do the appropriate
            *  comparisons.
            */
            if (State & (STATE_DW | STATE_REVERSE_DW))
            {
                /*
                *  Get the diacritic weights.
                */
                dw1 = GET_DIACRITIC( &Weight1 );
                dw2 = GET_DIACRITIC( &Weight2 );

                if (dw1 != dw2)
                {
                    /*
                    *  Look ahead to see if diacritic follows a
                    *  minimum diacritic weight.  If so, get the
                    *  diacritic weight of the nonspace mark.
                    */
                    while (!AT_STRING_END( ctr1 - 1, pString1 + 1, cchCount1 ))
                    {
                        GetDwordWeight( *(pString1 + 1), &Wt );
                        if (GET_SCRIPT_MEMBER( &Wt ) == NONSPACE_MARK)
                        {
                            dw1 = static_cast<BYTE>(dw1 + GET_DIACRITIC( &Wt ));
                            pString1++;
                            ctr1--;
                        }
                        else
                        {
                            break;
                        }
                    }

                    while (!AT_STRING_END( ctr2 - 1, pString2 + 1, cchCount2 ))
                    {
                        GetDwordWeight( *(pString2 + 1), &Wt );
                        if (GET_SCRIPT_MEMBER( &Wt ) == NONSPACE_MARK)
                        {
                            dw2 = static_cast<BYTE>(dw2 + GET_DIACRITIC( &Wt ));
                            pString2++;
                            ctr2--;
                        }
                        else
                        {
                            break;
                        }
                    }

                    /*
                    *  Save which string has the smaller diacritic
                    *  weight if the diacritic weights are still
                    *  different.
                    */
                    if (dw1 != dw2)
                    {
                        WhichDiacritic = (dw1 < dw2) ? 1 : 3;

                        /*
                        *  Remove state from state machine.
                        */
                        REMOVE_STATE( STATE_DW );
                    }
                }
            }
            if (State & STATE_CW)
            {
                /*
                *  Get the case weights.
                */
                if (GET_CASE( &Weight1 ) != GET_CASE( &Weight2 ))
                {
                    /*
                    *  Save which string has the smaller case weight.
                    */
                    WhichCase = (GET_CASE( &Weight1 ) < GET_CASE( &Weight2 )) ? 1 : 3;

                    /*
                    *  Remove state from state machine.
                    */
                    REMOVE_STATE( STATE_CW );
                }
            }
        }

        /*
        *  Fixup the pointers and counters.
        */
        POINTER_FIXUP();
        ctr1--;
        ctr2--;

        /*
        *  Reset the weights to be invalid.
        */
        Weight1 = CMP_INVALID_WEIGHT;
        Weight2 = CMP_INVALID_WEIGHT;
    }

    /*
    *  If the end of BOTH strings has been reached, then the unicode
    *  weights match exactly.  Check the diacritic, case and special
    *  weights.  If all are zero, then return success.  Otherwise,
    *  return the result of the weight difference.
    *
    *  NOTE:  The following checks MUST REMAIN IN THIS ORDER:
    *			Diacritic, Case, Punctuation.
    */
    if (AT_STRING_END( ctr1, pString1, cchCount1 ))
    {
        if (AT_STRING_END( ctr2, pString2, cchCount2 ))
        {
            if (WhichDiacritic)
            {
                return ( WhichDiacritic );
            }
            if (WhichCase)
            {
                return ( WhichCase );
            }
            if (WhichExtra)
            {
                if (!fIgnoreDiacritic)
                {
                    if (GET_WT_FOUR( &WhichExtra ))
                    {
                        return ( GET_WT_FOUR( &WhichExtra ) );
                    }
                    if (GET_WT_FIVE( &WhichExtra ))
                    {
                        return ( GET_WT_FIVE( &WhichExtra ) );
                    }
                }
                if (GET_WT_SIX( &WhichExtra ))
                {
                    return ( GET_WT_SIX( &WhichExtra ) );
                }
                if (GET_WT_SEVEN( &WhichExtra ))
                {
                    return ( GET_WT_SEVEN( &WhichExtra ) );
                }
            }
            if (WhichPunct1)
            {
                return ( WhichPunct1 );
            }
            if (WhichPunct2)
            {
                return ( WhichPunct2 );
            }

            return ( 2 );
        }
        else
        {
            /*
            *  String 2 is longer.
            */
            pString1 = pString2;
            ctr1 = ctr2;
            cchCount1 = cchCount2;
            fEnd1 = 1;
        }
    }
    else
    {
        fEnd1 = 3;
    }

    /*
    *  Scan to the end of the longer string.
    */
    return SCAN_LONGER_STRING( ctr1, pString1, cchCount1, fEnd1, &Weight1, fIgnoreDiacritic, wExcept, fIgnoreSymbol,
                               WhichCase, WhichExtra, WhichPunct1, WhichPunct2, WhichDiacritic );
}

#define MakeWordFirstSecond(b1, b2) ((WORD)(b1) << 8 | (WORD)(uchar)(b2))

// TODO: see if LOBYTE works
#define LowByte(w)			static_cast<unsigned char>((w) & 0x00ff)

/*------------------------------------------------------------------------
	MsoMultiByteToWideChar

	This function provides our own layer to WideCharToMultiByte() call.  We
	provide our own lookup table if it's not available from NLS subsystem.
	Currently works for 20 SBCS Win and 5 Mac code pages only.
-------------------------------------------------------------- SHAMIKB -*/
MSOAPI_(int) MsoMultiByteToWideChar(UINT CodePage, DWORD  dwFlags, LPCSTR lpMultiByteStr, INT_PTR cchMultiByte, _Out_opt_cap_(cchWideChar)
						            LPWSTR lpWideCharStr, int cchWideChar)
{
	int cch, cchT;
	const WCHAR *pwchConv;
	int iRet = -1;

	if (lpWideCharStr)
    {
		MsoDebugFill(lpWideCharStr, cchWideChar * sizeof(WCHAR), msomfBuffer);
    }

	// added to check that length of MultiByte String can be 0 with
	// no harmful side effects thereafter (i.e., no conversion takes place)
	MsoAssertTag((lpMultiByteStr != NULL) || (cchMultiByte==0), ASSERTTAG('eqyq'));
	
    // Win32 docs state that these must be different buffers
	// note: can be same if both point to null values
	MsoAssertTag(((char*)lpWideCharStr != (char*)lpMultiByteStr) || (cchMultiByte==0), ASSERTTAG('eqyr'));

    if (CodePage == CP_ACP || CodePage == CP_OEMCP || IsValidCodePage(CodePage))
    {
        iRet = MultiByteToWideChar(CodePage, dwFlags, lpMultiByteStr, cchMultiByte, lpWideCharStr, cchWideChar);
    }

    if (iRet > 0)
    {
        return iRet;
    }

    if (!iRet)
    {
        if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
        {
            return cchWideChar;
        }

        return 0;
    }

	if (cchMultiByte < 0)
    {
		cch = MsoCchSzLen(lpMultiByteStr) + 1;  // include null terminator
    }
	else
    {
		cch = cchMultiByte;
    }

	// Special handling for UTF8
	if (CodePage == CP_UTF8)
    {
		return UTF8ToUnicode(lpMultiByteStr, &cch, lpWideCharStr, cchWideChar);
    }

	if (cchWideChar == 0)
    {
		return (cch);
    }

	cchT = cch = min(cch, cchWideChar);

    switch (CodePage)
    {
    case CP_EASTEUROPE:
        {
            pwchConv = mpwchConv1250;
            break;
        }
    case CP_RUSSIAN:
        {
            pwchConv = mpwchConv1251;
            break;
        }
    case CP_WESTEUROPE:
        {
            pwchConv = mpwchConv1252;
            break;
        }
    case CP_GREEK:
        {
            pwchConv = mpwchConv1253;
            break;
        }
    case CP_TURKISH:
        {
            pwchConv = mpwchConv1254;
            break;
        }
    case CP_HEBREW:        
        {
            pwchConv = mpwchConv1255;
            break;
        }
    case CP_ARABIC:
        {
            pwchConv = mpwchConv1256;
            break;
        }
    case CP_BALTIC:
        {
            pwchConv = mpwchConv1257;
            break;
        }
    case CP_VIETNAMESE:
        {
            pwchConv = mpwchConv1258;
            break;
        }
    case CP_RUSSIANKOI8R:
        {
            pwchConv = mpwchConv20866;
            break;
        }
    case CP_ASCII:
    case CP_ISOLATIN1:
        {
            pwchConv = mpwchConv28591;
            break;
        }
    case CP_ISOEASTEUROPE:
        {
            pwchConv = mpwchConv28592;
            break;
        }
    case CP_ISOTURKISH:
        {
            pwchConv = mpwchConv28593;
            break;
        }
    case CP_ISOBALTIC:
        {
            pwchConv = mpwchConv28594;
            break;
        }
    case CP_ISORUSSIAN:
        {
            pwchConv = mpwchConv28595;
            break;
        }
    case CP_ISOARABIC:
        {
            pwchConv = mpwchConv28596;
            break;
        }
    case CP_ISOGREEK:
        {
            pwchConv = mpwchConv28597;
            break;
        }
    case CP_ISOHEBREW:
        {
            pwchConv = mpwchConv28598;
            break;
        }
    case CP_ISOTURKISH2:
        {
            pwchConv = mpwchConv28599;
            break;
        }
    case CP_ISOLATIN9:
        {
            pwchConv = mpwchConv28605;
            break;
        }
    #ifdef MAC  // REVIEW nobuyah: bad assumption?
    case CP_ACP:
    case CP_OEMCP:
#endif // MAC
    case CP_MACCP:
    case CP_MAC_ROMAN:
        {
            pwchConv = mpwchConv10000;
            break;
        }
    case CP_MAC_GREEK:
        {
            pwchConv = mpwchConv10006;
            break;
        }
    case CP_MAC_CYRILLIC:
        {
            pwchConv = mpwchConv10007;
            break;
        }
    case CP_MAC_LATIN2:
        {
            pwchConv = mpwchConv10029;
            break;
        }
    case CP_MAC_TURKISH:
        {
            pwchConv = mpwchConv10081;
            break;
        }
    case CP_SYMBOL:
        {
            if (lpWideCharStr != NULL)
            {
                while (cchT--)
                {
                    *(lpWideCharStr++) = static_cast<WCHAR>(MakeWordFirstSecond(*(uchar *)lpMultiByteStr >= 0x20 ? 0xF0 : 0, *(lpMultiByteStr)));
                    ++lpMultiByteStr;
                }
            }
            return (cch);
        }
        //SOUTHASIA
    case CP_THAI:
        {
            pwchConv = mpwchConv874;
            break;
        }
        // SOUTHASIA
    default:
        {
            // codepage not supported but if all ASCII then convert anyway.
            pwchConv = NULL;
            break;
        }
    }

    if (lpWideCharStr != NULL)
    {
        while (cchT--)
        {
            uchar ch;

            ch = *(lpMultiByteStr++);
            if (ch >= 0x80)
            {
                if (pwchConv)
                {
                    *(lpWideCharStr++) = pwchConv[ch - 0x80];
                }
                else
                {
                    return (0);
                }
            }
            else
                *(lpWideCharStr++) = static_cast<WCHAR>(MakeWordFirstSecond(0, ch));
        }
    }

    return (cch);
}

/*----------------------------------------------------------------------------
	UTF8ToUnicode

	Maps a UTF-8 character string to its wide character string counterpart.
	Created 02-06-96 by JulieB. Surrogate pairs support added by WalterPu.
	See UnicodeToUTF8 comment for info on Unicode surrogate pairs.

	UTF8 (0zzzzzzz)                   -> Unicode (00000000 0zzzzzzz)
	UTF8 (110xxxyy 10zzzzzz)          -> Unicode (00000xxx yyzzzzzz)
	UTF8 (1110wwww 10xxxxyy 10zzzzzz) -> Unicode (wwwwxxxx yyzzzzzz)

	UTF8 (11110bbb 10bbvvvv 10wwxxyy 10zzzzzz) ->
		Unicode pair (110110aa aavvvvww) (110111xx yyzzzzzz)
		where aaaa == bbbbb - 1
----------------------------------------------------------------- WALTERPU -*/
MSOAPI_(int) UTF8ToUnicode(LPCSTR lpSrcStr, int *pcchSrc, _Out_cap_(cchDest) LPWSTR lpDestStr, int cchDest)
{
	LPCSTR pb = lpSrcStr, pbMax;
	LPWSTR pw = lpDestStr;
	LPCWSTR pwMax;

	MsoAssertTag(pcchSrc != NULL, ASSERTTAG('eqys'));
	MsoDebugFill(lpDestStr, cchDest*sizeof(WCHAR), msomfBuffer);

    pbMax = lpSrcStr + *pcchSrc;
    pwMax = lpDestStr + cchDest;
    while (pb < pbMax && (pw < pwMax || cchDest == 0))
    {
        #if X86
        if (((UINT_PTR)pb & 0x3) == 0)
        {
            while (pb + 8 <= pbMax && (pw + 8 <= pwMax || cchDest == 0))
            {
                MsoAssertTag(((UINT_PTR)pb & 0x3) == 0, ASSERTTAG('ndtx'));

                DWORD *pdw = reinterpret_cast<DWORD *>(pw);
                const ULARGE_INTEGER *pul = reinterpret_cast<const ULARGE_INTEGER *>(pb);
                const ULARGE_INTEGER ulMask = { 0x80808080, 0x80808080 };
                if (ulMask.QuadPart & pul->QuadPart)
                {
                    break;
                }
                if (cchDest)
                {
                    pdw[0] = HIBYTE(LOWORD(pul->LowPart))  << 16 | LOBYTE(LOWORD(pul->LowPart));
                    pdw[1] = HIBYTE(HIWORD(pul->LowPart))  << 16 | LOBYTE(HIWORD(pul->LowPart));
                    pdw[2] = HIBYTE(LOWORD(pul->HighPart)) << 16 | LOBYTE(LOWORD(pul->HighPart));
                    pdw[3] = HIBYTE(HIWORD(pul->HighPart)) << 16 | LOBYTE(HIWORD(pul->HighPart));
                }

                pw += 8;
                pb += 8;
            }

            if (!(pb < pbMax && (pw < pwMax || cchDest == 0)))
            {
                break;
            }
        }
        #endif // X86

        if (BIT7(*pb) == 0)
        {
            // Found ASCII.
            if (cchDest)
            {
                *pw = static_cast<WCHAR>(*pb);
            }

            pw++;
            pb++;
        }
        else if (HIGH_3_MASK(*pb) == 0xc0)
        {
            // Found 2 byte sequence.
            if (pb + 1 >= pbMax)
            {
                break;
            }

            if (cchDest)
            {
                *pw = static_cast<WCHAR>((LOWER_5_BIT(*pb) << 6) | LOWER_6_BIT(*(pb + 1)));
            }

            pw++;
            pb += 2;
        }
        else if (HIGH_4_MASK(*pb) == 0xe0)
        {
            // Found 3 byte sequence.
            if (pb + 2 >= pbMax)
            {
                break;
            }

            if (cchDest)
            {
                *pw = static_cast<WCHAR>((LOWER_4_BIT(*pb) << 12) | (LOWER_6_BIT(*(pb + 1)) << 6) | LOWER_6_BIT(*(pb + 2)));
            }

            pw++;
            pb += 3;
        }
        else if (HIGH_5_MASK(*pb) == 0xf0)
        {
            // Found 4 byte sequence forming surrogate pair.
            if (pb + 3 >= pbMax)
            {
                break;
            }

            if (cchDest)
            {
                if (pw + 1 >= pwMax)
                {
                    return 0;
                }

                int wPlane = (((*pb & 7) << 2) | ((*(pb + 1) & 0x30) >> 4)) - 1;
                
                *pw = static_cast<WCHAR>(UCS2_1ST_OF_2 | (wPlane << 6) | (LOWER_4_BIT(*(pb + 1)) << 2) | ((*(pb + 2) & 0x30) >> 4));
                *(pw + 1) = static_cast<WCHAR>(UCS2_2ND_OF_2 | (LOWER_4_BIT(*(pb + 2)) << 6) | LOWER_6_BIT(*(pb + 3)));
            }

            pw += 2;
            pb += 4;
        }
        else
        {
            pb++;
        }
    }

    // Make sure the destination buffer was large enough.
    if (cchDest && pw >= pwMax && pb < pbMax)
    {
        return 0;
    }

    *pcchSrc = static_cast<int>(pb - lpSrcStr);
    return (pw - lpDestStr);
}

/*------------------------------------------------------------------------
	MsoWideCharToMultiByte

	This function provides our own layer to WideCharToMultiByte() call.  We
	provide our own lookup table if it's not available from NLS subsystem.
	Currently works for 20 SBCS Win and 5 Mac code pages only.
-------------------------------------------------------------- SHAMIKB -*/
MSOAPI_(int) MsoWideCharToMultiByte(UINT CodePage, DWORD dwFlags, LPCWSTR lpWideChar, int cchWideChar,
						            _Out_cap_(cchMultiByte) LPSTR lpMultiByte, INT_PTR cchMultiByte, LPCSTR lpDefaultChar,
						            LPBOOL lpUsedDefaultChar)
{
	int cch, cchT;
	const WCHAR *pwchConv = NULL;
	static UINT CodePageLast = CP_ACP;
	static BOOL fValidCodePageLast = TRUE;
	BOOL fValidCodePage = FALSE;

	if (lpMultiByte)
    {
		MsoDebugFill(lpMultiByte, cchMultiByte, msomfBuffer);
    }

	// PERF: calling IsValidCodePage can be about a 2 billion cycle operation.
	//   let's optimize that out in the case where someone calls MsoWideCharToMultiByte
	//   multiple times with the same code page.
	if (CodePage == CP_ACP)
    {
		fValidCodePage = TRUE;
    }
    else
    {
        if (CodePage != CodePageLast)
        {
            fValidCodePageLast = IsValidCodePage(CodePage);
            CodePageLast = CodePage;
        }

        fValidCodePage = fValidCodePageLast;
    }

	MsoAssertTag(lpWideChar != NULL, ASSERTTAG('eqyt'));

	// Win32 docs state that these must be different buffers
	MsoAssertTag(((char *)lpWideChar != (char*)lpMultiByte), ASSERTTAG('eqyu'));

    if (CodePage == CP_ACP || CodePage == CP_OEMCP || fValidCodePage)
    {
        LPCWSTR lpwch;
        LPCWSTR lpwchLim;
        int fSysCall = fTrue;
        int cwch;
        int cchOut = 0;
        int iRet = -1;

        if (cchWideChar < 0)
        {
            cwch = MsoCchWzLen(lpWideChar) + 1;  // include null terminator
        }
        else
        {
            cwch = cchWideChar;
        }

        if (cchMultiByte>0 && CodePage!=CP_UTF8)
        {
            for (lpwch = lpWideChar, lpwchLim = lpwch + cwch; lpwch < lpwchLim; lpwch++)
            {
                if (*lpwch >= 0x80 && *lpwch <= 0x9f)
                {
                    fSysCall = fFalse;
                    break;
                }
            }
        }

        if (fSysCall)
        {
            iRet = WideCharToMultiByte(CodePage, dwFlags, lpWideChar, cchWideChar, lpMultiByte, cchMultiByte, lpDefaultChar, lpUsedDefaultChar);

            if (iRet > 0)
            {
                return iRet;
            }
            if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
            {
                return iRet;
            }
        }

		if (lpUsedDefaultChar)
        {
			*lpUsedDefaultChar = fFalse;
        }

        while (cwch--)
        {
            WCHAR wch;
            BYTE rgch[4];
            BOOL fDef;
            int cchMB;

            wch = *(lpWideChar++);
            if (wch >= 0x80 && wch <= 0x9f)
            {
                if (cchOut+1 > cchMultiByte)
                {
                    return cchOut;
                }

                if (lpUsedDefaultChar)
                {
                    *lpUsedDefaultChar = fTrue;
                }

                if (lpMultiByte)
                {
                    *(lpMultiByte++) = LowByte(wch);
                }

                cchOut++;
            }
            else
            {
                cchMB = WideCharToMultiByte(CodePage, dwFlags, &wch, 1, reinterpret_cast<CHAR *>(&rgch[0]), sizeof(rgch), lpDefaultChar, &fDef);

                if (cchOut+cchMB > cchMultiByte)
                {
                    return (cchOut);
                }

                if (cchOut + cchMB < cchOut) // integer overflow check
                {
                    return (0);
                }

                cchOut += cchMB;
                
                if (cchMultiByte)
                {
                    if (lpMultiByte)
                    {
                        memcpy(lpMultiByte, rgch, cchMB);
                        lpMultiByte += cchMB;
                    }
                }

                if (lpUsedDefaultChar)
                {
                    *lpUsedDefaultChar |= fDef;
                }
            }
        }

        return (cchOut);
    }

    if (cchWideChar < 0)
    {
        cch = MsoCchWzLen(lpWideChar) + 1;  // include null terminator
    }
    else
    {
        cch = cchWideChar;
    }

    // Special handling for UTF8
    if (CodePage == CP_UTF8)
    {
        return UnicodeToUTF8(lpWideChar, cch, lpMultiByte, cchMultiByte);
    }

    if (cchMultiByte == 0)
    {
        return (cch);
    }

    cchT = cch = min(cch, cchMultiByte);

    if (lpUsedDefaultChar)
    {
        *lpUsedDefaultChar = fFalse;
    }

    switch (CodePage)
    {
    case CP_EASTEUROPE:
        {
            pwchConv = mpwchConv1250;
            break;
        }
    case CP_RUSSIAN:
        {
            pwchConv = mpwchConv1251;
            break;
        }
    case CP_WESTEUROPE:
        {
            pwchConv = mpwchConv1252;
            break;
        }
    case CP_GREEK:
        {
            pwchConv = mpwchConv1253;
            break;
        }
    case CP_TURKISH:
        {
            pwchConv = mpwchConv1254;
            break;
        }
    case CP_HEBREW:
        {
            pwchConv = mpwchConv1255;
            break;
        }
    case CP_ARABIC:
        {
            pwchConv = mpwchConv1256;
            break;
        }
    case CP_BALTIC:
        {
            pwchConv = mpwchConv1257;
            break;
        }
    case CP_VIETNAMESE:
        {
            pwchConv = mpwchConv1258;
            break;
        }
    case CP_RUSSIANKOI8R:
        {
            pwchConv = mpwchConv20866;
            break;
        }
    case CP_ISOLATIN1:
        {
            pwchConv = mpwchConv28591;
            break;
        }
    case CP_ISOEASTEUROPE:
        {
            pwchConv = mpwchConv28592;
            break;
        }
    case CP_ISOTURKISH:
        {
            pwchConv = mpwchConv28593;
            break;
        }
    case CP_ISOBALTIC:
        {
            pwchConv = mpwchConv28594;
            break;
        }
    case CP_ISORUSSIAN:
        {
            pwchConv = mpwchConv28595;
            break;
        }
    case CP_ISOARABIC:
        {
            pwchConv = mpwchConv28596;
            break;
        }
    case CP_ISOGREEK:
        {
            pwchConv = mpwchConv28597;
            break;
        }
    case CP_ISOHEBREW:
        {
            pwchConv = mpwchConv28598;
            break;
        }
    case CP_ISOTURKISH2:
        {
            pwchConv = mpwchConv28599;
            break;
        }
    case CP_ISOLATIN9:
        {
            pwchConv = mpwchConv28605;
            break;
        }
    case CP_MACCP:
    case CP_MAC_ROMAN:
        {
            pwchConv = mpwchConv10000;
            break;
        }
    case CP_MAC_GREEK:
        {
            pwchConv = mpwchConv10006;
            break;
        }
    case CP_MAC_CYRILLIC:
        {
            pwchConv = mpwchConv10007;
            break;
        }
    case CP_MAC_LATIN2:
        {
            pwchConv = mpwchConv10029;
            break;
        }
    case CP_MAC_TURKISH:
        {
            pwchConv = mpwchConv10081;
            break;
        }
    case CP_SYMBOL:
        {
            if (lpMultiByte)
            {
                while (cchT--)
                {
                    *(lpMultiByte++) = static_cast<CHAR>(*(lpWideChar++) & 0xFF);
                }
            }
            return cch;
        }
        //SOUTHASIA
    case CP_THAI:
        {
            pwchConv = mpwchConv874;
            break;
        }
        // SOUTHASIA
    default:
        {
            // code page not supported, but if all ASCII then convert anyway.
            if (lpMultiByte)
            {
                while (cchT--)
                {
                    WCHAR wch;

                    wch = *(lpWideChar++);
                    if (wch < 0x80)
                    {
                        *(lpMultiByte++) = static_cast<char>(wch);
                    }
                    else
                    {
                        return (0);
                    }
                }
            }
            return (cch);
            break;
        }
    }

    if (lpMultiByte)
    {
        while (cchT--)
        {
            WCHAR wch;
            BYTE ch = 0;

            wch = *(lpWideChar++);

            if (wch < 0x80)
            {
                *(lpMultiByte++) = static_cast<char>(wch);
            }
            else if (wch >= 0x80 && wch <= 0x9f)
            {
                if (lpUsedDefaultChar)
                {
                    *lpUsedDefaultChar = fTrue;
                }

                *(lpMultiByte++) = LowByte(wch);
            }
            else
            {
                // REVIEW nobuyah: very inefficient; speed up later
                for (int ich = 0; ich < 128; ich++)
                {
                    if (pwchConv[ich] == wch)
                    {
                        ch = static_cast<BYTE>(ich + 128);
                        break;
                    }
                }

                if (ch == 0)
                {
                    if (lpUsedDefaultChar)
                    {
                        *lpUsedDefaultChar = fTrue;
                    }

                    ch = (lpDefaultChar ? *lpDefaultChar : '?');
                }

                *(lpMultiByte++) = ch;
            }
        }
    }

    return (cch);
}


/*----------------------------------------------------------------------------
	UnicodeToUTF8

	Simple version of UnicodeToUTF8Core that always converts the whole input.
	Use this for a single independent buffer, and Core for when converting
	a large buffer through a sequence of calls. This behaves differently than
	Core when the input ends with the first character of a Unicode surrogate
	pair. This will just treat it as a single Unicode character. However Core
	will avoid splitting the surrogate, and return 1 in the cchSrcLeft out
	parameter, meaning the last character of the input was not converted, and
	should be recycled as the first character of the next call to Core.
----------------------------------------------------------------- WALTERPU -*/
MSOAPI_(int) UnicodeToUTF8(LPCWSTR lpSrcStr, int cchSrc,
	_Out_cap_(cchDest) LPSTR lpDestStr, int cchDest)
{
	return UnicodeToUTF8Core(lpSrcStr, cchSrc, NULL, lpDestStr, cchDest);
}

/*----------------------------------------------------------------------------
	UnicodeToUTF8Core

	Maps a Unicode character string to its UTF-8 string counterpart. Created
	02-06-96 by JulieB. Unicode surrogate pairs support added by WalterPu.

	Unicode (00000000 0zzzzzzz) -> UTF8 (0zzzzzzz)
	Unicode (00000xxx yyzzzzzz) -> UTF8 (110xxxyy 10zzzzzz)
	Unicode (wwwwxxxx yyzzzzzz) -> UTF8 (1110wwww 10xxxxyy 10zzzzzz)

	Unicode pair (110110aa aavvvvww) (110111xx yyzzzzzz) ->
		UTF8 (11110bbb 10bbvvvv 10wwxxyy 10zzzzzz) where bbbbb == aaaa + 1

	Note: A Unicode surrogate pair is where two Unicode words are used to
	represent one 32 bit UCS-4 value. The UCS-4 range that may be represented
	as a surrogate pair is from 0x00010000 to 0x0010ffff. We don't convert
	between UCS-4 and Unicode surrogate pairs here, only between existing
	surrogate pairs and UTF-8. For informational purposes only:

	UCS-4 (00000000 000bbbbb xxxxxxyy zzzzzzzz) ->
		Unicode (110110aa aaxxxxxx) (110111yy zzzzzzzz) where aaaa == bbbbb - 1.
----------------------------------------------------------------- WALTERPU -*/
MSOAPIX_(int) UnicodeToUTF8Core(LPCWSTR lpSrcStr, int cchSrc, int *pcchSrcLeft, _Out_cap_(cchDest) LPSTR lpDestStr, int cchDest)
{
	LPCWSTR lpWC = lpSrcStr;
	int cchU8 = 0;           // Number of UTF8 bytes generated
	int wPlane;

	MsoDebugFill(lpDestStr, cchDest, msomfBuffer);

    while (cchSrc && (cchU8 < cchDest || cchDest == 0))
    {
        #if X86
        if (((UINT_PTR)&lpDestStr[cchU8] & 0x3) == 0)
        {
            while (cchSrc >= 4 && (cchU8 + 4 <= cchDest || cchDest == 0))
            {
				MsoAssertTag(((UINT_PTR)&lpDestStr[cchU8] & 0x3) == 0, ASSERTTAG('ndty'));

				DWORD *pdw = reinterpret_cast<DWORD *>(&lpDestStr[cchU8]);
				DWORD dw = 0;
				const ULARGE_INTEGER *pul = reinterpret_cast<const ULARGE_INTEGER *>(lpWC);
				const ULARGE_INTEGER ulMask = { 0xff80ff80, 0xff80ff80 };
				
                if (ulMask.QuadPart & pul->QuadPart)
                {
					break;
                }

				dw |= LOBYTE(HIWORD(pul->HighPart));
				dw <<= 8;
				dw |= LOBYTE(pul->HighPart);
				dw <<= 8;
				dw |= LOBYTE(HIWORD(pul->LowPart));
				dw <<= 8;
				dw |= LOBYTE(pul->LowPart);
				
                if (cchDest)
                {
					*pdw = dw;
                }

				cchU8 += 4;
				lpWC += 4;
				cchSrc -= 4;
            }

            if (!(cchSrc && (cchU8 < cchDest || cchDest == 0)))
            {
                break;
            }
        }
        #endif // X86

        if (*lpWC <= ASCII)
        {
            // Generate 1 byte sequence (ASCII) if <= 0x7F (7 bits).
            if (cchDest && cchU8 >= 0)
            {
                lpDestStr[cchU8] = static_cast<CHAR>(*lpWC);
            }

            cchU8++;
            lpWC++;
            cchSrc--;
        }
        else if (*lpWC <= UTF8_2_MAX)
        {
            // Generate 2 byte sequence if <= 0x07ff (11 bits).
            if (cchDest)
            {
                if (cchU8 + 2 <= cchDest)
                {
                    // Use upper 5 bits in first byte.
                    // Use lower 6 bits in second byte.
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_1ST_OF_2 | (*lpWC >> 6));
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_TRAIL    | LOWER_6_BIT(*lpWC));
                }
                else
                {
                    // Error - buffer too small.
                    break;
                }
            }
            else
            {
                cchU8 += 2;
            }

            lpWC++;
            cchSrc--;
        }
        else if (HIGH_6_MASK(*lpWC) != UCS2_1ST_OF_2 || (cchSrc >= 2 ? HIGH_6_MASK(*(lpWC + 1)) != UCS2_2ND_OF_2 : pcchSrcLeft == NULL))
        {
            // Generate 3 byte sequence.
            if (cchDest)
            {
                if (cchU8 + 3 <= cchDest)
                {
                    // Use upper  4 bits in first byte.
                    // Use middle 6 bits in second byte.
                    // Use lower  6 bits in third byte.
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_1ST_OF_3 | (*lpWC >> 12));
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_TRAIL    | MIDDLE_6_BIT(*lpWC));
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_TRAIL    | LOWER_6_BIT(*lpWC));
                }
                else
                {
                    // Error - buffer too small.
                    break;
                }
            }
            else
            {
                cchU8 += 3;
            }

            lpWC++;
            cchSrc--;
        }
        else
        {
            // Generate 4 byte sequence involving two source Unicode characters.
            if (cchSrc < 2)
            {
                // Split surrogate pair at end of input buffer.
                MsoAssertTag(pcchSrcLeft != NULL, ASSERTTAG('xpim'));
                break;
            }

            MsoAssertTag(HIGH_6_MASK(*lpWC) == UCS2_1ST_OF_2, ASSERTTAG('eqyw'));
            MsoAssertTag(HIGH_6_MASK(*(lpWC + 1)) == UCS2_2ND_OF_2, ASSERTTAG('eqyx'));

            if (cchDest)
            {
                if (cchU8 + 4 <= cchDest)
                {
                    wPlane = MIDDLE_4_BIT(*lpWC) + 1;

                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_1ST_OF_4 | (wPlane >> 2));
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_TRAIL | (LOWER_2_BIT(wPlane) << 4) | MIDLOW_4_BIT(*lpWC));
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_TRAIL | (LOWER_2_BIT(*lpWC) << 4) | MIDDLE_4_BIT(*(lpWC + 1)));
                    lpDestStr[cchU8++] = static_cast<CHAR>(UTF8_TRAIL | LOWER_6_BIT(*(lpWC + 1)));
                }
                else
                {
                    // Error - buffer too small.
                    break;
                }
            }
            else
            {
                cchU8 += 4;
            }

            lpWC += 2;
            cchSrc -= 2;
        }
    }

    // Make sure the destination buffer was large enough.
    if (pcchSrcLeft == NULL)
    {
        if (cchDest && cchSrc > 0)
        {
            SetLastError(ERROR_INSUFFICIENT_BUFFER);
            return 0;
        }
    }
    else
    {
        MsoAssertTag(cchSrc == 0 || cchSrc == 1, ASSERTTAG('xpin'));

        *pcchSrcLeft = cchSrc;
    }

	// Return the number of UTF-8 bytes written.
	return cchU8;
}

/*------------------------------------------------------------------------
	MsoWchToUpperInvariant

	This function performs upper case conversion for a given
	Unicode character. It is invariant with respect to lcids and
	must be kept this way because word needs an invariant upper casing.
	Returns the input Unicode character if no conversion is found.
	Works only for the 5 Win and 5 Mac code pages needed for the
	Pan European version of the apps.
	The code is ugly but I decided not to use a
	binary search to avoid paging costs as much as possible. Profiling
	shows this works faster in most cases.
-------------------------------------------------------------- SHAMIKB -*/
__inline WCHAR inl_WchToUpperLid(WCHAR wch, LID lid)
{
	const WCHAR *pTbl = &adjUpper[3-3];

    do
    {
        pTbl += 3;
    }while (wch > pTbl[1-3]);

    if (wch >= pTbl[0-3]) 
    {
        wch = static_cast<WCHAR>(wch - pTbl[2-3]);

        if (static_cast<short>(pTbl[2-3]) > 0x00FF)		// signed compare
        {
            wch = mpccToUpper[wch];
        }
        else if (pTbl[2-3] <= 1) 			// unsigned compare
        {
            wch &= ~1;
            wch = static_cast<WCHAR>(wch + pTbl[2-3]);
        }
        else if (ilLangBase(lid) == PRIMARYLANGID(lidTurkish) && wch == 0x0069-0x20)
        {
            wch = 0x0130;
        }
    }

    return wch;
}

MSOAPI_(WCHAR) MsoWchToUpperInvariant(WCHAR wch)
{
	return inl_WchToUpperLid(wch, 0);
}

/*------------------------------------------------------------------------
	_WGetCType?Wch

	These static functions are built because of the modification below in
	the API MsoGetStringTypeExW. These functions are used to be in the file
	unistgen.inc, which was generated automatilly during thge build process
	We keep them to minimize the modification because of deleting that file
------------------------------------------------------------------------*/
namespace
{
    WORD _WGetCType1Wch(WCHAR wch)
    {
        WORD wCharType;
        return GetStringTypeExW(LOCALE_USER_DEFAULT, CT_CTYPE1, &wch, 1, &wCharType) ? wCharType : (WORD)0;
    };

    WORD _WGetCType2Wch(WCHAR wch)
    {
        WORD wCharType;
        return GetStringTypeExW(LOCALE_USER_DEFAULT, CT_CTYPE2, &wch, 1, &wCharType) ? wCharType : (WORD)0;
    };

    WORD _WGetCType3Wch(WCHAR wch)
    {
        WORD wCharType;
        return GetStringTypeExW(LOCALE_USER_DEFAULT, CT_CTYPE3, &wch, 1, &wCharType) ? wCharType : (WORD)0;
    };
}

/*----------------------------------------------------------------------------
	MsoFWchSpace
	MsoFChSpace

	code stolen from word,
------------------------------------------------------------------ t-hailiu--*/
MSOAPI_(BOOL) MsoFSpaceWch(WCHAR wch)
{
	if (wch <= 0x7F)	// we're ANSI
    {
		return InBiasSetOf6(wch, 9,10,11,12,13,32);	//InSetOf6(wch-1, 31,8,9,10,11,12);
    }
	else
    {
		return _WGetCType1Wch(wch) & C1_SPACE;
    }
}

MSOAPIXX_(BOOL) MsoFSpaceCh(CHAR ch)
{
	return InBiasSetOf6(ch, 9,10,11,12,13,32);
}

/*---------------------------------------------------------------------------
   MsoFDigitWch

	Returns fTrue if and only if wch is a digit.
----------------------------------------------------------------- CMoore --*/
MSOAPI_(BOOL) MsoFDigitWch(WCHAR wch)
{
    if (wch <= 0x7F) // we're ANSI
    {
        MsoAssertTag(!(_WGetCType1Wch(wch) & C1_DIGIT) == !FBetween(wch, '0', '9'), ASSERTTAG('eqzm'));
        return (FBetween(wch, '0', '9'));
    }
	else
    {
		return _WGetCType1Wch(wch) & C1_DIGIT;
    }
}

/*--------------------------------------------------------------------------
   MsoFDigitCh

	Returns TRUE if and only if ch is a digit.
------------------------------------------------------------------ RICKP -*/
MSOAPI_(BOOL) MsoFDigitCh(int ch)
{
	WCHAR wchT = 0;
	CHAR rgch[2] = {0};

    #ifdef DEBUG
	int fCache = -1;

	if (ch <= 0x7F)  // we're ANSI
    {
		fCache = FBetween(ch, '0', '9');
    }

    if (IsDBCSLeadByte(HIBYTE(LOWORD(ch))))
    {
        rgch[0] = HIBYTE(LOWORD(ch));
        rgch[1] = LOBYTE(LOWORD(ch));

        if (!MsoMultiByteToWideChar(CP_ACP, 0, rgch, 2, &wchT, 1))
        {
            MsoAssertTag(fCache != 1, ASSERTTAG('eqzn'));
            return FALSE;
        }
    }
    else
    {
        rgch[0] = LOBYTE(LOWORD(ch));
        if (!MsoMultiByteToWideChar(CP_ACP, 0, rgch, 1, &wchT, 1))
        {
            MsoAssertTag(fCache != 1, ASSERTTAG('eqzo'));
            return FALSE;
        }
    }

    if(MsoFDigitWch(wchT))
    {
        MsoAssertTag(fCache != 0, ASSERTTAG('eqzp'));
        return TRUE;
    }
    else
    {
        MsoAssertTag(fCache != 1, ASSERTTAG('eqzq'));
        return FALSE;
    }
    #else  //!DEBUG
	if (ch <= 0x7F)  // we're ANSI
    {
		return (FBetween(ch, '0', '9'));
    }

    if (IsDBCSLeadByte(HIBYTE(LOWORD(ch))))
    {
        rgch[0] = HIBYTE(LOWORD(ch));
        rgch[1] = LOBYTE(LOWORD(ch));
        if (!MsoMultiByteToWideChar(CP_ACP, 0, rgch, 2, &wchT, 1))
        {
            return FALSE;
        }
    }
    else
    {
        rgch[0] = LOBYTE(LOWORD(ch));

        if (!MsoMultiByteToWideChar(CP_ACP, 0, rgch, 1, &wchT, 1))
        {
            return FALSE;
        }
    }

	return MsoFDigitWch(wchT);
    #endif
}

/*-----------------------------------------------------------------------------
	MsoSgnRgwchCompareInvariant

	Compares the sort order of the two unicode strings rgwch1
	(# of chars cch1) and rgwch2 (# of chars cch2).  Returns negative if
	rgwch1 < rgwch2, 0 if they sort equally, and positive if rgwch1 > rgwch2.
	The string compare is based on Unicode order.  MsoWchToUpperInvariant is
	used as an invariant ToUpper.  It must behave the same no matter which
	language we run on.

-------------------------------------------------------------------- PETERO -*/
//  The invariant sort routine.  We will want to use this in other situations within Word.
MSOAPI_(int) MsoSgnRgwchCompareInvariant(const WCHAR* rgwch1, int cch1, const WCHAR* rgwch2, int cch2, int cs)
{
    int cchCommon = cch1 > cch2 ? cch2 : cch1;
    int sgn = 0;

    MsoAssertTag(sgnEQ == 0, ASSERTTAG('eqzr'));
    if (cchCommon) 
    {
        if (!(cs & msocsCase))	// case insensitive?
        {
            do 
            {
                sgn = MsoWchToUpperInvariant(*rgwch1++) - MsoWchToUpperInvariant(*rgwch2++);
            }while (sgn == 0 && --cchCommon > 0);
        }
        else
        {
            do
            {
                sgn = (*rgwch1++) - (*rgwch2++);
            }while (sgn == 0 && --cchCommon > 0);
        }

        if (sgn == 0)			// equal so far, which string is longer?
        {
            sgn = cch1 - cch2;
        }

        #ifdef FUTURE		//  REVIEW:  PETERO:  Current uses simply rely on negative/positive/zero
        if (sgn)
        {
            sgn = (sgn > 0) ? sgnGT : sgnLT;
        }
        #endif
    }

    return sgn;
}

#undef ASCII_TOUPPER