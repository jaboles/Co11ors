/****************************************************************************
	msodstr.h

	Owner: AndrewQ
	Copyright (c) 1999 Microsoft Corporation

	MSO's dynamic-size string handling.
****************************************************************************/

#pragma once

#if !defined(MSODSTR_H)
#define MSODSTR_H

#ifndef CONST_METHOD
// to declare const methods in C++ */
#if defined(__cplusplus) && !defined(CINTERFACE)
	#define CONST_METHOD const
#else
	#define CONST_METHOD
#endif
#endif

#define MSOSTR_NOT_FOUND (-1) // IchFind...() result: no match

/****************************************************************************
	IMsoString

	All "cch" references are to the count of characters, which does not
	include the trailing '\0'. The exception (of course) is FGetData(),
	which is for emulating system APIs, which DO include the '\0'.
*****************************************************************************/

#undef  INTERFACE
#define INTERFACE IMsoString
DECLARE_INTERFACE_(IMsoString, ISimpleUnknown)
{
	// ----- ISimpleUnknown methods

	MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

	// ----- IMsoString methods

	// FDebugMessage method
	MSODEBUGMETHOD

	// Get rid of an IMsoString.
	MSOMETHOD_(VOID, Free) (THIS) PURE;

	// Operations
	MSOMETHOD_(VOID, Empty) (THIS) PURE; // same as this = L"";

	// Attributes
	MSOMETHOD_(LPCWSTR, WzGetValue) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FGetData) (THIS_ _Out_opt_z_cap_(*pcchBuf) LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD_(WCHAR, WchGetAt) (THIS_ int ich) CONST_METHOD PURE;
	MSOMETHOD_(int, CchGetLength) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FIsEmpty) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FIsNotEmpty) (THIS) CONST_METHOD PURE;

	// Assignment
	MSOMETHOD_(BOOL, FCopyWz) (THIS_ LPCWSTR wz) PURE;
	MSOMETHOD_(BOOL, FCopyRgwch) (THIS_ LPCWSTR wz, int cch) PURE;
	MSOMETHOD_(BOOL, FCopyStr) (THIS_ const IMsoString *pstr) PURE;
	MSOMETHOD_(BOOL, FCopyWzCch) (THIS_ LPCWSTR wz, int cch) PURE;
	MSOMETHOD_(BOOL, FCopyFromResource) (THIS_ HINSTANCE hInstance, int nResourceID) PURE;
	MSOMETHOD_(BOOL, FCopyFromTmc) (THIS_ UINT tmc) PURE;
	MSOMETHOD_(HRESULT, HrCopyFromRegistry) (THIS_ int msorid) PURE;
	MSOMETHOD_(HRESULT, HrRegQueryValue) (THIS_ HKEY hKey, LPCWSTR pwzSubKey) PURE;
	MSOMETHOD_(BOOL, FSetAt) (THIS_ int ichPos, WCHAR wch) PURE;

	// Concatenation
	MSOMETHOD_(BOOL, FAppendWz) (THIS_ LPCWSTR wz) PURE;
	MSOMETHOD_(BOOL, FAppendRgwch) (THIS_ LPCWSTR wz, int cch) PURE;
	MSOMETHOD_(BOOL, FAppendStr) (THIS_ const IMsoString *pstr) PURE;
	MSOMETHOD_(BOOL, FAppendWch) (THIS_ WCHAR wch) PURE;
	MSOMETHOD_(BOOL, FAppendFromResource) (THIS_ HINSTANCE hInstance, int nResourceID) PURE;
	MSOMETHOD_(BOOL, FAppendFromResourceWz) (THIS_ HINSTANCE hInstance, int nResourceID, LPCWSTR wz) PURE;
	MSOMETHOD_(BOOL, FAppendInt) (THIS_ int w, int wBase) PURE;

	// Case conversion
	MSOMETHOD_(VOID, MakeUpper) (THIS) PURE;
	MSOMETHOD_(VOID, MakeLower) (THIS) PURE;

	// Comparisons ("cs" arguments are as per MsoSgnWzCompare(), etc. - see <msostr.h>)
	MSOMETHOD_(int, SgnCompareWz) (THIS_ LPCWSTR wz, int cs) CONST_METHOD PURE;
	MSOMETHOD_(int, SgnCompareWzSubstr) (THIS_ LPCWSTR wz, int cs, int ichStart) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FEqualWz) (THIS_ LPCWSTR wz, int cs) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FEqualStr) (THIS_ const IMsoString *pstr, int cs) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FEqualWzSubstr) (THIS_ LPCWSTR wz, int cs, int ichStart) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FEqualWzTail) (THIS_ LPCWSTR wz, int cs) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FEqualWzSubstrTail) (THIS_ LPCWSTR wz, int cs, int ichStart, int cchSubstr) CONST_METHOD PURE;

	// Find functions, return index or MSOSTR_NOT_FOUND
	MSOMETHOD_(int, IchFindWz) (THIS_ LPCWSTR wz, int cs, int ichStart) CONST_METHOD PURE;
	MSOMETHOD_(int, IchFindWzSubstr) (THIS_ LPCWSTR wz, int cs, int ichStart, int cchSubstr) CONST_METHOD PURE;
	MSOMETHOD_(int, IchFindWch) (THIS_ WCHAR wch, int cs, int ichStart) CONST_METHOD PURE;
	MSOMETHOD_(int, IchFindWchSubstr) (THIS_ WCHAR wch, int cs, int ichStart, int cchSubstr) CONST_METHOD PURE;
	MSOMETHOD_(int, IchFindWchReverse) (THIS_ WCHAR wch, int cs) CONST_METHOD PURE;
	MSOMETHOD_(int, IchFindWchSubstrReverse) (THIS_ WCHAR wch, int cs, int ichStart, int cchSubstr) CONST_METHOD PURE;

	// Insertion
	MSOMETHOD_(BOOL, FInsertWz) (THIS_ LPCWSTR wzSubStr, int ichPos, int cchReplace) PURE;
	MSOMETHOD_(BOOL, FInsertWch) (THIS_ WCHAR wch, int ichPos, int cchReplace) PURE;

	// Search-and-replace
	MSOMETHOD_(VOID, ReplaceWchWithWch) (THIS_ WCHAR wchOld, WCHAR wchNew, int ichStart) PURE;
	MSOMETHOD_(VOID, ReplaceWchWithWchSubstr) (THIS_ WCHAR wchOld, WCHAR wchNew, int ichStart, int cchSubstr) PURE;

	// Removal and truncation
	MSOMETHOD_(BOOL, FTruncAt) (THIS_ int ichPos) PURE;
	MSOMETHOD_(BOOL, FTruncLeft) (THIS_ int cch) PURE;                  // remove the first n characters
	MSOMETHOD_(BOOL, FTruncRight) (THIS_ int cch) PURE;                 // remove the last n characters
	MSOMETHOD_(BOOL, FRemoveAt) (THIS_ int ichPos, int cchRemove) PURE; // remove n characters
	MSOMETHOD_(VOID, RemoveLeadingSpaces) (THIS) PURE;                  // remove all space characters at beginning of string
	MSOMETHOD_(VOID, RemoveTrailingSpaces) (THIS) PURE;                 // remove all space characters at end of string

	// FUTURE: AndrewQ: Ability to specify different dg's for different strings.
	// FUTURE: AndrewQ: Numeric conversions (ala MsoWzDecodeInt, etc.).

	// Size conversions
	MSOMETHOD_(LPWSTR, WzLockBuffer) (THIS_ int cch) PURE;
	MSOMETHOD_(VOID, ReleaseBuffer) (THIS) PURE;

	// Decoding
	MSOMETHOD_(int, CchWzDecodeInt) (THIS_ int w, int wBase) PURE;
	MSOMETHOD_(int, CchWzDecodeUint) (THIS_ unsigned int w, int wBase) PURE;
};

// Create an IMsoString
MSOAPI_(HRESULT) MsoHrMakeString(IMsoString **ppstr);

/****************************************************************************
	IMsoString C interface
*****************************************************************************/

#ifndef __cplusplus

// ----- ISimpleUnknown methods
#define IMsoString_QueryInterface(This, riid, ppvObj) \
	(This)->lpVtbl->QueryInterface(This, riid, ppvObj)

// ----- MSODEBUGMETHOD
#define IMsoString_FDebugMessage(This, hinst, dm, wParam, lParam) \
	(This)->lpVtbl->FDebugMessage(This, hinst, dm, wParam, lParam)

// ----- IMsoString methods
#define IMsoString_Free(This) \
	(This)->lpVtbl->Free(This)

#define IMsoString_Empty(This) \
	(This)->lpVtbl->Empty(This)

#define IMsoString_WzGetValue(This) \
	(This)->lpVtbl->WzGetValue(This)

#define IMsoString_FGetData(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->FGetData(This, wzBuf, pcchBuf)

#define IMsoString_WchGetAt(This, ich) \
	(This)->lpVtbl->WchGetAt(This, ich)

#define IMsoString_CchGetLength(This) \
	(This)->lpVtbl->CchGetLength(This)

#define IMsoString_FIsEmpty(This) \
	(This)->lpVtbl->FIsEmpty(This)

#define IMsoString_FIsNotEmpty(This) \
	(This)->lpVtbl->FIsNotEmpty(This)

#define IMsoString_FCopyWz(This, wz) \
	(This)->lpVtbl->FCopyWz(This, wz)

#define IMsoString_FCopyRgwch(This, pwch, cch) \
	(This)->lpVtbl->FCopyRgwch(This, pwch, cch)

#define IMsoString_FCopyStr(This, pstr) \
	(This)->lpVtbl->FCopyStr(This, pstr)

#define IMsoString_FCopyWzCch(This, wz, cch) \
	(This)->lpVtbl->FCopyWzCch(This, wz, cch)

#define IMsoString_FCopyFromResource(This, hInstance, nResourceID) \
	(This)->lpVtbl->FCopyFromResource(This, hInstance, nResourceID)

#define IMsoString_FCopyFromTmc(This, tmc) \
	(This)->lpVtbl->FCopyFromTmc(This, tmc)

#define IMsoString_HrCopyFromRegistry(This, msorid) \
	(This)->lpVtbl->HrCopyFromRegistry(This, msorid)

#define IMsoString_HrRegQueryValue(This, hKey, pwzSubKey) \
	(This)->lpVtbl->HrRegQueryValue(This, hKey, pwzSubKey)

#define IMsoString_FSetAt(This, ichPos, wch) \
	(This)->lpVtbl->FSetAt(This, ichPos, wch)

#define IMsoString_FAppendWz(This, wz) \
	(This)->lpVtbl->FAppendWz(This, wz)

#define IMsoString_FAppendRgwch(This, pwch, cch) \
	(This)->lpVtbl->FAppendRgwch(This, pwch, cch)

#define IMsoString_FAppendStr(This, pstr) \
	(This)->lpVtbl->FAppendStr(This, pstr)

#define IMsoString_FAppendWch(This, wch) \
	(This)->lpVtbl->FAppendWch(This, wch)

#define IMsoString_FAppendFromResource(This, hInstance, nResourceID) \
	(This)->lpVtbl->FAppendFromResource(This, hInstance, nResourceID)

#define IMsoString_FAppendFromResourceWz(This, hInstance, nResourceID, wz) \
	(This)->lpVtbl->FAppendFromResourceWz(This, hInstance, nResourceID, wz)

#define IMsoString_FAppendInt(This, w, wBase) \
	(This)->lpVtbl->FAppendInt(This, w, wBase)

#define IMsoString_MakeUpper(This) \
	(This)->lpVtbl->MakeUpper(This)

#define IMsoString_MakeLower(This) \
	(This)->lpVtbl->MakeLower(This)

#define IMsoString_SgnCompareWz(This, wz, cs) \
	(This)->lpVtbl->SgnCompareWz(This, wz, cs)

#define IMsoString_SgnCompareWzSubstr(This, wz, cs, ichStart) \
	(This)->lpVtbl->SgnCompareWzSubstr(This, wz, cs, ichStart)

#define IMsoString_FEqualWz(This, wz, cs) \
	(This)->lpVtbl->FEqualWz(This, wz, cs)

#define IMsoString_FEqualStr(This, pstr, cs) \
	(This)->lpVtbl->FEqualStr(This, pstr, cs)

#define IMsoString_FEqualWzSubstr(This, wz, cs, ichStart) \
	(This)->lpVtbl->FEqualWzSubstr(This, wz, cs, ichStart)

#define IMsoString_FEqualWzTail(This, wz, cs) \
	(This)->lpVtbl->FEqualWzTail(This, wz, cs)

#define IMsoString_FEqualWzSubstrTail(This, wz, cs, ichStart, cchSubstr) \
	(This)->lpVtbl->FEqualWzSubstrTail(This, wz, cs, ichStart, cchSubstr)

#define IMsoString_IchFindWz(This, wz, cs, ichStart) \
	(This)->lpVtbl->IchFindWz(This, wz, cs, ichStart)

#define IMsoString_IchFindWzSubstr(This, wz, ichStart, cchSubstr) \
	(This)->lpVtbl->IchFindWzSubstr(This, wz, ichStart, cchSubstr)

#define IMsoString_IchFindWch(This, wch, cs, ichStart) \
	(This)->lpVtbl->IchFindWch(This, wch, cs, ichStart)

#define IMsoString_IchFindWchSubstr(This, wch, ichStart, cchSubstr) \
	(This)->lpVtbl->IchFindWchSubstr(This, wch, ichStart, cchSubstr)

#define IMsoString_IchFindWchReverse(This, wch, cs) \
	(This)->lpVtbl->IchFindWchReverse(This, wch, cs)

#define IMsoString_IchFindWchSubstrReverse(This, wch, ichStart, cchSubstr) \
	(This)->lpVtbl->IchFindWchSubstrReverse(This, wch, ichStart, cchSubstr)

#define IMsoString_FInsertWz(This, wzSubStr, ichPos, cchReplace) \
	(This)->lpVtbl->FInsertWz(This, wzSubStr, ichPos, cchReplace)

#define IMsoString_FInsertWch(This, wch, ichPos, cchReplace) \
	(This)->lpVtbl->FInsertWch(This, wch, ichPos, cchReplace)

#define IMsoString_ReplaceWchWithWch(This, wchOld, wchNew, ichStart) \
	(This)->lpVtbl->ReplaceWchWithWch(This, wchOld, wchNew, ichStart)

#define IMsoString_ReplaceWchWithWchSubstr(This, wchOld, wchNew, ichStart, cchSubstr) \
	(This)->lpVtbl->ReplaceWchWithWchSubstr(This, wchOld, wchNew, ichStart, cchSubstr)

#define IMsoString_FTruncAt(This, ichPos) \
	(This)->lpVtbl->FTruncAt(This, ichPos)

#define IMsoString_FTruncLeft(This, cch) \
	(This)->lpVtbl->FTruncLeft(This, cch)

#define IMsoString_FTruncRight(This, cch) \
	(This)->lpVtbl->FTruncRight(This, cch)

#define IMsoString_FRemoveAt(This, ichPos, cchRemove) \
	(This)->lpVtbl->FRemoveAt(This, ichPos, cchRemove)

#define IMsoString_RemoveLeadingSpaces(This) \
	(This)->lpVtbl->RemoveLeadingSpaces(This)

#define IMsoString_RemoveTrailingSpaces(This) \
	(This)->lpVtbl->RemoveTrailingSpaces(This)

#define IMsoString_WzLockBuffer(This, cch) \
	(This)->lpVtbl->WzLockBuffer(This, cch)

#define IMsoString_ReleaseBuffer(This) \
	(This)->lpVtbl->ReleaseBuffer(This)

#define IMsoString_CchWzDecodeInt(This, w, wBase) \
	(This)->lpVtbl->CchWzDecodeInt(This, w, wBase)

#define IMsoString_CchWzDecodeUint(This, w, wBase) \
	(This)->lpVtbl->CchWzDecodeUint(This, w, wBase)

#endif // __cplusplus

#endif // MSOSTR_H
