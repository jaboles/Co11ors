/*
One-stop shopping for Office Asserts!
*/

/*
TODO: Move from msodebug.h:
	AssertInline
	AssertExp
	AssertMsgInline
	AssertMsgTemplate
TODO: kill iaw\debug.h, bln\debug.h, blnmgr\debug.h

TODO: change all the
	  do { blah blah } while (0)
	loops into
	     { blah blah }
    They're in msodbglg.h, at least.
REVIEW: (smueller) I wouldn't remove the do while(0).  If you do,
	if (condition)
		AssertishThingWithNoDoWhile;
	else
		whatever
becomes illegal.  The do while(0) makes the Assertish things look like
statements, not blocks, to the compiler.  This is a good thing, since
they look like statements, not blocks, to Office geeks.

TODO: What's MSO_ASSERT_EXP?
  Who uses it?  Who doesn't?  Can we eliminate
  the differences?

TODO: msodebug.h:
	Inline Assert stubs - these should only happen for people who
	define MSO_NO_ASSERTS and don't define the asserts we need.  If
	it happens -- screw 'em.  They don't get asserts.
Who gets this stuff?  Who loses their Asserts?  Fix them.

TODO: get rid of #pragma once, and add #error if include file
    has already been included

*/

#pragma once

#ifdef MSOASSERT_H
#pragma message ("MsoAssert.h is #included more than once")
#error
#else
#define MSOASSERT_H

#ifndef VSMSO
#define VSMSO
#endif

// **************************************************************************
// Skip the whole file if client does their own Asserts
//   (OSA, Outlook, WS, Access SOA)
#ifndef MSO_NO_ASSERTS

// Can we rip the stuff we want out of msodbglg.h, and put it in here?
#include "msodbglg.h"

#if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
#define UNTAGGED '0000'
#ifndef ZENSTAT_LIB_DEF
#ifdef OACR	
__declspec(noreturn)
#endif
#if 0 //$[VSMSO]
MSOAPI_(void) MsoULSShipAssertTagProc(DWORD dwTag, EXCEPTION_POINTERS *pep);
#else //$[VSMSO]
#define MsoULSShipAssertTagProc(dwTag, pep)
#endif //$[VSMSO]
#ifdef OACR	
__declspec(noreturn)
#endif

#if DEBUG
MSOCDECLAPI_(void) MsoULSShipAssertSzTagProcVar(DWORD dwTag, const char *sz, ...);
#endif

MSOAPI_(int) MsoULSReportExceptionTagProc(DWORD dwTag, int expr, EXCEPTION_POINTERS *pep);

// Allow apps to save their GetLastError, HRESULT or other error DWORD value
// prior to a ship assert or alert. Also leaves a possible bread trail for crashes.
// 

// Note Developers should call MsoULSSaveLastErrorSz this will get tagged by the build lab to be
// MsoULSSaveLastErrorSzTag. 
// NEVER directly call MsoULSSaveLastErrorTag() we want the szDescribe string added to the database
// to make it easier to understand what that error value means.
// You will be able to create SQM reports based on the data collected.
#if 0 //$[VSMSO]
MSOAPI_(void) MsoULSSaveLastErrorTag(DWORD dwErr, DWORD dwTag);
#else //$[VSMSO]
#define MsoULSSaveLastErrorTag(err, tag)
#endif //$[VSMSO]
#define MsoULSSaveLastErrorSz(dwErr, szDescribe) MsoULSSaveLastErrorTag(dwErr, UNTAGGED)
#define MsoULSSaveLastErrorSzTag(dwErr, szDescribe, tag) MsoULSSaveLastErrorTag(dwErr, tag)


/*-----------------------------------------------------------------------------
	MsoULSAddCvrEventTag

	Add the dwTag and dwData1 values to the Cvr_Tag_Event CVR stream.

	Use this API to record that a software event happened.  This data will
	then show up in the CVR log that is sent up with microdumps and
	minidumps from the user (crash, ship assert, watson alert).  These events
	can then be used to help reconstruct the repro steps that led up to an
	unexpected event.  This data will NOT appear in SQM reports, as it is
	never sent to the SQM server.

	To use this API, call MsoULSAddCvrEventSz with your data and a descriptive
	string, like below.
		MsoULSAddCvrEventSz(dwFormatCode, "Collected Format");

	The build lab will then process this API and give it a unique tag.
--------------------------------------------------------------------dgray----*/
MSOAPI_(void) MsoULSAddCvrEventTag(DWORD dwData1, DWORD dwTag);
#define MsoULSAddCvrEventSz(dwData1, szDescribe) MsoULSAddCvrEventTag(dwData1, UNTAGGED)
#define MsoULSAddCvrEventSzTag(dwData1, szDescribe, tag) MsoULSAddCvrEventTag(dwData1, tag)

#else
// these aren't available in ZEN static lib
#define MsoULSSaveLastErrorSz(dwErr, szDescribe) 
#define MsoULSSaveLastErrorSzTag(dwErr, szDescribe, tag) 
#define MsoULSAddCvrEventSz(dwData1, szDescribe) 
#define MsoULSAddCvrEventSzTag(dwData1, szDescribe, tag) 
#endif // ZENSTAT_LIB_DEF

#else
// these aren't available in MSO static lib
#define MsoULSSaveLastErrorSz(dwErr, szDescribe) 
#define MsoULSSaveLastErrorSzTag(dwErr, szDescribe, tag) 
#define MsoULSAddCvrEventSz(dwData1, szDescribe) 
#define MsoULSAddCvrEventSzTag(dwData1, szDescribe, tag) 

#endif // !STATIC_LIB_DEF || ZENSTAT_LIB_DEF



//---------------------------------------------------------------------------

/* FImplies must be a macro (as opposed to an inline function) so that
	short-circuiting works. */
#define FImplies(a, b)  (!(a) || (b))
#define FBiImplies(a, b)  (((a) != 0) == ((b) != 0))

//---------------------------------------------------------------------------

#define AssertData
#define DEBUGASSERTSZ
#define VSZASSERT

#define ASSERTTAG(x) x

// use the following to reserve an assert tag that gets passed to a tagged
// function in either ship or debug
// ie: dwTag_AssertTagIgnore = MsoReserveTag(ASSERTTAG('abcd'));
#define MsoReserveTag(x) x

//////////////////////////////////
// SHIP ONLY VERSIONS
///////////////////////////////////

// -- untagged
#define MsoULSShipOnlyVerify(f)					MsoULSShipOnlyAssertSz0((f), #f)
#define MsoULSShipOnlyVerifyMsg(f, sz)			MsoULSShipOnlyAssertSz0((f), sz)
// -- tagged
#define MsoULSShipOnlyVerifyTag(f, tag)			MsoULSShipOnlyAssertSz0Tag((f), #f, tag)
#define MsoULSShipOnlyVerifyMsgTag(f, sz, tag)	MsoULSShipOnlyAssertSz0Tag((f), sz, tag)

// Don't Report Exception Marker - just returns expression
#define MsoULSDontReportException(expr) expr


#if defined(STATIC_LIB_DEF) || defined(VSMSO)
// don't have ship asserts in the static lib
#define MsoULSShipOnlyAssert(f)
#define MsoULSShipOnlyAssertSz0(f, sz)
#define MsoULSShipOnlyAssertSz1(f, sz, a)
#define MsoULSShipOnlyAssertSz2(f, sz, a, b)
#define MsoULSShipOnlyAssertSz3(f, sz, a, b, c)
#define MsoULSShipOnlyAssertSz4(f, sz, a, b, c, d)
#define MsoULSShipOnlyAssertSz5(f, sz, a, b, c, d, e)

// don't have tagged asserts in the static lib, map them back to regular asserts
#define MsoULSShipOnlyAssertTag(f, tag)     MsoULSShipOnlyAssert(f)
#define MsoULSShipOnlyAssertSz0Tag(f, sz, tag)     MsoULSShipOnlyAssertSz0(f, sz)
#define MsoULSShipOnlyAssertSz1Tag(f, sz, a, tag)     MsoULSShipOnlyAssertSz1(f, sz, a)
#define MsoULSShipOnlyAssertSz2Tag(f, sz, a, b, tag)     MsoULSShipOnlyAssertSz2(f, sz, a, b)
#define MsoULSShipOnlyAssertSz3Tag(f, sz, a, b, c, tag)     MsoULSShipOnlyAssertSz3(f, sz, a, b, c)
#define MsoULSShipOnlyAssertSz4Tag(f, sz, a, b, c, d, tag)     MsoULSShipOnlyAssertSz4(f, sz, a, b, c, d)
#define MsoULSShipOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag)     MsoULSShipOnlyAssertSz5(f, sz, a, b, c, d, e)

#define MsoULSReportException(expr) expr
#define MsoULSReportExceptionTag(expr, tag) expr

#define MsoTryExceptBegin()
#define MsoTryExceptEnd()
#define MsoTryExceptEndTag(tag)


#else // !STATIC_LIB_DEF

#if DEBUG
#define MsoULSShipOnlyAssert(f) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz0(f, sz) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz1(f, sz, a) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz2(f, sz, a, b) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz3(f, sz, a, b, c) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b, c); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz4(f, sz, a, b, c, d) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b, c, d); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz5(f, sz, a, b, c, d, e) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b, c, d, e); \
			} \
		} while (0)

// Tagged Versions
#define MsoULSReportException(expr)  MsoULSReportExceptionTagProc(UNTAGGED, expr, GetExceptionInformation())

#define MsoTryExceptBegin() __try {
#define MsoTryExceptEnd()   } __except(MsoULSReportException(EXCEPTION_CONTINUE_SEARCH)) {}
#define MsoTryExceptEndTag(tag) } __except(MsoULSReportExceptionTag(EXCEPTION_CONTINUE_SEARCH, tag)) {}

#endif // !STATIC_LIB_DEF

#if defined(STATIC_LIB_DEF) || defined(VSMSO) //&& !defined(ZENSTAT_LIB_DEF)
// don't have tagged asserts in the static lib, map them back to regular asserts
#define MsoULSShipOnlyAssertTag(f, tag)     MsoULSShipOnlyAssert(f)
#define MsoULSShipOnlyAssertSz0Tag(f, sz, tag)     MsoULSShipOnlyAssertSz0(f, sz)
#define MsoULSShipOnlyAssertSz1Tag(f, sz, a, tag)     MsoULSShipOnlyAssertSz1(f, sz, a)
#define MsoULSShipOnlyAssertSz2Tag(f, sz, a, b, tag)     MsoULSShipOnlyAssertSz2(f, sz, a, b)
#define MsoULSShipOnlyAssertSz3Tag(f, sz, a, b, c, tag)     MsoULSShipOnlyAssertSz3(f, sz, a, b, c)
#define MsoULSShipOnlyAssertSz4Tag(f, sz, a, b, c, d, tag)     MsoULSShipOnlyAssertSz4(f, sz, a, b, c, d)
#define MsoULSShipOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag)     MsoULSShipOnlyAssertSz5(f, sz, a, b, c, d, e)
#define MsoULSReportExceptionTag(expr, tag) expr
#else // !STATIC_LIB_DEF || ZENSTAT_LIB_DEF

#define MsoULSShipOnlyAssertTag(f, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz0Tag(f, sz, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz1Tag(f, sz, a, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz2Tag(f, sz, a, b, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz3Tag(f, sz, a, b, c, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b, c); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz4Tag(f, sz, a, b, c, d, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b, c, d); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b, c, d, e); \
			} \
		} while (0)
#else /* !DEBUG */
#define MsoULSShipOnlyAssert(f) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz0(f, sz) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz1(f, sz, a) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz2(f, sz, a, b) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz3(f, sz, a, b, c) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz4(f, sz, a, b, c, d) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz5(f, sz, a, b, c, d, e) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			} \
		} while (0)

// Tagged Versions
#define MsoULSShipOnlyAssertTag(f, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz0Tag(f, sz, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz1Tag(f, sz, a, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz2Tag(f, sz, a, b, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz3Tag(f, sz, a, b, c, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz4Tag(f, sz, a, b, c, d, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#define MsoULSShipOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			} \
		} while (0)

#endif
#define MsoULSReportException(expr)  MsoULSReportExceptionTagProc(UNTAGGED, expr, GetExceptionInformation())
#define MsoULSReportExceptionTag(expr, tag)  MsoULSReportExceptionTagProc(tag, expr, GetExceptionInformation())

#define MsoTryExceptBegin() __try {
#define MsoTryExceptEnd()   } __except(MsoULSReportException(EXCEPTION_CONTINUE_SEARCH)) {}
#define MsoTryExceptEndTag(tag) } __except(MsoULSReportExceptionTag(EXCEPTION_CONTINUE_SEARCH, tag)) {}

#endif // !STATIC_LIB_DEF

//////////////////////////////////
// END SHIP ONLY VERSIONS
///////////////////////////////////

#ifdef DEBUG

#define vszAssertFile __FILE__

///////////////////////////////////
// DEBUG ONLY VERSIONS
///////////////////////////////////

// -- untagged
#define MsoULSDebugOnlyVerify(f)				MsoULSDebugOnlyAssertSz0((f), #f)
#define MsoULSDebugOnlyVerifyMsg(f, sz)			MsoULSDebugOnlyAssertSz0((f), sz)
// -- tagged
#define MsoULSDebugOnlyVerifyTag(f, tag)			MsoULSDebugOnlyAssertSz0Tag((f), #f, tag)
#define MsoULSDebugOnlyVerifyMsgTag(f, sz, tag)	MsoULSDebugOnlyAssertSz0Tag((f), sz, tag)

#define MsoULSDebugOnlyAssert(f) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, #f)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz0(f, sz) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz1(f, sz, a) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz2(f, sz, a, b) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz3(f, sz, a, b, c) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b, c)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz4(f, sz, a, b, c, d) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b, c, d)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz5(f, sz, a, b, c, d, e) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b, c, d, e)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#if (defined(STATIC_LIB_DEF) && !defined(ZENSTAT_LIB_DEF)) || defined(VSMSO)
// don't have tagged asserts in the static lib, map them back to regular asserts
// but we have them in ZEN static lib
#define MsoULSDebugOnlyAssertTag(f, tag)     MsoULSDebugOnlyAssert(f)
#define MsoULSDebugOnlyAssertSz0Tag(f, sz, tag)     MsoULSDebugOnlyAssertSz0(f, sz)
#define MsoULSDebugOnlyAssertSz1Tag(f, sz, a, tag)     MsoULSDebugOnlyAssertSz1(f, sz, a)
#define MsoULSDebugOnlyAssertSz2Tag(f, sz, a, b, tag)     MsoULSDebugOnlyAssertSz2(f, sz, a, b)
#define MsoULSDebugOnlyAssertSz3Tag(f, sz, a, b, c, tag)     MsoULSDebugOnlyAssertSz3(f, sz, a, b, c)
#define MsoULSDebugOnlyAssertSz4Tag(f, sz, a, b, c, d, tag)     MsoULSDebugOnlyAssertSz4(f, sz, a, b, c, d)
#define MsoULSDebugOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag)     MsoULSDebugOnlyAssertSz5(f, sz, a, b, c, d, e)
#else // !STATIC_LIB_DEF || ZENSTAT_LIB_DEF

#define MsoULSDebugOnlyAssertTag(f, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, #f)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz0Tag(f, sz, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz1Tag(f, sz, a, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz2Tag(f, sz, a, b, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz3Tag(f, sz, a, b, c, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b, c)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz4Tag(f, sz, a, b, c, d, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b, c, d)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSDebugOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag) \
	do { \
		if ((f) == 0) \
			{ \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b, c, d, e)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)
#endif // !STATIC_LIB_DEF || ZENSTAT_LIB_DEF

///////////////////////////////////
// END DEBUG ONLY VERSIONS
///////////////////////////////////

///////////////////////////////////
// DEBUG + SHIP VERSIONS
///////////////////////////////////

// -- untagged
#define MsoULSShipVerify(f)					MsoULSShipAssertSz0((f), #f)
#define MsoULSShipVerifyMsg(f, sz)			MsoULSShipAssertSz0((f), sz)
// -- tagged
#define MsoULSShipVerifyTag(f, tag)			MsoULSShipAssertSz0Tag((f), #f, tag)
#define MsoULSShipVerifyMsgTag(f, sz, tag)	MsoULSShipAssertSz0Tag((f), sz, tag)

#if defined(STATIC_LIB_DEF) || defined(VSMSO)
// don't have ship asserts in the static lib, map them to debug only
#define MsoULSShipAssert     	MsoULSDebugOnlyAssert
#define MsoULSShipAssertSz0	MsoULSDebugOnlyAssertSz0
#define MsoULSShipAssertSz1	MsoULSDebugOnlyAssertSz1
#define MsoULSShipAssertSz2	MsoULSDebugOnlyAssertSz2
#define MsoULSShipAssertSz3	MsoULSDebugOnlyAssertSz3
#define MsoULSShipAssertSz4	MsoULSDebugOnlyAssertSz4
#define MsoULSShipAssertSz5	MsoULSDebugOnlyAssertSz5
#else // !STATIC_LIB_DEF

#define MsoULSShipAssert(f) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(UNTAGGED, NULL); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, #f)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz0(f, sz) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz1(f, sz, a) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz2(f, sz, a, b) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz3(f, sz, a, b, c) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b, c); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b, c)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz4(f, sz, a, b, c, d) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b, c, d); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b, c, d)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz5(f, sz, a, b, c, d, e) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(UNTAGGED, sz, a, b, c, d, e); \
			if (MsoFAssertSzProcVar (__FILE__, __LINE__, sz, a, b, c, d, e)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)
#endif // !STATIC_LIB_DEF

#if (defined(STATIC_LIB_DEF) && !defined(ZENSTAT_LIB_DEF)) || defined(VSMSO)
// don't have tagged asserts in the static lib, map them back to regular asserts
#define MsoULSShipAssertTag(f, tag)     MsoULSShipAssert(f)
#define MsoULSShipAssertSz0Tag(f, sz, tag)     MsoULSShipAssertSz0(f, sz)
#define MsoULSShipAssertSz1Tag(f, sz, a, tag)     MsoULSShipAssertSz1(f, sz, a)
#define MsoULSShipAssertSz2Tag(f, sz, a, b, tag)     MsoULSShipAssertSz2(f, sz, a, b)
#define MsoULSShipAssertSz3Tag(f, sz, a, b, c, tag)     MsoULSShipAssertSz3(f, sz, a, b, c)
#define MsoULSShipAssertSz4Tag(f, sz, a, b, c, d, tag)     MsoULSShipAssertSz4(f, sz, a, b, c, d)
#define MsoULSShipAssertSz5Tag(f, sz, a, b, c, d, e, tag)     MsoULSShipAssertSz5(f, sz, a, b, c, d, e)
#elif defined(ZENSTAT_LIB_DEF)// STATIC_LIB_DEF && !ZENSTAT_LIB_DEF
//we have tagged asserts in ZEN, but only debug
#define MsoULSShipAssertTag(f, tag)     MsoULSDebugOnlyAssertTag(f, tag)
#define MsoULSShipAssertSz0Tag(f, sz, tag)     MsoULSDebugOnlyAssertSz0Tag(f, sz, tag)
#define MsoULSShipAssertSz1Tag(f, sz, a, tag)     MsoULSDebugOnlyAssertSz1Tag(f, sz, a, tag)
#define MsoULSShipAssertSz2Tag(f, sz, a, b, tag)     MsoULSDebugOnlyAssertSz2Tag(f, sz, a, b, tag)
#define MsoULSShipAssertSz3Tag(f, sz, a, b, c, tag)     MsoULSDebugOnlyAssertSz3Tag(f, sz, a, b, c, tag)
#define MsoULSShipAssertSz4Tag(f, sz, a, b, c, d, tag)     MsoULSDebugOnlyAssertSz4Tag(f, sz, a, b, c, d, tag)
#define MsoULSShipAssertSz5Tag(f, sz, a, b, c, d, e, tag)     MsoULSDebugOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag)
#else // ZENSTAT_LIB_DEF
#define MsoULSShipAssertTag(f, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertTagProc(tag, NULL); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, #f)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz0Tag(f, sz, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz1Tag(f, sz, a, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz2Tag(f, sz, a, b, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz3Tag(f, sz, a, b, c, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b, c); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b, c)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz4Tag(f, sz, a, b, c, d, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b, c, d); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b, c, d)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)

#define MsoULSShipAssertSz5Tag(f, sz, a, b, c, d, e, tag) \
	do { \
		if ((f) == 0) \
			{ \
			MsoULSShipAssertSzTagProcVar(tag, sz, a, b, c, d, e); \
			if (MsoFAssertSzTagProcVar (tag, __FILE__, __LINE__, sz, a, b, c, d, e)) \
				MsoDebugBreakInline(); \
			} \
		} while (0)
#endif // !STATIC_LIB_DEF || ZENSTAT_LIB_DEF

///////////////////////////////////
// END DEBUG + SHIP VERSIONS
///////////////////////////////////

///////////////////////////////////
// OLD MACRO VERSIONS
///////////////////////////////////

// -- untagged

#define Verify						MsoULSShipVerify
#define VerifyMsg					MsoULSShipVerifyMsg

#define Assert					MsoULSShipAssert
#define AssertSz0					MsoULSShipAssertSz0
#define AssertSz1					MsoULSShipAssertSz1
#define AssertSz2					MsoULSShipAssertSz2
#define AssertSz3					MsoULSShipAssertSz3
#define AssertSz4					MsoULSShipAssertSz4
#define AssertSz5					MsoULSShipAssertSz5

// -- tagged

#define MsoVerifyTag				MsoULSShipVerifyTag
#define MsoVerifyMsgTag			MsoULSShipVerifyMsgTag

#define MsoAssertTag				MsoULSShipAssertTag
#define MsoAssertSz0Tag			MsoULSShipAssertSz0Tag
#define MsoAssertSz1Tag			MsoULSShipAssertSz1Tag
#define MsoAssertSz2Tag			MsoULSShipAssertSz2Tag
#define MsoAssertSz3Tag			MsoULSShipAssertSz3Tag
#define MsoAssertSz4Tag 			MsoULSShipAssertSz4Tag
#define MsoAssertSz5Tag			MsoULSShipAssertSz5Tag

// -- never tagged
// These exist because some asserts are so generic we would have 1000s of bugs
// with the same tag in the database.
// These should RARELY be used.  If in doubt contact jeffmit.
#define MsoAssertNeverTag		MsoULSShipAssert
#define MsoAssertSz0NeverTag	MsoULSShipAssertSz0
#define MsoAssertSz1NeverTag	MsoULSShipAssertSz1
#define MsoAssertSz2NeverTag	MsoULSShipAssertSz2
#define MsoAssertSz3NeverTag	MsoULSShipAssertSz3
#define MsoAssertSz4NeverTag	MsoULSShipAssertSz4
#define MsoAssertSz5NeverTag	MsoULSShipAssertSz5

///////////////////////////////////
// END OLD MACRO VERSIONS
///////////////////////////////////

#else // !DEBUG

///////////////////////////////////
// DEBUG + SHIP VERSIONS
///////////////////////////////////

// -- untagged

#define MsoULSShipVerify			MsoULSShipOnlyVerify
#define MsoULSShipVerifyMsg		MsoULSShipOnlyVerifyMsg

#define MsoULSShipAssert			MsoULSShipOnlyAssert
#define MsoULSShipAssertSz0		MsoULSShipOnlyAssertSz0
#define MsoULSShipAssertSz1		MsoULSShipOnlyAssertSz1
#define MsoULSShipAssertSz2		MsoULSShipOnlyAssertSz2
#define MsoULSShipAssertSz3		MsoULSShipOnlyAssertSz3
#define MsoULSShipAssertSz4		MsoULSShipOnlyAssertSz4
#define MsoULSShipAssertSz5		MsoULSShipOnlyAssertSz5

// -- tagged

#define MsoULSShipVerifyTag		MsoULSShipOnlyVerifyTag
#define MsoULSShipVerifyMsgTag		MsoULSShipOnlyVerifyMsgTag

#define MsoULSShipAssertTag		MsoULSShipOnlyAssertTag
#define MsoULSShipAssertSz0Tag		MsoULSShipOnlyAssertSz0Tag
#define MsoULSShipAssertSz1Tag		MsoULSShipOnlyAssertSz1Tag
#define MsoULSShipAssertSz2Tag		MsoULSShipOnlyAssertSz2Tag
#define MsoULSShipAssertSz3Tag		MsoULSShipOnlyAssertSz3Tag
#define MsoULSShipAssertSz4Tag		MsoULSShipOnlyAssertSz4Tag
#define MsoULSShipAssertSz5Tag		MsoULSShipOnlyAssertSz5Tag

///////////////////////////////////
// END DEBUG + SHIP VERSIONS
///////////////////////////////////

///////////////////////////////////
// DEBUG ONLY VERSIONS
///////////////////////////////////

// -- untagged

#define MsoULSDebugOnlyVerify(f)			Verify(f)
#define MsoULSDebugOnlyVerifyMsg(f,sz)	Verify(f)

#define MsoULSDebugOnlyAssert(f)
#define MsoULSDebugOnlyAssertSz0(f, sz)
#define MsoULSDebugOnlyAssertSz1(f, sz, a)
#define MsoULSDebugOnlyAssertSz2(f, sz, a, b)
#define MsoULSDebugOnlyAssertSz3(f, sz, a, b, c)
#define MsoULSDebugOnlyAssertSz4(f, sz, a, b, c, d)
#define MsoULSDebugOnlyAssertSz5(f, sz, a, b, c, d, e)

// -- tagged

#define MsoULSDebugOnlyVerifyTag(f, tag)			Verify(f)
#define MsoULSDebugOnlyVerifyMsgTag(f, sz, tag)	Verify(f)

#define MsoULSDebugOnlyAssertTag(f, tag)
#define MsoULSDebugOnlyAssertSz0Tag(f, sz, tag)
#define MsoULSDebugOnlyAssertSz1Tag(f, sz, a, tag)
#define MsoULSDebugOnlyAssertSz2Tag(f, sz, a, b, tag)
#define MsoULSDebugOnlyAssertSz3Tag(f, sz, a, b, c, tag)
#define MsoULSDebugOnlyAssertSz4Tag(f, sz, a, b, c, d, tag)
#define MsoULSDebugOnlyAssertSz5Tag(f, sz, a, b, c, d, e, tag)

///////////////////////////////////
// END DEBUG ONLY VERSIONS
///////////////////////////////////

///////////////////////////////////
// OLD MACRO VERSIONS
///////////////////////////////////

// -- untagged

#if( !OACR )
#define Verify(f)					(f)
#else
#define Verify(f)					OACR_ASSERTDO(f)
#endif
#define VerifyMsg(f, sz)		Verify(f)

#define Assert(f)
#define AssertSz0(f, sz)
#define AssertSz1(f, sz, a)
#define AssertSz2(f, sz, a, b)
#define AssertSz3(f, sz, a, b, c)
#define AssertSz4(f, sz, a, b, c, d)
#define AssertSz5(f, sz, a, b, c, d, e)

// -- tagged

#define MsoVerifyTag(f, tag)			Verify(f)
#define MsoVerifyMsgTag(f, sz, tag)	Verify(f)

#define MsoAssertTag(f, tag)
#define MsoAssertSz0Tag(f, sz, tag)
#define MsoAssertSz1Tag(f, sz, a, tag)
#define MsoAssertSz2Tag(f, sz, a, b, tag)
#define MsoAssertSz3Tag(f, sz, a, b, c, tag)
#define MsoAssertSz4Tag(f, sz, a, b, c, d, tag)
#define MsoAssertSz5Tag(f, sz, a, b, c, d, e, tag)

// -- never tagged
// These exist because some asserts are so generic we would have 1000s of bugs
// with the same tag in the database.
// These should RARELY be used.  If in doubt contact jeffmit.
#define MsoAssertNeverTag(f)
#define MsoAssertSz0NeverTag(f, sz)
#define MsoAssertSz1NeverTag(f, sz, a)
#define MsoAssertSz2NeverTag(f, sz, a, b)
#define MsoAssertSz3NeverTag(f, sz, a, b, c)
#define MsoAssertSz4NeverTag(f, sz, a, b, c, d)
#define MsoAssertSz5NeverTag(f, sz, a, b, c, d, e)

///////////////////////////////////
// END OLD MACRO VERSIONS
///////////////////////////////////

#endif // !DEBUG

//---------------------------------------------------------------------------
/////////////////////////////////////////////////////////////////////////////
//                                                                         //
// READ ME                                                                 //
// MsoAssert*Tag functions are just like Assert* functions but with an     //
// extra parameter for identification purposes.                            //
// See http://officedev/asserttags.htm for more information.               //
//                                                                         //
/////////////////////////////////////////////////////////////////////////////
//---------------------------------------------------------------------------

// -- untagged
#define AssertSzSz(f, a, b)							AssertSz2(f, "%s %s", a, b)
#define AssertSzSzSz(f, a, b, c)						AssertSz3(f, "%s %s %s", a, b, c)
#define AssertSzSzSzSz(f, a, b, c, d)					AssertSz4(f, "%s %s %s %s", a, b, c, d)
#define AssertSzSzSzSzSz(f, a, b, c, d, e)				AssertSz5(f, "%s %s %s %s %s", a, b, c, d, e)
#define AssertSzIntInt(f, a, b, c)						AssertSz3(f, "%s %d %d", a, b, c)
#define AssertSzSzInt(f, a, b, c)						AssertSz3(f, "%s %s %d", a, b, c)
// -- tagged
#define MsoAssertSzSzTag(f, a, b, tag)				MsoAssertSz2Tag(f, "%s %s", a, b, tag)
#define MsoAssertSzSzSzTag(f, a, b, c, tag)			MsoAssertSz3Tag(f, "%s %s %s", a, b, c, tag)
#define MsoAssertSzSzSzSzTag(f, a, b, c, d, tag)		MsoAssertSz4Tag(f, "%s %s %s %s", a, b, c, d, tag)
#define MsoAssertSzSzSzSzSzTag(f, a, b, c, d, e, tag)	MsoAssertSz5Tag(f, "%s %s %s %s %s", a, b, c, d, e, tag)
#define MsoAssertSzIntIntTag(f, a, b, c, tag)			MsoAssertSz3Tag(f, "%s %d %d", a, b, c, tag)
#define MsoAssertSzSzIntTag(f, a, b, c, tag)			MsoAssertSz3Tag(f, "%s %s %d", a, b, c, tag)

// -- untagged
#define AssertEx							Assert
#define AssertExp							Assert
// -- tagged
#define MsoAssertExTag					MsoAssertTag
#define MsoAssertExpTag					MsoAssertTag

// -- untagged
#define AssertSz(f, sz)						AssertSz0((f), sz)
#define AssertMsg							AssertSz
// -- tagged
#define MsoAssertSzTag(f, sz, tag)			MsoAssertSz0Tag((f), sz, tag)
#define MsoAssertMsgTag					MsoAssertSzTag

// -- untagged
#define VerifyExp							Verify
#define AssertDo							Verify
#define SideAssert							Verify
// -- tagged
#define MsoVerifyExpTag					MsoVerifyTag
#define MsoAssertDoTag					MsoVerifyTag
#define MsoSideAssertTag					MsoVerifyTag

#if( OACR )

#if !defined(STATIC_LIB_DEF) && !defined(ZENSTAT_LIB_DEF)
	// -- untagged
	#define NotReached()						MsoULSShipAssertTagProc(UNTAGGED, NULL)
	#define NotReachedSz(sz)					MsoULSShipAssertTagProc(UNTAGGED, NULL)
	#define NotYetImplemented()				MsoULSShipAssertTagProc(UNTAGGED, NULL)
	// -- tagged
	#define MsoNotReachedTag(tag)				MsoULSShipAssertTagProc(tag, NULL)
	#define MsoNotReachedSzTag(sz, tag)			MsoULSShipAssertTagProc(tag, NULL)
	#define MsoNotYetImplementedTag(tag)		MsoULSShipAssertTagProc(tag, NULL)
#else
	// -- untagged
	#define NotReached()
	#define NotReachedSz(sz)
	#define NotYetImplemented()
	// -- tagged
	#define MsoNotReachedTag(tag)
	#define MsoNotReachedSzTag(sz, tag)
	#define MsoNotYetImplementedTag(tag)
#endif

#else // !OACR

#if defined(STATIC_LIB_DEF) || defined(DEBUG)
	// -- untagged
	#define NotReached()						AssertSz(FALSE, "Not Reached")
	#define NotReachedSz(sz)					AssertSz(FALSE, sz)
	#define NotYetImplemented()				AssertSz(FALSE, "Not Yet Implemented")
	// -- tagged
	#define MsoNotReachedTag(tag)				MsoAssertSzTag(FALSE, "Not Reached", tag)
	#define MsoNotReachedSzTag(sz, tag)			MsoAssertSzTag(FALSE, sz, tag)
	#define MsoNotYetImplementedTag(tag)		MsoAssertSzTag(FALSE, "Not Yet Implemented", tag)
#else
	// -- untagged
	#define NotReached()						MsoULSShipAssertTagProc(UNTAGGED, NULL)
	#define NotReachedSz(sz)					MsoULSShipAssertTagProc(UNTAGGED, NULL)
	#define NotYetImplemented()				MsoULSShipAssertTagProc(UNTAGGED, NULL)
	// -- tagged
	#define MsoNotReachedTag(tag)				MsoULSShipAssertTagProc(tag, NULL)
	#define MsoNotReachedSzTag(sz, tag)			MsoULSShipAssertTagProc(tag, NULL)
	#define MsoNotYetImplementedTag(tag)		MsoULSShipAssertTagProc(tag, NULL)
#endif

#endif // OACR

// -- untagged
#define AssertImplies(a, b)					AssertSz(FImplies(a, b), #a " => " #b)
#define AssertBiImplies(a, b)				AssertSz(FBiImplies(a, b), #a " <==> " #b)
// -- tagged
#define MsoAssertImpliesTag(a, b, tag)		MsoAssertSzTag(FImplies(a, b), #a " => " #b, tag)
#define MsoAssertBiImpliesTag(a, b, tag)		MsoAssertSzTag(FBiImplies(a, b), #a " <==> " #b, tag)

// -- ship assert versions
#define MsoULSShipAssertImpliesTag(a, b, tag)		MsoULSShipAssertSzTag(FImplies(a, b), #a " => " #b, tag)
#define MsoULSShipAssertBiImpliesTag(a, b, tag)		MsoULSShipAssertSzTag(FBiImplies(a, b), #a " <==> " #b, tag)
#define MsoULSShipAssertImpliesTag(a, b, tag)		MsoULSShipAssertSzTag(FImplies(a, b), #a " => " #b, tag)
#define MsoULSShipAssertBiImpliesTag(a, b, tag)		MsoULSShipAssertSzTag(FBiImplies(a, b), #a " <==> " #b, tag)

#if DEBUG
// -- untagged
#define MsoULSShipAssertSz(f, sz)			MsoULSShipAssertSz0((f), sz)
#define MsoULSShipOnlyAssertSz(f, sz)		MsoULSShipOnlyAssertSz0((f), sz)
// -- tagged
#define MsoULSShipAssertSzTag(f, sz, tag)		MsoULSShipAssertSz0Tag((f), sz, tag)
#define MsoULSShipOnlyAssertSzTag(f, sz, tag)	MsoULSShipOnlyAssertSz0Tag((f), sz, tag)
#else
// -- untagged
#define MsoULSShipAssertSz(f, sz)			MsoULSShipAssert((f))
#define MsoULSShipOnlyAssertSz(f, sz)		MsoULSShipOnlyAssert((f))
// -- tagged
#define MsoULSShipAssertSzTag(f, sz, tag)		MsoULSShipAssertTag((f), tag)
#define MsoULSShipOnlyAssertSzTag(f, sz, tag)	MsoULSShipOnlyAssertTag((f), tag)

#endif

#if 0 //$[VSMSO]
#if defined(DEBUG) && !defined(_WIN64)
// This API handles /RTC1 reports
MSOCDECLAPI_(int) MsoRTCAssert(int, const char *, int, const char *, const char *, ...);

// This forces a reference to __fInitMsoRTC (defined in msortc.obj, linked into mso(s)d.lib
// which does the initialziation at boot time.
#pragma comment(linker, "/include:___fInitMsoRTC")

// For bogus apps that include msoassert.h (you know who you are) but don't
// link with MSO, this dummy variable replaces the above-mentioned boot variable.
// The linker directive tells the linker to use this symbol as a backup if it
// can't find __fInitMsoRTC anywhere.
MSOEXTERN_C extern __declspec(selectany) void *__fNoMsoRTCAvailable = 0;
#pragma comment(linker, "/alternatename:___fInitMsoRTC=___fNoMsoRTCAvailable")
#endif
#endif //$[VSMSO]

#endif	// (End of) skip the whole file if client does their own Asserts

// **************************************************************************

// Compile time asserts can be used for conditions that use constant expressions.
// e.g. MsoCompileAssert( sizeof( myStruct ) == 24 );
//      MsoCompileAssert( thisConstant == thatConstant );
// Compile asserts do not produce any code
// C++ version can be used on global scope.

#ifdef __cplusplus
#define MsoCompileAssert( fCondition ) extern int __dummy[fCondition]
#else
#define MsoCompileAssert( fCondition ) do{ extern int __dummy[fCondition]; } while(0)
#endif

#endif // MSOASSERT_H
