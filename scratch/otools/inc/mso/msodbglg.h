#pragma once

/*----------------------------------------------------------------------------
	%%File: msodbglg.h
	%%Unit: Debug
	%%Contact: kirkg

	Debugging macros and functions.

	The debug log mechanism automatically creates and manages a log file
	"debug.log".  By controlling a single variable, vwDebugLogLvl, we
	can route printf()-style output to a log file, or to the Comm port
	AND the log file.

	The log is created and written to whenever the run-time log level,
	vwDebugLogLvl >= wTraceLvl.  wTraceLvl is the declared, or compile-
	time log level, passed as an argument to MsoDebugLog().  It will NOT print
	if vwDebugLogLvl is LESS than the declared trace level.  If
	vwDebugLogLvl is an ODD number, anything written to the log file
	will also appear on the port used by OutputDebugString().

	NOTE:
		THE SPRINTF STRING MUST NOT EXPAND BEYOND 256 CHARACTERS!

	USAGE:

		There are several macros, debugXX0() thru debugXX6(), etc.,
		that are the main interface to MsoDebugLog(), where "XX" is the
		subsystem mneumonic.  EACH SUBSYSTEM GROUP SHOULD HAVE ITS OWN SET
		OF MACROS.  Thus the Event Monitor group adds debugEM0 thru
		debugEM6() and the Hyperlink group might add debugHL0 thru
		debugHL6, etc.  EACH SUBSYSTEM GETS ITS OWN LOG FILTER BIT,
		DEFINED BELOW (see fDebugFilterEm), which allows developers to
		see their logging info without having to wade through other logging
		output.

		MsoDebugLog() has a variable number of arguments, and is
		similar to the standard C sprintf() routine, and in fact calls it.
		The macros remove the need to enclose calls to MsoDebugLog() with
		#ifdef's, but C does not allow variable numbers of arguments in
		macros.  To overcome this (and also due to problems with CsConst'ing
		the format strings), each debugxx() macro has a fixed number of
		arguments, where xx indicates the number of arguments past the format
		string.

		Examples:

			debugXX0((fTrace ? 0 : 6), "Value is NOT chosen\n");
			debugXX1(6, "NEW SECTION: %d\n", wSection);
			debugXX4(2, "Paragraph %d (CP %ld thru %ld) has style %d\n",
				     iaftag, lpaftag->cpFirst, (lpaftag + 1)->cpFirst, istd);

		The first argument to the macros is the wTraceLvl, the declared
		trace-level.  In the three examples above, we are saying: the first
		should not print unless vwDebugLogLvl >= (fTrace ? 0 : 6),
		the second unless vwDebugLogLvl >= 6, and the third unless
		vwDebugLogLvl >= 2.

	CHOICE OF DECLARED TRACE LEVELS:

		In general, low declared trace levels should be used within discrete
		subsystems that are *not* shared, and high levels should be used
		within shared code.  Note that only EVEN numbers are used.  The
		reason will be explained below.

		THE HIGHER THE LEVEL NUMBER, THE MORE GOO YOU ARE GOING TO OUTPUT.

			Level		Generally used when:
			------------------------------------------------------
			-1			Warning messages.

			0			Messages to be seen by testers.
			2			Message to be seen by rigorous testers.

			4			Structural changes, such as dialog boxes appearing,
						Windows created, features beginning.

			6			Normal logging of a subsystem.
			8			Detailed logging of a subsystem.

			10			Major subsystem I/O, such as scanning CHRs, or
						reading buffers.
			12			Minor subsystem I/O, such as reading characters.

			...

			20			Shared subsystem, such as FormatLine, CachePara
						or something else that we would call frequently,
						and be buried under detail.

	RUN-TIME USAGE:

		Set vwDebugLogLvl to an ODD number, if you wish output via
		OutputDebugString() as well as the log file.  Set vwDebugLogLvl to an
		EVEN number if you wish output only to the log file.

		Set vwDebugLogLvl to a low level using the dialog box.  For
		higher levels (e.g. to debug FormatLine()), you should set a
		breakpoint in the debugger, and set vwDebugLogLvl when in the
		targeted routine.

		Example (from Word):

			1.	We set vwDebugLogLvl to 2 from the dialog box in the
				"Preferences" memu.  We also set a breakpoint within
				TkLexText().
			2.	We run the AutoFormatter, and when we hit the breakpoint,
				we set vwDebugLogLvl from the watch window to 8,
				so we will get a lot of output.
			3.	When we approach the interesting point of a problem, we
				change the vwDebugLogLvl to 9, so it will appear on
				the OutputDebugStringA() terminal.

----------------------------------------------------------------------------*/

#ifdef MSODBGLG_H
#error
#else
#define MSODBGLG_H

#include <stdarg.h>


#ifdef STANDALONE
#error
#endif

#ifdef STANDALONE_WORD
#error
#endif

#ifdef UNIX
#error
#endif

#if _MSC_VER <= 600
#error
#endif


MSOEXTERN_C_BEGIN	// ***************** Begin extern "C" ********************

#ifndef UNIX
#if _MSC_VER <= 600
#ifndef __cdecl
#define __cdecl _cdecl
#endif /* !__cdecl */
#endif /* MSC_VER <= 600 */
#else
#define __cdecl
#define sprintf ansi_sprintf
#define vsprintf ansi_vsprintf
int ansi_sprintf(_Out_ char *lpchBuf, const char *sz, ...);
int ansi_vsprintf(_Out_ char *lpchBuf, const char *sz, va_list ap);
#define CommSz printf
#endif /* !UNIX */


// Define filters for MsoDebugLog()
// These values can overlap for different applications.
// However, Office reserves the lower 0x00FF bits.

#define fDebugFilterEm			0x0001U					// App Event Monitor
#define fDebugFilterWc			0x0002U					// Web Client
#define fDebugFilterAcb			0x0004U					// Active ClipBoard
#define fDebugFilterEc			0x0008U					// Exec Command (DOIT)
#define fDebugFilterMsoEm		0x0010U					// MSO Event Monitor
#define fDebugFilterLS			0x0020U					// Word LineServices
#define fDebugFilterGEL			0x0040U					// Office GEL

#if defined(WORD_BUILD)
#define fDebugFilterPrint		0x0100U					// Print
#define fDebugFilterXml			0x0200U					// XML (was Reconcile)
#define fDebugFilterHtmlIn		0x0400U					// HTML in

#elif defined(EXCEL_BUILD)
#define fDebugFilterUnused1		0x0100U					// Use me 1st in Excel

#elif defined(PPT_BUILD)
#define fDebugFilterUnused1		0x0100U					// Use me 1st in PPT

#elif defined(ACCESS_BUILD)
#define fDebugFilterUnused1		0x0100U					// Use me 1st in Access
#endif

#define fDebugFilterAll			0xFFFFU


#ifdef DEBUG

#ifndef WORD_BUILD
extern int vwDebugLogFilter;							// Debug trace filter
extern int vwDebugLogLvl;								// Debug trace level
#endif /* !WORD_BUILD */


// Return the sz, or if null, return "(null)"
#define SzOrNull(sz) \
			((sz) != NULL ? (sz) : "(null)")

MSOCDECLAPI_(void) MsoDebugLog(							// Print debug msg
	int					wTraceLvl,
	unsigned int		grfFilter,
	const char *sz,
	...
	);
MSOCDECLAPI_(BOOL) MsoFAssertSzProcVar(					// Assert with sprintf
	const char	       *szFile,
	int					line,
	const char	       *sz,
	...
	);
#if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
MSOCDECLAPI_(BOOL) MsoFAssertSzTagProcVar(     // tagged Assert with sprintf
	DWORD              dwTag,
	const char         *szFile,
	int                line,
	const char         *sz,
	...
	);
#endif // !STATIC_LIB_DEF || ZENSTAT_LIB_DEF
MSOCDECLAPI_(int) MsoFReportSzProcVar(					// ReportSz w/ sprintf
	const char         *szFile,
	int					line,
	const char	       *sz,
	...
	);
#if 0 //$[VSMSO]
MSOCDECLAPI_(int) MsoFFeatureReportSzProcVar(			// Report per Feature
	DWORD					dwTag,
	unsigned int		bitfFeatureFilter,
	const char         *szFile,
	int					line,
	const char	       *sz,
	...
	);
#endif //$[VSMSO]
MSOCDECLAPI_(void) MsoCommSzVar(const char *sz, ...);	// CommSz with sprintf
MSOAPI_(void) MsoDebugLogAp(							// Print debug msg
	int					wTraceLvl,
	unsigned int		grfFilter,
	const char *sz,
	va_list				ap
	);
MSOAPI_(void) MsoDebugLogPch(							// Print large msg
	int					wTraceLvl,
	unsigned int		grfFilter,
	_Out_cap_(cchLen)
	char			   *pch,
	int					cchLen,
	int					fIsRgxch
	);
MSOAPI_(int) MsoFDebugLogCloseFile(void);				// Close log file
MSOAPI_(int *) MsoPwDebugLogLvl(int **ppwDebugLogFilter);// Return ptr to vars

// These *debug* routines are XCHAR (WCHAR on unicode builds and char on ANSI)
#ifdef ANSI_XCHAR
#error
#endif

typedef unsigned short XCHAR;

MSOAPI_(XCHAR *) MsoXszFromRgxchDebug(					// Convert pwch to wz
	const XCHAR			   *rgxch,
	int					cch
	);
MSOAPI_(char *) MsoSzFromRgxchDebug(					// Convert pwch to sz
	const XCHAR		   *rgxch,
	int					cch
	);
MSOAPI_(char *) MsoSzFromXszDebug(const XCHAR *xsz);	// Convert wz to sz


// NOTE: debugvar has to be called with 2 parenthesis...
#define debugvar(a)		MsoDebugLog a
#define Debug(e)		e
#define DebugElse(s, t)	s
#define debuglog0(wLevel, grfFilter, sz) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug); \
			} while (0)
#define debuglog1(wLevel, grfFilter, sz, a) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug, a); \
			} while (0)
#define debuglog2(wLevel, grfFilter, sz, a, b) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug, a, b); \
			} while (0)
#define debuglog3(wLevel, grfFilter, sz, a, b, c) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug, a, b, c); \
			} while (0)
#define debuglog4(wLevel, grfFilter, sz, a, b, c, d) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug, a, b, c, d); \
			} while (0)
#define debuglog5(wLevel, grfFilter, sz, a, b, c, d, e) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug, a, b, c, d, e); \
			} while (0)
#define debuglog6(wLevel, grfFilter, sz, a, b, c, d, e, f) \
			do { \
			static const char 		szDebug[] = sz; \
			MsoDebugLog(wLevel, grfFilter, szDebug, a, b, c, d, e, f); \
			} while (0)
#define debuglogPch(wLevel, grfFilter, pch, cch) \
			MsoDebugLogPch(wLevel, grfFilter, pch, cch, fFalse)
#define debuglogPwch(wLevel, grfFilter, pwch, cwch) \
			MsoDebugLogPch(wLevel, grfFilter, (char *)pwch, cwch, fTrue)


#define ReportSz(sz) 	MsoReportSz(sz)

#define MsoReportSz(sz) \
			do { \
			static const char szXXXXXXXFar[] = sz; \
			if (!MsoFReportSzProcVar(vszAssertFile, __LINE__, szXXXXXXXFar)) \
				MsoDebugBreakInline(); \
			} while (0)

#define ReportSz0If(f, sz) \
			if (!(f)) \
				{ \
				static const char 		szReport[] = sz; \
				if (!MsoFReportSzProcVar (vszAssertFile, __LINE__, szReport)) \
					MsoDebugBreakInline(); \
				} \
			else

#define ReportSz1If(f, sz, a) \
			if (!(f)) \
				{ \
				static const char 		szReport[] = sz; \
				if (!MsoFReportSzProcVar (vszAssertFile, __LINE__, szReport, a)) \
					MsoDebugBreakInline(); \
				} \
			else

#define ReportSz2If(f, sz, a, b) \
			if (!(f)) \
				{ \
				static const char 		szReport[] = sz; \
				if (!MsoFReportSzProcVar (vszAssertFile, __LINE__, szReport, a, b)) \
					MsoDebugBreakInline(); \
				} \
			else

#define ReportSz3If(f, sz, a, b, c) \
			if (!(f)) \
				{ \
				static const char 		szReport[] = sz; \
				if (!MsoFReportSzProcVar (vszAssertFile, __LINE__, szReport, a, b, c)) \
					MsoDebugBreakInline(); \
				} \
			else

#if 0 //$[VSMSO]
#define MsoEnableFeatureReportsWC(fEnable) \
			MsoEnableFeatureReports(fDebugFilterWc, fEnable)

#define MsoFFeatureReportsEnabledWC() \
			MsoFFeatureReportsEnabled(fDebugFilterWc)

#define MsoReportWCSz(szMsg) \
			do { \
				if (!MsoFFeatureReportSzProcVar(0, fDebugFilterWc, \
												vszAssertFile, __LINE__, \
												"%s", (const CHAR *)(szMsg))) \
				MsoDebugBreakInline(); \
				} while (0)

#define MsoReportWCSzTag(szMsg, tag) \
			do { \
				if (!MsoFFeatureReportSzProcVar(tag, fDebugFilterWc, \
												vszAssertFile, __LINE__, \
												"%s", (const CHAR *)(szMsg))) \
				MsoDebugBreakInline(); \
				} while (0)

#define MsoReportWCFSz(f, szMsg) \
			do { \
				if (!(f) \
					 &&  !MsoFFeatureReportSzProcVar(0, fDebugFilterWc, \
					 								 vszAssertFile, __LINE__, \
													 (const CHAR *)(szMsg))) \
					MsoDebugBreakInline(); \
				} while (0)

#define MsoReportWCFSzTag(f, szMsg, tag) \
			do { \
				if (!(f) \
					 &&  !MsoFFeatureReportSzProcVar(tag, fDebugFilterWc, \
					 								 vszAssertFile, __LINE__, \
													 (const CHAR *)(szMsg))) \
					MsoDebugBreakInline(); \
				} while (0)

#define MsoReportWCFSz1(f, szMsg, a) \
			do { \
				if (!(f) \
					 &&  !MsoFFeatureReportSzProcVar(0, fDebugFilterWc, \
					 								 vszAssertFile, __LINE__, \
													 (const CHAR *)(szMsg), \
													 (a))) \
					MsoDebugBreakInline(); \
				} while (0)

#define MsoReportWCFSz1Tag(f, szMsg, a, tag) \
			do { \
				if (!(f) \
					 &&  !MsoFFeatureReportSzProcVar(tag, fDebugFilterWc, \
					 								 vszAssertFile, __LINE__, \
													 (const CHAR *)(szMsg), \
													 (a))) \
					MsoDebugBreakInline(); \
				} while (0)
#endif //$[VSMSO]
				
#define CommSz0(sz) \
			do { \
			static const char 		szComm[] = sz; \
			MsoCommSzVar(szComm); \
			} while (0)

#define CommSz1(sz, a) \
			do { \
			static const char 		szComm[] = sz; \
			MsoCommSzVar(szComm, a); \
			} while (0)

#define CommSz2(sz, a, b) \
			do { \
			static const char 		szComm[] = sz; \
			MsoCommSzVar(szComm, a, b); \
			} while (0)

#define CommSz3(sz, a, b, c) \
			do { \
			static const char 		szComm[] = sz; \
			MsoCommSzVar(szComm, a, b, c); \
			} while (0)

#define CommSz4(sz, a, b, c, d) \
			do { \
			static const char 		szComm[] = sz; \
			MsoCommSzVar(szComm, a, b, c, d); \
			} while (0)

#define CommSz5(sz, a, b, c, d, e) \
			do { \
			static const char 		szComm[] = sz; \
			MsoCommSzVar(szComm, a, b, c, d, e); \
			} while (0)

#else /* !DEBUG */

#define debugvar(a)
#define Debug(e)
#define DebugElse(s, t) t
#define debuglog0(wLevel, grfFilter, sz)
#define debuglog1(wLevel, grfFilter, sz, a)
#define debuglog2(wLevel, grfFilter, sz, a, b)
#define debuglog3(wLevel, grfFilter, sz, a, b, c)
#define debuglog4(wLevel, grfFilter, sz, a, b, c, d)
#define debuglog5(wLevel, grfFilter, sz, a, b, c, d, e)
#define debuglog6(wLevel, grfFilter, sz, a, b, c, d, e, f)
#define debuglogPch(wLevel, grfFilter, pch, cch)
#define debuglogPwch(wLevel, grfFilter, pwch, cwch)

#define ReportSz(sz)
#define MsoReportSz(sz)
#define ReportSz0If(f, sz)
#define ReportSz1If(f, sz, a)
#define ReportSz2If(f, sz, a, b)
#define ReportSz3If(f, sz, a, b, c)

#define MsoEnableFeatureReportsWC(fEnable) (0)
#define MsoFFeatureReportsEnabledWC() (FALSE)
#define MsoReportWCSz(sz)
#define MsoReportWCSzTag(sz, tag)
#define MsoReportWCFSz(f, sz)
#define MsoReportWCFSzTag(f, sz, tag)
#define MsoReportWCFSz1(f, sz, a)
#define MsoReportWCFSz1Tag(f, sz, a, tag)

#define CommSz0(sz)
#define CommSz1(sz, a)
#define CommSz2(sz, a, b)
#define CommSz3(sz, a, b, c)
#define CommSz4(sz, a, b, c, d)
#define CommSz5(sz, a, b, c, d, e)

#endif /* DEBUG */

// Generic debug log macros - WARNING: THIS WILL BE SEEN BY EVERYONE
#define debug0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterAll, sz) 
#define debug1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterAll, sz, a) 
#define debug2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterAll, sz, a, b) 
#define debug3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterAll, sz, a, b, c) 
#define debug4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterAll, sz, a, b, c, d) 
#define debug5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterAll, sz, a, b, c, d, e) 
#define debug6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterAll, sz, a, b, c, d, e, f) 
#define debugPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterAll, pch, cch)
#define debugPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterAll, pwch, cwch)

// Application Event Monitor debug log macros
#define debugEM0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterEm, sz) 
#define debugEM1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterEm, sz, a) 
#define debugEM2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterEm, sz, a, b) 
#define debugEM3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterEm, sz, a, b, c) 
#define debugEM4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterEm, sz, a, b, c, d) 
#define debugEM5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterEm, sz, a, b, c, d, e) 
#define debugEM6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterEm, sz, a, b, c, d, e, f) 
#define debugEMPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterEm, pch, cch)
#define debugEMPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterEm, pwch, cwch)

// Application Rule Engine debug log macros
#define debugRL0(wLevel, sz) \
 			debuglog0(wLevel, vlpruls->grfDebugLogFilter, sz) 
#define debugRL1(wLevel, sz, a) \
 			debuglog1(wLevel, vlpruls->grfDebugLogFilter, sz, a) 
#define debugRL2(wLevel, sz, a, b) \
 			debuglog2(wLevel, vlpruls->grfDebugLogFilter, sz, a, b) 
#define debugRL3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, vlpruls->grfDebugLogFilter, sz, a, b, c) 
#define debugRL4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, vlpruls->grfDebugLogFilter, sz, a, b, c, d) 
#define debugRL5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, vlpruls->grfDebugLogFilter, sz, a, b, c, d, e) 
#define debugRL6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel,vlpruls->grfDebugLogFilter, sz, a, b, c, d, e, f) 
#define debugRLPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, vlpruls->grfDebugLogFilter, pch, cch)
#define debugRLPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, vlpruls->grfDebugLogFilter, pwch, cwch)

// Web Client debug log macros
#define debugWC0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterWc, sz) 
#define debugWC1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterWc, sz, a) 
#define debugWC2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterWc, sz, a, b) 
#define debugWC3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterWc, sz, a, b, c) 
#define debugWC4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterWc, sz, a, b, c, d) 
#define debugWC5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterWc, sz, a, b, c, d, e) 
#define debugWC6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterWc, sz, a, b, c, d, e, f) 
#define debugWCPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterWc, pch, cch)
#define debugWCPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterWc, pwch, cwch)

// Active Clip Board (C & C) debug log macros
#define debugACB0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterAcb, sz) 
#define debugACB1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterAcb, sz, a) 
#define debugACB2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterAcb, sz, a, b) 
#define debugACB3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterAcb, sz, a, b, c) 
#define debugACB4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterAcb, sz, a, b, c, d) 
#define debugACB5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterAcb, sz, a, b, c, d, e) 
#define debugACB6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterAcb, sz, a, b, c, d, e, f) 
#define debugACBPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterAcb, pch, cch)
#define debugACBPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterAcb, pwch, cwch)

// Toolbar debug log macros
#define debugEC0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterEc, sz) 
#define debugEC1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterEc, sz, a) 
#define debugEC2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterEc, sz, a, b) 
#define debugEC3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterEc, sz, a, b, c) 
#define debugEC4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterEc, sz, a, b, c, d) 
#define debugEC5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterEc, sz, a, b, c, d, e) 
#define debugEC6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterEc, sz, a, b, c, d, e, f) 
#define debugECPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterEc, pch, cch)
#define debugECPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterEc, pwch, cwch)

// MSO Internal Event Monitor debug log macros
#define debugMsoEM0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterMsoEm, sz) 
#define debugMsoEM1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterMsoEm, sz, a) 
#define debugMsoEM2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterMsoEm, sz, a, b) 
#define debugMsoEM3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterMsoEm, sz, a, b, c) 
#define debugMsoEM4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterMsoEm, sz, a, b, c, d) 
#define debugMsoEM5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterMsoEm, sz, a, b, c, d, e) 
#define debugMsoEM6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterMsoEm, sz, a, b, c, d, e, f) 
#define debugMsoEMPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterMsoEm, pch, cch)
#define debugMsoEMPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterMsoEm, pwch, cwch)

// MSO Internal Event Monitor debug log macros
#define debugLS0(wLevel, sz) \
 			debuglog0(wLevel, fDebugFilterLS, sz) 
#define debugLS1(wLevel, sz, a) \
 			debuglog1(wLevel, fDebugFilterLS, sz, a) 
#define debugLS2(wLevel, sz, a, b) \
 			debuglog2(wLevel, fDebugFilterLS, sz, a, b) 
#define debugLS3(wLevel, sz, a, b, c) \
 			debuglog3(wLevel, fDebugFilterLS, sz, a, b, c) 
#define debugLS4(wLevel, sz, a, b, c, d) \
 			debuglog4(wLevel, fDebugFilterLS, sz, a, b, c, d) 
#define debugLS5(wLevel, sz, a, b, c, d, e) \
 			debuglog5(wLevel, fDebugFilterLS, sz, a, b, c, d, e) 
#define debugLS6(wLevel, sz, a, b, c, d, e, f) \
 			debuglog6(wLevel, fDebugFilterLS, sz, a, b, c, d, e, f) 
#define debugLSPch(wLevel, pch, cch)\
 			debuglogPch(wLevel, fDebugFilterLS, pch, cch)
#define debugLSPwch(wLevel, pwch, cwch)\
 			debuglogPwch(wLevel, fDebugFilterLS, pwch, cwch)

#ifdef DEBUG
#define DebugStandalone(e)		e
#else
#define DebugStandalone(e)
#endif

MSOEXTERN_C_END		// ****************** End extern "C" *********************

#endif // !MSODBGLG_H
