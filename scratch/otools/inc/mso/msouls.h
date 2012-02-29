#pragma once

/****************************************************************************
	MsoULS.h

	Owner: TCoon
	Copyright (c) 2001 Microsoft Corporation

	This file contains the exported interfaces and declarations for
	Office Logging.
****************************************************************************/

#ifndef MSOULS_H
#define MSOULS_H

#include "msoulscat.h"     // Include Developer defined Categories
#if 0 //$[VSMSO]
#ifdef DEBUG
// HACK: This assumes the UTIL project includes are in a "util" subdirectory.
#include "util\operfinc.h" // Include Developer defined PerfMon counters
#endif
#endif //$[VSMSO]


/****************************************************************************
	Logging enums and constants
******************************************************************* TCOON **/

// The following Category and Event IDs are the default until the event call
// has been tagged.
extern const WORD msoulscatIdUnknown; // Unknown Category ID
extern const DWORD msoulsevtIdUnknown; // Unknown Event ID

// Logging Levels
typedef enum _msoulslevel
{
	msoulslUnassigned = 0,         // No logging
	msoulslCriticalEvent = 1,      // Major problem (crash)is about to occur Reserved
	msoulslException = 4,          // Major problem (crash)is about to occur Reserved
	msoulslAssert = 7,             // Assert fired Reserved
	msoulslUnexpected = 10,        // Unexpected condition has happened
	msoulslMonitorable = 15,       // Anything else that should be required by default.
	msoulslHigh = 20,              // High level User action/major funcitonal area
	msoulslMedium = 50,            // Mid level functional area
	msoulslVerbose = 100,          // Excruciating detail
} MSOULSLEVEL;

#ifdef DEBUG
// PCVO (Performance Log Value Operation) constants
#define msopcvoSet       0
#define msopcvoIncrement 1
#define msopcvoSetIfMax  2
#define msopcvoSetIfMin  3
#define msopcvoMax       4
#endif

/****************************************************************************
	Global DLL API's
****************************************************************************/
/****************************************************************************
	Initalization API's which Developers will call once per process	
****************************************************************************/
MSOEXTERN_C_BEGIN
/****************************************************************************
	Logging API's which Developers will call
	These will be transformed by the Event Tagging mechanism into Tagged Logging API's.
	No calls to these should appear an any build machine released builds.
****************************************************************************/

#if DEBUG
MSOAPI_(void) MsoULSTrace(MSOULSLEVEL bLevel);
MSOAPI_(void) MsoULSTraceTag(DWORD idEvent, WORD wCat, MSOULSLEVEL bLevel);
#else
#define MsoULSTrace(bLevel) 
#define MsoULSTraceTag(idEvent, wCat, bLevel) 
#endif

MSOCDECLAPI_(void) MsoULSTraceWz(MSOULSLEVEL bLevel, const WCHAR *wzFormat, ...);
MSOCDECLAPI_(void) MsoULSTraceWzTag(DWORD idEvent, WORD wCat, MSOULSLEVEL bLevel, const WCHAR *wzFormat, ...);

/****************************************************************************
	PerfMon Counter API's
****************************************************************************/
#if DEBUG
MSOAPI_(HRESULT) MsoULSPerfmon(int pcid, DWORD dwValue, int pcvo);
#define MsoULSPerfmonSet(pcid, dwNew) MsoULSPerfmon(pcid, dwNew, msopcvoSet)
#define MsoULSPerfmonEvent(pcid) MsoULSPerfmonIncrement(pcid, 1)
#define MsoULSPerfmonIncrement(pcid, dwAdd) MsoULSPerfmon(pcid, dwAdd, msopcvoIncrement)
#endif

MSOEXTERN_C_END
#endif // MSOULS_H
