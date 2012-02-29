#pragma once

/****************************************************************************
	msotip.h

	Owner: rolandr
	Copyright (c) 1999 Microsoft Corporation

	This file contains the exported interfaces and declarations for
	Office Tooltips and related components.
****************************************************************************/

#ifndef MSOTT_H
#define MSOTT_H

#include "msosdm.h"

/****************************************************************************
	Tooltip structures and constants
***************************************************************** rolandr **/

// Size of tooltip buffers and maximum number of chars
#define TOOLTIP_BUFSIZE (1024)
#define MAX_TT_CHARS    (TOOLTIP_BUFSIZE-2)

// The members of this enum can be passed as the fourth parameter of
// MsoFShowTooltipEx. They determine the way the wtz argument is
// interpreted.
enum msofmtTooltipFormat {
	msofmtNone,			// no formatting
	msofmtString,     // determine formatting from first character in string
	msofmtBest,       // determine width/height based on golden cut
	msofmtManual,     // width is determined by longest substring
	msofmtFixedWidth, // width is determined by specified rectangle

	msofmtMax,        // from here on: flages

	msofmtEnforceMaxWidth = 0x4000, // Strictly enforce a maximum width
	msofmtDontCorrectBadWraps = 0x8000 // Dont attempt to improve the layout
};

// These flages determine the conversion behaviour of MsoCchPrepareTooltipUrl
enum ttcTooltipConversion
{
	ttcDontEscapeEscapes           = 0x01,
	ttcDontBreakBeforeDelimiters   = 0x02,
	ttcDoBreakAfterDelimiters      = 0x04,
	ttcDontProtectSpaces           = 0x08,
	ttcDontUnderline               = 0x10,
	ttcForceRealMaxLenBreaks       = 0x20,
};

// Tooltip APIs

MSOAPI_(void) MsoSetupTooltipText(HDC hdc, const WCHAR *wtzText, RECT *prc, int fmt);
MSOAPI_(void) MsoDrawTooltipText(HDC hdc, const WCHAR *wtzText, RECT *prc);
MSOAPI_(void *) MsoPsqlNew(void);
MSOAPI_(void) MsoDeletePsql(void *pv);
MSOAPI_(void) MsoSetupTooltipTextPsql(void *psql, HDC hdc, const WCHAR *wtzText, RECT *prc, int fmt);
MSOAPI_(void) MsoDrawTooltipTextPsql(void *psql, HDC hdc, const WCHAR *wtzText, RECT *prc);

// Set/Get the tooltip for an SDM control
// [in sdmapiw.c]
MSOAPI_(void) MsoSetSDMTooltip(TMC tmc, const WCHAR *wzTip, int fmt);
MSOAPIX_(LPCWSTR) MsoWzGetSDMTooltip(TMC tmc);
// Show a tooltip with formatting and for a particular parent window
// [in tbutil.cpp]
MSOAPIX_(BOOL) MsoFShowTooltipEx(const WCHAR *wtz, int tt, LONG_PTR lid, int fmt, HWND hwnd);
// Prepare a url or file for display in a tooltip
MSOAPIX_(int)  MsoCchPrepareTooltipUrl(const WCHAR *wzSrc, _Out_cap_(cMax) WCHAR *wzDst, int cMax, int ttc, int iBreak);
// Forbid
MSOAPI_(BOOL) MsoDisableTooltips(BOOL fDisable);

#endif // MSOTT_H
