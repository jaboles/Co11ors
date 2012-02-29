#pragma once

/****************************************************************************
	MsoTheme.h

	Owner: AndChang
	Copyright (c) 2001 Microsoft Corporation

	MSO clients can include this file and get the Windows XP theming API at their fingertips,
	as well as some handy help functions for deploying themed applications.
****************************************************************************/
#ifndef MSOTHEME_H
#define MSOTHEME_H

// Because we're not compiling with the _WIN32_WINNT=0x0501
#ifndef WM_THEMECHANGED
#define WM_THEMECHANGED                 0x031A
#endif 

#include <uxtheme.h>
#include <tmschema.h>
#include "vscolors.h"

// ----------------------------------------------------------------------------
// Enumeration
// YOU ADD SOMETHING HERE, MAKE SURE TO ADD IT TO THEME.CPP
// - MsoThemeGetHTheme() 
// - MsoThemeDeinit() 
// - struct MSOHTHEMES
enum ENUM_MSOHTHEMES { 
	msohtInvalid = -1,
	msohtToolBar = 0, 
	msohtComboBox, 
	msohtSpin, 
	msohtProgress, 
	msohtScrollBar,
	msohtButton, 
	msohtExplorerBar,
	msohtHeader,
	msohtListView,
	msohtEdit,
	msohtTab,
	msohtWindow,
	msohtGlobals,
	msohtTreeview,
	msohtCount // not a valid theme
	};

// ----------------------------------------------------------------------------
// Stuff for getting the current theme name
extern const WCHAR* const MSOVWZUNTHEMED;
typedef struct
{
	WCHAR pwzThemeFileName[MAX_PATH];
	WCHAR pwzColorBuff[MAX_PATH];
	WCHAR pwzSizeBuff[MAX_PATH];
} MSOTHEMENAME;

// struct with info needed to set up an offscreen buffer.
typedef struct
{
	HDC hdcMem;
	HDC hdcScreen;
	HBITMAP hbmpScreen;
	HBITMAP hbmpSav;
	HFONT hfontSav; 
	POINT   ptOrgSav;
	RECT    rc;
} MSOTHEMEOFFSCREENBUF;

// ----------------------------------------------------------------------------
// Constants 
// flags for MsoDrawThemeTabButtonBack
#define MSOTHEMETABBUTTON_CURRENT       0x00000001
#define MSOTHEMETABBUTTON_FIRST         0x00000002
#define MSOTHEMETABBUTTON_LAST          0x00000004
#define MSOTHEMETABBUTTON_HASMOUSE      0x00000008
#define MSOTHEMETABBUTTON_DISABLED      0x00000010
#define MSOTHEMETABBUTTON_XEXPAND       0x00000020
#define MSOTHEMETABBUTTON_YEXPAND       0x00000040
#define MSOTHEMETABBUTTON_MULTIPLEROWS  0x00000080

// ----------------------------------------------------------------------------
// Exported APIs
MSOAPI_(BOOL) MsoFWhistlerOrLater();
MSOAPI_(BOOL) MsoFWhistlerSP1OrLater();
// Tell MSO that you have a manifest. 
MSOAPI_(void) MsoThemeHaveManifest(BOOL);
// Tell MSO you want theming on Windows XP RTM. Don't call this before checking with AndChang.
MSOAPI_(void) MsoThemeXPRTM();
MSOAPI_(BOOL) MsoThemeFActiveInSystem();
MSOAPI_(BOOL) MsoThemeFActive(); // Are themes active?
MSOAPI_(HTHEME) MsoThemeGetHTheme(enum ENUM_MSOHTHEMES); // Returns the HTHEME needed for calling Theming APIs
MSOAPI_(MSOTHEMENAME*) MsoThemeGetCurrentThemeName();
MSOAPI_(VSACTIVETHEME) MsoThemeGetCurrentTheme();
MSOAPI_(int) MsoThemeGetBorderWidth();
MSOAPI_(int) MsoThemeGetMarginDelta(HTHEME hTheme, HDC hdc, int iPartId, int iStateId, int iPropId, int nUnthemedMargin);
MSOAPI_(void) MsoDrawThemeBorder(HDC hdc, RECT* prc);  // uses Listview hTheme
MSOAPI_(void) MsoDrawThemeBorderInset(HDC hdc, RECT* prc, int nInset);
MSOAPI_(void) MsoDrawThemeBorderEx(HTHEME hTheme, HDC hdc, int iPart, int iState, const RECT *prc, int nInset);
MSOAPI_(void) MsoDrawThemeNCPaintBorder(HWND hwnd, int nInset);
MSOAPI_(void) MsoDrawThemeGroupLine(HDC hdc, RECT* prc);
MSOAPI_(void) MsoThemeOnThemeChanged();

	// For handling themed tab sheets
MSOAPI_(void) MsoDrawThemeTabButtonBack(HDC hdc, int xLeft, int yTop,
		int xRight, int yBottom, int xRightTabBody, DWORD grfTabStyles);
MSOAPI_(VOID) MsoDrawThemedTabSheetRectForGcc(BOOL fSDM, HDC hdc, HWND hwnd, RECT* prect); // defined in sdmtab.c
// Win32 tab sheets
MSOAPI_(void) MsoSetWsExTransparent(HWND hwndCtl, BOOL fTransparent);
MSOAPI_(HRESULT) MsoDrawThemeTabSheetRect(HDC hdc, RECT* prcTabSheet, RECT* prcClip);
MSOAPI_(HRESULT) MsoDrawThemeTabSheetBkgnd(HDC hdc, HWND hwndTabControl, HWND hwnd);
MSOAPI_(HRESULT) MsoDrawThemeTabSheetFrame(HDC hdc, HWND hwndTabControl, HWND hwnd);
MSOAPI_(void) MsoThemeShowAcceleratorsAndFocus(HWND hwnd);
	// Outlook needs this for inspector windows.
MSOAPI_(int) MsoThemeGetShadowWidth();
MSOAPI_(int) MsoThemeGetTabSheetExpansion(); 

// SDM tabsheets
MSOAPI_(BOOL) MsoThemeFSDMTabSwitcher(); // defined in sdmtab.c
MSOAPI_(VOID) MsoDrawThemedSDMTabSheetRect(HDC hdc, RECT* prect);
MSOAPI_(VOID) MsoDrawThemedSDMTabSheetFrame(HDC hdc, RECT* prect);
MSOAPI_(BOOL) MsoThemeFSDMRectOnTabSheet(RECT *prect);

// including sdm headers here would be a bad idea...
#ifndef TMC
typedef UINT_PTR	TMC;			// Item codes. 
#endif // TMC
MSOAPI_(void) MsoThemeSDMDisableMouseOver(TMC tmc); // defined in sdmtab.c


	// For getting the padding needed for pushbuttons.
MSOAPI_(int) MsoThemeGetContentMargin(enum ENUM_MSOHTHEMES emsot, int iPartId, int iStateId);
MSOAPI_(void) MsoThemeDrawTransparentIcon(HDC hdc, int tcid, RECT* prc, COLORREF crBkgnd, BOOL fDisabled);

MSOAPI_(BOOL) MsoThemeFDrawScrollGripper(HDC hdc, RECT* prcThumb, BOOL fVert);

// For drawing into an offscreen bitmap
MSOAPI_(BOOL) MsoThemeInitOffscreenBuffer(MSOTHEMEOFFSCREENBUF* pMsotob, HDC hdcScreen, RECT* prc, OPTIONAL HFONT hFont);
MSOAPI_(void) MsoThemeFlushOffscreenBuffer(MSOTHEMEOFFSCREENBUF* pMsotob);

// ----------------------------------------------------------------------------
// Activation Context stuff
#if (_WIN32_WINNT < 0x0500) && (_WIN32_FUSION < 0x0100) && !ISOLATION_AWARE_ENABLED
// Copied from winbase.h for apps that don't compile with (_WIN32_WINNT >= 0x0500), e.g. Excel, Cag, Publisher, Xdocs as of 3/19/2002
#define ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID    (0x00000001)
#define ACTCTX_FLAG_LANGID_VALID                    (0x00000002)
#define ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID        (0x00000004)
#define ACTCTX_FLAG_RESOURCE_NAME_VALID             (0x00000008)
#define ACTCTX_FLAG_SET_PROCESS_DEFAULT             (0x00000010)
#define ACTCTX_FLAG_APPLICATION_NAME_VALID          (0x00000020)
#define ACTCTX_FLAG_SOURCE_IS_ASSEMBLYREF           (0x00000040)
#define ACTCTX_FLAG_HMODULE_VALID                   (0x00000080)

typedef struct tagACTCTXA {
    ULONG       cbSize;
    DWORD       dwFlags;
    LPCSTR      lpSource;
    USHORT      wProcessorArchitecture;
    LANGID      wLangId;
    LPCSTR      lpAssemblyDirectory;
    LPCSTR      lpResourceName;
    LPCSTR      lpApplicationName;
    HMODULE     hModule;
} ACTCTXA, *PACTCTXA;
typedef struct tagACTCTXW {
    ULONG       cbSize;
    DWORD       dwFlags;
    LPCWSTR     lpSource;
    USHORT      wProcessorArchitecture;
    LANGID      wLangId;
    LPCWSTR     lpAssemblyDirectory;
    LPCWSTR     lpResourceName;
    LPCWSTR     lpApplicationName;
    HMODULE     hModule;
} ACTCTXW, *PACTCTXW;
#ifdef UNICODE
typedef ACTCTXW ACTCTX;
typedef PACTCTXW PACTCTX;
#else // !UNICODE
typedef ACTCTXA ACTCTX;
typedef PACTCTXA PACTCTX;
#endif // !UNICODE

typedef const ACTCTXA *PCACTCTXA;
typedef const ACTCTXW *PCACTCTXW;
#ifdef UNICODE
typedef ACTCTXW ACTCTX;
typedef PCACTCTXW PCACTCTX;
#else // !UNICODE
typedef ACTCTXA ACTCTX;
typedef PCACTCTXA PCACTCTX;
#endif // !UNICODE
#endif //  (_WIN32_WINNT < 0x0500) && (_WIN32_FUSION < 0x0100) && !ISOLATION_AWARE_ENABLED

MSOAPI_(BOOL) MsoActivateActCtx(HANDLE hActCtx, ULONG_PTR *lpCookie);
MSOAPI_(BOOL) MsoDeactivateActCtx(DWORD dwFlags, ULONG_PTR ulCookie);
MSOAPI_(HANDLE) MsoCreateActCtx(PACTCTXW pActCtx);
MSOAPI_(VOID) MsoReleaseActCtx(HANDLE hActCtx);
MSOAPI_(BOOL) MsoQueryActCtxW(DWORD dwFlags, HANDLE hActCtx, PVOID pvSubInstance,  
	ULONG ulInfoClass, PVOID pvBuffer, SIZE_T cbBuffer, SIZE_T *pcbWrittenOrRequired);

MSOAPI_(HWND) MsoUnthemedCreateWindowExA(DWORD dwExStyle, LPCSTR lpClassName, LPCSTR lpWindowName, DWORD dwStyle, int x,int y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam);
MSOAPI_(HWND) MsoUnthemedCreateWindowExW(DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName, DWORD dwStyle, int x,int y, int nWidth, int nHeight, HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam);

// ----------------------------------------------------------------------------
// Wrap the Theming APIs for MSO clients to use
MSOAPI_(HRESULT) MsoEnableThemeDialogTexture(HWND, DWORD);
MSOAPI_(HRESULT) MsoEnableTheming(BOOL fEnable);

MSOAPI_(HRESULT) MsoCloseThemeData(HTHEME);
MSOAPI_(HTHEME) MsoOpenThemeData(HWND, LPCWSTR);

MSOAPI_(HRESULT) MsoDrawThemeBackground(HTHEME, HDC, int, int, const RECT*, const RECT*);
MSOAPI_(HRESULT) MsoDrawThemeEdge(HTHEME, HDC, int, int, const RECT*, UINT, UINT, RECT*);
MSOAPI_(HRESULT) MsoDrawThemeIcon(HTHEME, HDC, int, int, const RECT*, HIMAGELIST, int);
MSOAPI_(HRESULT) MsoDrawThemeParentBackground(HWND, HDC, OPTIONAL RECT*);
MSOAPI_(HRESULT) MsoDrawThemeText(HTHEME, HDC, int, int, LPCWSTR, int, DWORD, DWORD, const RECT*);

MSOAPI_(DWORD) MsoGetThemeAppProperties();
MSOAPI_(HRESULT) MsoGetThemeBackgroundContentRect(HTHEME, OPTIONAL HDC, int, int,  const RECT*, RECT*);
MSOAPI_(HRESULT) MsoGetThemeBackgroundExtent(HTHEME, OPTIONAL HDC, int, int, const RECT*, OUT RECT*);
MSOAPI_(HRESULT) MsoGetThemeBackgroundRegion(HTHEME, OPTIONAL HDC, int, int, const RECT*, OUT HRGN*);
MSOAPI_(HRESULT) MsoGetThemeBool(HTHEME, int, int, int, OUT BOOL*);
MSOAPI_(HRESULT) MsoGetThemeColor(HTHEME, int, int, int, COLORREF*);
MSOAPI_(HRESULT) MsoGetThemeDocumentationProperty(LPCWSTR, LPCWSTR, _Out_cap_(cchMaxValChars) OUT LPWSTR pszValueBuff, int cchMaxValChars);
MSOAPI_(HRESULT) MsoGetThemeEnumValue(HTHEME, int, int, int, OUT int*);
MSOAPI_(HRESULT) MsoGetThemeFilename(HTHEME, int, int, int, _Out_cap_(cchMaxBuffChars) OUT LPWSTR pszThemeFileName, int cchMaxBuffChars);
MSOAPI_(HRESULT) MsoGetThemeFont(HTHEME, OPTIONAL HDC, int, int, int, LOGFONT*);
MSOAPI_(HRESULT) MsoGetThemeInt(HTHEME, int, int, int, OUT int*);
MSOAPI_(HRESULT) MsoGetThemeIntList(HTHEME, int, int, int, OUT INTLIST*);
MSOAPI_(HRESULT) MsoGetThemePartSize(HTHEME, HDC, int, int, RECT*, enum THEMESIZE, OUT SIZE*);
MSOAPI_(HRESULT) MsoGetThemePosition(HTHEME, int, int, int, OUT POINT*);
MSOAPI_(HRESULT) MsoGetThemePropertyOrigin(HTHEME, int, int, int, OUT enum PROPERTYORIGIN*);
MSOAPI_(HRESULT) MsoGetThemeRect(HTHEME, int, int, int, OUT RECT *);
MSOAPI_(HRESULT) MsoGetThemeString(HTHEME, int, int, int, _Out_cap_(cchMaxBuffChars) OUT LPWSTR pszBuff, int cchMaxBuffChars);
MSOAPI_(BOOL) MsoGetThemeSysBool(HTHEME, int);
MSOAPI_(HRESULT) MsoGetThemeSysFont(HTHEME, int, OUT LOGFONT*);
MSOAPI_(COLORREF) MsoGetThemeSysColor(HTHEME, int);
MSOAPI_(HBRUSH) MsoGetThemeSysColorBrush(HTHEME, int);
MSOAPI_(HRESULT) MsoGetThemeSysInt(HTHEME, int, int*);
MSOAPI_(int) MsoGetThemeSysSize(HTHEME, int);
MSOAPI_(HRESULT) MsoGetThemeSysString(HTHEME, int ,  _Out_cap_(cchMaxStringChars) OUT LPWSTR pszStringBuff, int cchMaxStringChars);
MSOAPI_(HRESULT) MsoGetThemeTextExtent(HTHEME, HDC, int, int, LPCWSTR, int, DWORD, OPTIONAL const RECT*, RECT*);
MSOAPI_(HRESULT) MsoGetThemeTextMetrics(HTHEME, OPTIONAL HDC, int, int, OUT TEXTMETRIC*);
MSOAPI_(HRESULT) MsoGetThemeMetric (HTHEME, HDC, int, int, int, int*);
MSOAPI_(HRESULT) MsoGetThemeMargins(HTHEME, HDC, int, int, int, RECT*, MARGINS*);
MSOAPI_(HTHEME) MsoGetWindowTheme(HWND);

MSOAPI_(HRESULT) MsoHitTestThemeBackground(HTHEME, OPTIONAL HDC, int, int, DWORD, const RECT*, OPTIONAL HRGN, POINT, OUT WORD*);
MSOAPI_(BOOL) MsoIsThemePartDefined(HTHEME, int, int);
MSOAPI_(BOOL) MsoIsThemeBackgroundPartiallyTransparent(HTHEME, int, int);

MSOAPI_(void) MsoSetThemeAppProperties(DWORD);
MSOAPI_(HRESULT) MsoSetWindowTheme(HWND, LPCWSTR, LPCWSTR);


#endif // !MSOTHEME_H

