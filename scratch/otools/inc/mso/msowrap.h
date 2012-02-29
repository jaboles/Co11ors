#pragma once

/****************************************************************************
	msowrap.h

	Owner: JBelt
 	Copyright (c) 1996 Microsoft Corporation

	Wrappers for Unicode functions that don't exist on Win95.

	Ideally, all the wrappers we use throughout Office would live in this
	file. But nobody is scheduled to do that. Please add any new wrappers
	you need in this file, and move old ones in here as you go.
****************************************************************************/
#ifndef MSOWRAP_H
#define MSOWRAP_H

#include <shellapi.h>
#include <commctrl.h>
#include <shlobj.h>
#include <wininet.h>
#include <winuser.h>

/*---------------------------------------------------------------------------
	MsoGetWindowsDirectoryW
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(UINT) MsoGetWindowsDirectoryW(_Out_cap_(cch) WCHAR *wz, UINT cch);


/*---------------------------------------------------------------------------
	MsoGetDriveTypeW

	UNICODE wrapper for GetDriveType() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(UINT) MsoGetDriveTypeW(const WCHAR *wzRoot);
MSOAPI_(BOOL) MsoGetComputerNameW(_Out_cap_(*nSize) WCHAR *lpBuffer, LPDWORD nSize);


// Return values from MsoGetFileAttributesW().  All except msofaError
// could be bitwise-OR'ed together.
#define msofaReadOnly FILE_ATTRIBUTE_READONLY
#define msofaHidden FILE_ATTRIBUTE_HIDDEN
#define msofaSystem FILE_ATTRIBUTE_SYSTEM
#define msofaDirectory FILE_ATTRIBUTE_DIRECTORY
#define msofaArchive FILE_ATTRIBUTE_ARCHIVE
#define msofaNormal FILE_ATTRIBUTE_NORMAL
#define msofaTemporary FILE_ATTRIBUTE_TEMPORARY
#define msofaCompressed FILE_ATTRIBUTE_COMPRESSED
#define msofaOffline FILE_ATTRIBUTE_OFFLINE
#define msofaError 0xFFFFFFFF


/*---------------------------------------------------------------------------
	MsoGet|SetFileAttributesW

	UNICODE wrapper for Get|SetFileAttributes() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(DWORD) MsoGetFileAttributesW(const WCHAR *wzPath);
MSOAPI_(DWORD) MsoSetFileAttributesW(const WCHAR *wzPath, DWORD dw); /*KBROWN*/

/*-----------------------------------------------------------------------------
	MsoGetDiskFreeSpaceW

	UNICODE wrapper for GetDiskFreeSpace() system call.
------------------------------------------------------------------ DARRELLA -*/
MSOAPI_(BOOL) MsoGetDiskFreeSpaceW(const WCHAR *lpRootPathName,
		LPDWORD lpSectorsPerCluster,
		LPDWORD lpBytesPerSector,
		LPDWORD lpNumberOfFreeClusters,
		LPDWORD lpTotalNumberOfClusters);

/*-----------------------------------------------------------------------------
	MsoGetVolumeInformationW

	UNICODE wrapper for GetVolumeInformationW() system call.
------------------------------------------------------------------ DARRELLA -*/
MSOAPI_(BOOL) MsoGetVolumeInformationW(LPCWSTR  lpRootPathName,
		_Out_opt_cap_(nVolumeNameSize) 
		LPWSTR lpVolumeNameBuffer,
		DWORD nVolumeNameSize,
		LPDWORD  lpVolumeSerialNumber,
		LPDWORD lpMaximumComponentLength,
		LPDWORD lpFileSystemFlags,
		_Out_opt_cap_(nFileSystemNameSize)
		LPWSTR lpFileSystemNameBuffer,
		DWORD nFileSystemNameSize);


/*-----------------------------------------------------------------------------
	MsoGetPrivateProfileString
	MsoGetProfileString
	MsoWritePrivateProfileString
	MsoGetPrivateProfileInt

	UNICODE wrapper for Private Profile APIs

	if fUTF8 is set, then use UnicodeToUTF8() or UTF8ToUnicode to convert
	and then use the ANSI functions.

-------------------------------------------------------------------- KBROWN -*/
MSOAPI_(DWORD) MsoGetPrivateProfileString(const WCHAR *lpAppName, const WCHAR *lpKeyName, int fUTF8, const WCHAR *lpDefault, _Out_cap_(cchSize) WCHAR *lpReturnedString, DWORD cchSize, const WCHAR *lpFileName);
MSOAPI_(BOOL)  MsoWritePrivateProfileString(const WCHAR *lpAppName, const WCHAR *lpKeyName, int fUTF8, const WCHAR *lpString, const WCHAR *lpFileName);

#define MsoGetPrivateProfileStringW(lpAppName, lpKeyName, lpDefault, lpReturnedString, cchSize, lpFileName) \
		MsoGetPrivateProfileString(lpAppName, lpKeyName, fFalse, lpDefault, lpReturnedString, cchSize, lpFileName)
#define MsoWritePrivateProfileStringW(lpAppName, lpKeyName, lpString, lpFileName) \
		MsoWritePrivateProfileString(lpAppName, lpKeyName, fFalse, lpString, lpFileName)

/*---------------------------------------------------------------------------
	MsoGetDateFormatW
	
	UNICODE wrapper for GetDateFormat() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(int) MsoGetDateFormatW(LCID Locale, DWORD dwFlags,
	const SYSTEMTIME *lpDate, const WCHAR *lpFormat, _Out_opt_cap_(cchDate) WCHAR *lpDateStr,
	int cchDate);

/*---------------------------------------------------------------------------
	MsoGetTimeFormatW
	
	UNICODE wrapper for GetTimeFormat() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(int) MsoGetTimeFormatW(LCID Locale, DWORD dwFlags,
	const SYSTEMTIME *lpTime, const WCHAR *lpFormat, _Out_opt_cap_(cchTime) WCHAR *lpTimeStr,
	int cchTime);

/*---------------------------------------------------------------------------
	MsoRegisterClassW

	UNICODE wrapper for RegisterClass() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(ATOM) MsoRegisterClassW(const WNDCLASSW *lpWndClass);

/*---------------------------------------------------------------------------
	MsoUnregisterClassW

	UNICODE wrapper for UnregisterClass() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPIX_(BOOL) MsoUnregisterClassW(const WCHAR *lpClassName,
	HINSTANCE hInstance);

/*---------------------------------------------------------------------------
	MsoFindResourceW

	UNICODE wrapper for FindResource() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(HRSRC) MsoFindResourceW(HMODULE hModule, const WCHAR *lpName,
	const WCHAR *lpType);

/*---------------------------------------------------------------------------
	MsoLoadStringW

	UNICODE wrapper for LoadString() system call
------------------------------------------------------------------- JBelt -*/
MSOAPI_(int) MsoLoadStringW(HINSTANCE hInstance, UINT uID, _Out_cap_(nBufferMax) WCHAR *wzBuffer, 
							int nBufferMax);


/*---------------------------------------------------------------------------
	MsoLoadIconExW

    Implements functionality of LoadIcon() and LoadCursor() system call.
------------------------------------------------------------------- DHsu -*/
MSOAPI_(HICON) MsoLoadIconExW(HINSTANCE hinst, LPCWSTR wzName, BOOL fIcon,
                               int cxDesired, int cyDesired, UINT uFlags);
#define MsoLoadIconW(hinst, wzIconName) \
        MsoLoadIconExW(hinst, wzIconName, TRUE, 0, 0, LR_DEFAULTSIZE|LR_DEFAULTCOLOR)
#define MsoLoadCursorW(hinst, wzCursorName) \
        MsoLoadIconExW(hinst, wzCursorName, FALSE, 0, 0, LR_DEFAULTSIZE|LR_DEFAULTCOLOR)


/*---------------------------------------------------------------------------
	MsoLoadBitmapExW

	Implements functionality of LoadBitmap() system call.
------------------------------------------------------------------- DHsu -*/
MSOAPI_(HBITMAP) MsoLoadBitmapExW(HINSTANCE hInstance, LPCWSTR wzBitmapName, UINT uFlags);
#define MsoLoadBitmapW(hinst, wzBitmapName) \
        MsoLoadBitmapExW(hinst, wzBitmapName, LR_DEFAULTCOLOR)


/*---------------------------------------------------------------------------
	MsoLoadMenuW

	UNICODE wrapper for LoadMenu() system call.
------------------------------------------------------------------- DHsu -*/
MSOAPI_(HMENU) MsoLoadMenuW(HINSTANCE hInstance, LPCWSTR wzMenuName);


/*---------------------------------------------------------------------------
	MsoLoadImageW

	UNICODE wrapper for LoadImage() system call.
------------------------------------------------------------ PavelK, DHsu -*/
MSOAPI_(HANDLE) MsoLoadImageW(HINSTANCE hinst, const WCHAR *wzName, UINT uType, 
							  int cxDesired, int cyDesired, UINT fuLoad);


/*---------------------------------------------------------------------------
	MsoSetWindowLongW

	UNICODE wrapper for SetWindowLong() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(LONG) MsoSetWindowLongW(HWND hWnd, int nIndex, LONG dwNewLong);

/*---------------------------------------------------------------------------
	MsoGetWindowLongW

	UNICODE wrapper for GetWindowLong() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(LONG) MsoGetWindowLongW(HWND hWnd, int nIndex);

#ifdef _WIN64
MSOAPIX_(LONG_PTR) MsoSetWindowLongPtrW(HWND hWnd, int nIndex, LONG_PTR dwNewLong);
MSOAPIX_(LONG_PTR) MsoGetWindowLongPtrW(HWND hWnd, int nIndex);
#else
#define MsoSetWindowLongPtrW(hwnd, nIndex, dwNewLong) MsoSetWindowLongW(hwnd, nIndex, dwNewLong)
#define MsoGetWindowLongPtrW(hwnd, nIndex) MsoGetWindowLongW(hwnd, nIndex)
#endif // _WIN64

/*---------------------------------------------------------------------------
	MsoGetClassLongW

	UNICODE wrapper for GetClassLong() system call.
------------------------------------------------------------------ KatsuN -*/
MSOAPIX_(LONG) MsoGetClassLongW(HWND hWnd, int nIndex);

#ifdef _WIN64
MSOAPIX_(LONG_PTR) MsoGetClassLongPtrW(HWND hWnd, int nIndex);
#else
#define MsoGetClassLongPtrW(hWnd, nIndex)	MsoGetClassLongW(hWnd, nIndex)
#endif // _WIN64

/*---------------------------------------------------------------------------
	MsoFindFirstChangeNotificationW

	UNICODE wrapper for FindFirstChangeNotification() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPIX_(HANDLE) MsoFindFirstChangeNotificationW(const WCHAR *lpPathName,
	BOOL bWatchSubtree, DWORD dwNotifyFilter);

/*---------------------------------------------------------------------------
	MsoDefWindowProcW

	UNICODE wrapper for DefWindowProc() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(LRESULT) MsoDefWindowProcW(HWND hWnd, UINT Msg, WPARAM wParam,
	LPARAM lParam);

/*---------------------------------------------------------------------------
	MsoGetClassNameW

	UNICODE wrapper for GetClassName() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPIX_(int) MsoGetClassNameW(HWND hWnd, _Out_cap_(cchClassName) WCHAR *wzClassName, int  cchClassName);

/*---------------------------------------------------------------------------
	MsoDialogBoxIndirectParamW

	UNICODE wrapper for DialogBoxIndirectParam() system call.
----------------------------------------------------------------- RGiesen -*/
MSOAPIX_(int) MsoDialogBoxIndirectParamW(HINSTANCE hInstance,
	LPCDLGTEMPLATEW lpTemplate, HWND hWndParent, DLGPROC lpDialogFunc,
	LPARAM lParamInit);

/*---------------------------------------------------------------------------
	MsoCreateDialogIndirectParamW

	UNICODE wrapper for CreateDialogIndirectParam() system call.
---------------------------------------------------------------- EricSchr -*/
MSOAPIX_(HWND) MsoCreateDialogIndirectParamW(HINSTANCE hInstance,
	LPCDLGTEMPLATEW lpTemplate, HWND hWndParent, DLGPROC lpDialogFunc,
	LPARAM lParamInit);

/*---------------------------------------------------------------------------
	MsoCreateDialogParamW

	UNICODE wrapper for CreateDialogParam() system call.
----------------------------------------------------------------- ShamikB -*/
MSOAPIX_(HWND) MsoCreateDialogParamW(HINSTANCE hInstance,
	int iResource, HWND hWndParent, DLGPROC lpDialogFunc,
	LPARAM lParamInit);

MSOAPI_(UINT) MsoGetDlgItemTextW(HWND hDlg, int nIDDlgItem, _Out_cap_(nMaxCount) LPWSTR lpString,
    int nMaxCount);

MSOAPI_(HANDLE) MsoFindFirstFileW(LPCWSTR lpFileName, LPWIN32_FIND_DATAW lpFindFileData);

/*---------------------------------------------------------------------------
	MsoFindNextFileW
------------------------------------------------------------------- IgorZ -*/
MSOAPI_(BOOL) MsoFindNextFileW(HANDLE hFindFile, LPWIN32_FIND_DATAW lpFindFileData);

MSOAPI_(HMODULE) MsoGetModuleHandleW(LPCWSTR lpModuleName);
MSOAPI_(DWORD) MsoGetModuleFileNameW(HMODULE hModule, _Out_cap_(nSize) LPWSTR lpFilename,
		DWORD nSize);

MSOAPI_(BOOL) MsoSetWindowTextW(HWND hwnd, LPCWSTR lpString);
MSOAPI_(BOOL) MsoSetWindowTextExW(HWND hWnd, LPCWSTR lpString, LPCWSTR lpStringDefault);

/*---------------------------------------------------------------------------
	MsoGetWindowTextW
------------------------------------------------------------------- IlyaV -*/
MSOAPI_(int) MsoGetWindowTextW(HWND hWnd, _Out_cap_(nMaxCount) LPWSTR lpString, int nMaxCount);

/*---------------------------------------------------------------------------
	MsoGetCommandLineW

	Doesn't follow the exact Windows syntax because needs space to return
	the converted string.
------------------------------------------------------------------- JBelt -*/
MSOAPIX_(void) MsoGetCommandLineW(_Out_cap_(cchMax) WCHAR *wzCmdLine, int cchMax);


/*---------------------------------------------------------------------------
	MsoCopyFileW
------------------------------------------------------------------- JBelt -*/
MSOAPI_(BOOL) MsoCopyFileW(LPCWSTR lpExistingFileName, LPCWSTR lpNewFileName,
		BOOL bFailIfExists);

/*---------------------------------------------------------------------------
	MsoCreateFileW

	Can only use it to open files (not named pipes, nor DOS devices)
------------------------------------------------------------------- IlyaV -*/
MSOAPI_(HANDLE) MsoCreateFileW (
	LPCWSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDistribution,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile);

/*---------------------------------------------------------------------------
	MsoCreateNamedPipeFileW

	Can only use it to open named pipes.
----------------------------------------------------------------- chakmli -*/
MSOAPI_(HANDLE) MsoCreateNamedPipeFileW (
	LPCWSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDistribution,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile);

/*	WARNING: these APIs are *not* Darwin aware. Make sure your code deals with
	Darwin before calling it. See msotc.h for MsoLoadLibrary[Ex], which are
	Darwin aware. */
MSOAPI_(HMODULE) _MsoLoadLibraryW(const WCHAR *wzLib);
MSOAPI_(HMODULE) _MsoLoadLibraryExW(const WCHAR *wzLib, HANDLE hFile, DWORD dwFlags);


/*---------------------------------------------------------------------------
	MsoIntializeCriticalSection

	Call either InitializeCriticalSection or InitializeCriticalSectionAndSpinCount
	if available.

------------------------------------------------------------------ BrianWen -*/
MSOAPIX_(BOOL) MsoInitializeCriticalSection (LPCRITICAL_SECTION pcs, DWORD dwSpin);


/*----------------------------------------------------------------------------
	MsoExpandEnvironmentStringsW
-------------------------------------------------------------------- JBelt --*/
MSOAPIX_(DWORD) MsoExpandEnvironmentStringsW(LPCWSTR lpSrc, _Out_opt_cap_(cchDst) LPWSTR  wzDst,
    DWORD  cchDst);

/*----------------------------------------------------------------------------
	MsoFormatMessageW
-------------------------------------------------------------------- JBelt --*/
MSOAPI_(DWORD) MsoFormatMessageW(
		DWORD  dwFlags,
		LPCVOID  lpSource,
		DWORD  dwMessageId,
		DWORD  dwLanguageId,
		_Out_cap_(nSize)
		LPWSTR  lpBuffer,
		DWORD  nSize);

/*----------------------------------------------------------------------------
	MsoGetTempFileNameW
-------------------------------------------------------------------- IlyaV --*/
MSOAPI_(UINT) MsoGetTempFileNameW(LPCWSTR lpPathName, LPCWSTR lpPrefixString, UINT uUnique, _Out_cap_(cchTempFileName) LPWSTR lpTempFileName, int cchTempFileName);

/*----------------------------------------------------------------------------
	MsoGetFullPathNameW
-------------------------------------------------------------------- IlyaV --*/
MSOAPI_(DWORD) MsoGetFullPathNameW(
	LPCWSTR lpFileName, DWORD nBufferLength, _Out_cap_(nBufferLength) LPWSTR lpBuffer, LPWSTR *lpFilePart);

// Flags to control and optimize MsoExtTextOutWEx, MsoGetCharWidthWEx,
//  MsoGetTextExtentPointWEx and MsoGetTextExtentExPointEx

#define MSO_FLDC_NOFL		0x0001 // Disable font linking
#define MSO_FLDC_FLUI		0x0002 // default
#define MSO_FLDC_FLALL		0x0004 // link for all fonts, use FLCS if set
#define MSO_FLDC_FLPARM		0x0008 // font link only fonts found in FLCS
#define MSO_FLDC_FLREPL		0x2000 // user supplied font list to *replace* default list
								   //  - default behavior is to *prepend* user
								   //    supplied font list to default list
#define MSO_FLDC_FLAPPEND	0x0100 // user supplied font list to *append* user supplied
								   //   font list to default list
#define MSO_FLDC_FLFORCE   0x20000 // force return of MSOFLINFO
#define MSO_FLDC_NOSIZEADJ 0x40000 // don't perform font size adjustment
#define MSO_FLDC_SKIPCOMMON	0x80000 // skips the common check for font-linking
#define MSO_FLDC_NOPRIORITY 0x100000 // turns off current font priority
#define MSO_FLDC_NOPRIORITY_WS 0x200000 // turns off current font priority for whitespaces

#define MSO_FLDC_FLMASK		(MSO_FLDC_FLUI | MSO_FLDC_FLALL | MSO_FLDC_FLPARM | \
							 MSO_FLDC_FLREPL | MSO_FLDC_FLAPPEND | MSO_FLDC_FLFORCE | MSO_FLDC_SKIPCOMMON)

// Use these flags to avoid runtime hdc check in MsoXXXXWEx routines
#define MSO_FLDC_EMETA			0x0010	// hdc is an enhanced metafile DC
#define MSO_FLDC_SCREEN			0x0020  // hdc is a screen DC
#define MSO_FLDC_PRINTER		0x0040
#define MSO_FLDC_META			0x0080  // metafile hdc, must call MsoSetMetafileReferenceHdc prior
#define MSO_FLDC_MEMORY			0x10000
#define MSO_FLDC_DCMASK		(MSO_FLDC_EMETA | MSO_FLDC_SCREEN | MSO_FLDC_PRINTER | MSO_FLDC_META | MSO_FLDC_MEMORY)

// Use these flags to avoid runtime device font check for printer DCs
#define MSO_FLDC_DEVICEFONT	 0x0200
#define MSO_FLDC_NOTDEVICEFONT 0x0400
#define MSO_FLDC_DEVICEMASK	(MSO_FLDC_DEVICEFONT | MSO_FLDC_NOTDEVICEFONT)

// Use these flags to avoid runtime check for symbol font
#define MSO_FLDC_NORMALFONT		0x0800
#define MSO_FLDC_SYMBOLFONT		0x1000
#define MSO_FLDC_FONTMASK	(MSO_FLDC_NORMALFONT | MSO_FLDC_SYMBOLFONT)

// Use this flag to indicate metafile DC has TA_UPDATECP set already.
#define MSO_FLDC_UPDATECP	0x4000

// Use this flag to to skip "ASCII only" optimizations
#define MSO_FLDC_SKIPASCII	0x8000

// One extra Device Technology bit
#define DT_USE_MAX_LINE_COUNT 0x10000000
#define DT_FLDC_NOPRI_WS 0x20000000
#define DT_FLDC_NOPRI	0x40000000
#define DT_FLDC_FULL	0x80000000

typedef struct _FLCSE // Font link control structure entry
	{
	WCHAR	wszName[LF_FACESIZE+1];  	// Font name in Unicode
	UINT		fsCpg;		// mask of codepages font supports (from non-OEM mask in FONTSIGNATURE)
	} FLCSE;

typedef struct _FLCSL // Font link control structure list
	{
	FLCSE *rgpflcse[8];	// list of flcse. NULL entry terminates list.
	} FLCSL;

typedef struct _FLCS
	{
	FLCSL *rgpflcsl[2];	// variable sized array of flcsl. NULL entry terminates list.
	} FLCS;

typedef struct _FLSI // Font link information structure
	{
	WCHAR	wzName[LF_FACESIZE+1];  	// Font name in Unicode
	int		chs;	   // charset of font
	int		fsFtcMask; // font signature supported codepage mask
	BYTE	*pbfcmap;  // 64kbit mask of supported Unicode chars
	struct _FLSI *pflsiNext;// Link to next item in cache
	HFONT   hfont;	   // cached font
	} FLSI;
	
typedef struct _FSCA // Font Signature CAche
	{
	char	szName[LF_FACESIZE+1];  	// Font name
	FONTSIGNATURE	fs; 			// font signature
	struct _FSCA *pfscaNext;// Link to next item in cache
	} FSCA;

/*---------------------------------------------------------------------------
	MsoFComplexStr
----------------------------------------------------------------- SHIRAZS -*/

inline BOOL MsoFRTLLayOut() { return FALSE; }
inline BOOL MsoSetRTLLayOut(BOOL fSet) { return FALSE; }
inline BOOL MsoFComplexStr(HDC hdc, LPCWSTR lpString, int cbString, 
                           UINT fuOptions, HDC hdcRef, UINT grfFlDc) { return FALSE; }

/*---------------------------------------------------------------------------
	MsoExtTextOutW
	
	Note: For usage of pflfl and cflfl parameters in MsoExtTextOutFontLinkW(),
	please consult header comments for MsoDoFontLinkRgwch() for corresponding
	parameters.
----------------------------------------------------------------- NobuyaH -*/
MSOAPI_(BOOL) MsoExtTextOutW(HDC hdc, int xp, int yp, UINT eto,
							 CONST RECT *lprect, LPCWSTR lpwch, UINT cLen,
							 CONST INT *lpdxp);

MSOAPI_(BOOL) MsoExtTextOutWEx(HDC hdc, int xp, int yp, UINT eto,
							 CONST RECT *lprect, LPCWSTR lpwch, UINT cLen,
							 CONST INT *lpdxp, UINT grfFlDc, FLCS *pflcs);
							 
/*---------------------------------------------------------------------------
	MsoTextOutW
----------------------------------------------------------------- NobuyaH -*/
MSOAPI_(BOOL) MsoTextOutW(HDC hdc, int xp, int yp, LPCWSTR lpwch, int cLen);

/*---------------------------------------------------------------------------
	MsoGetTextExtentPointW
	
	Note: For usage of pflfl and cflfl parameters in MsoGetTextExtentPointFontLinkW(),
	please consult header comments for MsoDoFontLinkRgwch() for corresponding
	parameters.
----------------------------------------------------------------- NobuyaH -*/
MSOAPI_(BOOL) MsoGetTextExtentPointW(HDC hdc, LPCWSTR lpwch, int cch,
									 LPSIZE lpSize);

/*---------------------------------------------------------------------------
	MsoGetLocaleInfoW
----------------------------------------------------------------- NitinG --*/
MSOAPI_(int) MsoGetLocaleInfoW(LCID  Locale, LCTYPE  LCType,
							   _Out_cap_(cchData) LPWSTR  lpLCData, int  cchData);

// This API does nothing at all
// TODO(JBelt/IgorZ): clean this up after VSE stops calling it
MSOAPIX_(void) MsoInitDarwin(_In_ WCHAR *wzGuidcCore);

MSOAPI_(LRESULT) MsoSetDlgItemText(HWND hDlg, int id, const WCHAR *wz, BOOL fRichEdit);

/*---------------------------------------------------------------------------
	LogFont/TextMetrics wrappers
----------------------------------------------------------------- ziyiw --*/
MSOAPI_(HFONT) MsoCreateFontIndirectW(const LOGFONTW *plfw);

MSOAPI_(int) MsoGetFontObjectW(HGDIOBJ hobj, int cbBuffer, _Out_cap_(cbBuffer) LPVOID lpvObj);

MSOAPI_(int) MsoGetTextFaceW(HDC hdc, int cch, _Out_cap_(cch) WCHAR *xszFace);

MSOAPI_(int) MsoEnumFontFamiliesExW(HDC hdc, LOGFONTW *plfw, FONTENUMPROC lpfn, LPARAM lParam);

MSOAPI_(int) MsoEnumFontFamiliesW(HDC hdc, const WCHAR *wzFontFace, FONTENUMPROC lpfn, LPARAM lParam);

MSOAPI_(int) MsoLogFontWToLogFontA(LOGFONTW *plfw, LOGFONTA *plfa);

MSOAPI_(int) MsoLogFontAToLogFontW(LOGFONTA *plfa, LOGFONTW *plfw);

MSOAPI_(int) MsoEnumLogFontAToEnumLogFontW(ENUMLOGFONTA *pelfa, ENUMLOGFONTW *pelfw, int fTrueType);

MSOAPIX_(int) MsoNewTextMetricExAToNewTextMetricExW(NEWTEXTMETRICEXA *ptma, NEWTEXTMETRICEXW *ptmw);

MSOAPI_(int) MsoTextMetricAToNewTextMetricExW(TEXTMETRICA *ptma, NEWTEXTMETRICEXW *ptmw);

MSOAPI_(int) MsoNewTextMetricAToNewTextMetricExW(NEWTEXTMETRICA *ptma, NEWTEXTMETRICEXW *ptmw);

/*---------------------------------------------------------------------------
	Common controls wrappers.
----------------------------------------------------------------- IlyaV --*/
MSOAPI_(int) ListView_InsertColumnW(HWND hwnd, int iCol, const LV_COLUMNW *pcol);

MSOAPI_(int) ListView_InsertItemW(HWND hwnd, const LV_ITEMW* pitem);

MSOAPI_(BOOL) TreeView_GetItemW(HWND hwnd, TV_ITEMW* pitem);

MSOAPI_(BOOL) TreeView_SetItemW(HWND hwnd, const TV_ITEMW* pitem);

MSOAPI_(HTREEITEM) TreeView_InsertItemW(HWND hwnd, const TV_INSERTSTRUCTW* pis);

MSOAPIX_(BOOL) TabCtrl_GetItemW(HWND hwnd, int iItem, TC_ITEMW* pitem);

MSOAPI_(int) ListView_FindItemW(HWND hwnd, int iStart, const LV_FINDINFOW* plvfi);

MSOAPI_(BOOL) ListView_SetItemW(HWND hwnd, const LV_ITEMW* pitem);

MSOAPI_(BOOL) ListView_SetItemTextW(HWND hwnd, int i, int iSubItem, const WCHAR *pszText);

MSOAPI_(BOOL) ListView_GetItemW(HWND hwnd, LV_ITEMW* pitem);

MSOAPIX_(BOOL) ListView_GetItemTextW(HWND hwnd, int i, int iSubItem, _Out_cap_(cchTextMax) WCHAR *pszText, int cchTextMax);


MSOAPIX_(int) TabCtrl_InsertItemW(HWND hwnd, int iItem, const TC_ITEMW* pitem);

MSOAPIX_(HWND) ListView_EditLabelW(HWND hwnd, int i);

// The predefined macro doesn't return the result.
#undef ListView_SetItemState
__inline BOOL ListView_SetItemState(HWND hwnd, int i, UINT data, UINT mask)
	{
	LV_ITEMA lvi;
	lvi.stateMask = mask;
	lvi.state = data;
	return (BOOL)SendMessageA(hwnd, LVM_SETITEMSTATE, i, (LPARAM)&lvi);
	}


MSOAPI_(int) MsoGetClipboardFormatNameW(UINT format, _Out_cap_(cchMaxCount) LPWSTR lpwzFormatName,
										int cchMaxCount);
MSOAPI_(UINT) MsoGetSystemDirectoryW(_Out_cap_(uSize) LPWSTR lpBuffer, UINT uSize);
MSOAPI_(DWORD) MsoGetEnvironmentVariableW(LPCWSTR lpName, _Out_cap_(nSize) LPWSTR lpBuffer,
										  DWORD nSize);
MSOAPIX_(DWORD) MsoSetEnvironmentVariableW(LPCWSTR lpName, LPCWSTR lpValue);
MSOAPI_(BOOL) MsoMoveFileW(LPCWSTR lpExistingFileName, LPCWSTR lpNewFileName);
MSOAPI_(DWORD) MsoSearchPathW(LPCWSTR lpPath, LPCWSTR lpFileName, LPCWSTR lpExtension,
    						  DWORD nBufferLength, _Out_cap_(nBufferLength) LPWSTR lpBuffer, LPWSTR *lpFilePart);
MSOAPI_(DWORD) MsoGetGlyphOutlineW(HDC hdc, UINT uchar, UINT uFormat, LPGLYPHMETRICS pgm,
								   DWORD cbBuf, _Out_cap_(cbBuf) LPVOID pvBuf, const MAT2 *pm2);
MSOAPI_(BOOL) MsoFGlyphAvailable(HDC hdc, LPCWSTR pwch, INT cch);

/*-----------------------------------------------------------------------------
	MsoWNetAddConnectionW
	MsoWNetCancelConnectionW
	MsoWNetGetConnectionW

	Unicode wrapper for WNet* system calls
------------------------------------------------------------------- NobuyaH -*/
MSOAPI_(DWORD) MsoWNetAddConnectionW(LPCWSTR lpRemoteName, LPCWSTR lpPassword,
									 LPCWSTR lpLocalName);
MSOAPI_(DWORD) MsoWNetCancelConnectionW(LPCWSTR lpName, BOOL fForce);
MSOAPI_(DWORD) MsoWNetGetConnectionW(LPCWSTR lpLocalName, _Out_cap_(*lpnLength) LPWSTR lpRemoteName,
									 LPDWORD lpnLength);

/*-----------------------------------------------------------------------------
	MsoGlobalAddAtomW
	MsoGlobalFindAtomW
	MsoGlobalGetAtomNameW

	Unicode wrapper for atom-related system calls
------------------------------------------------------------------- NobuyaH -*/
MSOAPI_(ATOM) MsoGlobalAddAtomW(LPCWSTR lpString);
MSOAPI_(ATOM) MsoGlobalFindAtomW(LPCWSTR lpString);
MSOAPI_(UINT) MsoGlobalGetAtomNameW(ATOM nAtom, _Out_cap_(nSize) LPWSTR lpBuffer, int nSize);

/*-----------------------------------------------------------------------------
	macro to declare delayed apis
-------------------------------------------------------------------- HAILIU -*/
#define DECLAREDELAYEDAPIVOID(rt, fn) \
	MSOAPI_(rt) Mso##fn();

#define DECLAREDELAYEDAPI1(rt, fn, t1) \
	MSOAPI_(rt) Mso##fn(t1 p1);

#define DECLAREDELAYEDAPI2(rt, fn, t1, t2) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2);

#define DECLAREDELAYEDAPI3(rt, fn, t1, t2, t3) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3);
	
#define DECLAREDELAYEDAPI4(rt, fn, t1, t2, t3, t4) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4);

#define DECLAREDELAYEDAPI5(rt, fn, t1, t2, t3, t4, t5) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5);

#define DECLAREDELAYEDAPI6(rt, fn, t1, t2, t3, t4, t5, t6) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6);
#define DECLAREDELAYEDAPIX6(rt, fn, t1, t2, t3, t4, t5, t6) \
	MSOAPIX_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6);

#define DECLAREDELAYEDAPI7(rt, fn, t1, t2, t3, t4, t5, t6, t7) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6, t7 p7);

#define DECLAREDELAYEDAPI8(rt, fn, t1, t2, t3, t4, t5, t6, t7, t8) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6, t7 p7, t8 p8);

#define DECLAREDELAYEDAPI10(rt, fn, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6, t7 p7, t8 p8, t9 p9, t10 p10);

#define DECLAREDELAYEDAPI11(rt, fn, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11) \
	MSOAPI_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6, t7 p7, t8 p8, t9 p9, t10 p10, t11 p11);
#define DECLAREDELAYEDAPIX11(rt, fn, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11) \
	MSOAPIX_(rt) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4, t5 p5, t6 p6, t7 p7, t8 p8, t9 p9, t10 p10, t11 p11);

#define DECLAREDELAYEDVOIDAPI0(fn) \
	MSOAPI_(VOID) Mso##fn();
	
#define DECLAREDELAYEDVOIDAPI1(fn, t1) \
	MSOAPI_(VOID) Mso##fn(t1 p1);

#define DECLAREDELAYEDVOIDAPI2(fn, t1, t2) \
	MSOAPI_(VOID) Mso##fn(t1 p1, t2 p2);

#define DECLAREDELAYEDVOIDAPI4(fn, t1, t2, t3, t4) \
	MSOAPI_(VOID) Mso##fn(t1 p1, t2 p2, t3 p3, t4 p4);

/*-----------------------------------------------------------------------------
	Winspool.drv APIs
-------------------------------------------------------------------- HAILIU -*/
#include <winspool.h>
DECLAREDELAYEDAPI5(DWORD, DeviceCapabilitiesW, LPCWSTR, LPCWSTR, WORD, _Out_opt_cap_(4096) LPWSTR, CONST DEVMODEW*)
DECLAREDELAYEDAPI3(BOOL, OpenPrinterW, LPCWSTR, LPHANDLE, LPPRINTER_DEFAULTSW)
DECLAREDELAYEDAPI6(LONG, DocumentPropertiesW, HWND, HANDLE, LPCWSTR, PDEVMODEW, PDEVMODEW, DWORD)
DECLAREDELAYEDAPI7(BOOL, EnumPrintersA, DWORD, LPSTR, DWORD, LPBYTE, DWORD, LPDWORD, LPDWORD)
DECLAREDELAYEDAPI8(BOOL, EnumJobsA, HANDLE, DWORD, DWORD, DWORD, LPBYTE, DWORD, LPDWORD, LPDWORD)
DECLAREDELAYEDAPI6(BOOL, GetPrinterDriverA, HANDLE, LPSTR, DWORD, LPBYTE, DWORD, LPDWORD)
DECLAREDELAYEDAPI1(BOOL, ClosePrinter, HANDLE)
DECLAREDELAYEDAPI1(BOOL, DeletePrinter, HANDLE)
MSOAPI_(BOOL) MsoGetDefaultPrinterW(_Out_cap_(*pcch) LPWSTR wz, LPDWORD pcch);


MSOAPI_(DWORD) MsoStartDocPrinterW(HANDLE hPrinter, DWORD Level, LPBYTE pDocInfo);
MSOAPI_(BOOL) MsoEndDocPrinter(HANDLE hPrinter);
MSOAPI_(BOOL) MsoStartPagePrinter(HANDLE hPrinter);
MSOAPI_(BOOL) MsoEndPagePrinter(HANDLE hPrinter);
MSOAPI_(BOOL) MsoWritePrinter(HANDLE hPrinter, _Out_bytecap_(cbBuf) LPVOID pBuf, DWORD cbBuf, LPDWORD pcWritten);

MSOAPI_(BOOL) MsoEnumPrintersW(DWORD dFlags, LPCWSTR wzName, DWORD dLevel,
		_Out_opt_cap_(cbBuf) LPBYTE pPrinterEnum, DWORD cbBuf, LPDWORD pcbNeeded, LPDWORD pcReturned);
MSOAPI_(BOOL) MsoEnumJobsW(HANDLE hPrinter, DWORD FirstJob, DWORD NoJobs, DWORD Level,
		_Out_opt_cap_(cbBuf) LPBYTE pJob, DWORD cbBuf, LPDWORD pcbNeeded, LPDWORD pcReturned);
MSOAPI_(BOOL) MsoGetPrinterDriverW(HANDLE hPrinter, LPCWSTR pEnvironment, DWORD Level,
		_Out_opt_cap_(cbBuf) LPBYTE pDriverInfo, DWORD cbBuf, LPDWORD pcbNeeded);
MSOAPI_(BOOL) MsoAddPrinterDriverW(LPCWSTR pName, DWORD Level, LPBYTE pDriverInfo);
MSOAPI_(BOOL) MsoGetPrinterW(HANDLE hPrinter, DWORD Level, _Out_opt_cap_(cbBuf) LPBYTE pPrinter,
		DWORD cbBuf, LPDWORD pcbNeeded);
MSOAPI_(BOOL) MsoGetPrinterDriverDirectoryW(LPCWSTR pName, LPCWSTR pEnvironment,
		DWORD Level, _Out_cap_(cbBuf) LPBYTE pDriverDirectory, DWORD cbBuf, LPDWORD pcbNeeded);
MSOAPI_(HANDLE) MsoAddPrinterW(LPCWSTR pName, DWORD Level, const BYTE *pPrinter);
MSOAPIX_(BOOL) MsoAddPrinterConnectionW(LPCWSTR pName);


/*-----------------------------------------------------------------------------
	Shell32.dll APIs
-------------------------------------------------------------------- HAILIU -*/
DECLAREDELAYEDVOIDAPI2(DragAcceptFiles, HWND, BOOL)
DECLAREDELAYEDVOIDAPI1(DragFinish, HDROP)
DECLAREDELAYEDAPI2(BOOL, DragQueryPoint, HDROP, LPPOINT)
DECLAREDELAYEDAPI1(HRESULT, SHGetDesktopFolder, LPSHELLFOLDER*)
DECLAREDELAYEDVOIDAPI4(SHChangeNotify, LONG, UINT, LPCVOID, LPCVOID)
DECLAREDELAYEDAPI3(HRESULT, SHGetSpecialFolderLocation, HWND, int, LPITEMIDLIST*)


#if defined(EXCEL_BUILD) || defined(GRAF)
#define DocumentPropertiesW     MsoDocumentPropertiesW
#define DeviceCapabilitiesW     MsoDeviceCapabilitiesW
#define CreateDCW               MsoCreateDCW
#define ResetDCW                MsoResetDCW
#define StartDocW               MsoStartDocW
#endif


/*-----------------------------------------------------------------------------
	Shell.dll wrapper functions.
------------------------------------------------------------------- toddzim -*/
MSOAPI_(BOOL) MsoSHGetPathFromIDListW(LPCITEMIDLIST pidl, _Out_cap_(cch) LPWSTR wzPathReturn, int cch);
MSOAPIX_(LPITEMIDLIST) MsoSHBrowseForFolderW(LPBROWSEINFOW lpbi);
MSOAPI_(HINSTANCE) MsoFindExecutableW(LPCWSTR lpwzFile, LPCWSTR lpwzDirectory, _Out_cap_(cch) LPWSTR lpwzResult, int cch);
MSOAPI_(DWORD_PTR) MsoSHGetFileInfoW(LPCWSTR lpwzPath, DWORD dwFileAttrib, SHFILEINFOW *psfi, UINT cbfi, UINT uFlags);
MSOAPI_(HINSTANCE) MsoShellExecuteW(HWND hwnd, const WCHAR *wzOp, const WCHAR *wzPath, const WCHAR *wzParam, const WCHAR *wzDir, int nShow);
MSOAPI_(BOOL) MsoShellExecuteExW(LPSHELLEXECUTEINFOW pei);

/*---------------------------------------------------------------------------
	MsoSHGetSpecialFolderPathW

	UNICODE wrapper for SHGetSpecialFolderPath() system call.
	Although the Win32 API says this function returns an HRESULT, through
	testing I found that it actually returns a BOOL.  In fact, it sometimes
	even returns TRUE with an empty string, so it is best to ignore the
	return value and look at the string itself.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(BOOL) MsoSHGetSpecialFolderPathW(HWND hwndOwner, _Out_cap_(cchMax) WCHAR *wzPath, int cchMax,
	int nFolder, BOOL fCreate);

/*-----------------------------------------------------------------------------
	MsoSHGetFileInfoW
-------------------------------------------------------------------- HAILIU -*/
MSOAPI_(DWORD_PTR) MsoSHGetFileInfoW(LPCWSTR wzPath, DWORD dwAttr,
	SHFILEINFOW *psfi, UINT cbfi, UINT grf);

/*-----------------------------------------------------------------------------
	This function is used to close an hkey opened with the ORAPI 
	MsoRegOpenKeyEx	or MsoRegCreateKeyEx functions.
--------------------------------------------------------------------dgray----*/
#ifdef DEBUG
MSOAPI_(LONG) MsoRegCloseKeyHkeyCore(HKEY hkey);
#define MsoRegCloseKeyHkey(hkey) MsoRegCloseKeyHkeyCore(hkey)
#else  // ! DEBUG
#define MsoRegCloseKeyHkey(hkey) RegCloseKey(hkey)
#endif // ! DEBUG

/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Office Registry API Wide Wrappers

	These are direct wrappers of the Win32 functions with all wide
	parameters.  Use these only if you can't use ORAPI.
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
MSOAPI_(LONG) MsoRegOpenKeyExW (HKEY hKey, LPCWSTR lpSubKey, REGSAM samDesired, PHKEY phkResult);
MSOAPI_(LONG) MsoRegQueryValueExW (HKEY hKey, 
                                   LPCWSTR lpValueName,
                                   LPDWORD lpType, 
                                   _Out_opt_cap_(*lpcbData) LPBYTE lpData, 
                                   LPDWORD lpcbData);
MSOAPI_(LONG) MsoRegSetValueExW(HKEY hKey, LPCWSTR lpSubKey, 
	DWORD dwType, const BYTE *lpData, DWORD cbData);
MSOAPI_(LONG) MsoRegDeleteValueW(HKEY hKey, LPCWSTR lpSubKey);
MSOAPI_(LONG) MsoRegEnumKeyW (HKEY hKey, DWORD dwIndex, _Out_cap_(cchName) LPWSTR lpName, DWORD cchName);
MSOAPI_(LONG) MsoRegEnumValueW (HKEY hKey, DWORD dwIndex, _Out_cap_(*lpcchValueName) LPWSTR lpValueName, 
		LPDWORD lpcchValueName, LPDWORD lpType, _Out_opt_cap_(*lpcbData) LPBYTE lpData, LPDWORD lpcbData);

/*-----------------------------------------------------------------------------
	OleAut32.dll stubs
-------------------------------------------------------------------- toddzim -*/
#if (defined(OUTLOOK_BUILD)) || defined(ZENSTAT_LIB_DEF)
#undef NO_MSO_OLEWRAP
#define NO_MSO_OLEWRAP 1
#endif	// OUTLOOK_BUILD

#include <oleauto.h>
#include <olectl.h>

/******************************************************************************
	Version APIs
******************************************************************** ToddZim **/

MSOAPI_(DWORD) MsoGetFileVersionInfoSizeW(LPCWSTR wzFilename, LPDWORD pdwHandle);
MSOAPI_(BOOL) MsoGetFileVersionInfoW(LPCWSTR wzFilename, DWORD dwHandle, DWORD dwLen, _Out_cap_(dwLen) LPVOID pData);
MSOAPI_(BOOL) MsoVerQueryValueW(const LPVOID pBlock, LPCWSTR pSubBlock, LPVOID *ppBuffer, PUINT puLen);


/******************************************************************************
	Commdlg APIs
******************************************************************** ToddZim **/

#include <commdlg.h>

MSOAPI_(short) MsoGetFileTitleW(LPCWSTR wzFile, _Out_cap_(cbuf) LPWSTR wzTitle, WORD cbuf);
MSOAPI_(HWND) MsoReplaceTextW(LPFINDREPLACEW lpFR);
MSOAPI_(HWND) MsoFindTextW(LPFINDREPLACEW lpFR);
DECLAREDELAYEDAPIVOID(DWORD, CommDlgExtendedError);


/******************************************************************************
	Commctrl APIs
******************************************************************** ToddZim **/

#include <commctrl.h>

MSOAPI_(HIMAGELIST) MsoImageList_LoadImageW(HINSTANCE hi, LPCWSTR lpbmp, int cx, int cGrow, COLORREF crMask, UINT uType, UINT uFlags);

#define MsoImageList_ExtractIcon(hi, himl, i) MsoImageList_GetIcon(himl, i, 0)
#define MsoImageList_LoadBitmapW(hi, lpbmp, cx, cGrow, crMask) MsoImageList_LoadImageW(hi, lpbmp, cx, cGrow, crMask, IMAGE_BITMAP, 0)


/******************************************************************************
	Kernel32 APIs
******************************************************************************/

/*---------------------------------------------------------------------------
	MsoLoadLibraryW
-------------------------------------------------------------------ToddZim-*/
MSOAPI_(HINSTANCE) MsoLoadLibraryW(LPCWSTR lpLibFile);


/******************************************************************************
	User32 APIs
******************************************************************************/

/******************************************************************************
	AdvApi32 APIs
******************************************************************************/
#include <wincrypt.h>

// no longer necessary to wrap these, available in NT 4.0+
#define MsoCryptAcquireContextW CryptAcquireContextW
#define MsoCryptReleaseContext CryptReleaseContext
#define MsoCryptCreateHash CryptCreateHash
#define MsoCryptDecrypt CryptDecrypt
#define MsoCryptDeriveKey CryptDeriveKey
#define MsoCryptDestroyHash CryptDestroyHash
#define MsoCryptDestroyKey CryptDestroyKey
#define MsoCryptEncrypt CryptEncrypt
#define MsoCryptExportKey CryptExportKey
#define MsoCryptGenKey CryptGenKey
#define MsoCryptGetHashParam CryptGetHashParam
#define MsoCryptGetProvParam CryptGetProvParam
#define MsoCryptHashData CryptHashData
#define MsoCryptImportKey CryptImportKey
#define MsoCryptSignHash(hHash, dwKeySpec, dwFlags, pbSignature, pdwSigLen) CryptSignHashW(hHash, dwKeySpec, NULL, dwFlags, pbSignature, pdwSigLen)
#define MsoCryptVerifySignature(hHash, pbSignature, dwSigLen, hPubKey, dwFlags) CryptVerifySignatureW(hHash, pbSignature, dwSigLen, hPubKey, NULL, dwFlags)

// no longer necessary to wrap this, available in NT 5.0+
#define MsoCryptEnumProvidersW CryptEnumProvidersW



/*-----------------------------------------------------------------------------
	Const correct wrapper for IDispatch::GetIDsOfNames
--------------------------------------------------------------------hannesr---*/
__inline HRESULT MsoIDispatchGetIDOfName( IDispatch* pIDispatch, REFIID riid, const WCHAR* wzName, LCID lcid, DISPID *pDispId )
{
#if defined(__cplusplus) && !defined(CINTERFACE)
   return pIDispatch->GetIDsOfNames( riid, (LPOLESTR*)(&wzName), 1, lcid, pDispId );
#else
   return pIDispatch->lpVtbl ->GetIDsOfNames( pIDispatch, riid, (LPOLESTR*)(&wzName), 1, lcid, pDispId );
#endif
}

#endif // MSOWRAP_H

