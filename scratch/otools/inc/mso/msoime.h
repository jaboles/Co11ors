#pragma once

/****************************************************************************
	MSOIME.H

	Owner: Motos
	Copyright (c) 1994 Microsoft Corporation

	Miscellanous IME stuff to be used by Drawing Helper APIs
****************************************************************************/

#ifndef MSOIME_H

#define MSOIME_H

#include <imm.h>
//#include <ime.h>
#include <imeshare.h>

#ifdef __cplusplus
extern "C" {
#endif // __cplusplus

typedef HIMC (WINAPI * T_LpfnImmAssociateContext)(HWND ,HIMC);
typedef BOOL (WINAPI * T_LpfnImmConfigureIME)(HKL, HWND ,DWORD, LPVOID);
typedef HIMC (WINAPI * T_LpfnImmCreateContext)(void);
typedef BOOL (WINAPI * T_LpfnImmDestroyContext)(HIMC);
typedef LRESULT (WINAPI * T_LpfnImmEscape)(HKL, HIMC, UINT, LPVOID);
typedef BOOL (WINAPI * T_LpfnImmGetCandidateWindow)(HIMC, DWORD, LPCANDIDATEFORM);
typedef BOOL (WINAPI * T_LpfnImmGetCompositionFontA)(HIMC, LPLOGFONTA);
typedef BOOL (WINAPI * T_LpfnImmGetCompositionFontW)(HIMC, LPLOGFONTW);
typedef LONG (WINAPI * T_LpfnImmGetCompositionString)(HIMC, DWORD, LPVOID, DWORD);
typedef BOOL (WINAPI * T_LpfnImmGetCompositionWindow)(HIMC, LPCOMPOSITIONFORM);
typedef HIMC (WINAPI * T_LpfnImmGetContext)(HWND);
typedef BOOL (WINAPI * T_LpfnImmGetConversionStatus)(HIMC, LPDWORD, LPDWORD);
typedef HWND (WINAPI * T_LpfnImmGetDefaultIMEWnd)(HWND);
typedef UINT (WINAPI * T_LpfnImmGetDescriptionA)(HKL, LPSTR, UINT);
typedef UINT (WINAPI * T_LpfnImmGetDescriptionW)(HKL, LPWSTR, UINT);
typedef BOOL (WINAPI * T_LpfnImmGetIMEFileNameA)(HKL, LPSTR, UINT);
typedef BOOL (WINAPI * T_LpfnImmGetIMEFileNameW)(HKL, LPWSTR, UINT);
typedef BOOL (WINAPI * T_LpfnImmGetOpenStatus)(HIMC);
typedef DWORD (WINAPI * T_LpfnImmGetProperty)(HKL, DWORD);
typedef UINT (WINAPI * T_LpfnImmGetVirtualKey)(HWND);
typedef BOOL (WINAPI * T_LpfnImmIsIME)(HKL);
typedef BOOL (WINAPI * T_LpfnImmIsUIMessage)(HWND, UINT, WPARAM, LPARAM);
typedef BOOL (WINAPI * T_LpfnImmNotifyIME)(HIMC, DWORD, DWORD, DWORD);
typedef BOOL (WINAPI * T_LpfnImmRegisterWordA)(HKL, LPCSTR, DWORD, LPCSTR);
typedef BOOL (WINAPI * T_LpfnImmRegisterWordW)(HKL, LPCWSTR, DWORD, LPCWSTR);
typedef BOOL (WINAPI * T_LpfnImmReleaseContext)(HWND ,HIMC);
typedef BOOL (WINAPI * T_LpfnImmSetCandidateWindow)(HIMC, LPCANDIDATEFORM);
typedef BOOL (WINAPI * T_LpfnImmSetCompositionFontA)(HIMC, LPLOGFONTA);
typedef BOOL (WINAPI * T_LpfnImmSetCompositionFontW)(HIMC, LPLOGFONTW);
typedef BOOL (WINAPI * T_LpfnImmSetCompositionString)(HIMC, DWORD, LPCVOID, DWORD, LPCVOID, DWORD);
typedef BOOL (WINAPI * T_LpfnImmSetCompositionWindow)(HIMC, LPCOMPOSITIONFORM);
typedef BOOL (WINAPI * T_LpfnImmSetConversionStatus)(HIMC, DWORD, DWORD);
typedef BOOL (WINAPI * T_LpfnImmSetOpenStatus)(HIMC, BOOL);
typedef BOOL (WINAPI * T_LpfnImmSetStatusWindowPos)(HIMC, LPPOINT);
typedef HRESULT (WINAPI * T_LpfnAImmActivate)(int);
typedef HRESULT (WINAPI * T_LpfnAImmDeactivate)();
typedef HRESULT (WINAPI * T_LpfnAImmStart)();
typedef HRESULT (WINAPI * T_LpfnAImmEnd)();
typedef LRESULT (WINAPI * T_LpfnAImmOnTranslateMessage)(MSG *);
typedef HRESULT (WINAPI * T_LpfnAImmOnDefWindowProc)(HWND, UINT, WPARAM, LPARAM, LRESULT *);
typedef HRESULT (WINAPI * T_LpfnAImmFilterClientWindows)(ATOM *, UINT);

typedef struct _imm_data
	{
	int fAIMM;
	int fAIMM12;
	int fAIMMActive;
	int fAIMMGoing;
	int cRefAIMM;


	T_LpfnImmAssociateContext	m_pfnImmAssociateContext;
	T_LpfnImmConfigureIME		m_pfnImmConfigureIMEA;
	T_LpfnImmCreateContext		m_pfnImmCreateContext;
	T_LpfnImmDestroyContext		m_pfnImmDestroyContext;
	T_LpfnImmEscape	m_pfnImmEscapeA;
	T_LpfnImmGetCandidateWindow	m_pfnImmGetCandidateWindow;
	T_LpfnImmGetCompositionFontA	m_pfnImmGetCompositionFontA;
	T_LpfnImmGetCompositionString	m_pfnImmGetCompositionStringA;
	T_LpfnImmGetCompositionWindow	m_pfnImmGetCompositionWindow;
	T_LpfnImmGetContext	m_pfnImmGetContext;
	T_LpfnImmGetConversionStatus	m_pfnImmGetConversionStatus;
	T_LpfnImmGetDefaultIMEWnd	m_pfnImmGetDefaultIMEWnd;
	T_LpfnImmGetDescriptionA	m_pfnImmGetDescriptionA;
	T_LpfnImmGetIMEFileNameA	m_pfnImmGetIMEFileNameA;
	T_LpfnImmGetOpenStatus		m_pfnImmGetOpenStatus;
	T_LpfnImmGetProperty		m_pfnImmGetProperty;
	T_LpfnImmGetVirtualKey		m_pfnImmGetVirtualKey;
	T_LpfnImmIsIME				m_pfnImmIsIME;
	T_LpfnImmIsUIMessage		m_pfnImmIsUIMessageA;
	T_LpfnImmNotifyIME			m_pfnImmNotifyIME;
	T_LpfnImmRegisterWordA 		m_pfnImmRegisterWordA;
	T_LpfnImmReleaseContext		m_pfnImmReleaseContext;
	T_LpfnImmSetCandidateWindow	m_pfnImmSetCandidateWindow;
	T_LpfnImmSetCompositionFontA	m_pfnImmSetCompositionFontA;
	T_LpfnImmSetCompositionString	m_pfnImmSetCompositionStringA;
	T_LpfnImmSetCompositionWindow	m_pfnImmSetCompositionWindow;
	T_LpfnImmSetConversionStatus	m_pfnImmSetConversionStatus;
	T_LpfnImmSetOpenStatus		m_pfnImmSetOpenStatus;
	T_LpfnImmSetStatusWindowPos	m_pfnImmSetStatusWindowPos;

	// we keep all the unicode API together to make our check faster
	T_LpfnImmConfigureIME		m_pfnImmConfigureIMEW;
	T_LpfnImmEscape	m_pfnImmEscapeW;
	T_LpfnImmGetCompositionFontW	m_pfnImmGetCompositionFontW;
	T_LpfnImmGetCompositionString	m_pfnImmGetCompositionStringW;
	T_LpfnImmGetDescriptionW	m_pfnImmGetDescriptionW;
	T_LpfnImmGetIMEFileNameW	m_pfnImmGetIMEFileNameW;
	T_LpfnImmIsUIMessage		m_pfnImmIsUIMessageW;
	T_LpfnImmRegisterWordW 		m_pfnImmRegisterWordW;
	T_LpfnImmSetCompositionFontW	m_pfnImmSetCompositionFontW;
	T_LpfnImmSetCompositionString	m_pfnImmSetCompositionStringW;

	// Here are all the AIMM functions 
	T_LpfnAImmActivate			m_pfnAImmActivate;
	T_LpfnAImmDeactivate		m_pfnAImmDeactivate;
	T_LpfnAImmStart				m_pfnAImmStart;
	T_LpfnAImmEnd				m_pfnAImmEnd;
    T_LpfnAImmOnTranslateMessage m_pfnAImmOnTranslateMessage;
    T_LpfnAImmOnDefWindowProc 	m_pfnAImmOnDefWindowProc;
    T_LpfnAImmFilterClientWindows m_pfnAImmFilterClientWindows;

	IUnknown *punkAIMM;

	int fUnicode;	// ImmGetCompositionStringW is available or not
	} IMM_DATA;

#ifdef __cplusplus
}; // extern "C"
#endif // __cplusplus

typedef struct _IMEATTR {
	DWORD attr;
	DWORD iwchStart;
	DWORD iwchEnd;
} IMEATTR;

typedef struct _imestring {
	DWORD cwch;
	WCHAR *wz;
	DWORD	cAttr;
	IMEATTR *rgAttr;
	HIMC	himc;
} IMESTRING;

MSOAPI_(void) MsoUIMSetModeBias(HWND hwnd, int);

#endif //!IME_H

