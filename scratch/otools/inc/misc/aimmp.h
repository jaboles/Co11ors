/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Wed Oct 04 13:17:04 2000
 */
/* Compiler settings for aimmp.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __aimmp_h__
#define __aimmp_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IActiveIMMIME_Private_FWD_DEFINED__
#define __IActiveIMMIME_Private_FWD_DEFINED__
typedef interface IActiveIMMIME_Private IActiveIMMIME_Private;
#endif 	/* __IActiveIMMIME_Private_FWD_DEFINED__ */


#ifndef __IActiveIME_Private_FWD_DEFINED__
#define __IActiveIME_Private_FWD_DEFINED__
typedef interface IActiveIME_Private IActiveIME_Private;
#endif 	/* __IActiveIME_Private_FWD_DEFINED__ */


/* header files for imported files */
#include "unknwn.h"
#include "dimm.h"
#include "msctf.h"

__post __maybenull
__post __writableTo(byteCount(size))
void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t size);
void __RPC_USER MIDL_user_free( __inout void __RPC_FAR * ); 

/* interface __MIDL_itf_aimmp_0000 */
/* [local] */ 

//=--------------------------------------------------------------------------=
// aimmp.h
//=--------------------------------------------------------------------------=
// (C) Copyright 1995-1997 Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=

#pragma comment(lib,"uuid.lib")

//--------------------------------------------------------------------------
// IActiveIMM Private (internal) Interfaces.

#include <windows.h>

#ifdef __cplusplus
extern "C" {
#endif /* __cplusplus */

#define AIMMP_AFF_SETFOCUS     0x00000001
#define AIMMP_AFF_SETNULLDIM   0x00000002
#define AIMMP_SE_SELECT     0x00000001
#define AIMMP_SE_ISPRESENT  0x00000002
#ifdef __cplusplus
}
#endif  /* __cplusplus */


extern RPC_IF_HANDLE __MIDL_itf_aimmp_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_aimmp_0000_v0_0_s_ifspec;

#ifndef __IActiveIMMIME_Private_INTERFACE_DEFINED__
#define __IActiveIMMIME_Private_INTERFACE_DEFINED__

/* interface IActiveIMMIME_Private */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IActiveIMMIME_Private;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("08C03411-F96B-11d0-A475-00AA006BCC59")
    IActiveIMMIME_Private : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE AssociateContext( 
            /* [in] */ HWND hWnd,
            /* [in] */ HIMC hIME,
            /* [out] */ HIMC __RPC_FAR *phPrev) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ConfigureIMEA( 
            /* [in] */ HKL hKL,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwMode,
            /* [in] */ REGISTERWORDA __RPC_FAR *pData) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ConfigureIMEW( 
            /* [in] */ HKL hKL,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwMode,
            /* [in] */ REGISTERWORDW __RPC_FAR *pData) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE CreateContext( 
            /* [out] */ HIMC __RPC_FAR *phIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DestroyContext( 
            /* [in] */ HIMC hIME) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EnumRegisterWordA( 
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPSTR szRegister,
            /* [in] */ LPVOID pData,
            /* [out] */ IEnumRegisterWordA __RPC_FAR *__RPC_FAR *pEnum) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EnumRegisterWordW( 
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szRegister,
            /* [in] */ LPVOID pData,
            /* [out] */ IEnumRegisterWordW __RPC_FAR *__RPC_FAR *pEnum) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EscapeA( 
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uEscape,
            /* [out][in] */ LPVOID pData,
            /* [out] */ LRESULT __RPC_FAR *plResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EscapeW( 
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uEscape,
            /* [out][in] */ LPVOID pData,
            /* [out] */ LRESULT __RPC_FAR *plResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCandidateListA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ UINT uBufLen,
            /* [out] */ CANDIDATELIST __RPC_FAR *pCandList,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCandidateListW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ UINT uBufLen,
            /* [out] */ CANDIDATELIST __RPC_FAR *pCandList,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCandidateListCountA( 
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwListSize,
            /* [out] */ DWORD __RPC_FAR *pdwBufLen) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCandidateListCountW( 
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwListSize,
            /* [out] */ DWORD __RPC_FAR *pdwBufLen) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCandidateWindow( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [out] */ CANDIDATEFORM __RPC_FAR *pCandidate) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCompositionFontA( 
            /* [in] */ HIMC hIMC,
            /* [out] */ LOGFONTA __RPC_FAR *plf) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCompositionFontW( 
            /* [in] */ HIMC hIMC,
            /* [out] */ LOGFONTW __RPC_FAR *plf) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCompositionStringA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ LONG __RPC_FAR *plCopied,
            /* [out] */ LPVOID pBuf) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCompositionStringW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ LONG __RPC_FAR *plCopied,
            /* [out] */ LPVOID pBuf) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCompositionWindow( 
            /* [in] */ HIMC hIMC,
            /* [out] */ COMPOSITIONFORM __RPC_FAR *pCompForm) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetContext( 
            /* [in] */ HWND hWnd,
            /* [out] */ HIMC __RPC_FAR *phIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetConversionListA( 
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ _In_z_ LPSTR pSrc,
            /* [in] */ UINT uBufLen,
            /* [in] */ UINT uFlag,
            /* [out] */ CANDIDATELIST __RPC_FAR *pDst,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetConversionListW( 
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ _In_z_ LPWSTR pSrc,
            /* [in] */ UINT uBufLen,
            /* [in] */ UINT uFlag,
            /* [out] */ CANDIDATELIST __RPC_FAR *pDst,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetConversionStatus( 
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pfdwConversion,
            /* [out] */ DWORD __RPC_FAR *pfdwSentence) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetDefaultIMEWnd( 
            /* [in] */ HWND hWnd,
            /* [out] */ HWND __RPC_FAR *phDefWnd) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetDescriptionA( 
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ LPSTR szDescription,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetDescriptionW( 
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ LPWSTR szDescription,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetGuideLineA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ LPSTR pBuf,
            /* [out] */ DWORD __RPC_FAR *pdwResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetGuideLineW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ LPWSTR pBuf,
            /* [out] */ DWORD __RPC_FAR *pdwResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetIMEFileNameA( 
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPSTR szFileName,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetIMEFileNameW( 
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPWSTR szFileName,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetOpenStatus( 
            /* [in] */ HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetProperty( 
            /* [in] */ HKL hKL,
            /* [in] */ DWORD fdwIndex,
            /* [out] */ DWORD __RPC_FAR *pdwProperty) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetRegisterWordStyleA( 
            /* [in] */ HKL hKL,
            /* [in] */ UINT nItem,
            /* [out] */ STYLEBUFA __RPC_FAR *pStyleBuf,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetRegisterWordStyleW( 
            /* [in] */ HKL hKL,
            /* [in] */ UINT nItem,
            /* [out] */ STYLEBUFW __RPC_FAR *pStyleBuf,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetStatusWindowPos( 
            /* [in] */ HIMC hIMC,
            /* [out] */ POINT __RPC_FAR *pptPos) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetVirtualKey( 
            /* [in] */ HWND hWnd,
            /* [out] */ UINT __RPC_FAR *puVirtualKey) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE InstallIMEA( 
            /* [in] */ _In_z_ LPSTR szIMEFileName,
            /* [in] */ _In_z_ LPSTR szLayoutText,
            /* [out] */ HKL __RPC_FAR *phKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE InstallIMEW( 
            /* [in] */ _In_z_ LPWSTR szIMEFileName,
            /* [in] */ _In_z_ LPWSTR szLayoutText,
            /* [out] */ HKL __RPC_FAR *phKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE IsIME( 
            /* [in] */ HKL hKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE IsUIMessageA( 
            /* [in] */ HWND hWndIME,
            /* [in] */ UINT msg,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE IsUIMessageW( 
            /* [in] */ HWND hWndIME,
            /* [in] */ UINT msg,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE NotifyIME( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwAction,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwValue) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RegisterWordA( 
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPSTR szRegister) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RegisterWordW( 
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szRegister) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ReleaseContext( 
            /* [in] */ HWND hWnd,
            /* [in] */ HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCandidateWindow( 
            /* [in] */ HIMC hIMC,
            /* [in] */ CANDIDATEFORM __RPC_FAR *pCandidate) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCompositionFontA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ LOGFONTA __RPC_FAR *plf) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCompositionFontW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ LOGFONTW __RPC_FAR *plf) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCompositionStringA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ LPVOID pComp,
            /* [in] */ DWORD dwCompLen,
            /* [in] */ LPVOID pRead,
            /* [in] */ DWORD dwReadLen) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCompositionStringW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ LPVOID pComp,
            /* [in] */ DWORD dwCompLen,
            /* [in] */ LPVOID pRead,
            /* [in] */ DWORD dwReadLen) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCompositionWindow( 
            /* [in] */ HIMC hIMC,
            /* [in] */ COMPOSITIONFORM __RPC_FAR *pCompForm) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetConversionStatus( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD fdwConversion,
            /* [in] */ DWORD fdwSentence) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetOpenStatus( 
            /* [in] */ HIMC hIMC,
            /* [in] */ BOOL fOpen) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetStatusWindowPos( 
            /* [in] */ HIMC hIMC,
            /* [in] */ POINT __RPC_FAR *pptPos) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SimulateHotKey( 
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwHotKeyID) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UnregisterWordA( 
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPSTR szUnregister) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UnregisterWordW( 
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szUnregister) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GenerateMessage( 
            /* [in] */ HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE LockIMC( 
            /* [in] */ HIMC hIMC,
            /* [out] */ INPUTCONTEXT __RPC_FAR *__RPC_FAR *ppIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UnlockIMC( 
            /* [in] */ HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetIMCLockCount( 
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwLockCount) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE CreateIMCC( 
            /* [in] */ DWORD dwSize,
            /* [out] */ HIMCC __RPC_FAR *phIMCC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DestroyIMCC( 
            /* [in] */ HIMCC hIMCC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE LockIMCC( 
            /* [in] */ HIMCC hIMCC,
            /* [out] */ void __RPC_FAR *__RPC_FAR *ppv) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UnlockIMCC( 
            /* [in] */ HIMCC hIMCC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ReSizeIMCC( 
            /* [in] */ HIMCC hIMCC,
            /* [in] */ DWORD dwSize,
            /* [out] */ HIMCC __RPC_FAR *phIMCC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetIMCCSize( 
            /* [in] */ HIMCC hIMCC,
            /* [out] */ DWORD __RPC_FAR *pdwSize) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetIMCCLockCount( 
            /* [in] */ HIMCC hIMCC,
            /* [out] */ DWORD __RPC_FAR *pdwLockCount) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetHotKey( 
            /* [in] */ DWORD dwHotKeyID,
            /* [out] */ UINT __RPC_FAR *puModifiers,
            /* [out] */ UINT __RPC_FAR *puVKey,
            /* [out] */ HKL __RPC_FAR *phKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetHotKey( 
            /* [in] */ DWORD dwHotKeyID,
            /* [in] */ UINT uModifiers,
            /* [in] */ UINT uVKey,
            /* [in] */ HKL hKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE AssociateContextEx( 
            /* [in] */ HWND hWnd,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DisableIME( 
            /* [in] */ DWORD idThread) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetImeMenuItemsA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags,
            /* [in] */ DWORD dwType,
            /* [in] */ IMEMENUITEMINFOA __RPC_FAR *pImeParentMenu,
            /* [out] */ IMEMENUITEMINFOA __RPC_FAR *pImeMenu,
            /* [in] */ DWORD dwSize,
            /* [out] */ DWORD __RPC_FAR *pdwResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetImeMenuItemsW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags,
            /* [in] */ DWORD dwType,
            /* [in] */ IMEMENUITEMINFOW __RPC_FAR *pImeParentMenu,
            /* [out] */ IMEMENUITEMINFOW __RPC_FAR *pImeMenu,
            /* [in] */ DWORD dwSize,
            /* [out] */ DWORD __RPC_FAR *pdwResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EnumInputContext( 
            /* [in] */ DWORD idThread,
            /* [out] */ IEnumInputContext __RPC_FAR *__RPC_FAR *ppEnum) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RequestMessageA( 
            /* [in] */ HIMC hIMC,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam,
            /* [out] */ LRESULT __RPC_FAR *plResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RequestMessageW( 
            /* [in] */ HIMC hIMC,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam,
            /* [out] */ LRESULT __RPC_FAR *plResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE FilterClientWindowsEx( 
            /* [in] */ HWND hWnd,
            /* [in] */ BOOL fGuidMap) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE FilterClientWindows( 
            /* [in] */ ATOM __RPC_FAR *aaClassList,
            /* [in] */ UINT uSize,
            /* [in] */ BOOL __RPC_FAR *aaGildMap) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetGuidAtom( 
            /* [in] */ HIMC hImc,
            /* [in] */ BYTE bAttr,
            /* [out] */ TfGuidAtom __RPC_FAR *pGuidAtom) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IActiveIMMIME_PrivateVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IActiveIMMIME_Private __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IActiveIMMIME_Private __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AssociateContext )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [in] */ HIMC hIME,
            /* [out] */ HIMC __RPC_FAR *phPrev);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ConfigureIMEA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwMode,
            /* [in] */ REGISTERWORDA __RPC_FAR *pData);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ConfigureIMEW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwMode,
            /* [in] */ REGISTERWORDW __RPC_FAR *pData);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateContext )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [out] */ HIMC __RPC_FAR *phIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DestroyContext )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIME);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EnumRegisterWordA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPSTR szRegister,
            /* [in] */ LPVOID pData,
            /* [out] */ IEnumRegisterWordA __RPC_FAR *__RPC_FAR *pEnum);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EnumRegisterWordW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szRegister,
            /* [in] */ LPVOID pData,
            /* [out] */ IEnumRegisterWordW __RPC_FAR *__RPC_FAR *pEnum);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EscapeA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uEscape,
            /* [out][in] */ LPVOID pData,
            /* [out] */ LRESULT __RPC_FAR *plResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EscapeW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uEscape,
            /* [out][in] */ LPVOID pData,
            /* [out] */ LRESULT __RPC_FAR *plResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCandidateListA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ UINT uBufLen,
            /* [out] */ CANDIDATELIST __RPC_FAR *pCandList,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCandidateListW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ UINT uBufLen,
            /* [out] */ CANDIDATELIST __RPC_FAR *pCandList,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCandidateListCountA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwListSize,
            /* [out] */ DWORD __RPC_FAR *pdwBufLen);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCandidateListCountW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwListSize,
            /* [out] */ DWORD __RPC_FAR *pdwBufLen);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCandidateWindow )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [out] */ CANDIDATEFORM __RPC_FAR *pCandidate);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCompositionFontA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ LOGFONTA __RPC_FAR *plf);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCompositionFontW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ LOGFONTW __RPC_FAR *plf);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCompositionStringA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ LONG __RPC_FAR *plCopied,
            /* [out] */ LPVOID pBuf);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCompositionStringW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ LONG __RPC_FAR *plCopied,
            /* [out] */ LPVOID pBuf);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCompositionWindow )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ COMPOSITIONFORM __RPC_FAR *pCompForm);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetContext )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [out] */ HIMC __RPC_FAR *phIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetConversionListA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ _In_z_ LPSTR pSrc,
            /* [in] */ UINT uBufLen,
            /* [in] */ UINT uFlag,
            /* [out] */ CANDIDATELIST __RPC_FAR *pDst,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetConversionListW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HIMC hIMC,
            /* [in] */ _In_z_ LPWSTR pSrc,
            /* [in] */ UINT uBufLen,
            /* [in] */ UINT uFlag,
            /* [out] */ CANDIDATELIST __RPC_FAR *pDst,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetConversionStatus )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pfdwConversion,
            /* [out] */ DWORD __RPC_FAR *pfdwSentence);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetDefaultIMEWnd )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [out] */ HWND __RPC_FAR *phDefWnd);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetDescriptionA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPSTR szDescription,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetDescriptionW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPWSTR szDescription,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGuideLineA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ _Out_z_cap_post_count_(dwBufLen, *pdwResult) LPSTR pBuf,
            /* [out] */ DWORD __RPC_FAR *pdwResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGuideLineW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwBufLen,
            /* [out] */ _Out_z_cap_post_count_(dwBufLen, *pdwResult) LPWSTR pBuf,
            /* [out] */ DWORD __RPC_FAR *pdwResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIMEFileNameA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPSTR szFileName,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIMEFileNameW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ UINT uBufLen,
            /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPWSTR szFileName,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetOpenStatus )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetProperty )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ DWORD fdwIndex,
            /* [out] */ DWORD __RPC_FAR *pdwProperty);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetRegisterWordStyleA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ UINT nItem,
            /* [out] */ STYLEBUFA __RPC_FAR *pStyleBuf,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetRegisterWordStyleW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ UINT nItem,
            /* [out] */ STYLEBUFW __RPC_FAR *pStyleBuf,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetStatusWindowPos )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ POINT __RPC_FAR *pptPos);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetVirtualKey )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [out] */ UINT __RPC_FAR *puVirtualKey);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *InstallIMEA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ _In_z_ LPSTR szIMEFileName,
            /* [in] */ _In_z_ LPSTR szLayoutText,
            /* [out] */ HKL __RPC_FAR *phKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *InstallIMEW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ _In_z_ LPWSTR szIMEFileName,
            /* [in] */ _In_z_ LPWSTR szLayoutText,
            /* [out] */ HKL __RPC_FAR *phKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *IsIME )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *IsUIMessageA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWndIME,
            /* [in] */ UINT msg,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *IsUIMessageW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWndIME,
            /* [in] */ UINT msg,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NotifyIME )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwAction,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwValue);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RegisterWordA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPSTR szRegister);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RegisterWordW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szRegister);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ReleaseContext )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [in] */ HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCandidateWindow )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ CANDIDATEFORM __RPC_FAR *pCandidate);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCompositionFontA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ LOGFONTA __RPC_FAR *plf);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCompositionFontW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ LOGFONTW __RPC_FAR *plf);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCompositionStringA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ LPVOID pComp,
            /* [in] */ DWORD dwCompLen,
            /* [in] */ LPVOID pRead,
            /* [in] */ DWORD dwReadLen);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCompositionStringW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ LPVOID pComp,
            /* [in] */ DWORD dwCompLen,
            /* [in] */ LPVOID pRead,
            /* [in] */ DWORD dwReadLen);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCompositionWindow )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ COMPOSITIONFORM __RPC_FAR *pCompForm);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetConversionStatus )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD fdwConversion,
            /* [in] */ DWORD fdwSentence);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetOpenStatus )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ BOOL fOpen);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetStatusWindowPos )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ POINT __RPC_FAR *pptPos);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SimulateHotKey )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwHotKeyID);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UnregisterWordA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPSTR szUnregister);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UnregisterWordW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szUnregister);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GenerateMessage )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LockIMC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ INPUTCONTEXT __RPC_FAR *__RPC_FAR *ppIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UnlockIMC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIMCLockCount )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwLockCount);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateIMCC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ DWORD dwSize,
            /* [out] */ HIMCC __RPC_FAR *phIMCC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DestroyIMCC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMCC hIMCC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LockIMCC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMCC hIMCC,
            /* [out] */ void __RPC_FAR *__RPC_FAR *ppv);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UnlockIMCC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMCC hIMCC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ReSizeIMCC )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMCC hIMCC,
            /* [in] */ DWORD dwSize,
            /* [out] */ HIMCC __RPC_FAR *phIMCC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIMCCSize )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMCC hIMCC,
            /* [out] */ DWORD __RPC_FAR *pdwSize);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIMCCLockCount )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMCC hIMCC,
            /* [out] */ DWORD __RPC_FAR *pdwLockCount);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetHotKey )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ DWORD dwHotKeyID,
            /* [out] */ UINT __RPC_FAR *puModifiers,
            /* [out] */ UINT __RPC_FAR *puVKey,
            /* [out] */ HKL __RPC_FAR *phKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetHotKey )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ DWORD dwHotKeyID,
            /* [in] */ UINT uModifiers,
            /* [in] */ UINT uVKey,
            /* [in] */ HKL hKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AssociateContextEx )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DisableIME )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ DWORD idThread);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetImeMenuItemsA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags,
            /* [in] */ DWORD dwType,
            /* [in] */ IMEMENUITEMINFOA __RPC_FAR *pImeParentMenu,
            /* [out] */ IMEMENUITEMINFOA __RPC_FAR *pImeMenu,
            /* [in] */ DWORD dwSize,
            /* [out] */ DWORD __RPC_FAR *pdwResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetImeMenuItemsW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags,
            /* [in] */ DWORD dwType,
            /* [in] */ IMEMENUITEMINFOW __RPC_FAR *pImeParentMenu,
            /* [out] */ IMEMENUITEMINFOW __RPC_FAR *pImeMenu,
            /* [in] */ DWORD dwSize,
            /* [out] */ DWORD __RPC_FAR *pdwResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EnumInputContext )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ DWORD idThread,
            /* [out] */ IEnumInputContext __RPC_FAR *__RPC_FAR *ppEnum);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RequestMessageA )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam,
            /* [out] */ LRESULT __RPC_FAR *plResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RequestMessageW )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ WPARAM wParam,
            /* [in] */ LPARAM lParam,
            /* [out] */ LRESULT __RPC_FAR *plResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FilterClientWindowsEx )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HWND hWnd,
            /* [in] */ BOOL fGuidMap);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FilterClientWindows )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ ATOM __RPC_FAR *aaClassList,
            /* [in] */ UINT uSize,
            /* [in] */ BOOL __RPC_FAR *aaGildMap);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGuidAtom )( 
            IActiveIMMIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hImc,
            /* [in] */ BYTE bAttr,
            /* [out] */ TfGuidAtom __RPC_FAR *pGuidAtom);
        
        END_INTERFACE
    } IActiveIMMIME_PrivateVtbl;

    interface IActiveIMMIME_Private
    {
        CONST_VTBL struct IActiveIMMIME_PrivateVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IActiveIMMIME_Private_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IActiveIMMIME_Private_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IActiveIMMIME_Private_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IActiveIMMIME_Private_AssociateContext(This,hWnd,hIME,phPrev)	\
    (This)->lpVtbl -> AssociateContext(This,hWnd,hIME,phPrev)

#define IActiveIMMIME_Private_ConfigureIMEA(This,hKL,hWnd,dwMode,pData)	\
    (This)->lpVtbl -> ConfigureIMEA(This,hKL,hWnd,dwMode,pData)

#define IActiveIMMIME_Private_ConfigureIMEW(This,hKL,hWnd,dwMode,pData)	\
    (This)->lpVtbl -> ConfigureIMEW(This,hKL,hWnd,dwMode,pData)

#define IActiveIMMIME_Private_CreateContext(This,phIMC)	\
    (This)->lpVtbl -> CreateContext(This,phIMC)

#define IActiveIMMIME_Private_DestroyContext(This,hIME)	\
    (This)->lpVtbl -> DestroyContext(This,hIME)

#define IActiveIMMIME_Private_EnumRegisterWordA(This,hKL,szReading,dwStyle,szRegister,pData,pEnum)	\
    (This)->lpVtbl -> EnumRegisterWordA(This,hKL,szReading,dwStyle,szRegister,pData,pEnum)

#define IActiveIMMIME_Private_EnumRegisterWordW(This,hKL,szReading,dwStyle,szRegister,pData,pEnum)	\
    (This)->lpVtbl -> EnumRegisterWordW(This,hKL,szReading,dwStyle,szRegister,pData,pEnum)

#define IActiveIMMIME_Private_EscapeA(This,hKL,hIMC,uEscape,pData,plResult)	\
    (This)->lpVtbl -> EscapeA(This,hKL,hIMC,uEscape,pData,plResult)

#define IActiveIMMIME_Private_EscapeW(This,hKL,hIMC,uEscape,pData,plResult)	\
    (This)->lpVtbl -> EscapeW(This,hKL,hIMC,uEscape,pData,plResult)

#define IActiveIMMIME_Private_GetCandidateListA(This,hIMC,dwIndex,uBufLen,pCandList,puCopied)	\
    (This)->lpVtbl -> GetCandidateListA(This,hIMC,dwIndex,uBufLen,pCandList,puCopied)

#define IActiveIMMIME_Private_GetCandidateListW(This,hIMC,dwIndex,uBufLen,pCandList,puCopied)	\
    (This)->lpVtbl -> GetCandidateListW(This,hIMC,dwIndex,uBufLen,pCandList,puCopied)

#define IActiveIMMIME_Private_GetCandidateListCountA(This,hIMC,pdwListSize,pdwBufLen)	\
    (This)->lpVtbl -> GetCandidateListCountA(This,hIMC,pdwListSize,pdwBufLen)

#define IActiveIMMIME_Private_GetCandidateListCountW(This,hIMC,pdwListSize,pdwBufLen)	\
    (This)->lpVtbl -> GetCandidateListCountW(This,hIMC,pdwListSize,pdwBufLen)

#define IActiveIMMIME_Private_GetCandidateWindow(This,hIMC,dwIndex,pCandidate)	\
    (This)->lpVtbl -> GetCandidateWindow(This,hIMC,dwIndex,pCandidate)

#define IActiveIMMIME_Private_GetCompositionFontA(This,hIMC,plf)	\
    (This)->lpVtbl -> GetCompositionFontA(This,hIMC,plf)

#define IActiveIMMIME_Private_GetCompositionFontW(This,hIMC,plf)	\
    (This)->lpVtbl -> GetCompositionFontW(This,hIMC,plf)

#define IActiveIMMIME_Private_GetCompositionStringA(This,hIMC,dwIndex,dwBufLen,plCopied,pBuf)	\
    (This)->lpVtbl -> GetCompositionStringA(This,hIMC,dwIndex,dwBufLen,plCopied,pBuf)

#define IActiveIMMIME_Private_GetCompositionStringW(This,hIMC,dwIndex,dwBufLen,plCopied,pBuf)	\
    (This)->lpVtbl -> GetCompositionStringW(This,hIMC,dwIndex,dwBufLen,plCopied,pBuf)

#define IActiveIMMIME_Private_GetCompositionWindow(This,hIMC,pCompForm)	\
    (This)->lpVtbl -> GetCompositionWindow(This,hIMC,pCompForm)

#define IActiveIMMIME_Private_GetContext(This,hWnd,phIMC)	\
    (This)->lpVtbl -> GetContext(This,hWnd,phIMC)

#define IActiveIMMIME_Private_GetConversionListA(This,hKL,hIMC,pSrc,uBufLen,uFlag,pDst,puCopied)	\
    (This)->lpVtbl -> GetConversionListA(This,hKL,hIMC,pSrc,uBufLen,uFlag,pDst,puCopied)

#define IActiveIMMIME_Private_GetConversionListW(This,hKL,hIMC,pSrc,uBufLen,uFlag,pDst,puCopied)	\
    (This)->lpVtbl -> GetConversionListW(This,hKL,hIMC,pSrc,uBufLen,uFlag,pDst,puCopied)

#define IActiveIMMIME_Private_GetConversionStatus(This,hIMC,pfdwConversion,pfdwSentence)	\
    (This)->lpVtbl -> GetConversionStatus(This,hIMC,pfdwConversion,pfdwSentence)

#define IActiveIMMIME_Private_GetDefaultIMEWnd(This,hWnd,phDefWnd)	\
    (This)->lpVtbl -> GetDefaultIMEWnd(This,hWnd,phDefWnd)

#define IActiveIMMIME_Private_GetDescriptionA(This,hKL,uBufLen,szDescription,puCopied)	\
    (This)->lpVtbl -> GetDescriptionA(This,hKL,uBufLen,szDescription,puCopied)

#define IActiveIMMIME_Private_GetDescriptionW(This,hKL,uBufLen,szDescription,puCopied)	\
    (This)->lpVtbl -> GetDescriptionW(This,hKL,uBufLen,szDescription,puCopied)

#define IActiveIMMIME_Private_GetGuideLineA(This,hIMC,dwIndex,dwBufLen,pBuf,pdwResult)	\
    (This)->lpVtbl -> GetGuideLineA(This,hIMC,dwIndex,dwBufLen,pBuf,pdwResult)

#define IActiveIMMIME_Private_GetGuideLineW(This,hIMC,dwIndex,dwBufLen,pBuf,pdwResult)	\
    (This)->lpVtbl -> GetGuideLineW(This,hIMC,dwIndex,dwBufLen,pBuf,pdwResult)

#define IActiveIMMIME_Private_GetIMEFileNameA(This,hKL,uBufLen,szFileName,puCopied)	\
    (This)->lpVtbl -> GetIMEFileNameA(This,hKL,uBufLen,szFileName,puCopied)

#define IActiveIMMIME_Private_GetIMEFileNameW(This,hKL,uBufLen,szFileName,puCopied)	\
    (This)->lpVtbl -> GetIMEFileNameW(This,hKL,uBufLen,szFileName,puCopied)

#define IActiveIMMIME_Private_GetOpenStatus(This,hIMC)	\
    (This)->lpVtbl -> GetOpenStatus(This,hIMC)

#define IActiveIMMIME_Private_GetProperty(This,hKL,fdwIndex,pdwProperty)	\
    (This)->lpVtbl -> GetProperty(This,hKL,fdwIndex,pdwProperty)

#define IActiveIMMIME_Private_GetRegisterWordStyleA(This,hKL,nItem,pStyleBuf,puCopied)	\
    (This)->lpVtbl -> GetRegisterWordStyleA(This,hKL,nItem,pStyleBuf,puCopied)

#define IActiveIMMIME_Private_GetRegisterWordStyleW(This,hKL,nItem,pStyleBuf,puCopied)	\
    (This)->lpVtbl -> GetRegisterWordStyleW(This,hKL,nItem,pStyleBuf,puCopied)

#define IActiveIMMIME_Private_GetStatusWindowPos(This,hIMC,pptPos)	\
    (This)->lpVtbl -> GetStatusWindowPos(This,hIMC,pptPos)

#define IActiveIMMIME_Private_GetVirtualKey(This,hWnd,puVirtualKey)	\
    (This)->lpVtbl -> GetVirtualKey(This,hWnd,puVirtualKey)

#define IActiveIMMIME_Private_InstallIMEA(This,szIMEFileName,szLayoutText,phKL)	\
    (This)->lpVtbl -> InstallIMEA(This,szIMEFileName,szLayoutText,phKL)

#define IActiveIMMIME_Private_InstallIMEW(This,szIMEFileName,szLayoutText,phKL)	\
    (This)->lpVtbl -> InstallIMEW(This,szIMEFileName,szLayoutText,phKL)

#define IActiveIMMIME_Private_IsIME(This,hKL)	\
    (This)->lpVtbl -> IsIME(This,hKL)

#define IActiveIMMIME_Private_IsUIMessageA(This,hWndIME,msg,wParam,lParam)	\
    (This)->lpVtbl -> IsUIMessageA(This,hWndIME,msg,wParam,lParam)

#define IActiveIMMIME_Private_IsUIMessageW(This,hWndIME,msg,wParam,lParam)	\
    (This)->lpVtbl -> IsUIMessageW(This,hWndIME,msg,wParam,lParam)

#define IActiveIMMIME_Private_NotifyIME(This,hIMC,dwAction,dwIndex,dwValue)	\
    (This)->lpVtbl -> NotifyIME(This,hIMC,dwAction,dwIndex,dwValue)

#define IActiveIMMIME_Private_RegisterWordA(This,hKL,szReading,dwStyle,szRegister)	\
    (This)->lpVtbl -> RegisterWordA(This,hKL,szReading,dwStyle,szRegister)

#define IActiveIMMIME_Private_RegisterWordW(This,hKL,szReading,dwStyle,szRegister)	\
    (This)->lpVtbl -> RegisterWordW(This,hKL,szReading,dwStyle,szRegister)

#define IActiveIMMIME_Private_ReleaseContext(This,hWnd,hIMC)	\
    (This)->lpVtbl -> ReleaseContext(This,hWnd,hIMC)

#define IActiveIMMIME_Private_SetCandidateWindow(This,hIMC,pCandidate)	\
    (This)->lpVtbl -> SetCandidateWindow(This,hIMC,pCandidate)

#define IActiveIMMIME_Private_SetCompositionFontA(This,hIMC,plf)	\
    (This)->lpVtbl -> SetCompositionFontA(This,hIMC,plf)

#define IActiveIMMIME_Private_SetCompositionFontW(This,hIMC,plf)	\
    (This)->lpVtbl -> SetCompositionFontW(This,hIMC,plf)

#define IActiveIMMIME_Private_SetCompositionStringA(This,hIMC,dwIndex,pComp,dwCompLen,pRead,dwReadLen)	\
    (This)->lpVtbl -> SetCompositionStringA(This,hIMC,dwIndex,pComp,dwCompLen,pRead,dwReadLen)

#define IActiveIMMIME_Private_SetCompositionStringW(This,hIMC,dwIndex,pComp,dwCompLen,pRead,dwReadLen)	\
    (This)->lpVtbl -> SetCompositionStringW(This,hIMC,dwIndex,pComp,dwCompLen,pRead,dwReadLen)

#define IActiveIMMIME_Private_SetCompositionWindow(This,hIMC,pCompForm)	\
    (This)->lpVtbl -> SetCompositionWindow(This,hIMC,pCompForm)

#define IActiveIMMIME_Private_SetConversionStatus(This,hIMC,fdwConversion,fdwSentence)	\
    (This)->lpVtbl -> SetConversionStatus(This,hIMC,fdwConversion,fdwSentence)

#define IActiveIMMIME_Private_SetOpenStatus(This,hIMC,fOpen)	\
    (This)->lpVtbl -> SetOpenStatus(This,hIMC,fOpen)

#define IActiveIMMIME_Private_SetStatusWindowPos(This,hIMC,pptPos)	\
    (This)->lpVtbl -> SetStatusWindowPos(This,hIMC,pptPos)

#define IActiveIMMIME_Private_SimulateHotKey(This,hWnd,dwHotKeyID)	\
    (This)->lpVtbl -> SimulateHotKey(This,hWnd,dwHotKeyID)

#define IActiveIMMIME_Private_UnregisterWordA(This,hKL,szReading,dwStyle,szUnregister)	\
    (This)->lpVtbl -> UnregisterWordA(This,hKL,szReading,dwStyle,szUnregister)

#define IActiveIMMIME_Private_UnregisterWordW(This,hKL,szReading,dwStyle,szUnregister)	\
    (This)->lpVtbl -> UnregisterWordW(This,hKL,szReading,dwStyle,szUnregister)

#define IActiveIMMIME_Private_GenerateMessage(This,hIMC)	\
    (This)->lpVtbl -> GenerateMessage(This,hIMC)

#define IActiveIMMIME_Private_LockIMC(This,hIMC,ppIMC)	\
    (This)->lpVtbl -> LockIMC(This,hIMC,ppIMC)

#define IActiveIMMIME_Private_UnlockIMC(This,hIMC)	\
    (This)->lpVtbl -> UnlockIMC(This,hIMC)

#define IActiveIMMIME_Private_GetIMCLockCount(This,hIMC,pdwLockCount)	\
    (This)->lpVtbl -> GetIMCLockCount(This,hIMC,pdwLockCount)

#define IActiveIMMIME_Private_CreateIMCC(This,dwSize,phIMCC)	\
    (This)->lpVtbl -> CreateIMCC(This,dwSize,phIMCC)

#define IActiveIMMIME_Private_DestroyIMCC(This,hIMCC)	\
    (This)->lpVtbl -> DestroyIMCC(This,hIMCC)

#define IActiveIMMIME_Private_LockIMCC(This,hIMCC,ppv)	\
    (This)->lpVtbl -> LockIMCC(This,hIMCC,ppv)

#define IActiveIMMIME_Private_UnlockIMCC(This,hIMCC)	\
    (This)->lpVtbl -> UnlockIMCC(This,hIMCC)

#define IActiveIMMIME_Private_ReSizeIMCC(This,hIMCC,dwSize,phIMCC)	\
    (This)->lpVtbl -> ReSizeIMCC(This,hIMCC,dwSize,phIMCC)

#define IActiveIMMIME_Private_GetIMCCSize(This,hIMCC,pdwSize)	\
    (This)->lpVtbl -> GetIMCCSize(This,hIMCC,pdwSize)

#define IActiveIMMIME_Private_GetIMCCLockCount(This,hIMCC,pdwLockCount)	\
    (This)->lpVtbl -> GetIMCCLockCount(This,hIMCC,pdwLockCount)

#define IActiveIMMIME_Private_GetHotKey(This,dwHotKeyID,puModifiers,puVKey,phKL)	\
    (This)->lpVtbl -> GetHotKey(This,dwHotKeyID,puModifiers,puVKey,phKL)

#define IActiveIMMIME_Private_SetHotKey(This,dwHotKeyID,uModifiers,uVKey,hKL)	\
    (This)->lpVtbl -> SetHotKey(This,dwHotKeyID,uModifiers,uVKey,hKL)

#define IActiveIMMIME_Private_AssociateContextEx(This,hWnd,hIMC,dwFlags)	\
    (This)->lpVtbl -> AssociateContextEx(This,hWnd,hIMC,dwFlags)

#define IActiveIMMIME_Private_DisableIME(This,idThread)	\
    (This)->lpVtbl -> DisableIME(This,idThread)

#define IActiveIMMIME_Private_GetImeMenuItemsA(This,hIMC,dwFlags,dwType,pImeParentMenu,pImeMenu,dwSize,pdwResult)	\
    (This)->lpVtbl -> GetImeMenuItemsA(This,hIMC,dwFlags,dwType,pImeParentMenu,pImeMenu,dwSize,pdwResult)

#define IActiveIMMIME_Private_GetImeMenuItemsW(This,hIMC,dwFlags,dwType,pImeParentMenu,pImeMenu,dwSize,pdwResult)	\
    (This)->lpVtbl -> GetImeMenuItemsW(This,hIMC,dwFlags,dwType,pImeParentMenu,pImeMenu,dwSize,pdwResult)

#define IActiveIMMIME_Private_EnumInputContext(This,idThread,ppEnum)	\
    (This)->lpVtbl -> EnumInputContext(This,idThread,ppEnum)

#define IActiveIMMIME_Private_RequestMessageA(This,hIMC,wParam,lParam,plResult)	\
    (This)->lpVtbl -> RequestMessageA(This,hIMC,wParam,lParam,plResult)

#define IActiveIMMIME_Private_RequestMessageW(This,hIMC,wParam,lParam,plResult)	\
    (This)->lpVtbl -> RequestMessageW(This,hIMC,wParam,lParam,plResult)

#define IActiveIMMIME_Private_FilterClientWindowsEx(This,hWnd,fGuidMap)	\
    (This)->lpVtbl -> FilterClientWindowsEx(This,hWnd,fGuidMap)

#define IActiveIMMIME_Private_FilterClientWindows(This,aaClassList,uSize,aaGildMap)	\
    (This)->lpVtbl -> FilterClientWindows(This,aaClassList,uSize,aaGildMap)

#define IActiveIMMIME_Private_GetGuidAtom(This,hImc,bAttr,pGuidAtom)	\
    (This)->lpVtbl -> GetGuidAtom(This,hImc,bAttr,pGuidAtom)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_AssociateContext_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [in] */ HIMC hIME,
    /* [out] */ HIMC __RPC_FAR *phPrev);


void __RPC_STUB IActiveIMMIME_Private_AssociateContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_ConfigureIMEA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HWND hWnd,
    /* [in] */ DWORD dwMode,
    /* [in] */ REGISTERWORDA __RPC_FAR *pData);


void __RPC_STUB IActiveIMMIME_Private_ConfigureIMEA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_ConfigureIMEW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HWND hWnd,
    /* [in] */ DWORD dwMode,
    /* [in] */ REGISTERWORDW __RPC_FAR *pData);


void __RPC_STUB IActiveIMMIME_Private_ConfigureIMEW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_CreateContext_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [out] */ HIMC __RPC_FAR *phIMC);


void __RPC_STUB IActiveIMMIME_Private_CreateContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_DestroyContext_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIME);


void __RPC_STUB IActiveIMMIME_Private_DestroyContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_EnumRegisterWordA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ LPSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ LPSTR szRegister,
    /* [in] */ LPVOID pData,
    /* [out] */ IEnumRegisterWordA __RPC_FAR *__RPC_FAR *pEnum);


void __RPC_STUB IActiveIMMIME_Private_EnumRegisterWordA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_EnumRegisterWordW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ LPWSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ LPWSTR szRegister,
    /* [in] */ LPVOID pData,
    /* [out] */ IEnumRegisterWordW __RPC_FAR *__RPC_FAR *pEnum);


void __RPC_STUB IActiveIMMIME_Private_EnumRegisterWordW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_EscapeA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HIMC hIMC,
    /* [in] */ UINT uEscape,
    /* [out][in] */ LPVOID pData,
    /* [out] */ LRESULT __RPC_FAR *plResult);


void __RPC_STUB IActiveIMMIME_Private_EscapeA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_EscapeW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HIMC hIMC,
    /* [in] */ UINT uEscape,
    /* [out][in] */ LPVOID pData,
    /* [out] */ LRESULT __RPC_FAR *plResult);


void __RPC_STUB IActiveIMMIME_Private_EscapeW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCandidateListA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ UINT uBufLen,
    /* [out] */ CANDIDATELIST __RPC_FAR *pCandList,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetCandidateListA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCandidateListW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ UINT uBufLen,
    /* [out] */ CANDIDATELIST __RPC_FAR *pCandList,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetCandidateListW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCandidateListCountA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ DWORD __RPC_FAR *pdwListSize,
    /* [out] */ DWORD __RPC_FAR *pdwBufLen);


void __RPC_STUB IActiveIMMIME_Private_GetCandidateListCountA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCandidateListCountW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ DWORD __RPC_FAR *pdwListSize,
    /* [out] */ DWORD __RPC_FAR *pdwBufLen);


void __RPC_STUB IActiveIMMIME_Private_GetCandidateListCountW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCandidateWindow_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [out] */ CANDIDATEFORM __RPC_FAR *pCandidate);


void __RPC_STUB IActiveIMMIME_Private_GetCandidateWindow_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCompositionFontA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ LOGFONTA __RPC_FAR *plf);


void __RPC_STUB IActiveIMMIME_Private_GetCompositionFontA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCompositionFontW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ LOGFONTW __RPC_FAR *plf);


void __RPC_STUB IActiveIMMIME_Private_GetCompositionFontW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCompositionStringA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ DWORD dwBufLen,
    /* [out] */ LONG __RPC_FAR *plCopied,
    /* [out] */ LPVOID pBuf);


void __RPC_STUB IActiveIMMIME_Private_GetCompositionStringA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCompositionStringW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ DWORD dwBufLen,
    /* [out] */ LONG __RPC_FAR *plCopied,
    /* [out] */ LPVOID pBuf);


void __RPC_STUB IActiveIMMIME_Private_GetCompositionStringW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetCompositionWindow_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ COMPOSITIONFORM __RPC_FAR *pCompForm);


void __RPC_STUB IActiveIMMIME_Private_GetCompositionWindow_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetContext_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [out] */ HIMC __RPC_FAR *phIMC);


void __RPC_STUB IActiveIMMIME_Private_GetContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetConversionListA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HIMC hIMC,
    /* [in] */ _In_z_ LPSTR pSrc,
    /* [in] */ UINT uBufLen,
    /* [in] */ UINT uFlag,
    /* [out] */ CANDIDATELIST __RPC_FAR *pDst,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetConversionListA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetConversionListW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HIMC hIMC,
    /* [in] */ _In_z_ LPWSTR pSrc,
    /* [in] */ UINT uBufLen,
    /* [in] */ UINT uFlag,
    /* [out] */ CANDIDATELIST __RPC_FAR *pDst,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetConversionListW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetConversionStatus_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ DWORD __RPC_FAR *pfdwConversion,
    /* [out] */ DWORD __RPC_FAR *pfdwSentence);


void __RPC_STUB IActiveIMMIME_Private_GetConversionStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetDefaultIMEWnd_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [out] */ HWND __RPC_FAR *phDefWnd);


void __RPC_STUB IActiveIMMIME_Private_GetDefaultIMEWnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetDescriptionA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ UINT uBufLen,
    /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPSTR szDescription,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetDescriptionA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetDescriptionW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ UINT uBufLen,
    /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPWSTR szDescription,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetDescriptionW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetGuideLineA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ DWORD dwBufLen,
    /* [out] */ _Out_z_cap_post_count_(dwBufLen, *pdwResult) LPSTR pBuf,
    /* [out] */ DWORD __RPC_FAR *pdwResult);


void __RPC_STUB IActiveIMMIME_Private_GetGuideLineA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetGuideLineW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ DWORD dwBufLen,
    /* [out] */ _Out_z_cap_post_count_(dwBufLen, *pdwResult) LPWSTR pBuf,
    /* [out] */ DWORD __RPC_FAR *pdwResult);


void __RPC_STUB IActiveIMMIME_Private_GetGuideLineW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetIMEFileNameA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ UINT uBufLen,
    /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPSTR szFileName,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetIMEFileNameA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetIMEFileNameW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ UINT uBufLen,
    /* [out] */ _Out_z_cap_post_count_(uBufLen, *puCopied) LPWSTR szFileName,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetIMEFileNameW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetOpenStatus_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC);


void __RPC_STUB IActiveIMMIME_Private_GetOpenStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetProperty_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ DWORD fdwIndex,
    /* [out] */ DWORD __RPC_FAR *pdwProperty);


void __RPC_STUB IActiveIMMIME_Private_GetProperty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetRegisterWordStyleA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ UINT nItem,
    /* [out] */ STYLEBUFA __RPC_FAR *pStyleBuf,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetRegisterWordStyleA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetRegisterWordStyleW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ UINT nItem,
    /* [out] */ STYLEBUFW __RPC_FAR *pStyleBuf,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIMMIME_Private_GetRegisterWordStyleW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetStatusWindowPos_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ POINT __RPC_FAR *pptPos);


void __RPC_STUB IActiveIMMIME_Private_GetStatusWindowPos_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetVirtualKey_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [out] */ UINT __RPC_FAR *puVirtualKey);


void __RPC_STUB IActiveIMMIME_Private_GetVirtualKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_InstallIMEA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ _In_z_ LPSTR szIMEFileName,
    /* [in] */ _In_z_ LPSTR szLayoutText,
    /* [out] */ HKL __RPC_FAR *phKL);


void __RPC_STUB IActiveIMMIME_Private_InstallIMEA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_InstallIMEW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ _In_z_ LPWSTR szIMEFileName,
    /* [in] */ _In_z_ LPWSTR szLayoutText,
    /* [out] */ HKL __RPC_FAR *phKL);


void __RPC_STUB IActiveIMMIME_Private_InstallIMEW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_IsIME_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL);


void __RPC_STUB IActiveIMMIME_Private_IsIME_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_IsUIMessageA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWndIME,
    /* [in] */ UINT msg,
    /* [in] */ WPARAM wParam,
    /* [in] */ LPARAM lParam);


void __RPC_STUB IActiveIMMIME_Private_IsUIMessageA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_IsUIMessageW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWndIME,
    /* [in] */ UINT msg,
    /* [in] */ WPARAM wParam,
    /* [in] */ LPARAM lParam);


void __RPC_STUB IActiveIMMIME_Private_IsUIMessageW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_NotifyIME_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwAction,
    /* [in] */ DWORD dwIndex,
    /* [in] */ DWORD dwValue);


void __RPC_STUB IActiveIMMIME_Private_NotifyIME_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_RegisterWordA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ _In_z_ LPSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPSTR szRegister);


void __RPC_STUB IActiveIMMIME_Private_RegisterWordA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_RegisterWordW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ _In_z_ LPWSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPWSTR szRegister);


void __RPC_STUB IActiveIMMIME_Private_RegisterWordW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_ReleaseContext_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [in] */ HIMC hIMC);


void __RPC_STUB IActiveIMMIME_Private_ReleaseContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetCandidateWindow_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ CANDIDATEFORM __RPC_FAR *pCandidate);


void __RPC_STUB IActiveIMMIME_Private_SetCandidateWindow_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetCompositionFontA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ LOGFONTA __RPC_FAR *plf);


void __RPC_STUB IActiveIMMIME_Private_SetCompositionFontA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetCompositionFontW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ LOGFONTW __RPC_FAR *plf);


void __RPC_STUB IActiveIMMIME_Private_SetCompositionFontW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetCompositionStringA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ LPVOID pComp,
    /* [in] */ DWORD dwCompLen,
    /* [in] */ LPVOID pRead,
    /* [in] */ DWORD dwReadLen);


void __RPC_STUB IActiveIMMIME_Private_SetCompositionStringA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetCompositionStringW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ LPVOID pComp,
    /* [in] */ DWORD dwCompLen,
    /* [in] */ LPVOID pRead,
    /* [in] */ DWORD dwReadLen);


void __RPC_STUB IActiveIMMIME_Private_SetCompositionStringW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetCompositionWindow_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ COMPOSITIONFORM __RPC_FAR *pCompForm);


void __RPC_STUB IActiveIMMIME_Private_SetCompositionWindow_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetConversionStatus_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD fdwConversion,
    /* [in] */ DWORD fdwSentence);


void __RPC_STUB IActiveIMMIME_Private_SetConversionStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetOpenStatus_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ BOOL fOpen);


void __RPC_STUB IActiveIMMIME_Private_SetOpenStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetStatusWindowPos_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ POINT __RPC_FAR *pptPos);


void __RPC_STUB IActiveIMMIME_Private_SetStatusWindowPos_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SimulateHotKey_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [in] */ DWORD dwHotKeyID);


void __RPC_STUB IActiveIMMIME_Private_SimulateHotKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_UnregisterWordA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ _In_z_ LPSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPSTR szUnregister);


void __RPC_STUB IActiveIMMIME_Private_UnregisterWordA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_UnregisterWordW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ _In_z_ LPWSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPWSTR szUnregister);


void __RPC_STUB IActiveIMMIME_Private_UnregisterWordW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GenerateMessage_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC);


void __RPC_STUB IActiveIMMIME_Private_GenerateMessage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_LockIMC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ INPUTCONTEXT __RPC_FAR *__RPC_FAR *ppIMC);


void __RPC_STUB IActiveIMMIME_Private_LockIMC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_UnlockIMC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC);


void __RPC_STUB IActiveIMMIME_Private_UnlockIMC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetIMCLockCount_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [out] */ DWORD __RPC_FAR *pdwLockCount);


void __RPC_STUB IActiveIMMIME_Private_GetIMCLockCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_CreateIMCC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ DWORD dwSize,
    /* [out] */ HIMCC __RPC_FAR *phIMCC);


void __RPC_STUB IActiveIMMIME_Private_CreateIMCC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_DestroyIMCC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMCC hIMCC);


void __RPC_STUB IActiveIMMIME_Private_DestroyIMCC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_LockIMCC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMCC hIMCC,
    /* [out] */ void __RPC_FAR *__RPC_FAR *ppv);


void __RPC_STUB IActiveIMMIME_Private_LockIMCC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_UnlockIMCC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMCC hIMCC);


void __RPC_STUB IActiveIMMIME_Private_UnlockIMCC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_ReSizeIMCC_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMCC hIMCC,
    /* [in] */ DWORD dwSize,
    /* [out] */ HIMCC __RPC_FAR *phIMCC);


void __RPC_STUB IActiveIMMIME_Private_ReSizeIMCC_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetIMCCSize_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMCC hIMCC,
    /* [out] */ DWORD __RPC_FAR *pdwSize);


void __RPC_STUB IActiveIMMIME_Private_GetIMCCSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetIMCCLockCount_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMCC hIMCC,
    /* [out] */ DWORD __RPC_FAR *pdwLockCount);


void __RPC_STUB IActiveIMMIME_Private_GetIMCCLockCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetHotKey_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ DWORD dwHotKeyID,
    /* [out] */ UINT __RPC_FAR *puModifiers,
    /* [out] */ UINT __RPC_FAR *puVKey,
    /* [out] */ HKL __RPC_FAR *phKL);


void __RPC_STUB IActiveIMMIME_Private_GetHotKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_SetHotKey_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ DWORD dwHotKeyID,
    /* [in] */ UINT uModifiers,
    /* [in] */ UINT uVKey,
    /* [in] */ HKL hKL);


void __RPC_STUB IActiveIMMIME_Private_SetHotKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_AssociateContextEx_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwFlags);


void __RPC_STUB IActiveIMMIME_Private_AssociateContextEx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_DisableIME_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ DWORD idThread);


void __RPC_STUB IActiveIMMIME_Private_DisableIME_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetImeMenuItemsA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD dwType,
    /* [in] */ IMEMENUITEMINFOA __RPC_FAR *pImeParentMenu,
    /* [out] */ IMEMENUITEMINFOA __RPC_FAR *pImeMenu,
    /* [in] */ DWORD dwSize,
    /* [out] */ DWORD __RPC_FAR *pdwResult);


void __RPC_STUB IActiveIMMIME_Private_GetImeMenuItemsA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetImeMenuItemsW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwFlags,
    /* [in] */ DWORD dwType,
    /* [in] */ IMEMENUITEMINFOW __RPC_FAR *pImeParentMenu,
    /* [out] */ IMEMENUITEMINFOW __RPC_FAR *pImeMenu,
    /* [in] */ DWORD dwSize,
    /* [out] */ DWORD __RPC_FAR *pdwResult);


void __RPC_STUB IActiveIMMIME_Private_GetImeMenuItemsW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_EnumInputContext_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ DWORD idThread,
    /* [out] */ IEnumInputContext __RPC_FAR *__RPC_FAR *ppEnum);


void __RPC_STUB IActiveIMMIME_Private_EnumInputContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_RequestMessageA_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ WPARAM wParam,
    /* [in] */ LPARAM lParam,
    /* [out] */ LRESULT __RPC_FAR *plResult);


void __RPC_STUB IActiveIMMIME_Private_RequestMessageA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_RequestMessageW_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ WPARAM wParam,
    /* [in] */ LPARAM lParam,
    /* [out] */ LRESULT __RPC_FAR *plResult);


void __RPC_STUB IActiveIMMIME_Private_RequestMessageW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_FilterClientWindowsEx_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HWND hWnd,
    /* [in] */ BOOL fGuidMap);


void __RPC_STUB IActiveIMMIME_Private_FilterClientWindowsEx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_FilterClientWindows_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ ATOM __RPC_FAR *aaClassList,
    /* [in] */ UINT uSize,
    /* [in] */ BOOL __RPC_FAR *aaGildMap);


void __RPC_STUB IActiveIMMIME_Private_FilterClientWindows_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIMMIME_Private_GetGuidAtom_Proxy( 
    IActiveIMMIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hImc,
    /* [in] */ BYTE bAttr,
    /* [out] */ TfGuidAtom __RPC_FAR *pGuidAtom);


void __RPC_STUB IActiveIMMIME_Private_GetGuidAtom_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IActiveIMMIME_Private_INTERFACE_DEFINED__ */


#ifndef __IActiveIME_Private_INTERFACE_DEFINED__
#define __IActiveIME_Private_INTERFACE_DEFINED__

/* interface IActiveIME_Private */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IActiveIME_Private;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("5276cbe5-2130-4a60-bb76-8a5135448dec")
    IActiveIME_Private : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE ConnectIMM( 
            /* [in] */ IActiveIMMIME_Private __RPC_FAR *pActiveIMM) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UnconnectIMM( void) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Inquire( 
            /* [in] */ DWORD dwSystemInfoFlags,
            /* [out] */ IMEINFO __RPC_FAR *pIMEInfo,
            /* [out] */ _Out_z_cap_(*pdwPrivate) LPWSTR szWndClass,
            /* [out] */ DWORD __RPC_FAR *pdwPrivate) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ConversionList( 
            /* [in] */ HIMC hIMC,
            /* [in] */ _In_z_ LPWSTR szSource,
            /* [in] */ UINT uFlag,
            /* [in] */ UINT uBufLen,
            /* [out] */ CANDIDATELIST __RPC_FAR *pDest,
            /* [out] */ UINT __RPC_FAR *puCopied) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Configure( 
            /* [in] */ HKL hKL,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwMode,
            /* [in] */ REGISTERWORDW __RPC_FAR *pRegisterWord) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Destroy( 
            /* [in] */ UINT uReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Escape( 
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uEscape,
            /* [out][in] */ void __RPC_FAR *pData,
            /* [out] */ LRESULT __RPC_FAR *plResult) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ProcessKey( 
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uVirKey,
            /* [in] */ DWORD lParam,
            /* [in] */ BYTE __RPC_FAR *pbKeyState) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Notify( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwAction,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwValue) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SelectEx( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetCompositionString( 
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ void __RPC_FAR *pComp,
            /* [in] */ DWORD dwCompLen,
            /* [in] */ void __RPC_FAR *pRead,
            /* [in] */ DWORD dwReadLen) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ToAsciiEx( 
            /* [in] */ UINT uVirKey,
            /* [in] */ UINT uScanCode,
            /* [in] */ BYTE __RPC_FAR *pbKeyState,
            /* [in] */ UINT fuState,
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwTransBuf,
            /* [out] */ UINT __RPC_FAR *puSize) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RegisterWord( 
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szString) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UnregisterWord( 
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szString) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetRegisterWordStyle( 
            /* [in] */ UINT nItem,
            /* [out] */ STYLEBUFW __RPC_FAR *pStyleBuf,
            /* [out] */ UINT __RPC_FAR *puBufSize) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EnumRegisterWord( 
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szRegister,
            /* [in] */ LPVOID pData,
            /* [out] */ IEnumRegisterWordW __RPC_FAR *__RPC_FAR *ppEnum) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCodePageA( 
            /* [out] */ UINT __RPC_FAR *uCodePage) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetLangId( 
            /* [out] */ LANGID __RPC_FAR *plid) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE AssociateFocus( 
            /* [in] */ HWND hwnd,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IActiveIME_PrivateVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IActiveIME_Private __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IActiveIME_Private __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ConnectIMM )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ IActiveIMMIME_Private __RPC_FAR *pActiveIMM);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UnconnectIMM )( 
            IActiveIME_Private __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Inquire )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ DWORD dwSystemInfoFlags,
            /* [out] */ IMEINFO __RPC_FAR *pIMEInfo,
            /* [out] */ _Out_z_cap_(*pdwPrivate) LPWSTR szWndClass,
            /* [out] */ DWORD __RPC_FAR *pdwPrivate);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ConversionList )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ _In_z_ LPWSTR szSource,
            /* [in] */ UINT uFlag,
            /* [in] */ UINT uBufLen,
            /* [out] */ CANDIDATELIST __RPC_FAR *pDest,
            /* [out] */ UINT __RPC_FAR *puCopied);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Configure )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ HWND hWnd,
            /* [in] */ DWORD dwMode,
            /* [in] */ REGISTERWORDW __RPC_FAR *pRegisterWord);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Destroy )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ UINT uReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Escape )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uEscape,
            /* [out][in] */ void __RPC_FAR *pData,
            /* [out] */ LRESULT __RPC_FAR *plResult);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ProcessKey )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ UINT uVirKey,
            /* [in] */ DWORD lParam,
            /* [in] */ BYTE __RPC_FAR *pbKeyState);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Notify )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwAction,
            /* [in] */ DWORD dwIndex,
            /* [in] */ DWORD dwValue);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SelectEx )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetCompositionString )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwIndex,
            /* [in] */ void __RPC_FAR *pComp,
            /* [in] */ DWORD dwCompLen,
            /* [in] */ void __RPC_FAR *pRead,
            /* [in] */ DWORD dwReadLen);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ToAsciiEx )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ UINT uVirKey,
            /* [in] */ UINT uScanCode,
            /* [in] */ BYTE __RPC_FAR *pbKeyState,
            /* [in] */ UINT fuState,
            /* [in] */ HIMC hIMC,
            /* [out] */ DWORD __RPC_FAR *pdwTransBuf,
            /* [out] */ UINT __RPC_FAR *puSize);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RegisterWord )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szString);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UnregisterWord )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szString);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetRegisterWordStyle )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ UINT nItem,
            /* [out] */ STYLEBUFW __RPC_FAR *pStyleBuf,
            /* [out] */ UINT __RPC_FAR *puBufSize);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EnumRegisterWord )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ _In_z_ LPWSTR szReading,
            /* [in] */ DWORD dwStyle,
            /* [in] */ _In_z_ LPWSTR szRegister,
            /* [in] */ LPVOID pData,
            /* [out] */ IEnumRegisterWordW __RPC_FAR *__RPC_FAR *ppEnum);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCodePageA )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *uCodePage);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLangId )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [out] */ LANGID __RPC_FAR *plid);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AssociateFocus )( 
            IActiveIME_Private __RPC_FAR * This,
            /* [in] */ HWND hwnd,
            /* [in] */ HIMC hIMC,
            /* [in] */ DWORD dwFlags);
        
        END_INTERFACE
    } IActiveIME_PrivateVtbl;

    interface IActiveIME_Private
    {
        CONST_VTBL struct IActiveIME_PrivateVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IActiveIME_Private_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IActiveIME_Private_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IActiveIME_Private_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IActiveIME_Private_ConnectIMM(This,pActiveIMM)	\
    (This)->lpVtbl -> ConnectIMM(This,pActiveIMM)

#define IActiveIME_Private_UnconnectIMM(This)	\
    (This)->lpVtbl -> UnconnectIMM(This)

#define IActiveIME_Private_Inquire(This,dwSystemInfoFlags,pIMEInfo,szWndClass,pdwPrivate)	\
    (This)->lpVtbl -> Inquire(This,dwSystemInfoFlags,pIMEInfo,szWndClass,pdwPrivate)

#define IActiveIME_Private_ConversionList(This,hIMC,szSource,uFlag,uBufLen,pDest,puCopied)	\
    (This)->lpVtbl -> ConversionList(This,hIMC,szSource,uFlag,uBufLen,pDest,puCopied)

#define IActiveIME_Private_Configure(This,hKL,hWnd,dwMode,pRegisterWord)	\
    (This)->lpVtbl -> Configure(This,hKL,hWnd,dwMode,pRegisterWord)

#define IActiveIME_Private_Destroy(This,uReserved)	\
    (This)->lpVtbl -> Destroy(This,uReserved)

#define IActiveIME_Private_Escape(This,hIMC,uEscape,pData,plResult)	\
    (This)->lpVtbl -> Escape(This,hIMC,uEscape,pData,plResult)

#define IActiveIME_Private_ProcessKey(This,hIMC,uVirKey,lParam,pbKeyState)	\
    (This)->lpVtbl -> ProcessKey(This,hIMC,uVirKey,lParam,pbKeyState)

#define IActiveIME_Private_Notify(This,hIMC,dwAction,dwIndex,dwValue)	\
    (This)->lpVtbl -> Notify(This,hIMC,dwAction,dwIndex,dwValue)

#define IActiveIME_Private_SelectEx(This,hIMC,dwFlags)	\
    (This)->lpVtbl -> SelectEx(This,hIMC,dwFlags)

#define IActiveIME_Private_SetCompositionString(This,hIMC,dwIndex,pComp,dwCompLen,pRead,dwReadLen)	\
    (This)->lpVtbl -> SetCompositionString(This,hIMC,dwIndex,pComp,dwCompLen,pRead,dwReadLen)

#define IActiveIME_Private_ToAsciiEx(This,uVirKey,uScanCode,pbKeyState,fuState,hIMC,pdwTransBuf,puSize)	\
    (This)->lpVtbl -> ToAsciiEx(This,uVirKey,uScanCode,pbKeyState,fuState,hIMC,pdwTransBuf,puSize)

#define IActiveIME_Private_RegisterWord(This,szReading,dwStyle,szString)	\
    (This)->lpVtbl -> RegisterWord(This,szReading,dwStyle,szString)

#define IActiveIME_Private_UnregisterWord(This,szReading,dwStyle,szString)	\
    (This)->lpVtbl -> UnregisterWord(This,szReading,dwStyle,szString)

#define IActiveIME_Private_GetRegisterWordStyle(This,nItem,pStyleBuf,puBufSize)	\
    (This)->lpVtbl -> GetRegisterWordStyle(This,nItem,pStyleBuf,puBufSize)

#define IActiveIME_Private_EnumRegisterWord(This,szReading,dwStyle,szRegister,pData,ppEnum)	\
    (This)->lpVtbl -> EnumRegisterWord(This,szReading,dwStyle,szRegister,pData,ppEnum)

#define IActiveIME_Private_GetCodePageA(This,uCodePage)	\
    (This)->lpVtbl -> GetCodePageA(This,uCodePage)

#define IActiveIME_Private_GetLangId(This,plid)	\
    (This)->lpVtbl -> GetLangId(This,plid)

#define IActiveIME_Private_AssociateFocus(This,hwnd,hIMC,dwFlags)	\
    (This)->lpVtbl -> AssociateFocus(This,hwnd,hIMC,dwFlags)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IActiveIME_Private_ConnectIMM_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ IActiveIMMIME_Private __RPC_FAR *pActiveIMM);


void __RPC_STUB IActiveIME_Private_ConnectIMM_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_UnconnectIMM_Proxy( 
    IActiveIME_Private __RPC_FAR * This);


void __RPC_STUB IActiveIME_Private_UnconnectIMM_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_Inquire_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ DWORD dwSystemInfoFlags,
    /* [out] */ IMEINFO __RPC_FAR *pIMEInfo,
    /* [out] */ _Out_z_cap_(*pdwPrivate) LPWSTR szWndClass,
    /* [out] */ DWORD __RPC_FAR *pdwPrivate);


void __RPC_STUB IActiveIME_Private_Inquire_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_ConversionList_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ _In_z_ LPWSTR szSource,
    /* [in] */ UINT uFlag,
    /* [in] */ UINT uBufLen,
    /* [out] */ CANDIDATELIST __RPC_FAR *pDest,
    /* [out] */ UINT __RPC_FAR *puCopied);


void __RPC_STUB IActiveIME_Private_ConversionList_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_Configure_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ HWND hWnd,
    /* [in] */ DWORD dwMode,
    /* [in] */ REGISTERWORDW __RPC_FAR *pRegisterWord);


void __RPC_STUB IActiveIME_Private_Configure_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_Destroy_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ UINT uReserved);


void __RPC_STUB IActiveIME_Private_Destroy_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_Escape_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ UINT uEscape,
    /* [out][in] */ void __RPC_FAR *pData,
    /* [out] */ LRESULT __RPC_FAR *plResult);


void __RPC_STUB IActiveIME_Private_Escape_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_ProcessKey_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ UINT uVirKey,
    /* [in] */ DWORD lParam,
    /* [in] */ BYTE __RPC_FAR *pbKeyState);


void __RPC_STUB IActiveIME_Private_ProcessKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_Notify_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwAction,
    /* [in] */ DWORD dwIndex,
    /* [in] */ DWORD dwValue);


void __RPC_STUB IActiveIME_Private_Notify_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_SelectEx_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwFlags);


void __RPC_STUB IActiveIME_Private_SelectEx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_SetCompositionString_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwIndex,
    /* [in] */ void __RPC_FAR *pComp,
    /* [in] */ DWORD dwCompLen,
    /* [in] */ void __RPC_FAR *pRead,
    /* [in] */ DWORD dwReadLen);


void __RPC_STUB IActiveIME_Private_SetCompositionString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_ToAsciiEx_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ UINT uVirKey,
    /* [in] */ UINT uScanCode,
    /* [in] */ BYTE __RPC_FAR *pbKeyState,
    /* [in] */ UINT fuState,
    /* [in] */ HIMC hIMC,
    /* [out] */ DWORD __RPC_FAR *pdwTransBuf,
    /* [out] */ UINT __RPC_FAR *puSize);


void __RPC_STUB IActiveIME_Private_ToAsciiEx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_RegisterWord_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ _In_z_ LPWSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPWSTR szString);


void __RPC_STUB IActiveIME_Private_RegisterWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_UnregisterWord_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ _In_z_ LPWSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPWSTR szString);


void __RPC_STUB IActiveIME_Private_UnregisterWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_GetRegisterWordStyle_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ UINT nItem,
    /* [out] */ STYLEBUFW __RPC_FAR *pStyleBuf,
    /* [out] */ UINT __RPC_FAR *puBufSize);


void __RPC_STUB IActiveIME_Private_GetRegisterWordStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_EnumRegisterWord_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ _In_z_ LPWSTR szReading,
    /* [in] */ DWORD dwStyle,
    /* [in] */ _In_z_ LPWSTR szRegister,
    /* [in] */ LPVOID pData,
    /* [out] */ IEnumRegisterWordW __RPC_FAR *__RPC_FAR *ppEnum);


void __RPC_STUB IActiveIME_Private_EnumRegisterWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_GetCodePageA_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [out] */ UINT __RPC_FAR *uCodePage);


void __RPC_STUB IActiveIME_Private_GetCodePageA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_GetLangId_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [out] */ LANGID __RPC_FAR *plid);


void __RPC_STUB IActiveIME_Private_GetLangId_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IActiveIME_Private_AssociateFocus_Proxy( 
    IActiveIME_Private __RPC_FAR * This,
    /* [in] */ HWND hwnd,
    /* [in] */ HIMC hIMC,
    /* [in] */ DWORD dwFlags);


void __RPC_STUB IActiveIME_Private_AssociateFocus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IActiveIME_Private_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
