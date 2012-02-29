/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Aug 10 15:14:28 2000
 */
/* Compiler settings for msimetip.idl:
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

#ifndef __msimetip_h__
#define __msimetip_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IIMECheckDefaultInputProfile_FWD_DEFINED__
#define __IIMECheckDefaultInputProfile_FWD_DEFINED__
typedef interface IIMECheckDefaultInputProfile IIMECheckDefaultInputProfile;
#endif 	/* __IIMECheckDefaultInputProfile_FWD_DEFINED__ */


#ifndef __IMECheckDefaultInputProfile_FWD_DEFINED__
#define __IMECheckDefaultInputProfile_FWD_DEFINED__

#ifdef __cplusplus
typedef class IMECheckDefaultInputProfile IMECheckDefaultInputProfile;
#else
typedef struct IMECheckDefaultInputProfile IMECheckDefaultInputProfile;
#endif /* __cplusplus */

#endif 	/* __IMECheckDefaultInputProfile_FWD_DEFINED__ */


#ifndef __IIMEFNIMEPad_FWD_DEFINED__
#define __IIMEFNIMEPad_FWD_DEFINED__
typedef interface IIMEFNIMEPad IIMEFNIMEPad;
#endif 	/* __IIMEFNIMEPad_FWD_DEFINED__ */


/* header files for imported files */
#include "unknwn.h"

__post __maybenull
__post __writableTo(byteCount(size))
void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t size);
void __RPC_USER MIDL_user_free( __inout void __RPC_FAR * ); 

/* interface __MIDL_itf_msimetip_0000 */
/* [local] */ 

//=--------------------------------------------------------------------------=
// msimetip.h


// SATORI TIP Defaulting declarations.

//=--------------------------------------------------------------------------=
// (C) Copyright 1995-1999 Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=

#ifndef _MSIMETIP_
#define _MSIMETIP_

#include <windows.h>

#ifdef __cplusplus
extern "C" {
#endif /* __cplusplus */

#ifdef __cplusplus
}
#endif  /* __cplusplus */

//
// error type (IsTipActive)
//
typedef /* [public] */ 
enum __MIDL___MIDL_itf_msimetip_0000_0001
    {	ISTIPACTIVEERR_ERRORUNKNOWN	= 1,
	ISTIPACTIVEERR_DIFFERENTVERSION	= ISTIPACTIVEERR_ERRORUNKNOWN + 1,
	ISTIPACTIVEERR_CURLANGISDIFFERENT	= ISTIPACTIVEERR_DIFFERENTVERSION + 1,
	ISTIPACTIVEERR_NOTDEFAULT	= ISTIPACTIVEERR_CURLANGISDIFFERENT + 1,
	ISTIPACTIVEERR_NOTFOUND	= ISTIPACTIVEERR_NOTDEFAULT + 1
    }	ISTIPACTIVATE;

//
// TIP type (IsTipActive, Activate)
//
typedef /* [public] */ 
enum __MIDL___MIDL_itf_msimetip_0000_0002
    {	TIPTYPE_DEFAULT	= 0,
	TIPTYPE_SEAMLESS	= 0x1,
	TIPTYPE_CLASSIC	= 0x2
    }	TIPTYPE;

//
// HKEY_CLASSES_ROOT values
//
#define SZIMECheckDefaultInputProfile	"IMECheckDefaultInputProfile.Japan"
#define SZIMECheckDefaultInputProfile_1	"IMECheckDefaultInputProfile.Japan.1"


extern RPC_IF_HANDLE __MIDL_itf_msimetip_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_msimetip_0000_v0_0_s_ifspec;

#ifndef __IIMECheckDefaultInputProfile_INTERFACE_DEFINED__
#define __IIMECheckDefaultInputProfile_INTERFACE_DEFINED__

/* interface IIMECheckDefaultInputProfile */
/* [helpstring][unique][uuid][object] */ 

typedef WORD LANGID;


EXTERN_C const IID IID_IIMECheckDefaultInputProfile;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("828C6122-9FE3-4ae4-8988-E673DC7D27AF")
    IIMECheckDefaultInputProfile : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE IsTipActive( 
            /* [in] */ REFCLSID rclsid,
            /* [out] */ DWORD __RPC_FAR *pdwReason) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ActivateTip( 
            /* [in] */ DWORD dwTIPType) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetDescription( 
            /* [in] */ LANGID langid,
            /* [out] */ _Out_z_cap_(*pcch) WCHAR __RPC_FAR *wsz,
            /* [out] */ UINT __RPC_FAR *pcch) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetIcon( 
            /* [in] */ BOOL fSmall,
            /* [out] */ HICON __RPC_FAR *phIcon) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IIMECheckDefaultInputProfileVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *IsTipActive )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This,
            /* [in] */ REFCLSID rclsid,
            /* [out] */ DWORD __RPC_FAR *pdwReason);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ActivateTip )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This,
            /* [in] */ DWORD dwTIPType);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetDescription )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This,
            /* [in] */ LANGID langid,
            /* [out] */ WCHAR __RPC_FAR *wsz,
            /* [out] */ UINT __RPC_FAR *pcch);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIcon )( 
            IIMECheckDefaultInputProfile __RPC_FAR * This,
            /* [in] */ BOOL fSmall,
            /* [out] */ HICON __RPC_FAR *phIcon);
        
        END_INTERFACE
    } IIMECheckDefaultInputProfileVtbl;

    interface IIMECheckDefaultInputProfile
    {
        CONST_VTBL struct IIMECheckDefaultInputProfileVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IIMECheckDefaultInputProfile_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IIMECheckDefaultInputProfile_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IIMECheckDefaultInputProfile_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IIMECheckDefaultInputProfile_IsTipActive(This,rclsid,pdwReason)	\
    (This)->lpVtbl -> IsTipActive(This,rclsid,pdwReason)

#define IIMECheckDefaultInputProfile_ActivateTip(This,dwTIPType)	\
    (This)->lpVtbl -> ActivateTip(This,dwTIPType)

#define IIMECheckDefaultInputProfile_GetDescription(This,langid,wsz,pcch)	\
    (This)->lpVtbl -> GetDescription(This,langid,wsz,pcch)

#define IIMECheckDefaultInputProfile_GetIcon(This,fSmall,phIcon)	\
    (This)->lpVtbl -> GetIcon(This,fSmall,phIcon)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IIMECheckDefaultInputProfile_IsTipActive_Proxy( 
    IIMECheckDefaultInputProfile __RPC_FAR * This,
    /* [in] */ REFCLSID rclsid,
    /* [out] */ DWORD __RPC_FAR *pdwReason);


void __RPC_STUB IIMECheckDefaultInputProfile_IsTipActive_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IIMECheckDefaultInputProfile_ActivateTip_Proxy( 
    IIMECheckDefaultInputProfile __RPC_FAR * This,
    /* [in] */ DWORD dwTIPType);


void __RPC_STUB IIMECheckDefaultInputProfile_ActivateTip_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IIMECheckDefaultInputProfile_GetDescription_Proxy( 
    IIMECheckDefaultInputProfile __RPC_FAR * This,
    /* [in] */ LANGID langid,
    /* [out] */ _Out_z_cap_(*pcch) WCHAR __RPC_FAR *wsz,
    /* [out] */ UINT __RPC_FAR *pcch);


void __RPC_STUB IIMECheckDefaultInputProfile_GetDescription_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IIMECheckDefaultInputProfile_GetIcon_Proxy( 
    IIMECheckDefaultInputProfile __RPC_FAR * This,
    /* [in] */ BOOL fSmall,
    /* [out] */ HICON __RPC_FAR *phIcon);


void __RPC_STUB IIMECheckDefaultInputProfile_GetIcon_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IIMECheckDefaultInputProfile_INTERFACE_DEFINED__ */



#ifndef __IMECheckDefaultInputProfile_LIBRARY_DEFINED__
#define __IMECheckDefaultInputProfile_LIBRARY_DEFINED__

/* library IMECheckDefaultInputProfile */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_IMECheckDefaultInputProfile;

EXTERN_C const CLSID CLSID_IMECheckDefaultInputProfile;

#ifdef __cplusplus

class DECLSPEC_UUID("1BA11148-1D90-49ab-AF4F-3688219D579C")
IMECheckDefaultInputProfile;
#endif
#endif /* __IMECheckDefaultInputProfile_LIBRARY_DEFINED__ */

#ifndef __IIMEFNIMEPad_INTERFACE_DEFINED__
#define __IIMEFNIMEPad_INTERFACE_DEFINED__

/* interface IIMEFNIMEPad */
/* [helpstring][unique][uuid][object] */ 


EXTERN_C const IID IID_IIMEFNIMEPad;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("90A5B5B4-E2C2-48d7-A90A-4384B369AF7D")
    IIMEFNIMEPad : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Activate( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IIMEFNIMEPadVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IIMEFNIMEPad __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IIMEFNIMEPad __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IIMEFNIMEPad __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Activate )( 
            IIMEFNIMEPad __RPC_FAR * This);
        
        END_INTERFACE
    } IIMEFNIMEPadVtbl;

    interface IIMEFNIMEPad
    {
        CONST_VTBL struct IIMEFNIMEPadVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IIMEFNIMEPad_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IIMEFNIMEPad_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IIMEFNIMEPad_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IIMEFNIMEPad_Activate(This)	\
    (This)->lpVtbl -> Activate(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IIMEFNIMEPad_Activate_Proxy( 
    IIMEFNIMEPad __RPC_FAR * This);


void __RPC_STUB IIMEFNIMEPad_Activate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IIMEFNIMEPad_INTERFACE_DEFINED__ */


/* interface __MIDL_itf_msimetip_0107 */
/* [local] */ 

//
//  Japanese Keyboard TIP
//
DEFINE_GUID( CLSID_JapaneseKbdTip,                	0xe9ba4710,	0x1d6a,	0x11d3,	0x99, 0x83, 0x00, 0xc0, 0x4f, 0x7a, 0xd1, 0xa3);
DEFINE_GUID( CLSID_IMECheckDefaultInputProfile,                	0x828C6122,0x9FE3,0x4ae4,0x89,0x88,0xE6,0x73,0xDC,0x7D,0x27,0xAF);
DEFINE_GUID( IID_IIMECheckDefaultInputProfile,                	0x1BA11148,0x1D90,0x49ab,0xAF,0x4F,0x36,0x88,0x21,0x9D,0x57,0x9C);
DEFINE_GUID( IID_IIMEFNIMEPad,                	0x90a5b5b4, 0xe2c2, 0x48d7, 0xa9, 0xa, 0x43, 0x84, 0xb3, 0x69, 0xaf, 0x7d);
#endif // _MSIMETIP_


extern RPC_IF_HANDLE __MIDL_itf_msimetip_0107_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_msimetip_0107_v0_0_s_ifspec;

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  HICON_UserSize(     unsigned long __RPC_FAR *, unsigned long            , HICON __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  HICON_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, HICON __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  HICON_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, HICON __RPC_FAR * ); 
void                      __RPC_USER  HICON_UserFree(     unsigned long __RPC_FAR *, HICON __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
