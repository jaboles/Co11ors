/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Wed Sep 13 22:33:23 2000
 */
/* Compiler settings for msuimw32.idl:
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

#ifndef __msuimw32_h__
#define __msuimw32_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IAImeProfile_FWD_DEFINED__
#define __IAImeProfile_FWD_DEFINED__
typedef interface IAImeProfile IAImeProfile;
#endif 	/* __IAImeProfile_FWD_DEFINED__ */


#ifndef __IAImeContext_FWD_DEFINED__
#define __IAImeContext_FWD_DEFINED__
typedef interface IAImeContext IAImeContext;
#endif 	/* __IAImeContext_FWD_DEFINED__ */


/* header files for imported files */
#include "unknwn.h"
#include "msctf.h"
#include "aimmp.h"

__post __maybenull
__post __writableTo(byteCount(size))
void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t size);
void __RPC_USER MIDL_user_free( __inout void __RPC_FAR * ); 

/* interface __MIDL_itf_msuimw32_0000 */
/* [local] */ 

//=--------------------------------------------------------------------------=
// msuimw32.h
//=--------------------------------------------------------------------------=
// (C) Copyright 1995-2000 Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=

#pragma comment(lib,"uuid.lib")

//--------------------------------------------------------------------------
// Win32 Layer Semi Private Interfaces.

#if 0
#endif


extern RPC_IF_HANDLE __MIDL_itf_msuimw32_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_msuimw32_0000_v0_0_s_ifspec;

#ifndef __IAImeProfile_INTERFACE_DEFINED__
#define __IAImeProfile_INTERFACE_DEFINED__

/* interface IAImeProfile */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IAImeProfile;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("B2AA53DF-21AB-40f2-B386-ED048CFC1C9D")
    IAImeProfile : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Activate( void) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Deactivate( void) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ChangeCurrentKeyboardLayout( 
            HKL hKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetLangId( 
            LANGID __RPC_FAR *plid) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCodePageA( 
            UINT __RPC_FAR *puCodePage) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetKeyboardLayout( 
            HKL __RPC_FAR *phkl) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE IsIME( 
            HKL hKL) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetActiveLanguageProfile( 
            /* [in] */ HKL hKL,
            /* [in] */ GUID catid,
            /* [out] */ TF_LANGUAGEPROFILE __RPC_FAR *pLanguageProfile) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IAImeProfileVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IAImeProfile __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IAImeProfile __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IAImeProfile __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Activate )( 
            IAImeProfile __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Deactivate )( 
            IAImeProfile __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ChangeCurrentKeyboardLayout )( 
            IAImeProfile __RPC_FAR * This,
            HKL hKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLangId )( 
            IAImeProfile __RPC_FAR * This,
            LANGID __RPC_FAR *plid);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCodePageA )( 
            IAImeProfile __RPC_FAR * This,
            UINT __RPC_FAR *puCodePage);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetKeyboardLayout )( 
            IAImeProfile __RPC_FAR * This,
            HKL __RPC_FAR *phkl);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *IsIME )( 
            IAImeProfile __RPC_FAR * This,
            HKL hKL);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetActiveLanguageProfile )( 
            IAImeProfile __RPC_FAR * This,
            /* [in] */ HKL hKL,
            /* [in] */ GUID catid,
            /* [out] */ TF_LANGUAGEPROFILE __RPC_FAR *pLanguageProfile);
        
        END_INTERFACE
    } IAImeProfileVtbl;

    interface IAImeProfile
    {
        CONST_VTBL struct IAImeProfileVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IAImeProfile_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IAImeProfile_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IAImeProfile_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IAImeProfile_Activate(This)	\
    (This)->lpVtbl -> Activate(This)

#define IAImeProfile_Deactivate(This)	\
    (This)->lpVtbl -> Deactivate(This)

#define IAImeProfile_ChangeCurrentKeyboardLayout(This,hKL)	\
    (This)->lpVtbl -> ChangeCurrentKeyboardLayout(This,hKL)

#define IAImeProfile_GetLangId(This,plid)	\
    (This)->lpVtbl -> GetLangId(This,plid)

#define IAImeProfile_GetCodePageA(This,puCodePage)	\
    (This)->lpVtbl -> GetCodePageA(This,puCodePage)

#define IAImeProfile_GetKeyboardLayout(This,phkl)	\
    (This)->lpVtbl -> GetKeyboardLayout(This,phkl)

#define IAImeProfile_IsIME(This,hKL)	\
    (This)->lpVtbl -> IsIME(This,hKL)

#define IAImeProfile_GetActiveLanguageProfile(This,hKL,catid,pLanguageProfile)	\
    (This)->lpVtbl -> GetActiveLanguageProfile(This,hKL,catid,pLanguageProfile)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IAImeProfile_Activate_Proxy( 
    IAImeProfile __RPC_FAR * This);


void __RPC_STUB IAImeProfile_Activate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_Deactivate_Proxy( 
    IAImeProfile __RPC_FAR * This);


void __RPC_STUB IAImeProfile_Deactivate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_ChangeCurrentKeyboardLayout_Proxy( 
    IAImeProfile __RPC_FAR * This,
    HKL hKL);


void __RPC_STUB IAImeProfile_ChangeCurrentKeyboardLayout_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_GetLangId_Proxy( 
    IAImeProfile __RPC_FAR * This,
    LANGID __RPC_FAR *plid);


void __RPC_STUB IAImeProfile_GetLangId_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_GetCodePageA_Proxy( 
    IAImeProfile __RPC_FAR * This,
    UINT __RPC_FAR *puCodePage);


void __RPC_STUB IAImeProfile_GetCodePageA_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_GetKeyboardLayout_Proxy( 
    IAImeProfile __RPC_FAR * This,
    HKL __RPC_FAR *phkl);


void __RPC_STUB IAImeProfile_GetKeyboardLayout_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_IsIME_Proxy( 
    IAImeProfile __RPC_FAR * This,
    HKL hKL);


void __RPC_STUB IAImeProfile_IsIME_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeProfile_GetActiveLanguageProfile_Proxy( 
    IAImeProfile __RPC_FAR * This,
    /* [in] */ HKL hKL,
    /* [in] */ GUID catid,
    /* [out] */ TF_LANGUAGEPROFILE __RPC_FAR *pLanguageProfile);


void __RPC_STUB IAImeProfile_GetActiveLanguageProfile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IAImeProfile_INTERFACE_DEFINED__ */


/* interface __MIDL_itf_msuimw32_0239 */
/* [local] */ 

#if 0
#endif


extern RPC_IF_HANDLE __MIDL_itf_msuimw32_0239_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_msuimw32_0239_v0_0_s_ifspec;

#ifndef __IAImeContext_INTERFACE_DEFINED__
#define __IAImeContext_INTERFACE_DEFINED__

/* interface IAImeContext */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IAImeContext;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("5F5B4ACB-D55D-492c-B596-F6390E1AD798")
    IAImeContext : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE CreateAImeContext( 
            HIMC hIMC,
            IActiveIME_Private __RPC_FAR *pActiveIME) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DestroyAImeContext( 
            HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE UpdateAImeContext( 
            HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE MapAttributes( 
            HIMC hIMC) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetGuidAtom( 
            HIMC hIMC,
            BYTE bAttr,
            TfGuidAtom __RPC_FAR *pGuidAtom) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IAImeContextVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IAImeContext __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IAImeContext __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IAImeContext __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateAImeContext )( 
            IAImeContext __RPC_FAR * This,
            HIMC hIMC,
            IActiveIME_Private __RPC_FAR *pActiveIME);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DestroyAImeContext )( 
            IAImeContext __RPC_FAR * This,
            HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UpdateAImeContext )( 
            IAImeContext __RPC_FAR * This,
            HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *MapAttributes )( 
            IAImeContext __RPC_FAR * This,
            HIMC hIMC);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGuidAtom )( 
            IAImeContext __RPC_FAR * This,
            HIMC hIMC,
            BYTE bAttr,
            TfGuidAtom __RPC_FAR *pGuidAtom);
        
        END_INTERFACE
    } IAImeContextVtbl;

    interface IAImeContext
    {
        CONST_VTBL struct IAImeContextVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IAImeContext_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IAImeContext_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IAImeContext_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IAImeContext_CreateAImeContext(This,hIMC,pActiveIME)	\
    (This)->lpVtbl -> CreateAImeContext(This,hIMC,pActiveIME)

#define IAImeContext_DestroyAImeContext(This,hIMC)	\
    (This)->lpVtbl -> DestroyAImeContext(This,hIMC)

#define IAImeContext_UpdateAImeContext(This,hIMC)	\
    (This)->lpVtbl -> UpdateAImeContext(This,hIMC)

#define IAImeContext_MapAttributes(This,hIMC)	\
    (This)->lpVtbl -> MapAttributes(This,hIMC)

#define IAImeContext_GetGuidAtom(This,hIMC,bAttr,pGuidAtom)	\
    (This)->lpVtbl -> GetGuidAtom(This,hIMC,bAttr,pGuidAtom)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IAImeContext_CreateAImeContext_Proxy( 
    IAImeContext __RPC_FAR * This,
    HIMC hIMC,
    IActiveIME_Private __RPC_FAR *pActiveIME);


void __RPC_STUB IAImeContext_CreateAImeContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeContext_DestroyAImeContext_Proxy( 
    IAImeContext __RPC_FAR * This,
    HIMC hIMC);


void __RPC_STUB IAImeContext_DestroyAImeContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeContext_UpdateAImeContext_Proxy( 
    IAImeContext __RPC_FAR * This,
    HIMC hIMC);


void __RPC_STUB IAImeContext_UpdateAImeContext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeContext_MapAttributes_Proxy( 
    IAImeContext __RPC_FAR * This,
    HIMC hIMC);


void __RPC_STUB IAImeContext_MapAttributes_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IAImeContext_GetGuidAtom_Proxy( 
    IAImeContext __RPC_FAR * This,
    HIMC hIMC,
    BYTE bAttr,
    TfGuidAtom __RPC_FAR *pGuidAtom);


void __RPC_STUB IAImeContext_GetGuidAtom_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IAImeContext_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
