
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 5.02.0235 */
/* at Mon Apr 05 13:25:23 1999
 */
/* Compiler settings for qexport.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32 (32b run), ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
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

#ifndef __qexport_h__
#define __qexport_h__

/* Forward Declarations */ 

#ifndef __IPhraseSink_FWD_DEFINED__
#define __IPhraseSink_FWD_DEFINED__
typedef interface IPhraseSink IPhraseSink;
#endif 	/* __IPhraseSink_FWD_DEFINED__ */


#ifndef __IWordSink_FWD_DEFINED__
#define __IWordSink_FWD_DEFINED__
typedef interface IWordSink IWordSink;
#endif 	/* __IWordSink_FWD_DEFINED__ */


#ifndef __IWordBreaker_FWD_DEFINED__
#define __IWordBreaker_FWD_DEFINED__
typedef interface IWordBreaker IWordBreaker;
#endif 	/* __IWordBreaker_FWD_DEFINED__ */


#ifndef __IStemSink_FWD_DEFINED__
#define __IStemSink_FWD_DEFINED__
typedef interface IStemSink IStemSink;
#endif 	/* __IStemSink_FWD_DEFINED__ */


#ifndef __IStemmer_FWD_DEFINED__
#define __IStemmer_FWD_DEFINED__
typedef interface IStemmer IStemmer;
#endif 	/* __IStemmer_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "filter.h"

#ifdef __cplusplus
extern "C"{
#endif 

__post __maybenull
__post __writableTo(byteCount(size))
void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t size);
void __RPC_USER MIDL_user_free( __inout void __RPC_FAR * ); 

#ifndef __IPhraseSink_INTERFACE_DEFINED__
#define __IPhraseSink_INTERFACE_DEFINED__

/* interface IPhraseSink */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IPhraseSink;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CC906FF0-C058-101A-B554-08002B33B0E6")
    IPhraseSink : public IUnknown
    {
    public:
        virtual SCODE STDMETHODCALLTYPE PutSmallPhrase( 
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcNoun,
            /* [in] */ ULONG cwcNoun,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcModifier,
            /* [in] */ ULONG cwcModifier,
            /* [in] */ ULONG ulAttachmentType) = 0;
        
        virtual SCODE STDMETHODCALLTYPE PutPhrase( 
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcPhrase,
            /* [in] */ ULONG cwcPhrase) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IPhraseSinkVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IPhraseSink __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IPhraseSink __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IPhraseSink __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutSmallPhrase )( 
            IPhraseSink __RPC_FAR * This,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcNoun,
            /* [in] */ ULONG cwcNoun,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcModifier,
            /* [in] */ ULONG cwcModifier,
            /* [in] */ ULONG ulAttachmentType);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutPhrase )( 
            IPhraseSink __RPC_FAR * This,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcPhrase,
            /* [in] */ ULONG cwcPhrase);
        
        END_INTERFACE
    } IPhraseSinkVtbl;

    interface IPhraseSink
    {
        CONST_VTBL struct IPhraseSinkVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPhraseSink_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IPhraseSink_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IPhraseSink_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IPhraseSink_PutSmallPhrase(This,pwcNoun,cwcNoun,pwcModifier,cwcModifier,ulAttachmentType)	\
    (This)->lpVtbl -> PutSmallPhrase(This,pwcNoun,cwcNoun,pwcModifier,cwcModifier,ulAttachmentType)

#define IPhraseSink_PutPhrase(This,pwcPhrase,cwcPhrase)	\
    (This)->lpVtbl -> PutPhrase(This,pwcPhrase,cwcPhrase)

#endif /* COBJMACROS */


#endif 	/* C style interface */



SCODE STDMETHODCALLTYPE IPhraseSink_PutSmallPhrase_Proxy( 
    IPhraseSink __RPC_FAR * This,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcNoun,
    /* [in] */ ULONG cwcNoun,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcModifier,
    /* [in] */ ULONG cwcModifier,
    /* [in] */ ULONG ulAttachmentType);


void __RPC_STUB IPhraseSink_PutSmallPhrase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IPhraseSink_PutPhrase_Proxy( 
    IPhraseSink __RPC_FAR * This,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcPhrase,
    /* [in] */ ULONG cwcPhrase);


void __RPC_STUB IPhraseSink_PutPhrase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IPhraseSink_INTERFACE_DEFINED__ */


#ifndef __IWordSink_INTERFACE_DEFINED__
#define __IWordSink_INTERFACE_DEFINED__

/* interface IWordSink */
/* [unique][uuid][object][local] */ 

#ifndef _tagWORDREP_BREAK_TYPE_DEFINED
typedef 
enum tagWORDREP_BREAK_TYPE
    {	WORDREP_BREAK_EOW	= 0,
	WORDREP_BREAK_EOS	= 1,
	WORDREP_BREAK_EOP	= 2,
	WORDREP_BREAK_EOC	= 3
    }	WORDREP_BREAK_TYPE;

#define _tagWORDREP_BREAK_TYPE_DEFINED
#define _WORDREP_BREAK_TYPE_DEFINED
#endif

EXTERN_C const IID IID_IWordSink;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CC907054-C058-101A-B554-08002B33B0E6")
    IWordSink : public IUnknown
    {
    public:
        virtual SCODE STDMETHODCALLTYPE PutWord( 
            /* [in] */ ULONG cwc,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwcSrcLen,
            /* [in] */ ULONG cwcSrcPos) = 0;
        
        virtual SCODE STDMETHODCALLTYPE PutAltWord( 
            /* [in] */ ULONG cwc,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwcSrcLen,
            /* [in] */ ULONG cwcSrcPos) = 0;
        
        virtual SCODE STDMETHODCALLTYPE StartAltPhrase( void) = 0;
        
        virtual SCODE STDMETHODCALLTYPE EndAltPhrase( void) = 0;
        
        virtual SCODE STDMETHODCALLTYPE PutBreak( 
            /* [in] */ WORDREP_BREAK_TYPE breakType) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IWordSinkVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IWordSink __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IWordSink __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IWordSink __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutWord )( 
            IWordSink __RPC_FAR * This,
            /* [in] */ ULONG cwc,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwcSrcLen,
            /* [in] */ ULONG cwcSrcPos);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutAltWord )( 
            IWordSink __RPC_FAR * This,
            /* [in] */ ULONG cwc,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwcSrcLen,
            /* [in] */ ULONG cwcSrcPos);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *StartAltPhrase )( 
            IWordSink __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *EndAltPhrase )( 
            IWordSink __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutBreak )( 
            IWordSink __RPC_FAR * This,
            /* [in] */ WORDREP_BREAK_TYPE breakType);
        
        END_INTERFACE
    } IWordSinkVtbl;

    interface IWordSink
    {
        CONST_VTBL struct IWordSinkVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWordSink_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IWordSink_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IWordSink_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IWordSink_PutWord(This,cwc,pwcInBuf,cwcSrcLen,cwcSrcPos)	\
    (This)->lpVtbl -> PutWord(This,cwc,pwcInBuf,cwcSrcLen,cwcSrcPos)

#define IWordSink_PutAltWord(This,cwc,pwcInBuf,cwcSrcLen,cwcSrcPos)	\
    (This)->lpVtbl -> PutAltWord(This,cwc,pwcInBuf,cwcSrcLen,cwcSrcPos)

#define IWordSink_StartAltPhrase(This)	\
    (This)->lpVtbl -> StartAltPhrase(This)

#define IWordSink_EndAltPhrase(This)	\
    (This)->lpVtbl -> EndAltPhrase(This)

#define IWordSink_PutBreak(This,breakType)	\
    (This)->lpVtbl -> PutBreak(This,breakType)

#endif /* COBJMACROS */


#endif 	/* C style interface */



SCODE STDMETHODCALLTYPE IWordSink_PutWord_Proxy( 
    IWordSink __RPC_FAR * This,
    /* [in] */ ULONG cwc,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
    /* [in] */ ULONG cwcSrcLen,
    /* [in] */ ULONG cwcSrcPos);


void __RPC_STUB IWordSink_PutWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordSink_PutAltWord_Proxy( 
    IWordSink __RPC_FAR * This,
    /* [in] */ ULONG cwc,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
    /* [in] */ ULONG cwcSrcLen,
    /* [in] */ ULONG cwcSrcPos);


void __RPC_STUB IWordSink_PutAltWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordSink_StartAltPhrase_Proxy( 
    IWordSink __RPC_FAR * This);


void __RPC_STUB IWordSink_StartAltPhrase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordSink_EndAltPhrase_Proxy( 
    IWordSink __RPC_FAR * This);


void __RPC_STUB IWordSink_EndAltPhrase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordSink_PutBreak_Proxy( 
    IWordSink __RPC_FAR * This,
    /* [in] */ WORDREP_BREAK_TYPE breakType);


void __RPC_STUB IWordSink_PutBreak_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IWordSink_INTERFACE_DEFINED__ */


/* interface __MIDL_itf_qexport_0121 */
/* [local] */ 

#ifndef _tagTEXT_SOURCE_DEFINED

typedef HRESULT ( __stdcall __RPC_FAR *PFNFILLTEXTBUFFER )( 
    struct tagTEXT_SOURCE __RPC_FAR *pTextSource);

typedef struct tagTEXT_SOURCE
    {
    PFNFILLTEXTBUFFER pfnFillTextBuffer;
    const WCHAR __RPC_FAR *awcBuffer;
    ULONG iEnd;
    ULONG iCur;
    }	TEXT_SOURCE;

#define _tagTEXT_SOURCE_DEFINED
#define _TEXT_SOURCE_DEFINED
#endif


extern RPC_IF_HANDLE __MIDL_itf_qexport_0121_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_qexport_0121_v0_0_s_ifspec;

#ifndef __IWordBreaker_INTERFACE_DEFINED__
#define __IWordBreaker_INTERFACE_DEFINED__

/* interface IWordBreaker */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IWordBreaker;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D53552C8-77E3-101A-B552-08002B33B0E6")
    IWordBreaker : public IUnknown
    {
    public:
        virtual SCODE STDMETHODCALLTYPE Init( 
            /* [in] */ BOOL fQuery,
            /* [in] */ ULONG ulMaxTokenSize,
            /* [out] */ BOOL __RPC_FAR *pfLicense) = 0;
        
        virtual SCODE STDMETHODCALLTYPE BreakText( 
            /* [in] */ TEXT_SOURCE __RPC_FAR *pTextSource,
            /* [in] */ IWordSink __RPC_FAR *pWordSink,
            /* [in] */ IPhraseSink __RPC_FAR *pPhraseSink) = 0;
        
        virtual SCODE STDMETHODCALLTYPE ComposePhrase( 
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcNoun,
            /* [in] */ ULONG cwcNoun,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcModifier,
            /* [in] */ ULONG cwcModifier,
            /* [in] */ ULONG ulAttachmentType,
            /* [size_is][out] */ _Out_z_cap_(*pcwcPhrase) WCHAR __RPC_FAR *pwcPhrase,
            /* [out][in] */ ULONG __RPC_FAR *pcwcPhrase) = 0;
        
        virtual SCODE STDMETHODCALLTYPE GetLicenseToUse( 
            /* [string][out] */ const WCHAR __RPC_FAR *__RPC_FAR *ppwcsLicense) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IWordBreakerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IWordBreaker __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IWordBreaker __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IWordBreaker __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *Init )( 
            IWordBreaker __RPC_FAR * This,
            /* [in] */ BOOL fQuery,
            /* [in] */ ULONG ulMaxTokenSize,
            /* [out] */ BOOL __RPC_FAR *pfLicense);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *BreakText )( 
            IWordBreaker __RPC_FAR * This,
            /* [in] */ TEXT_SOURCE __RPC_FAR *pTextSource,
            /* [in] */ IWordSink __RPC_FAR *pWordSink,
            /* [in] */ IPhraseSink __RPC_FAR *pPhraseSink);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *ComposePhrase )( 
            IWordBreaker __RPC_FAR * This,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcNoun,
            /* [in] */ ULONG cwcNoun,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcModifier,
            /* [in] */ ULONG cwcModifier,
            /* [in] */ ULONG ulAttachmentType,
            /* [size_is][out] */ WCHAR __RPC_FAR *pwcPhrase,
            /* [out][in] */ ULONG __RPC_FAR *pcwcPhrase);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *GetLicenseToUse )( 
            IWordBreaker __RPC_FAR * This,
            /* [string][out] */ const WCHAR __RPC_FAR *__RPC_FAR *ppwcsLicense);
        
        END_INTERFACE
    } IWordBreakerVtbl;

    interface IWordBreaker
    {
        CONST_VTBL struct IWordBreakerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWordBreaker_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IWordBreaker_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IWordBreaker_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IWordBreaker_Init(This,fQuery,ulMaxTokenSize,pfLicense)	\
    (This)->lpVtbl -> Init(This,fQuery,ulMaxTokenSize,pfLicense)

#define IWordBreaker_BreakText(This,pTextSource,pWordSink,pPhraseSink)	\
    (This)->lpVtbl -> BreakText(This,pTextSource,pWordSink,pPhraseSink)

#define IWordBreaker_ComposePhrase(This,pwcNoun,cwcNoun,pwcModifier,cwcModifier,ulAttachmentType,pwcPhrase,pcwcPhrase)	\
    (This)->lpVtbl -> ComposePhrase(This,pwcNoun,cwcNoun,pwcModifier,cwcModifier,ulAttachmentType,pwcPhrase,pcwcPhrase)

#define IWordBreaker_GetLicenseToUse(This,ppwcsLicense)	\
    (This)->lpVtbl -> GetLicenseToUse(This,ppwcsLicense)

#endif /* COBJMACROS */


#endif 	/* C style interface */



SCODE STDMETHODCALLTYPE IWordBreaker_Init_Proxy( 
    IWordBreaker __RPC_FAR * This,
    /* [in] */ BOOL fQuery,
    /* [in] */ ULONG ulMaxTokenSize,
    /* [out] */ BOOL __RPC_FAR *pfLicense);


void __RPC_STUB IWordBreaker_Init_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordBreaker_BreakText_Proxy( 
    IWordBreaker __RPC_FAR * This,
    /* [in] */ TEXT_SOURCE __RPC_FAR *pTextSource,
    /* [in] */ IWordSink __RPC_FAR *pWordSink,
    /* [in] */ IPhraseSink __RPC_FAR *pPhraseSink);


void __RPC_STUB IWordBreaker_BreakText_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordBreaker_ComposePhrase_Proxy( 
    IWordBreaker __RPC_FAR * This,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcNoun,
    /* [in] */ ULONG cwcNoun,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcModifier,
    /* [in] */ ULONG cwcModifier,
    /* [in] */ ULONG ulAttachmentType,
    /* [size_is][out] */ _Out_z_cap_(*pcwcPhrase) WCHAR __RPC_FAR *pwcPhrase,
    /* [out][in] */ ULONG __RPC_FAR *pcwcPhrase);


void __RPC_STUB IWordBreaker_ComposePhrase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IWordBreaker_GetLicenseToUse_Proxy( 
    IWordBreaker __RPC_FAR * This,
    /* [string][out] */ const WCHAR __RPC_FAR *__RPC_FAR *ppwcsLicense);


void __RPC_STUB IWordBreaker_GetLicenseToUse_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IWordBreaker_INTERFACE_DEFINED__ */


#ifndef __IStemSink_INTERFACE_DEFINED__
#define __IStemSink_INTERFACE_DEFINED__

/* interface IStemSink */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IStemSink;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("fe77c330-7f42-11ce-be57-00aa0051fe20")
    IStemSink : public IUnknown
    {
    public:
        virtual SCODE STDMETHODCALLTYPE PutAltWord( 
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwc) = 0;
        
        virtual SCODE STDMETHODCALLTYPE PutWord( 
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwc) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IStemSinkVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IStemSink __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IStemSink __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IStemSink __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutAltWord )( 
            IStemSink __RPC_FAR * This,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwc);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *PutWord )( 
            IStemSink __RPC_FAR * This,
            /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwc);
        
        END_INTERFACE
    } IStemSinkVtbl;

    interface IStemSink
    {
        CONST_VTBL struct IStemSinkVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IStemSink_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IStemSink_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IStemSink_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IStemSink_PutAltWord(This,pwcInBuf,cwc)	\
    (This)->lpVtbl -> PutAltWord(This,pwcInBuf,cwc)

#define IStemSink_PutWord(This,pwcInBuf,cwc)	\
    (This)->lpVtbl -> PutWord(This,pwcInBuf,cwc)

#endif /* COBJMACROS */


#endif 	/* C style interface */



SCODE STDMETHODCALLTYPE IStemSink_PutAltWord_Proxy( 
    IStemSink __RPC_FAR * This,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
    /* [in] */ ULONG cwc);


void __RPC_STUB IStemSink_PutAltWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IStemSink_PutWord_Proxy( 
    IStemSink __RPC_FAR * This,
    /* [size_is][in] */ const WCHAR __RPC_FAR *pwcInBuf,
    /* [in] */ ULONG cwc);


void __RPC_STUB IStemSink_PutWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IStemSink_INTERFACE_DEFINED__ */


#ifndef __IStemmer_INTERFACE_DEFINED__
#define __IStemmer_INTERFACE_DEFINED__

/* interface IStemmer */
/* [unique][uuid][object][local] */ 


EXTERN_C const IID IID_IStemmer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("efbaf140-7f42-11ce-be57-00aa0051fe20")
    IStemmer : public IUnknown
    {
    public:
        virtual SCODE STDMETHODCALLTYPE Init( 
            /* [in] */ ULONG ulMaxTokenSize,
            /* [out] */ BOOL __RPC_FAR *pfLicense) = 0;
        
        virtual SCODE STDMETHODCALLTYPE StemWord( 
            /* [in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwc,
            /* [in] */ IStemSink __RPC_FAR *pStemSink) = 0;
        
        virtual SCODE STDMETHODCALLTYPE GetLicenseToUse( 
            /* [string][out] */ const WCHAR __RPC_FAR *__RPC_FAR *ppwcsLicense) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IStemmerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IStemmer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IStemmer __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IStemmer __RPC_FAR * This);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *Init )( 
            IStemmer __RPC_FAR * This,
            /* [in] */ ULONG ulMaxTokenSize,
            /* [out] */ BOOL __RPC_FAR *pfLicense);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *StemWord )( 
            IStemmer __RPC_FAR * This,
            /* [in] */ const WCHAR __RPC_FAR *pwcInBuf,
            /* [in] */ ULONG cwc,
            /* [in] */ IStemSink __RPC_FAR *pStemSink);
        
        SCODE ( STDMETHODCALLTYPE __RPC_FAR *GetLicenseToUse )( 
            IStemmer __RPC_FAR * This,
            /* [string][out] */ const WCHAR __RPC_FAR *__RPC_FAR *ppwcsLicense);
        
        END_INTERFACE
    } IStemmerVtbl;

    interface IStemmer
    {
        CONST_VTBL struct IStemmerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IStemmer_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IStemmer_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IStemmer_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IStemmer_Init(This,ulMaxTokenSize,pfLicense)	\
    (This)->lpVtbl -> Init(This,ulMaxTokenSize,pfLicense)

#define IStemmer_StemWord(This,pwcInBuf,cwc,pStemSink)	\
    (This)->lpVtbl -> StemWord(This,pwcInBuf,cwc,pStemSink)

#define IStemmer_GetLicenseToUse(This,ppwcsLicense)	\
    (This)->lpVtbl -> GetLicenseToUse(This,ppwcsLicense)

#endif /* COBJMACROS */


#endif 	/* C style interface */



SCODE STDMETHODCALLTYPE IStemmer_Init_Proxy( 
    IStemmer __RPC_FAR * This,
    /* [in] */ ULONG ulMaxTokenSize,
    /* [out] */ BOOL __RPC_FAR *pfLicense);


void __RPC_STUB IStemmer_Init_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IStemmer_StemWord_Proxy( 
    IStemmer __RPC_FAR * This,
    /* [in] */ const WCHAR __RPC_FAR *pwcInBuf,
    /* [in] */ ULONG cwc,
    /* [in] */ IStemSink __RPC_FAR *pStemSink);


void __RPC_STUB IStemmer_StemWord_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


SCODE STDMETHODCALLTYPE IStemmer_GetLicenseToUse_Proxy( 
    IStemmer __RPC_FAR * This,
    /* [string][out] */ const WCHAR __RPC_FAR *__RPC_FAR *ppwcsLicense);


void __RPC_STUB IStemmer_GetLicenseToUse_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IStemmer_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


