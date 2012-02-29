/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* Compiler settings for C:\vsmso\office\dev\mso\sv.odl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __msosvtl_tmp__
#define __msosvtl_tmp__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IWebWizardHost_FWD_DEFINED__
#define __IWebWizardHost_FWD_DEFINED__
typedef interface IWebWizardHost IWebWizardHost;
#endif 	/* __IWebWizardHost_FWD_DEFINED__ */


#ifndef __INewWDEvents_FWD_DEFINED__
#define __INewWDEvents_FWD_DEFINED__
typedef interface INewWDEvents INewWDEvents;
#endif 	/* __INewWDEvents_FWD_DEFINED__ */


#ifndef __IMsoWDWizard_FWD_DEFINED__
#define __IMsoWDWizard_FWD_DEFINED__
typedef interface IMsoWDWizard IMsoWDWizard;
#endif 	/* __IMsoWDWizard_FWD_DEFINED__ */


/* header files for imported files */
#include "ShlDisp.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t size);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __SOW_LIBRARY_DEFINED__
#define __SOW_LIBRARY_DEFINED__

/* library SOW */
/* [version][helpcontext][helpstring][lcid][uuid] */ 


EXTERN_C const IID LIBID_SOW;

#ifndef __IWebWizardHost_INTERFACE_DEFINED__
#define __IWebWizardHost_INTERFACE_DEFINED__

/* interface IWebWizardHost */
/* [helpstring][dual][object][uuid] */ 


EXTERN_C const IID IID_IWebWizardHost;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("18bcc359-4990-4bfb-b951-3c83702be5f9")
    IWebWizardHost : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE FinalBack( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE FinalNext( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Cancel( void) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Caption( 
            /* [in] */ BSTR bstrCaption) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Caption( 
            /* [retval][out] */ BSTR __RPC_FAR *pbstrCaption) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Property( 
            /* [in] */ BSTR bstrPropertyName,
            /* [in] */ VARIANT __RPC_FAR *pvProperty) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Property( 
            /* [in] */ BSTR bstrPropertyName,
            /* [retval][out] */ VARIANT __RPC_FAR *pvProperty) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetWizardButtons( 
            /* [in] */ VARIANT_BOOL vfEnableBack,
            /* [in] */ VARIANT_BOOL vfEnableNext,
            /* [in] */ VARIANT_BOOL vfLastPage) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetHeaderText( 
            /* [in] */ BSTR bstrHeaderTitle,
            /* [in] */ BSTR bstrHeaderSubtitle) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IWebWizardHostVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IWebWizardHost __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IWebWizardHost __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IWebWizardHost __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ _In_count_(cNames) LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ _Out_cap_(cNames) DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FinalBack )( 
            IWebWizardHost __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FinalNext )( 
            IWebWizardHost __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            IWebWizardHost __RPC_FAR * This);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Caption )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ BSTR bstrCaption);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Caption )( 
            IWebWizardHost __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pbstrCaption);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Property )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ BSTR bstrPropertyName,
            /* [in] */ VARIANT __RPC_FAR *pvProperty);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Property )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ BSTR bstrPropertyName,
            /* [retval][out] */ VARIANT __RPC_FAR *pvProperty);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetWizardButtons )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL vfEnableBack,
            /* [in] */ VARIANT_BOOL vfEnableNext,
            /* [in] */ VARIANT_BOOL vfLastPage);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetHeaderText )( 
            IWebWizardHost __RPC_FAR * This,
            /* [in] */ BSTR bstrHeaderTitle,
            /* [in] */ BSTR bstrHeaderSubtitle);
        
        END_INTERFACE
    } IWebWizardHostVtbl;

    interface IWebWizardHost
    {
        CONST_VTBL struct IWebWizardHostVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWebWizardHost_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IWebWizardHost_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IWebWizardHost_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IWebWizardHost_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IWebWizardHost_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IWebWizardHost_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IWebWizardHost_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IWebWizardHost_FinalBack(This)	\
    (This)->lpVtbl -> FinalBack(This)

#define IWebWizardHost_FinalNext(This)	\
    (This)->lpVtbl -> FinalNext(This)

#define IWebWizardHost_Cancel(This)	\
    (This)->lpVtbl -> Cancel(This)

#define IWebWizardHost_put_Caption(This,bstrCaption)	\
    (This)->lpVtbl -> put_Caption(This,bstrCaption)

#define IWebWizardHost_get_Caption(This,pbstrCaption)	\
    (This)->lpVtbl -> get_Caption(This,pbstrCaption)

#define IWebWizardHost_put_Property(This,bstrPropertyName,pvProperty)	\
    (This)->lpVtbl -> put_Property(This,bstrPropertyName,pvProperty)

#define IWebWizardHost_get_Property(This,bstrPropertyName,pvProperty)	\
    (This)->lpVtbl -> get_Property(This,bstrPropertyName,pvProperty)

#define IWebWizardHost_SetWizardButtons(This,vfEnableBack,vfEnableNext,vfLastPage)	\
    (This)->lpVtbl -> SetWizardButtons(This,vfEnableBack,vfEnableNext,vfLastPage)

#define IWebWizardHost_SetHeaderText(This,bstrHeaderTitle,bstrHeaderSubtitle)	\
    (This)->lpVtbl -> SetHeaderText(This,bstrHeaderTitle,bstrHeaderSubtitle)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_FinalBack_Proxy( 
    IWebWizardHost __RPC_FAR * This);


void __RPC_STUB IWebWizardHost_FinalBack_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_FinalNext_Proxy( 
    IWebWizardHost __RPC_FAR * This);


void __RPC_STUB IWebWizardHost_FinalNext_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_Cancel_Proxy( 
    IWebWizardHost __RPC_FAR * This);


void __RPC_STUB IWebWizardHost_Cancel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput][id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_put_Caption_Proxy( 
    IWebWizardHost __RPC_FAR * This,
    /* [in] */ BSTR bstrCaption);


void __RPC_STUB IWebWizardHost_put_Caption_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget][id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_get_Caption_Proxy( 
    IWebWizardHost __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pbstrCaption);


void __RPC_STUB IWebWizardHost_get_Caption_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput][id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_put_Property_Proxy( 
    IWebWizardHost __RPC_FAR * This,
    /* [in] */ BSTR bstrPropertyName,
    /* [in] */ VARIANT __RPC_FAR *pvProperty);


void __RPC_STUB IWebWizardHost_put_Property_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget][id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_get_Property_Proxy( 
    IWebWizardHost __RPC_FAR * This,
    /* [in] */ BSTR bstrPropertyName,
    /* [retval][out] */ VARIANT __RPC_FAR *pvProperty);


void __RPC_STUB IWebWizardHost_get_Property_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_SetWizardButtons_Proxy( 
    IWebWizardHost __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL vfEnableBack,
    /* [in] */ VARIANT_BOOL vfEnableNext,
    /* [in] */ VARIANT_BOOL vfLastPage);


void __RPC_STUB IWebWizardHost_SetWizardButtons_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IWebWizardHost_SetHeaderText_Proxy( 
    IWebWizardHost __RPC_FAR * This,
    /* [in] */ BSTR bstrHeaderTitle,
    /* [in] */ BSTR bstrHeaderSubtitle);


void __RPC_STUB IWebWizardHost_SetHeaderText_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IWebWizardHost_INTERFACE_DEFINED__ */


#ifndef __INewWDEvents_INTERFACE_DEFINED__
#define __INewWDEvents_INTERFACE_DEFINED__

/* interface INewWDEvents */
/* [helpstring][dual][object][uuid] */ 


EXTERN_C const IID IID_INewWDEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0751c551-7568-41c9-8e5b-e22e38919236")
    INewWDEvents : public IWebWizardHost
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE PassportAuthenticate( 
            /* [in] */ BSTR bstrSignInUrl,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pvfAuthenitcated) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct INewWDEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            INewWDEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            INewWDEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            INewWDEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ _In_count_(cNames) LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ _Out_cap_(cNames) DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FinalBack )( 
            INewWDEvents __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FinalNext )( 
            INewWDEvents __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            INewWDEvents __RPC_FAR * This);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Caption )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ BSTR bstrCaption);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Caption )( 
            INewWDEvents __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pbstrCaption);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Property )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ BSTR bstrPropertyName,
            /* [in] */ VARIANT __RPC_FAR *pvProperty);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Property )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ BSTR bstrPropertyName,
            /* [retval][out] */ VARIANT __RPC_FAR *pvProperty);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetWizardButtons )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL vfEnableBack,
            /* [in] */ VARIANT_BOOL vfEnableNext,
            /* [in] */ VARIANT_BOOL vfLastPage);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetHeaderText )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ BSTR bstrHeaderTitle,
            /* [in] */ BSTR bstrHeaderSubtitle);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PassportAuthenticate )( 
            INewWDEvents __RPC_FAR * This,
            /* [in] */ BSTR bstrSignInUrl,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pvfAuthenitcated);
        
        END_INTERFACE
    } INewWDEventsVtbl;

    interface INewWDEvents
    {
        CONST_VTBL struct INewWDEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define INewWDEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define INewWDEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define INewWDEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define INewWDEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define INewWDEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define INewWDEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define INewWDEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define INewWDEvents_FinalBack(This)	\
    (This)->lpVtbl -> FinalBack(This)

#define INewWDEvents_FinalNext(This)	\
    (This)->lpVtbl -> FinalNext(This)

#define INewWDEvents_Cancel(This)	\
    (This)->lpVtbl -> Cancel(This)

#define INewWDEvents_put_Caption(This,bstrCaption)	\
    (This)->lpVtbl -> put_Caption(This,bstrCaption)

#define INewWDEvents_get_Caption(This,pbstrCaption)	\
    (This)->lpVtbl -> get_Caption(This,pbstrCaption)

#define INewWDEvents_put_Property(This,bstrPropertyName,pvProperty)	\
    (This)->lpVtbl -> put_Property(This,bstrPropertyName,pvProperty)

#define INewWDEvents_get_Property(This,bstrPropertyName,pvProperty)	\
    (This)->lpVtbl -> get_Property(This,bstrPropertyName,pvProperty)

#define INewWDEvents_SetWizardButtons(This,vfEnableBack,vfEnableNext,vfLastPage)	\
    (This)->lpVtbl -> SetWizardButtons(This,vfEnableBack,vfEnableNext,vfLastPage)

#define INewWDEvents_SetHeaderText(This,bstrHeaderTitle,bstrHeaderSubtitle)	\
    (This)->lpVtbl -> SetHeaderText(This,bstrHeaderTitle,bstrHeaderSubtitle)


#define INewWDEvents_PassportAuthenticate(This,bstrSignInUrl,pvfAuthenitcated)	\
    (This)->lpVtbl -> PassportAuthenticate(This,bstrSignInUrl,pvfAuthenitcated)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id] */ HRESULT STDMETHODCALLTYPE INewWDEvents_PassportAuthenticate_Proxy( 
    INewWDEvents __RPC_FAR * This,
    /* [in] */ BSTR bstrSignInUrl,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pvfAuthenitcated);


void __RPC_STUB INewWDEvents_PassportAuthenticate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __INewWDEvents_INTERFACE_DEFINED__ */


#ifndef __IMsoWDWizard_INTERFACE_DEFINED__
#define __IMsoWDWizard_INTERFACE_DEFINED__

/* interface IMsoWDWizard */
/* [object][uuid][nonextensible][dual] */ 


EXTERN_C const IID IID_IMsoWDWizard;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("000C0513-0000-0000-C000-000000000046")
    IMsoWDWizard : public INewWDEvents
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IMsoWDWizardVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IMsoWDWizard __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IMsoWDWizard __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ _In_count_(cNames) LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FinalBack )( 
            IMsoWDWizard __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FinalNext )( 
            IMsoWDWizard __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            IMsoWDWizard __RPC_FAR * This);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Caption )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ BSTR bstrCaption);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Caption )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pbstrCaption);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Property )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ BSTR bstrPropertyName,
            /* [in] */ VARIANT __RPC_FAR *pvProperty);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Property )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ BSTR bstrPropertyName,
            /* [retval][out] */ VARIANT __RPC_FAR *pvProperty);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetWizardButtons )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL vfEnableBack,
            /* [in] */ VARIANT_BOOL vfEnableNext,
            /* [in] */ VARIANT_BOOL vfLastPage);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetHeaderText )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ BSTR bstrHeaderTitle,
            /* [in] */ BSTR bstrHeaderSubtitle);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PassportAuthenticate )( 
            IMsoWDWizard __RPC_FAR * This,
            /* [in] */ BSTR bstrSignInUrl,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pvfAuthenitcated);
        
        END_INTERFACE
    } IMsoWDWizardVtbl;

    interface IMsoWDWizard
    {
        CONST_VTBL struct IMsoWDWizardVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMsoWDWizard_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IMsoWDWizard_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IMsoWDWizard_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IMsoWDWizard_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IMsoWDWizard_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IMsoWDWizard_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IMsoWDWizard_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IMsoWDWizard_FinalBack(This)	\
    (This)->lpVtbl -> FinalBack(This)

#define IMsoWDWizard_FinalNext(This)	\
    (This)->lpVtbl -> FinalNext(This)

#define IMsoWDWizard_Cancel(This)	\
    (This)->lpVtbl -> Cancel(This)

#define IMsoWDWizard_put_Caption(This,bstrCaption)	\
    (This)->lpVtbl -> put_Caption(This,bstrCaption)

#define IMsoWDWizard_get_Caption(This,pbstrCaption)	\
    (This)->lpVtbl -> get_Caption(This,pbstrCaption)

#define IMsoWDWizard_put_Property(This,bstrPropertyName,pvProperty)	\
    (This)->lpVtbl -> put_Property(This,bstrPropertyName,pvProperty)

#define IMsoWDWizard_get_Property(This,bstrPropertyName,pvProperty)	\
    (This)->lpVtbl -> get_Property(This,bstrPropertyName,pvProperty)

#define IMsoWDWizard_SetWizardButtons(This,vfEnableBack,vfEnableNext,vfLastPage)	\
    (This)->lpVtbl -> SetWizardButtons(This,vfEnableBack,vfEnableNext,vfLastPage)

#define IMsoWDWizard_SetHeaderText(This,bstrHeaderTitle,bstrHeaderSubtitle)	\
    (This)->lpVtbl -> SetHeaderText(This,bstrHeaderTitle,bstrHeaderSubtitle)


#define IMsoWDWizard_PassportAuthenticate(This,bstrSignInUrl,pvfAuthenitcated)	\
    (This)->lpVtbl -> PassportAuthenticate(This,bstrSignInUrl,pvfAuthenitcated)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMsoWDWizard_INTERFACE_DEFINED__ */

#endif /* __SOW_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
