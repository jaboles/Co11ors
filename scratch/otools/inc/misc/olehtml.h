/////////////////////////////////////////////////////////////////////////////
//
// olehtml.h    OLE HTML Control interfaces
//
//              OLE Version 2.0
//
//              Copyright (c) 1992-1994, Microsoft Corp. All rights reserved.
//
//      Author:         Jim Francis
//
/////////////////////////////////////////////////////////////////////////////

// {049948D1-5686-11cf-8E0D-00AA00A74C5C}
DEFINE_GUID(IID_IPersistHTML,
	0x49948d1, 0x5686, 0x11cf, 0x8e, 0xd, 0x0, 0xaa, 0x0, 0xa7, 0x4c, 0x5c);

//marquee, bgsound, inline video HTML controls
//{8422DAE3-9929-11CF-B8D3-004033373DA8}
DEFINE_GUID(CLSID_WHTMLBgsound, 
	0x8422DAE3,
	0x9929, 0x11CF,
	0xB8, 0xD3, 0x00, 0x40, 0x33, 0x37, 0x3D, 0xA8);

//{8422DAE7-9929-11CF-B8D3-004033373DA8}	
DEFINE_GUID(CLSID_WHTMLVideo,
	0x8422DAE7,
	0x9929, 0x11CF,
	0xB8, 0xD3, 0x00, 0x40, 0x33, 0x37, 0x3D, 0xA8);

//{250770F3-6AF2-11CF-A915-008029E31FCD}
DEFINE_GUID(CLSID_WHTMLMarquee, 
	0x250770F3,
	0x6AF2, 0x11CF,
	0xA9, 0x15, 0x00, 0x80, 0x29, 0xE3, 0x1F, 0xCD);

#ifndef __IPersistHTML_INTERFACE_DEFINED__
#define __IPersistHTML_INTERFACE_DEFINED__

// typedef IPersistHTML __RPC_FAR *LPPERSISTHTML;
typedef interface IPersistHTML IPersistHTML;

typedef IPersistHTML  *LPPERSISTHTML;

EXTERN_C const IID IID_IPersistHTML;

#if defined(__cplusplus) && !defined(CINTERFACE)

    interface IPersistHTML : public IPersist
    {
    public:
        virtual HRESULT __stdcall IsDirty( void) = 0;

        virtual HRESULT __stdcall LoadHTML(
            /* [unique][in] */ IStream __RPC_FAR *pStm) = 0;

        virtual HRESULT __stdcall SaveHTML(
            /* [unique][in] */ IStream __RPC_FAR *pStm,
            /* [in] */ BOOL fClearDirty) = 0;

        virtual HRESULT __stdcall GetSizeMaxHTML(
            /* [out] */ ULARGE_INTEGER __RPC_FAR *pcbSize) = 0;

        virtual HRESULT __stdcall InitNewHTML( void) = 0;
    };

#else   /* C style interface */

    typedef struct IPersistHTMLVtbl
    {

        HRESULT ( __stdcall __RPC_FAR *QueryInterface )(
            IPersistHTML __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [out] */ void __RPC_FAR *__RPC_FAR *ppvObject);

        ULONG ( __stdcall __RPC_FAR *AddRef )(
            IPersistHTML __RPC_FAR * This);

        ULONG ( __stdcall __RPC_FAR *Release )(
            IPersistHTML __RPC_FAR * This);

        HRESULT ( __stdcall __RPC_FAR *GetClassID )(
            IPersistHTML __RPC_FAR * This,
            /* [out] */ CLSID __RPC_FAR *pClassID);

        HRESULT ( __stdcall __RPC_FAR *IsDirty )(
            IPersistHTML __RPC_FAR * This);

        HRESULT ( __stdcall __RPC_FAR *LoadHTML )(
            IPersistHTML __RPC_FAR * This,
            /* [unique][in] */ IStream __RPC_FAR *pStm);

        HRESULT ( __stdcall __RPC_FAR *SaveHTML )(
            IPersistHTML __RPC_FAR * This,
            /* [unique][in] */ IStream __RPC_FAR *pStm,
            /* [in] */ BOOL fClearDirty);

        HRESULT ( __stdcall __RPC_FAR *GetSizeMaxHTML )(
            IPersistHTML __RPC_FAR * This,
            /* [out] */ ULARGE_INTEGER __RPC_FAR *pcbSize);

        HRESULT ( __stdcall __RPC_FAR *InitNewHTML )(
            IPersistHTML __RPC_FAR * This);
			
    } IPersistHTMLVtbl;

    interface IPersistHTML
    {
        CONST_VTBL struct IPersistHTMLVtbl __RPC_FAR *lpVtbl;
    };
	
#ifdef COBJMACROS

#define IPersistHTML_QueryInterface(This,riid,ppvObject)        \
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IPersistHTML_AddRef(This)       \
    (This)->lpVtbl -> AddRef(This)

#define IPersistHTML_Release(This)      \
    (This)->lpVtbl -> Release(This)

#define IPersistHTML_GetClassID(This,pClassID)  \
    (This)->lpVtbl -> GetClassID(This,pClassID)

#define IPersistHTML_IsDirty(This)      \
    (This)->lpVtbl -> IsDirty(This)

#define IPersistHTML_LoadHTML(This,pStm)    \
    (This)->lpVtbl -> LoadHTML(This,pStm)

#define IPersistHTML_SaveHTML(This,pStm,fClearDirty)        \
    (This)->lpVtbl -> SaveHTML(This,pStm,fClearDirty)

#define IPersistHTML_GetSizeMaxHTML(This,pcbSize)   \
    (This)->lpVtbl -> GetSizeMaxHTML(This,pcbSize)

#endif /* COBJMACROS */
#endif  /* C style interface */
#endif /* __IPersistHTML_INTERFACE_DEFINED__ */
