/****************************************************************************
	msopkmcl.h

	Owner: ShaneM
 	Copyright (c) 2000 Microsoft Corporation

	PKM Client for commications with Tahoe server

	API's that are minimally required to use the PKM Client interface
		in order to communication on a Tahoe server.  Once communication is
		open, the interface allows a user to 1) Checkout a document 2) Checkin
		a document, 3) Make a document public, 4) Query allowable operations on
		a document
	
	Interface: IMsoPKMClient
****************************************************************************/

#pragma once

#ifndef MSOPKMCL_H
#define MSOPKMCL_H
#pragma pack( push, msopkmcl, 4 )

#define fpkmCheckout   1
#define fpkmCheckin    2
#define fpkmMakePublic 4
#define fpkmUndoCheckout 8


#define fpkmdlgCancel   0
#define fpkmdlgCheckout 1
#define fpkmdlgCheckin  2
#define fpkmdlgOpen     3
#define fpkmdlgSave     4
#define fpkmdlgUndoCheckout 5
#define fpkmdlgDiscard	6


#define fpkmurlPublic	0
#define fpkmurlPrivate  1
#define fpkmurlPublicNameOnly 2
#define fpkmurlCheckedIn fpkmurlPublic
#define fpkmurlCheckedOut fpkmurlPublic
#define fpkmurlRedirected 3

#define fpkmflagCheckinOnClose      0x0001
#define fpkmflagUndoCheckoutOnClose 0x0002
#define fpkmflagMakePublicOnClose   0x0004
#define fpkmflagOperationCancelled  0x0008
#define fpkmflagReadOnly            0x0010
#define fpkmflagPropsAreInSync      0x0020
#define fpkmflagShowUI              0x0040
#define fpkmflagNoPromptOnClose     0x0080
#define fpkmflagSTSv2               0x0100

#define fpkmmdNone      0x00000000
#define fpkmmdUseCache  0x00000001

typedef enum 
{
	fpkmdsNone,
	fpkmdsCreated,
	fpkmdsCheckedOut,
	fpkmdsCheckedIn,
	fpkmdsApproved,
} PKMDocState;

// {565E6152-38EA-49c5-BC5A-EE4C6FDD84FA}
DEFINE_GUID(IID_IMsoPKMClient, 0x565e6152, 0x38ea, 0x49c5, 0xbc, 0x5a, 0xee, 0x4c, 0x6f, 0xdd, 0x84, 0xfa);
#ifdef INTERFACE
#undef INTERFACE
#endif //INTERFACE
#define INTERFACE IMsoPKMClient

DECLARE_INTERFACE_(IMsoPKMClient, ISimpleUnknown)
{
	MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
	// FDebugMessage method
	MSODEBUGMETHOD
	MSOMETHOD (SetURL) (THIS_ LPWSTR pwzURL) PURE;
	MSOMETHOD (GetValidOperations) (THIS_ int *piValidOps) PURE;
	MSOMETHOD (Checkout) (THIS) PURE;
	MSOMETHOD (Checkin) (THIS_ BOOL fUseForm) PURE;
	MSOMETHOD (MakePublic) (THIS) PURE;
	MSOMETHOD (UndoCheckout) (THIS) PURE;
	MSOMETHOD (Free) (THIS) PURE;
	MSOMETHOD (DoOpenCheckOutDlg) (THIS_ HWND pParent, HMSOINST pinst, LPWSTR pwzDisp, int *pResult) PURE;
	MSOMETHOD (DoCloseCheckInDlg) (THIS_ HWND pParent, HMSOINST pinst, LPWSTR pwzDisp, int *pResult) PURE;
	MSOMETHOD (DoCloseCheckInDlgND) (THIS_ HWND pParent, HMSOINST pinst, LPWSTR pwzDisp, int *pResult) PURE;
	MSOMETHOD (DoUndoCheckoutDlg) (THIS_ HWND pParent, HMSOINST pinst, LPWSTR pwzDisp, int *pResult) PURE;
	MSOMETHOD (AddRef) (THIS) PURE;
	MSOMETHOD (GetURL) (THIS_ int iurltype, LPWSTR pwzURL, int cch) PURE;
	MSOMETHOD (SetFlags) (THIS_ int iFlags) PURE;
	MSOMETHOD (GetFlags) (THIS_ int *piFlags) PURE;
	MSOMETHOD (SetCheckinVersionComments) (THIS_ const WCHAR* pwzComments) PURE;
	MSOMETHOD (GetCheckinVersionComments) (THIS_ WCHAR* pwzComments, int cwch) PURE;
	MSOMETHOD (InvalidateAllowedOpsCache) (THIS) PURE;
	MSOMETHOD (GetDocumentState) (THIS_ PKMDocState *ppkmdocstate) PURE;
	MSOMETHOD (DoAlreadyCheckedOutDlg) (THIS) PURE;
	MSOMETHOD (VersioningEnabled)(THIS_ BOOL *) PURE;
	MSOMETHOD (Release) (THIS) PURE;
	MSOMETHOD_(HRESULT, GetLastHResult)(THIS);
	MSOMETHOD_(void, ClearLastHResult)(THIS);
};

/****************************************************************************
	IMsoFileNewProp C interface
*****************************************************************************/

#ifndef __cplusplus

// ----- IUnknown methods
#define IMsoPKMClient_QueryInterface(This, riid, ppvObj) \
	(This)->lpVtbl->QueryInterface(This, riid, ppvObj)

#define IMsoPKMClient_SetURL(This, pwzURL) \
	(This)->lpVtbl->SetURL(This, pwzURL)

#define IMsoPKMClient_GetValidOperations(This, iValidOps) \
	(This)->lpVtbl->GetValidOperations(This, iValidOps)

#define IMsoPKMClient_Checkout(This) \
	(This)->lpVtbl->Checkout(This)

#define IMsoPKMClient_UndoCheckout(This) \
	(This)->lpVtbl->UndoCheckout(This)

#define IMsoPKMClient_UndoCheckout(This) \
	(This)->lpVtbl->UndoCheckout(This)

#define IMsoPKMClient_Checkin(This, fUseForm) \
	(This)->lpVtbl->Checkin(This, fUseForm)

#define IMsoPKMClient_MakePublic(This) \
	(This)->lpVtbl->MakePublic(This)

#define IMsoPKMClient_Free(This) \
	(This)->lpVtbl->Free(This)

#define IMsoPKMClient_Release(This) \
	(This)->lpVtbl->Release(This)

#define IMsoPKMClient_DoOpenCheckOutDlg(This,hwnd,hinst,pwzDisp,pResult) \
	(This)->lpVtbl->DoOpenCheckOutDlg(This,hwnd,hinst,pwzDisp,pResult)

#define IMsoPKMClient_DoCloseCheckInDlg(This,hwnd,hinst,pwzDisp,pResult) \
	(This)->lpVtbl->DoCloseCheckInDlg(This,hwnd,hinst,pwzDisp,pResult)

#define IMsoPKMClient_DoCloseCheckInDlgND(This,hwnd,hinst,pwzDisp,pResult) \
	(This)->lpVtbl->DoCloseCheckInDlgND(This,hwnd,hinst,pwzDisp,pResult)

#define IMsoPKMClient_DoUndoCheckoutDlg(This,hwnd,hinst,pwzDisp,pResult) \
	(This)->lpVtbl->DoUndoCheckoutDlg(This,hwnd,hinst,pwzDisp,pResult)

#define IMsoPKMClient_AddRef(This) \
	(This)->lpVtbl->AddRef(This)

#define IMsoPKMClient_GetURL(This, iurltype, pwzURL, cch) \
	(This)->lpVtbl->GetURL(This, iurltype, pwzURL, cch)

#define IMsoPKMClient_SetFlags(This, flags) \
	(This)->lpVtbl->SetFlags(This, flags)

#define IMsoPKMClient_GetFlags(This, flags) \
	(This)->lpVtbl->GetFlags(This, flags)

#define IMsoPKMClient_SetCheckinVersionComments(This, Comments) \
	(This)->lpVtbl->SetCheckinVersionComments(This,Comments)

#define IMsoPKMClient_GetCheckinVersionComments(This, Comments, cch) \
	(This)->lpVtbl->GetCheckinVersionComments(This,Comments,cch)

#define IMsoPKMClient_InvalidateAllowedOpsCache(This) \
	(This)->lpVtbl->InvalidateAllowedOpsCache(This)

#define IMsoPKMClient_GetDocumentState(This, ppkmdocstate) \
	(This)->lpVtbl->GetDocumentState(This, ppkmdocstate)

#define IMsoPKMClient_DoAlreadyCheckedOutDlg(This) \
	(This)->lpVtbl->DoAlreadyCheckedOutDlg(This)

#define IMsoPKMClient_GetLastHResult(This) \
	(This)->lpVtbl->GetLastHResult(This)

#define IMsoPKMClient_ClearLastHResult(This) \
	(This)->lpVtbl->ClearLastHResult(This)

#endif  // !__cplusplus


// {2B82E7EF-0FD9-4728-937C-20A857D1B7EA}
DEFINE_GUID(IID_IMsoPKMClient2, 
	0x2b82e7ef, 0xfd9, 0x4728, 0x93, 0x7c, 0x20, 0xa8, 0x57, 0xd1, 0xb7, 0xea);
#ifdef INTERFACE
#undef INTERFACE
#endif //INTERFACE
#define INTERFACE IMsoPKMClient2

DECLARE_INTERFACE_(IMsoPKMClient2, IMsoPKMClient)
{
	MSOMETHOD (GetCheckedOutUser) (THIS_ LPWSTR wzUser, int *pcchUser) PURE;
	MSOMETHOD (SetReadWriteOnSave) (THIS_ BOOL fReadWriteOnSave) PURE;
	MSOMETHOD (IsReadWriteOnSave) (THIS_ BOOL *pfReadWriteOnSave) PURE;
	MSOMETHOD (ChangesSinceOpen) (THIS_ BOOL *pfChanges) PURE;
	MSOMETHOD (CheckoutShort) (THIS) PURE;
	MSOMETHOD (GetVirusInfo) (THIS_ long *piStatus, WCHAR *wzInfo, DWORD cchInfo) PURE;
	MSOMETHOD (GetLastModifiedTime) (THIS_ FILETIME *pFT) PURE;
	MSOMETHOD (GetLastModifiedUser)(WCHAR *wzUser, int *pcchUser);	
	MSOMETHOD (GetPropString)(const WCHAR *wzKey, WCHAR *wzValue, ULONG *pcchValue);
	MSOMETHOD (SetPropString)(const WCHAR *wzKey, const WCHAR *wzValue);
	MSOMETHOD (InvalidateAllowedOpsCacheEx)(DWORD grfpkmmd);
};


/*-----------------------------------------------------------------------------
	MsoFCreate_MsoPKMClient

	This call will create the MsoPKMClient object.  It is used in place of
	CoCreateInstance because the client object is used internally and there
	is no CLSID registered for it in the registry.  Also, to boost performance.

	Arguments -
			ppfnp : address of newly allocated MsoFilePKMClient instance

------------------------------------------------------------------ SHANEM -*/
MSOAPI_(BOOL) MsoFCreate_MsoPKMClient(IMsoPKMClient **ppfnp);


/*-----------------------------------------------------------------------------
	HrTahoeGetFolderProperties

	Gets specific properties about a Tahoe folder.  

	Arguments -
			wzUrlPath - path to folder
			hWndAuth - parent hwnd for any authentication windows issued during bind
			pfEnhance - returns whether Tahoe folder is enhanced
			
			---- cut --- 
			 - O10_237145, O10_258639:  This was a check for properties which would exist on future versions of
				Tahoe, to disable redirection within office for versioned files, and to enable saving HTML
				thickets to classic folders.
			pfRedirectionSupported - whether Tahoe supports redirection on the server or not
			pfThicketsSupported - pfThicketsSupported

------------------------------------------------------------------ SHANEM -*/

HRESULT HrTahoeGetFolderProperties(LPCWSTR wzUrlPath, HWND hWndAuth, BOOL *pfEnhanced 
//O10_237145:                      , BOOL *pfRedirectionSupported
//O10_258639:                      , BOOL *pfThicketsSupported
                                   );

/*-----------------------------------------------------------------------------
	FreePKM
	
	Free/Release global pointers

------------------------------------------------------------------ SHANEM -*/
void FreePKM();

#pragma pack( pop, msopkmcl )

#endif  // MSOPKMCL_H
