/****************************************************************************
	Msosmartdoc.h

	Owner: PaulBrom
	Copyright (c) 2001 Microsoft Corporation

	This file contains the exported interfaces and declarations for
	the Smart Documents feature.
	
****************************************************************************/

#pragma once

#ifndef MSOSMARTDOC_H
#define MSOSMARTDOC_H
#if 0 //$[VSMSO]

#include <facapi.h>
#include <msomanifest.h>

#pragma pack(push, 4)

// solution prop lookup values
typedef enum
{
	msoispSolutionURL = 0,
	msoispSolutionID,
	msoispTemplateID,
	msoispWhiteRabbitID,
	msoispMax					
} MSOISP;

#define msoispSmartDocMax	(int)msoispTemplateID + 1

// solution download behavior types
typedef enum
{
	msodbtAsk = 0,	// need to ask
	msodbtYes,
	msodbtAlways,
	msodbtNo,
	msodbtNever,
	msodbtMax
} MSODBT;

// Smart Document Schema Span Data structure
typedef struct _MSOSDSSD
{
	IDispatch* pdispTarget;
	BSTR bstrText;
	BSTR bstrXml;
} MSOSDSSD;	

// Smart Document Schema Element structure
#ifdef __cplusplus
typedef struct _MSOSDSE : public MSOSDSSD
{
	WCHAR* wzUri;
	WCHAR* wzName;
} MSOSDSE;
#else
typedef struct _MSOSDSE
{
	union
	{
		MSOSDSSD sdssd;
		MSOSDSSD;		
	};
	
	WCHAR* wzUri;
	WCHAR* wzName;
} MSOSDSE;
#endif // __cplusplus

// smart document solution data
typedef struct _MSOSDSD
{
	WCHAR* wzName;
	WCHAR* wzManifestURL;
	WCHAR* wzSolutionID;
	WCHAR* wzVersion;
	WCHAR* wzSchemaURI;
	WCHAR* wzSolutionPath;
	BOOL fCanDelete;
	BOOL fShared;
} MSOSDSD;

#define cchMaxFilter	255

// solution file enumeration structure
typedef struct _MSOSLFE
{
	UINT uiSolution;
	UINT uiFileNext;
	HKEY hkeyEnum;
	WCHAR wzFilter[cchMaxFilter];
	WCHAR* wzFilePath;
	WCHAR* wzFileName;
	WCHAR* wzAlias;
	DWORD dwVersion;
	DWORD dwRemote;
	WCHAR* wzTemplateID;
	BOOL fTemplatesOnly;
} MSOSLFE;

#pragma pack(pop)


/****************************************************************************
	Defines the IMsoSmartDocument interface

**************************************************************** PAULBROM **/
#undef  INTERFACE
#define INTERFACE  IMsoSmartDocument

DECLARE_INTERFACE_(IMsoSmartDocument, ISimpleUnknown)
{
	// ISimpleUnknown methods
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;

	/* Standard FDebugMessage method */
	MSODEBUGMETHOD		

	MSOMETHOD_(BOOL, FGetSdsd) (THIS_ MSOSDSD** ppsdsd) PURE;
	
	MSOMETHOD_(void, BeginSchemaUpdateEx) (THIS_ IDispatch* pdispWholeDoc, int fAddDefault, int fAddActionsForEntireSchema) PURE;
	MSOMETHOD_(void, BeginSchemaUpdate) (THIS_ IDispatch* pdispWholeDoc) PURE;
	MSOMETHOD_(IFactoidPropertyBag*, PfpbAddSchemaElement) (THIS_ MSOSDSE* psdse) PURE;
	MSOMETHOD_(void, EndSchemaUpdate) (THIS_ IMsoToolbarSet* pitbs) PURE;
	MSOMETHOD_(BOOL, FShowSmartDocPane) (THIS_ IMsoToolbarSet* pitbs) PURE;
	MSOMETHOD_(BOOL, FGetFactoidPresentInPane) (THIS_ IFactoidPropertyBag* pfpb,
												int iAction) PURE;
	MSOMETHOD_(void, Refresh) (THIS_ IMsoToolbarSet* pitbs, BOOL fForce) PURE;
	MSOMETHOD_(BOOL, FPaneShowsOtherSmartDoc) (THIS_ IMsoToolbarSet* pitbs) PURE;

	// link/button invoking functions
	MSOMETHOD_(void, InvokeButtonOrLink) (THIS_ IFactoidPropertyBag* pfpb, 
										  int iAction) PURE;

	// help item manipulation functions	
	MSOMETHOD_(void, SetExpandHelpItem) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction, 
									     BOOL fExpand) PURE;
	MSOMETHOD_(BOOL, FGetExpandHelpItem) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction) PURE;

	// checkbox manipulation functions	
	MSOMETHOD_(void, SetCheckboxState) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction, BOOL fChecked) PURE;
	MSOMETHOD_(BOOL, FGetCheckboxState) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction) PURE;

	// textbox manipulation functions	
	MSOMETHOD_(void, SetTextboxText) (THIS_ IFactoidPropertyBag* pfpb, 
									  int iAction, BSTR bstrText) PURE;
	MSOMETHOD_(BOOL, FGetTextboxText) (THIS_ IFactoidPropertyBag* pfpb, 
									   int iAction, BSTR* pbstrText) PURE;	

	// textbox manipulation functions	
	MSOMETHOD_(BOOL, FSetListboxComboSel) (THIS_ IFactoidPropertyBag* pfpb, 
									  	   int iAction, int iSel) PURE;
	MSOMETHOD_(int, IGetListboxComboSel) (THIS_ IFactoidPropertyBag* pfpb, 
									      int iAction) PURE;		
	
	// documentfragment item manipulation functions	
	MSOMETHOD_(void, SetExpandDocumentFragmentItem) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction, 
									     BOOL fExpand) PURE;
	MSOMETHOD_(BOOL, FGetExpandDocumentFragmentItem) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction) PURE;

	// radio group manipulation functions	
	MSOMETHOD_(BOOL, FSetRadioGroupSel) (THIS_ IFactoidPropertyBag* pfpb, 
									     int iAction, int iSel) PURE;
	MSOMETHOD_(int, IGetRadioGroupSel) (THIS_ IFactoidPropertyBag* pfpb, 
									    int iAction) PURE;	

	// activeX manipulation functions
	MSOMETHOD_(IDispatch*, IDispatchGetActiveXControl) (THIS_ IFactoidPropertyBag* pfpb,
													int iAction) PURE;

	MSOMETHOD_(void, Delete) (THIS_ BOOL fAlsoDeleteProps) PURE;
};

// app specific factoid restart proc  - needed because FDeleteSolution may need to restart the factoid servers.
typedef void (CALLBACK *PFNFRESTARTPROC)(void);

/****************************************************************************
	Defines the IMsoSolutionLibrary interface

**************************************************************** PAULBROM **/
#undef  INTERFACE
#define INTERFACE  IMsoSolutionLibrary

DECLARE_INTERFACE_(IMsoSolutionLibrary, IUnknown)
{
	// IUnknown methods
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;
	MSOMETHOD_(ULONG, AddRef) (THIS) PURE;
	MSOMETHOD_(ULONG, Release) (THIS) PURE;

	/* Standard FDebugMessage method */
	MSODEBUGMETHOD		

	MSOMETHOD_(BOOL, FAddSchemaToLibrary) (THIS_ const WCHAR* wzSchemaUri,
										   BOOL fShared,
										   BOOL fIncludeSpecific) PURE;
	MSOMETHOD_(BOOL, FAddAllSchemasToLibrary) (THIS_ BOOL fIncludeSpecific) PURE;
	MSOMETHOD_(BOOL, FAddAttachedSolutionToLibrary) (THIS_ IMsoSmartDocument* psd) PURE;
	MSOMETHOD_(UINT, CSolutions) (THIS) PURE;
	MSOMETHOD_(BOOL, FGetSolution) (THIS_ UINT uiSolution, MSOSDSD** ppsdsd) PURE;
	MSOMETHOD_(BOOL, FChooseSolutionFromDlg) (THIS_ HMSOINST hinst, 
											  HWND hwndParent, 
											  MSOSDSD** ppsdsd);
	MSOMETHOD_(BOOL, FDeleteSolution) (THIS_ UINT uiSolution, BOOL fPrompt, PFNFRESTARTPROC pRestartProc, BOOL* pfUserAborted) PURE;
	MSOMETHOD_(BOOL, FFindSolution) (THIS_ const WCHAR* wzSolutionID, MSOSDSD** ppsdsd) PURE;
	MSOMETHOD_(BOOL, FEnumSolutionFiles) (THIS_ MSOSLFE* pslfe) PURE;	
	MSOMETHOD_(BOOL, FSolutionNeedsUpdate) (THIS_ MSOSDSD* psdsd) PURE;
};


/****************************************************************************
	Defines the IMsoSmartDocEventHandler interface

**************************************************************** PAULBROM **/
#undef  INTERFACE
#define INTERFACE  IMsoSmartDocEventHandler

DECLARE_INTERFACE_(IMsoSmartDocEventHandler, IMsoManifestEventHandler)
{
	// ISimpleUnknown methods
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;

	/* Standard FDebugMessage method */
	MSODEBUGMETHOD		

	// IMsoManifestEventHandler methods
	MSOMETHOD_(void, OnNeedSchemaBuggedList) (THIS_ void* pv, IMsoSchemaBuggedList** ppsbl) PURE;	
	MSOMETHOD_(void, OnInstallSolutionUri) (THIS_ void* pv, const WCHAR* wzUri, const WCHAR* wzSolutionID, const WCHAR* wzAlias, MSOMST mst) PURE;
	MSOMETHOD_(void, OnSchemaInstall) (THIS_ void* pv, const WCHAR* wzUri, const WCHAR* wzPath) PURE;
	MSOMETHOD_(void, OnSmartTagInstall) (THIS_ void* pv, const WCHAR* wzCLSID, const WCHAR* wzPath) PURE;
	MSOMETHOD_(void, OnCOMAddinInstall) (THIS_ void* pv, const WCHAR* wzCLSID, const WCHAR* wzPROGID, 
										 const WCHAR* wzPath) PURE;
	MSOMETHOD_(BOOL, OnTemplateInstallBefore) (THIS_ void* pv, const WCHAR* wzTemplateID, const WCHAR* wzPath) PURE;
	MSOMETHOD_(void, OnTemplateInstallAfter) (THIS_ void* pv, const WCHAR* wzTemplateID, const WCHAR* wzPath) PURE;
	MSOMETHOD_(void, OnInstallComplete) (THIS_ void* pv);	
	MSOMETHOD_(void, OnDestroyPv) (THIS_ void* pv) PURE;

	// called when we need to make sure the host document attaches a given schema.
	// sue me.)
	MSOMETHOD_(BOOL, OnNeedToAddSchemasToLibrary) (THIS_ void* pv, IMsoSolutionLibrary* psl, BOOL fIncludeSpecific) PURE;

	// called when we need to make sure the host document attaches a given schema.
	MSOMETHOD_(BOOL, OnNeedToAttachSchema) (THIS_ void* pv, const WCHAR* wzUri) PURE;

	// called when the smart doc is being destroyed
	MSOMETHOD_(void, OnSmartDocRevoke) (THIS_ void* pv, BOOL fDelete) PURE;

	// called when someone changes the active smart document
	MSOMETHOD_(void, OnNewSmartDoc) (THIS_ void* pv, IMsoSmartDocument* psd) PURE;

	// called when we have dirtied the doc by changing the lpudobj custom props
	MSOMETHOD_(void, OnDirtyDoc) (THIS_ void* pv) PURE;

	// called when we need to get an IDispatch referring to the document the smartdoc will be in (like Word::Document)
	MSOMETHOD_(void, OnNeedDocumentDispatch) (THIS_ void* pv, IDispatch** ppdisp) PURE;

	// called when we need a LPUDOBJ from the current doc.
	MSOMETHOD_(BOOL, OnNeedLpudobj) (THIS_ void* pv, LPUDOBJ* plpudobj) PURE;
};


// this function tells you if smart docs are disabled by policy.
MSOAPI_(BOOL) MsoFSmartDocsDisabledByPolicy();

// flags for MsoFLoadSmartDocSolution
#define msoflsdNormal		0x0000000
#define msoflsdInMacro		0x0000001
#define msoflsdNoSolPick	0x0000002
#define msoflsdNoSolDown	0x0000004
#define msoflsdOnlyLoadManifest	0x0000008
#define msoflsdForceReinstall	0x0000010

// this function loads a smart document solution described by the given strings
MSOAPI_(BOOL) MsoFLoadSmartDocSolution (HMSOINST hinst,
										HWND hwndParent,
										int lcid,
										LPUDOBJ lpudobj,
										void* pv,
										IMsoSmartDocEventHandler* psdeh,									
										IFactoidServer* pfs, 
										int grflsd,
										IMsoSmartDocument **ppsd);

// this functions manipulate solution props which are attached to a doc's lpudobj
// these props are used to attach smart document solutions to documents
MSOAPI_(BOOL) MsoFSolutionRefInDoc (LPUDOBJ lpudobj);
MSOAPI_(BOOL) MsoFSetSolutionPropForDoc (LPUDOBJ lpudobj, MSOISP isp, const WCHAR* wzPropVal);
MSOAPI_(void) MsoDeleteSolutionPropForDoc(LPUDOBJ lpudobj, MSOISP isp);
MSOAPI_(BOOL) MsoFSetAllSolutionPropsForDoc (LPUDOBJ lpudobj, MSOSDSD* psdsd);
MSOAPI_(WCHAR*) MsoWzGetSolutionPropInDoc (LPUDOBJ lpudobj, MSOISP isp);	
MSOAPI_(void) MsoClearSolutionPropsInDoc (LPUDOBJ lpudobj, BOOL fForce);
MSOAPI_(BOOL) MsoFClearSolutionPropsInDocAndSetNull(LPUDOBJ lpudobj);
MSOAPI_(BOOL) MsoFCloneSolutionPropsInDoc (LPUDOBJ lpudobjSrc, LPUDOBJ lpudobjDest);
MSOAPI_(BOOL) MsoFWzIsSmartDocPropName(const WCHAR* wzPropNameIn);
MSOAPI_(void) MsoIncrementSmartDocSQMCount(LPUDOBJ lpudobj, BOOL fOpen);

// this function displays a dialog to learn the download behavior the user wants for the
// smart doc.
MSOAPI_(BOOL) MsoFSetDownloadBehaviorForUri(const WCHAR* wzSchemaUri, 
											MSODBT dbt, 
											BOOL fOverwrite);
MSOAPI_(MSODBT) MsodbtGetDownloadBehaviorFromUri(const WCHAR* wzSchemaUri, 
											     const WCHAR* wzSolutionURL, 
											     HMSOINST hmsoinst, 
											     HWND hwndParent,
												 IMsoSchemaBuggedList* psbl,
											     BOOL fInMacro,
												 BOOL fUpdating);


// Used by the apps to retrieve MSO's SmartDocument IDispatch interface. This is used for
// object model support for the SmartDocument interface
MSOAPI_(BOOL) MsoFGetIDispatchSmartDoc(HMSOINST hmsoinst, 
									   HWND hwndParent,
									   int lcid,
									   LPUDOBJ lpudobj,
									   void* pv,
									   IMsoSmartDocEventHandler* psdeh,		
									   IFactoidServer* pfs,
									   IMsoSmartDocument* psdCur,
									   IDispatch** ppidisp);

// this function creates a manifest library (for C users)
MSOAPI_(BOOL) MsoFCreateSolutionLibrary (MSOMST mst, HMSOINST hinst, int lcid, IMsoSolutionLibrary **ppsl);

// these functions can manipulate MSOSDSD's created by the solution library.
MSOAPI_(void) MsoFreeSdsd(MSOSDSD* psdsd);
MSOAPI_(BOOL) MsoFCopySdsd(const MSOSDSD* psdsdIn, MSOSDSD** ppsdsdOut);
#ifdef DEBUG
MSOAPI_(void) MsoMarkSdsd(MSOSDSD* psdsd, HMSOINST hinst, UINT message, LPARAM lParam, WPARAM wParam);
#endif // DEBUG

MSOAPI_(HENHMETAFILE) MsoHemhFromWzPath(HDC hdc, const WCHAR* wzImage);
MSOAPI_(HBITMAP) MsoHBitmapFromWzPath(HDC hdc, const WCHAR* wzImage);
MSOAPI_(WCHAR*) MsoWzUnicodeFromStream(IStream* pstm, MLDETECTCP mldetectcp);
MSOAPI_(BOOL) MsoFSmartDocPaneVisible(IMsoToolbarSet* pitbs);

// white rabbit support
MSOAPI_(BOOL) MsoFLoadWhiteRabbitAssembly(const WCHAR *wzDocPath,
										  const WCHAR *wzDocName,
										  BSTR	bstrWRPath,
										  BSTR	bstrWRName,
										  IDispatch *pdispApp,
										  IDispatch *pdispDoc,
										  IUnknown **ppunkAssembly);

MSOAPI_(BOOL) MsoFUnloadWhiteRabbitAssembly(IUnknown *punkAssembly);
MSOAPI_(void) MsoUnloadWhiteRabbitLoader(void);
MSOAPI_(BOOL) MsoFWhiteRabbitPropertyPresent(LPUDOBJ lpudobj, 
											 const WCHAR *wzDocPath, 
											 const WCHAR *wzDocName, 
											 BSTR *pbstrWRPath, 
											 BSTR *pbstrWRName);
					  
#endif //$[VSMSO]
#endif // MSOSMARTDOC_H
