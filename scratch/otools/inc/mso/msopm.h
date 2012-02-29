#pragma once

/****************************************************************************
	MSOPM.H

	Owner: ShamikB
	Copyright (c) 1994 Microsoft Corporation

	Exported declarations for all Office 9 Programmability
****************************************************************************/

#ifndef MSOPM_H
#define MSOPM_H

#ifndef VBE

#if !defined(MSOSTD_H)
#include <msostd.h>
#endif

#include <msoalloc.h>
#if 0 //$[VSMSO]
#include <msopmi.h>
#endif //$[VSMSO]

#ifdef __cplusplus
interface IVbaProject;
#else  // __cplusplus
typedef interface IVbaProject IVbaProject;
#endif  // __cplusplus

#ifdef __cplusplus
#include "msotl.h"
#else  // __cplusplus
#include "msotlc.h"
#endif  // __cplusplus
#if 0 //$[VSMSO]
#include "msopmi.h"
#endif //$[VSMSO]

#endif // VBE

#if defined(__cplusplus)
extern "C" {
#endif

// minimum version of oleaut32.dll needed to match current drop of VBE/VBA
#define wOAVerMin		4265

// VBE version number
#define uVBEVerRef		9759
#define uVBEVerMajor	6
#define uVBEVerMin		0

#define usVBEStorageVerDelta	3

#if 0 //$[VSMSO]
// Office typelib version number
#define LIBID_Office_VerMajor	2
#define LIBID_Office_VerMinor	3
#else //$[VSMSO]
// VSMSO typelib version number
#define LIBID_Microsoft_VisualStudio_CommandBars_VerMajor	8
#define LIBID_Microsoft_VisualStudio_CommandBars_VerMinor	0
#endif //$[VSMSO]


#if 0 //$[VSMSO]
#ifndef VBE

#include "msowc.h"

MSOAPI_(IMsoBlip *) MsoScriptAnchorBlip(MsoScriptLanguage msoScript, BOOL fScriptAnchor);
MSOAPI_(IMsoBlip *) MsoAnchorBlip(MsoAnchorType msoat, BOOL fOnlyEnsure);

typedef struct SBCA	// script block information
{
	long	isb;		// script block ID
	long	cpFirst;	// cpFirst of HTML script block
	long	cpLim;		// cpLim of HTML script block
} SBCA;

MSOAPI_(BOOL) MsoFRemoveAllScriptsAlert();
MSOAPI_(BOOL) MsoFInsertScriptMenuEnabled(IMsoDrawingSelection *pidgsl);
MSOAPI_(BOOL) MsoFScriptAnchorSelected(IMsoDrawingSelection *pidgsl, BOOL fExclusive);
#define msoAttachedScript 1
#define msoAttachScript 2
#define msoDontAttachScript 3
#define msoAlreadyAttachedScript 4
MSOAPI_(int) MsoAttachScript(IMsoDrawingSelection *pidgsl);
MSOAPI_(VOID) MsoGetWebBotTip(WCHAR *wzWebBotTip, int cchMax);
MSOAPI_(void) MsoToggleAnchorVisibilityForPrintPreview(IMsoDrawing *pidg, BOOL fBefore);
MSOAPI_(BOOL) MsoFPlaceholderAnchor(MSOHSP hsp);

#ifdef __cplusplus
interface Scripts;
#else  // __cplusplus
typedef interface Scripts Scripts;
#endif  // __cplusplus

typedef struct MSOSCRIPTCA
{
	long cpFirst;
	long cpLim;
	void *pvDoc;
} MSOSCRIPTCA;

// Script Tag Descriptor
typedef struct
{
	MsoScriptLanguage Lang;
	DWORD cch;
	DWORD cchMac;
	WCHAR *wzIdAttr;
	WCHAR *wzLangAttr;
	WCHAR *wzExtAttr;
} MSOSCRIPTSTD;

#undef  INTERFACE
#define INTERFACE IMsoScripts
DECLARE_INTERFACE_(IMsoScripts, ISimpleUnknown)
{
	// ----- ISimpleUnknown methods

	MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;

	// ----- IMsoScripts methods

	/* FDebugMessage */
	MSODEBUGMETHOD

	MSOMETHOD_(int, Count) (THIS) PURE;
	MSOMETHOD_(IDispatch *, GetDispScripts) (THIS_ IMsoDrawingUserInterface *pdgui, IDispatch *pidispParent, void *pvprog) PURE;
	MSOMETHOD_(BOOL, FAdd) (THIS_ SBCA *psbca, void *pvSubDoc) PURE;
	MSOMETHOD_(BOOL, FRemove) (THIS_ SBCA *psbca) PURE;
	MSOMETHOD_(BOOL, FWriteStm) (THIS_ IStream *pistm) PURE;
	MSOMETHOD_(void, RefreshScriptsView) (THIS_ IMsoDrawing *pidg) PURE;
	MSOMETHOD_(void *, PLookupHsp) (THIS_ MSOHSP hsp) PURE;
	MSOMETHOD_(void *, GetPvDoc) (THIS) PURE;
	MSOMETHOD_(IDispatch *, GetDispRangeScripts) (THIS_ IMsoDrawingUserInterface *pdgui, IDispatch *pidispParent, MSOSCRIPTCA *pca) PURE;
	MSOMETHOD_(HRESULT, HrGetDispScript) (THIS_ MSOHSP hsp, Script **ppDispScript) PURE;
	MSOMETHOD_(void, SetRangeDirty) (THIS_ void *pvDRS) PURE;
	MSOMETHOD_(HRESULT, HrShowScriptAnchor)(THIS_ BOOL fFlag) PURE;
	MSOMETHOD_(HRESULT, HrSingleShowScriptAnchor)(THIS_ BOOL fFlag) PURE;
	MSOMETHOD_(BOOL, FScriptAnchorVisible) (THIS) PURE;
	MSOMETHOD_(void *, GetPvDgs) (THIS_ IMsoDrawing *pidg) PURE;
	MSOMETHOD_(void, ToggleVisibilityForPrintPreview) (THIS_ BOOL fBefore) PURE;
	MSOMETHOD_(BOOL, FVisibilityToggleDisabled) (THIS) PURE;
	MSOMETHOD_(void, SetDefScriptLang) (THIS_ MsoScriptLanguage DefLang, const WCHAR *wzDefLang) PURE;
	MSOMETHOD_(MsoScriptLanguage, GetDefScriptLang) (THIS_ WCHAR **pwzDefLang) PURE;
	MSOMETHOD_(void, RemoveAllHeadScripts) (THIS) PURE;
	MSOMETHOD_(BOOL, FAddHead) (THIS_ void *pvSubDoc, MSOSCRIPTSTD *pstd,
		const WCHAR *wzScript, void **ppvScript) PURE;
	MSOMETHOD_(void, HeadCopy) (THIS_ IMsoScripts *piscripts) PURE;

};
MSOAPI_(IMsoScripts *) MsoCreateScripts(void *pvDoc, HMSOINST hmsoinst);
MSOAPI_(void) MsoDestroyScripts(IMsoScripts **ppScripts);
MSOAPI_(void) MsoRemoveAllBodyScripts(IMsoDrawing *pidg);

#define REGKEY_VSEROOTW			L"Software\\Microsoft\\MSE\\9.0"

// VSE project types
#define VSEPROJTYPE_NULL	(-1)
#define VSEPROJTYPE_FIRST	(0)
#define VSEPROJTYPE_WORD	(0)
#define VSEPROJTYPE_EXCEL	(1)
#define VSEPROJTYPE_PPT		(2)
#define VSEPROJTYPE_ACCESS	(3)
#define VSEPROJTYPE_FP		(4)
#define VSEPROJTYPE_XDOCS	(5)
#define VSEPROJTYPE_LAST	(5)

typedef DWORD VSEPROJTYPE;


// VSE project item types
#define VSEPROJITEMTYPE_NULL		(-1)
#define VSEPROJITEMTYPE_DEFAULT		(0)
#define VSEPROJITEMTYPE_FRAMELINK	(1)

#define VSEPROJITEMTYPE_MASK		(0xFFFF)
#define VSEPROJITEMTYPE_INVARIANTITEMID	(0x00010000)

typedef DWORD VSEPROJITEMTYPE;


#define VSEFRAMENAVRESULT_FAIL		(0)
#define VSEFRAMENAVRESULT_SUCCEEDED	(1)
#define VSEFRAMENAVRESULT_NOTLOADED	(2)
#define	VSEFRAMENAVRESULT_OBSOLETE	(3)

typedef DWORD VSEFRAMENAVRESULT;


// VSE view type
#define VSEVIEW_NULL			(0)
#define VSEVIEW_VIEWSOURCE		(1)
#define VSEVIEW_HTML			(2)
#define VSEVIEW_HTMLDESIGN   	(3)
#define VSEVIEW_HTMLSOURCE   	(4)

#define VSEVIEW_MAYBE		   	(0x10000)
#define VSEVIEW_PROMPT		   	(0x20000)
#define VSEVIEW_MASK		   	(0x0FFFF)

typedef DWORD VSEVIEW;

MSOAPI_(DWORD) MsoDwMseValidateView(DWORD dwView, BOOL fAlert);

// VSE refresh type
#define	VSEREFRESH_NONE		(0x00000)
#define VSEREFRESH_FORCE	(0x00001)
#define	VSEREFRESH_RELOAD	(0x00002)
#define	VSEREFRESH_SAVE		(0x00004)
#define	VSEREFRESH_SAVETEMP	(0x00008)
#define	VSEREFRESH_IGNORE	(0x00010)

#define	VSEREFRESH_MASK		(0xFFFFF)
#define	VSEREFRESH_PROMPT	(0x10000)
#define	VSEREFRESH_AUTOMATION	(0x20000)

typedef DWORD VSEREFRESH;


// VSE project creation flags
#define VSEPROJCREATE_NULL			(0x0000)
#define VSEPROJCREATE_LOCK			(0x0001)
#define VSEPROJCREATE_CONNECT		(0x0002)
#define VSEPROJCREATE_MAYBE			(0x0004)
#define VSEPROJCREATE_VIEWSOURCE	(0x0008)

typedef DWORD VSEPROJCREATE;

MSOAPI_(HRESULT) MsoHrVseInitialize(HMSOINST hinst, VSEVIEW vvView, int fShowErrors, int fActivate, IMsoVsPackage **pIMsoVsPackage);
MSOAPI_(void) MsoVseUninit(HMSOINST hinst);
MSOAPI_(HRESULT) MsoHrUpdateProjectsList(void);

MSOAPI_(HRESULT) MsoHrUpdateScript(IMsoDrawing *pidg, MSOHSP hsp, const WCHAR *wszHtml);

MSOAPI_(BOOL) MsoFVseShown(void);

MSOAPI_(BOOL) MsoFInVseCall(void);

#ifdef DEBUG
MSOAPI_(BOOL) MsoFSaveBeVseProxies(LPARAM lParam);
#endif

MSOAPI_(void) MsoMseSetForegroundWindow(IUnknown *pSet, HWND hwnd);

/*-----------------------------------------------------------------------------

		Enabling/Disabling scripting features depending on a version
		of IE available.

------------------------------------------------------------------- IgorZ ---*/

#define msoscftMin			(-1)
#define msoscftNull			(-1)
#define msoscftViewSource	(0)
#define msoscftViewMse		(1)
#define msoscftMax			(msoscftViewMse+1)

MSOAPI_(BOOL) MsoFScriptingFeatureEnabled(int msoscft, int fUseDemandInstall);
MSOAPI_(BOOL) MsoFUseScriptingFeature(int msoscft);
MSOAPI_(BOOL) MsoFWebScriptingEnabled(void);

/*-----------------------------------------------------------------------------
------------------------------------------------------------------- IgorZ ---*/
MSOAPIX_(HRESULT) MsoHrGetHtmlEncodingWz(const WCHAR *wzFileName, int *pcpg);
MSOAPI_(HRESULT) MsoHrInitEncodedRead(IStream *pistm, int *pcpg);
MSOAPI_(HRESULT) MsoHrInitEncodedReadEx(IStream *pistm, int *pcpg, BOOL *pfUTF8BOM);
MSOAPI_(HRESULT) MsoHrReadEncoded(_Out_cap_(cchToRead) WCHAR *rgwch, int cchToRead, int *pcchRead);
MSOAPI_(void) MsoFinishEncodedRead(void);

MSOAPI_(HRESULT) MsoHrInitEncodedWrite(IStream *pistm, int cpg);
MSOAPI_(HRESULT) MsoHrInitEncodedWriteEx(IStream *pistm, int cpg, BOOL fUTF8BOM);
MSOAPI_(HRESULT) MsoHrWriteEncoded(const WCHAR *rgwch, int cchToWrite);
MSOAPI_(HRESULT) MsoHrFinishWriteEncoded(void);


/*-----------------------------------------------------------------------------
	MsoHrCreateVseProject

------------------------------------------------------------------- IgorZ ---*/
MSOAPI_(HRESULT) MsoHrCreateVseProject(
	HMSOINST hinst,
	VSEPROJCREATE projcreate,
	IMsoVsProjectSite *psite,
	const WCHAR *wzDisplayName, 				// project display name
	const WCHAR *wzPersistentName, 			// project presistent name
	DWORD dwProjType,
	IMsoVsProject **pproj);


/*-----------------------------------------------------------------------------
	MsoFVseBuildLockedErrorMessage

	Builds "buffer is locked" error message:

		The HTML source for the <document, workbook> is locked because the <document,
		workbook> has been modified in Microsoft <Word, Excel>.

		To modify the HTML source, first either click Refresh on the Refresh toolbar
		to incorporate the changes into the HTML source or click Do Not Refresh to
		ignore the changes in the <document, workbook>.

------------------------------------------------------------------- IgorZ ---*/
MSOAPI_(int) MsoFVseBuildLockedErrorMessage(WCHAR *wz, int cch, VSEPROJTYPE type, int fHost);

MSOAPI_(int) MsoFMseVBProjectChangedYesNoAlert(void);

MSOAPI_(void) MsoMseDisplayHostDocumentLockedAlert(VSEPROJTYPE type);

MSOAPI_(void) MsoGetStatusStringForMSE(WCHAR *wzStatus, int cch);

/*---------------------------------------------------------------------------


			VSE components name table

------------------------------------------------------------------- IgorZ -*/

MSOAPI_(int) MsoFGetNameFromVseNameTablePpx(MSOPX *ppx, long hetk, void *pvData, WCHAR *wtzName);
MSOAPIX_(MSOPX *) MsoPpxNewVseNameTable(void);
MSOAPIX_(void) MsoFreeVseNameTablePpx(MSOPX *ppx);
#ifdef DEBUG
MSOAPI_(void) MsoSaveBeVseNameTablePpx(MSOPX *ppxItems);
#endif

#undef  INTERFACE
#define INTERFACE  IMsoHTMLFileNameTable
DECLARE_INTERFACE_(IMsoHTMLFileNameTable, IUnknown)
{
	// ------------------------------------------------------------------------
	// Begin Interface
	//
	// This is important on the Mac.
	BEGIN_MSOINTERFACE

	// ------------------------------------------------------------------------
	// IUnknown Methods

	MSOMETHOD(QueryInterface) (THIS_ REFIID refiid, void * * ppvObject) PURE;

	MSOMETHOD_(ULONG, AddRef) (THIS) PURE;

	MSOMETHOD_(ULONG, Release) (THIS) PURE;

	// ------------------------------------------------------------------------
	// Standard Office Debug method
	MSODEBUGMETHOD

	// ------------------------------------------------------------------------
	// IMsoHTMLFileNameTable methods


	MSOMETHOD_(HRESULT, HrAddName) (THIS_ void *pvData, MSOHETN hetn, MSOHETN hetnPfx, const WCHAR *wzName, const WCHAR *wzPubID, MSOFNTENT entType) PURE;

	MSOMETHOD_(HRESULT, HrGetName) (THIS_ MSOHETK hetk, MSOHETK hetkPfx, void *pvData, WCHAR *wtzName, int cchWtzName) PURE;

	MSOMETHOD_(BOOL, FGetNameWz) (THIS_ MSOHESUI *phesui, const WCHAR *wzNameGet, MSOHEGDN hegdn, MSOVSEITEM *pvseItem) PURE;

	MSOMETHOD_(HRESULT, HrResetEnumTokens) (THIS) PURE;

	MSOMETHOD_(BOOL, FGetNextTokens) (THIS_ MSOVSEITEM *pvseItem) PURE;

	MSOMETHOD_(BOOL, FSetEntryType) (THIS_ MSOFNTENT entType) PURE;

	MSOMETHOD_(HRESULT, HrSetExportData) (THIS_ MSOHEGDN hegdn, const WCHAR *wzFolderSuffix) PURE;

	MSOMETHOD_(MSOPXUHI**, Pppxuhi) (THIS) PURE;

	MSOMETHOD_(HRESULT, HrLockFiles) (THIS_ BOOL fLock) PURE;

	MSOMETHOD_(HRESULT, HrGetFileHandle) (THIS_ MSOHETK hetk, MSOHETK hetkPfx, void *pvData, DWORD dwProcessId, HANDLE *phFile) PURE;
};

MSOAPI_(HRESULT) MsoHrCreateIMsoHTMLFileNameTable(IMsoHTMLFileNameTable **pphfnt, interface IMsoOLDocument *pioldoc);
#ifdef NEVER
MSOAPIX_(HRESULT) MsoHrCreateIMsoHTMLFileNameHashTable(IMsoHTMLFileNameTable **pphfnt, interface IMsoOLDocument *pioldoc);
#endif // NEVER

MSOAPIDBG_(void) MsoSaveBeIMsoHTMLFileNameTable(IMsoHTMLFileNameTable *phfnt);


/*---------------------------------------------------------------------------


			IMsoVsProject proxy

------------------------------------------------------------------- IgorZ -*/

#undef  INTERFACE
#define INTERFACE  IMsoVsProjectProxy
DECLARE_INTERFACE_(IMsoVsProjectProxy, IUnknown)
{
	// ------------------------------------------------------------------------
	// Begin Interface
	//
	// This is important on the Mac.
	BEGIN_MSOINTERFACE

	// ------------------------------------------------------------------------
	// IUnknown Methods

	MSOMETHOD(QueryInterface) (THIS_ REFIID refiid, void * * ppvObject) PURE;

	MSOMETHOD_(ULONG, AddRef) (THIS) PURE;

	MSOMETHOD_(ULONG, Release) (THIS) PURE;

	// ------------------------------------------------------------------------
	// Standard Office Debug method
	MSODEBUGMETHOD

	// ------------------------------------------------------------------------
	// IMsoVsProjectHost methods

	MSOMETHOD(GetProjectItemByIndex)(THIS_ int iIndex, long *plProjItem) PURE;

	MSOMETHOD(GetProjectItemByName)(THIS_ WCHAR *wzName, long *plProjItem) PURE;

	MSOMETHOD(ShowEditor)(THIS_ long lProjItem, VSEVIEW vvView) PURE;

	MSOMETHOD(GetProjectItems)(THIS_ long *rglProjItem, _Out_cap_(1) long *pcProjItems) PURE;

	MSOMETHOD(GetProjectItemInfo)(THIS_ long lProjItem, BSTR *pbstrName, BOOL *pfEditorOpened, BSTR *pbstrData) PURE;

	MSOMETHOD(SetProjectItemInfo)(THIS_ long lProjItem, BSTR bstrData) PURE;

	MSOMETHOD(GetNeedRefresh)(THIS_ BOOL *pfProjectNeedRefresh, BOOL *pfHostNeedRefresh) PURE;

	MSOMETHOD(RefreshProject)(THIS_ BOOL fRefresh) PURE;

	MSOMETHOD(RefreshHost)(THIS_ BOOL fRefresh) PURE;

	MSOMETHOD(SaveToFile)(THIS_  long lProjItem, BSTR bstrFileName) PURE;

	MSOMETHOD(LoadFromFile)(THIS_  long lProjItem, BSTR bstrFileName) PURE;

	MSOMETHOD_(DWORD, Lock)(THIS_ BOOL fLock) PURE;

	MSOMETHOD(ChangeProjectItemIDs)(THIS_ long *rgItemsOld, long *rgItemsNew, int cItems) PURE;

	MSOMETHOD(GetVseProjectAndProjItem)(THIS_ long lProjItem, IMsoVsProject **ppVseProject, long *pVseProjItem) PURE;

	MSOMETHOD(RebuildProxy)(THIS_ int fUpdateVse, int fUseFileName) PURE;

	MSOMETHOD(StartHTMLSave)(THIS_ const WCHAR *wzFileName, WCHAR *wzTempFileName, int cchTempFileName, IMsoHTMLFileNameTable **pphfnt, interface IMsoOLDocument **ppioldoc) PURE;
	MSOMETHOD(EndHTMLSave)(THIS_ IMsoHTMLFileNameTable *phfnt, HRESULT hr) PURE;

	MSOMETHOD(SetSiteForegroundWindow)(THIS_ HWND hwnd) PURE;

	MSOMETHOD(GetContainerDocument)(THIS_ IDispatch **ppDisp) PURE;
};


#undef  INTERFACE
#define INTERFACE  IMsoVsProjectProxySite
DECLARE_INTERFACE_(IMsoVsProjectProxySite, IUnknown)
{
	// ------------------------------------------------------------------------
	// Begin Interface
	//
	// This is important on the Mac.
	BEGIN_MSOINTERFACE

	// ------------------------------------------------------------------------
	// IUnknown Methods

	MSOMETHOD(QueryInterface) (THIS_ REFIID refiid, void * * ppvObject) PURE;

	MSOMETHOD_(ULONG, AddRef) (THIS) PURE;

	MSOMETHOD_(ULONG, Release) (THIS) PURE;

	// ------------------------------------------------------------------------
	// IMsoVsProjectProxySite methods

	MSOMETHOD(GetContainerDocument)(THIS_ IDispatch **ppDisp) PURE;
};


/*---------------------------------------------------------------------------


			IMsoScriptsUser

----------------------------------------------------------------------------*/

#define	msopmCmdInsertScipt		1
#define	msopmCmdViewScript		2
#define	msopmCmdDeleteScript	3

#undef  INTERFACE
#define INTERFACE  IMsoScriptsUser
DECLARE_INTERFACE(IMsoScriptsUser)
{
	MSOMETHOD_(BOOL, FGetDesignMode) (THIS) PURE;
	MSOMETHOD_(BOOL, FDocScriptsAdd) (THIS_ void *pvDoc, MSOHSP hsp, BOOL fOffice) PURE;
	MSOMETHOD_(IMsoScripts *, PiScripts) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(void, DocRemoveScripts) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(BOOL, FInsertScript) (THIS_ void *pvDoc, MsoScriptLanguage Lang, MSOHSP *phsp, IDispatch *pdispAnchor) PURE;
	MSOMETHOD_(BOOL, FHandleScriptDelete) (THIS_ MSOHSP hsp, void *pvDoc) PURE;
	MSOMETHOD_(BOOL, FBeginUndo) (THIS_ IDispatch *pidisp, IDispatch *pidispParent, void **ppvCtxt) PURE;
	MSOMETHOD_(void, EndUndo) (THIS_ void *pvCtxt) PURE;
	MSOMETHOD_(long, GetCpFromDocHsp) (THIS_ void *pvSubDoc, MSOHSP hsp) PURE;
	MSOMETHOD_(HRESULT, HrChangeSelection) (THIS_ IDispatch **ppdispNew, MSOSCRIPTCA *pcaNew) PURE;
	MSOMETHOD_(BOOL, FRestoreSelection) (THIS_ IDispatch *pdispNew) PURE;
	MSOMETHOD_(HRESULT, HrGetInlineShape) (THIS_ void *pvDoc, MSOHSP hsp, IDispatch **ppdisp) PURE;
	MSOMETHOD_(void *, PvRegRange) (THIS_ void *pvDRS, MSOSCRIPTCA *pca) PURE;
	MSOMETHOD_(void, UnregRange) (THIS_ void *pcaw) PURE;
	MSOMETHOD_(void, UpdateCa) (THIS_ MSOSCRIPTCA *pca, IDispatch *pidispParent) PURE;
	MSOMETHOD_(int, FShow) (THIS_ void *pvDoc, MSOHSP hsp) PURE;
	MSOMETHOD_(void, SaveCaBe) (THIS_ void *pca, int bt) PURE;
	MSOMETHOD_(BOOL, FCurDoc) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(BOOL, FAnchorVisible) (THIS) PURE;
	MSOMETHOD_(BOOL, FCommandEnabled) (THIS_ int iCommand, void *pvDgvs) PURE;
	MSOMETHOD_(void, LoadDoc) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(HRESULT, CreateScriptingProjects) (THIS) PURE;
	MSOMETHOD_(void, SetDocDefScriptLang) (THIS_ MSOHICD *phicd, MsoScriptLanguage DefLang,
			const WCHAR *wzDefLang) PURE;
	MSOMETHOD_(BOOL, FCanStoreData) (THIS_ void *pvClient) PURE;
	MSOMETHOD_(void, SetDocDirty) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(IMsoDrawing *, GetDg) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(BOOL, FDocHeadScriptsAdd) (THIS_ MSOSCRIPTSTD *pstd,
		const WCHAR *wzScript, void *pvDoc, void **ppvScript, BOOL fHtmlImport) PURE;
	MSOMETHOD_(void, ShowScripts) (THIS_ void *pvClient) PURE;
};

MSOAPI_(void) MsoInitScripts(IMsoScriptsUser *pScriptsUsr);
MSOAPI_(BOOL) MsoFCloneScripts(IMsoDrawing *pidg);
MSOAPIX_(void) MsoDeleteAllAnchors(IMsoDrawing *pidg);
MSOAPI_(BOOL) MsoFGetCScriptsVisible(BOOL fGetFromReg);
MSOAPI_(BOOL) MsoFSetCScriptsVisible(BOOL fVis);

// Digital Signatures

#undef  INTERFACE
#define INTERFACE  IMsoDigSigUser
DECLARE_INTERFACE(IMsoDigSigUser)
{
	MSOMETHOD_(void, SetProjDirty) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(void, GetDigSig) (THIS_ void *pvDigSigStore, PDIGSIGBLOB *ppDigSigBlob) PURE;
	MSOMETHOD_(BOOL, FSetDigSig) (THIS_ void *pvDigSigStore, void *pvDigSigBlobNew) PURE;
	MSOMETHOD_(void, FreeDigSig) (THIS_ void *pvDoc, BOOL fDirty) PURE;
	MSOMETHOD_(void, TrustInstalledFiles) (THIS_ BOOL fSet, DWORD *pdwValue) PURE;
	MSOMETHOD_(void, AlertWz) (THIS_ WCHAR *wzAlert) PURE;
	MSOMETHOD_(BOOL, FBootedEmbedded) (THIS) PURE;
	MSOMETHOD_(HMSOINST, HGetAppMsoInst) (THIS) PURE;
	MSOMETHOD_(BOOL, FAppUserControl) (THIS) PURE;
};

#undef  INTERFACE
#define INTERFACE  IMsoDigSigUserEx
DECLARE_INTERFACE_(IMsoDigSigUserEx, IMsoDigSigUser)
{
	// IMsoDigSigUser methods
	MSOMETHOD_(void, SetProjDirty) (THIS_ void *pvDoc) PURE;
	MSOMETHOD_(void, GetDigSig) (THIS_ void *pvDigSigStore, PDIGSIGBLOB *ppDigSigBlob) PURE;
	MSOMETHOD_(BOOL, FSetDigSig) (THIS_ void *pvDigSigStore, void *pvDigSigBlobNew) PURE;
	MSOMETHOD_(void, FreeDigSig) (THIS_ void *pvDoc, BOOL fDirty) PURE;
	MSOMETHOD_(void, TrustInstalledFiles) (THIS_ BOOL fSet, DWORD *pdwValue) PURE;
	MSOMETHOD_(void, AlertWz) (THIS_ WCHAR *wzAlert) PURE;
	MSOMETHOD_(BOOL, FBootedEmbedded) (THIS) PURE;
	MSOMETHOD_(HMSOINST, HGetAppMsoInst) (THIS) PURE;
	MSOMETHOD_(BOOL, FAppUserControl) (THIS) PURE;

	// IMsoDigSigUserEx methods
	MSOMETHOD_(BOOL, FGetHashOfDoc) (THIS_ void *pvDoc, BSTR *pbstrHash) PURE;
};

MSOAPI_(void) MsoGetDigSig(void *pvDigSigStore, PDIGSIGBLOB *ppDigSigBlob, BOOL fVba);
MSOAPI_(BOOL) MsoFFreeAndSetDigSig(void *pvDigSigStore, void *pvDigSigBlobNew,
		void *pvDoc, BOOL fVba);
MSOAPI_(void) MsoInitDigSig(IMsoDigSigUser *pDigSigUsr, int nAppID);
MSOAPI_(unsigned int) MsoVBADigSigCallDlg (void *pvDigSigStore, HMSOINST hinst, HWND hwnd,
		void *pvDoc, BOOL fVba);
MSOAPI_(unsigned int) MsoVBADigSigCallDlgEx (void *pvDigSigStore, HMSOINST hinst, HWND hwnd,
		void *pvDoc, BOOL fVba, const WCHAR *wzHelpFile, DWORD dwHelpId);
MSOAPI_(BOOL) MsoFVBADigSigSignProject (void *pvDigSigStore, HWND hwnd, void *pvVBProj,
		BOOL fStg, BOOL fVba, void *pvDoc);
MSOAPI_(BOOL) MsoFVBADigSigSignProjectEx (void *pvDigSigStore, HWND hwnd, void *pvVBProj,
		BOOL fStg, BOOL fVba, void *pvDoc, HRESULT *phrRetVal);

#define msoDigSigOffice 0
#define msoDigSigVBA 1
#define msoDigSigAddin 2	// COM addins
#define msoDigSigUntrustedAddin 3	// XLLs and WLLs being opened by File Open so they aren't trusted addins
#define msoDigSigXML 4	// XML files

MSOAPI_(int) MsoVBADigSigVerifyProject(void *pvDigSigStore, HMSOINST hmsoinst, HWND hwnd,
	void *pvMacro, BOOL fAppIsActive, int *pmsodsv, int iClient, int iSecurityLevel,
	const WCHAR *wzFileName, const WCHAR *wzHelpFile, DWORD dwHelpId, HANDLE hFileDLL,
	BOOL fUserControl, BOOL *pfLoadFromText);
MSOAPI_(BOOL) MsoFRevokedSignature (void *pvDigSigStore, int iClient, BOOL fShowAlert);
MSOAPI_(void) MsoFreeDigSig(void *pvDigSig, void *pvDoc, BOOL fDirty);
MSOAPI_(BOOL) MsoFCloneDigSig(void *pvDigSigOld, void *pvDigSigNew);
MSOAPI_(HRESULT) MsoHrSignDoc(const WCHAR *wzCert, void *pvDigSigStore, void *pvDoc);
//Get and set App.AutomationSecurity for VBE
MSOAPI_(MsoAutomationSecurity) MsoAppGetAutomationSecurity(void);
MSOAPI_(void) MsoAppSetAutomationSecurity(MsoAutomationSecurity AutomationSecurity);
MSOAPIDBG_(void) MsoDigSigSaveBe(void *pvDigSig);
MSOAPI_(void) MsoAppSetLMTrustVBAccess(DWORD dwTrustVBAccess);
MSOAPI_(void) MsoShowSignError (HRESULT hr);

#endif // VBE

// Wrapping for VB ISVs
#undef  INTERFACE
#define INTERFACE  IMsoVbaSecurity
DECLARE_INTERFACE(IMsoVbaSecurity)
{
	/*
	VerifyProject is called while opening a signed file. It checks the signature and
	puts up the signature verify dialog.

	in params:
	pStg - ptr to storage for VBA project
	pvSign - ptr to signature blob in doc
	iSecurityLevel - security level app is currently running under
	hmsoinst - VB has an msoinst - they'll have to give us this for the SDM dlg
	so in the final interface VB exposes to ISVs they'll have to remove this param
	fAppIsActive - not sure we need this param - we use it in Office to figure
	out whether to put up the SDM dlg or the regular Windows dlg (if the app is
	launching in the browser, the SDM dlg won't work)
	hwndParent - hwnd of parent of dlg we put up
	wzFileName - name of the file being opened - we display this in the dialog

	out params:
	pmsodsv - can take the following values (defined in msoav.h)
	#define	msodsvNoMacros	0
	#define msodsvUnsigned		1
	#define	msodsvPassedTrusted	2
	#define	msodsvPassedNoCert	3
	#define	msodsvCodeChanged	4
	#define	msodsvFailed		5
	#define	msodsvExpiredRevoked	6

	return values (defined in msoav.h):
	#define msoedmEnable	1
	#define	msoedmDisable	2
	#define	msoedmDontOpen	3
	depending on button user clicked in verify dlg
	*/
	MSOMETHOD_(int, VerifyProject) (THIS_ IStorage *pStg, void *pvSign, int iSecurityLevel,
		HMSOINST hmsoinst, BOOL fAppIsActive, int *pmsodsv, HWND hwndParent,
		const WCHAR *wzFileName, const WCHAR *wzHelpFile, DWORD dwHelpId) PURE;

	/*
	DlgSignProject is called when the Digital Signatures menu item is selected in VBE
	i.e. the user wants to pick a certificate to sign the project. The Sign dialog is
	shown. If the user selects a cert, IMsoVbaSecurityUser::FPutSign is called to put the
	cert info into the signature (signature is not complete yet because the actual
	project hash needs to be added - this will be done at doc save time).

	in params:
	hmsoinst - VB has an msoinst - they'll have to give us this for the SDM dlg
	so in the final interface VB exposes to ISVs they'll have to remove this param
	pvSign: ptr to the signature if doc is already signed, NULL if it is currently unsigned.
	dwDocId - some data that uniquely identifies a document for the client.
	hwndParent - hwnd of parent of dlg we put up
	*/
	MSOMETHOD_(void, DlgSignProject) (THIS_ void *pvSign, HMSOINST hmsoinst,
		DWORD dwDocId, HWND hwndParent, const WCHAR *wzHelpFile, DWORD dwHelpId) PURE;

	/*
	SignProject is called when a signed document with a VB project is being saved. It
	must be called after the VB project has been saved and IVbaProject::SaveCompleted has
	been called. SignProject actually computes the signature and calls
	IMsoVbaSecurityUser::FPutSign to save the signature.

	In params:
	pStg - ptr to storage for VBA project
	pvSign - ptr to signature blob currently in doc
	dwDocId - some data that uniquely identifies a document for the client.
	*/
	MSOMETHOD_(BOOL, SignProject) (THIS_ IStorage *pStg, void *pvSign, DWORD dwDocId) PURE;

	/*
	DlgSetSecurityLevel is called to show the Set Security level dialog. The
	client is responsible for keeping track of the current security level set in
	the app (Office does this through registry keys). The dlg also manages trust list
	of certificates (don't need any params for this - all code in Office) and
	a checkbox for whether installed files should be trusted or not (Office apps
	have the concept of trusted directories and if the box is checked we don't
	check files from those directories for signatures).

	in params:
	hmsoinst - VB has an msoinst - they'll have to give us this for the SDM dlg
	so in the final interface VB exposes to ISVs they'll have to remove this param
	hwndParent - hwnd of parent of dlg we put up

	in/out params:
	pSecurityLevel: client passes in the security level currently set in the app.
	I'll return the new security level set in the dialog. This param cannot be NULL
	because it defeats the whole purpose of the dialog otherwise.
	pfTrustInstalledFiles: client passes in following values -
	NULL if checkbox is to be hidden since client doesn't have the trusted files concept
	0 if checkbox is to be shown unchecked
	1 if checkbox is to be shown checked.
	If the checkbox is shown, I'll return the value set by user.
	*/
	MSOMETHOD_(void, DlgSetSecurityLevel) (THIS_ int *pSecurityLevel, HMSOINST hmsoinst,
		BOOL *pfTrustInstalledFiles, HWND hwndParent, BOOL fShowVirusCheckers,
		const WCHAR *wzHelpFile, DWORD dwHelpId) PURE;
};

/*
	The IMsoVbaSecurityUser interface is implemented by the VB ISV who wants to
	use the IMsoVbaSecurity interface. This interface is given to us in MsoVbaInitSecurity
	and it can't be NULL since signing won't work otherwise.
*/
#undef  INTERFACE
#define INTERFACE  IMsoVbaSecurityUser
DECLARE_INTERFACE(IMsoVbaSecurityUser)
{
	/*
	FPutSign is called by both DlgSignProject and SignProject. When it is called,
	FPutSign should free any existing signature blob it has for the doc specified
	by dwDocId, allocate a new blob of size cbSign and copy cbSign bytes over
	to the new memory from pvSign
	*/
	MSOMETHOD_(BOOL, FPutSign) (THIS_ DWORD dwDocId, void *pvSign, DWORD cbSign) PURE;
};

/*
	MsoVbaInitSecurity must be called by the VB ISV client app probably during boot
	(since they might load a signed file from the command line).

	pVbaSecUser should not be NULL - I'll return NULL for the IMsoVbaSecurity
	interface if this happens.
*/
MSOAPI_(IMsoVbaSecurity *) MsoVbaInitSecurity(IMsoVbaSecurityUser *pVbaSecUser);
MSOAPI_(void) MsoSetVBHashFuncPtr(void *pfnGetHash);

#endif //$[VSMSO]

#if defined(__cplusplus)
}
#endif

#if 0 //$[VSMSO]
#ifndef VBE

//#define  VBAHTMLDBGWRAPPERS // this is a debug code for testing the Save to Binary
typedef struct _OSTREAM {
	LPOLESTREAMVTBL 	 lpstbl;
	IStream *pistr;
} OSTREAM, FAR *LPOSTREAM;

typedef int HVBHVER;
typedef struct _HTMLVBBINHDR_V1
{
	DWORD cb;
	HVBHVER ver;
} HTMLVBBINHDR_V1;
typedef struct _HTMLVBBINHDR
{
	HTMLVBBINHDR_V1 vstruct;
	DWORD cbDigiSigData; // size of the digital signature blob

}HTMLVBBINHDR;

#define E_CBHVBH 1
#define E_VERHVBH 2
#define HVBHValidVer(p) (p!= E_VERHVBH)
#define HVBHValidCb(p) (p!= E_CBHVBH)

MSOAPI_(HRESULT)MsoHrReadHVBH(IStream *pstm,HVBHVER ver, _Out_bytecap_(cb) PVOID psaveInfo, DWORD cb,DWORD * pcbdigisigBlob);
MSOAPI_(HRESULT)MsoHrHTMLReadDigiSig(IStream *pstm,PDIGSIGBLOB *ppdigsigBlob,DWORD cbdigisigBlob);
MSOAPI_(HRESULT)MsoHrSaveHVBH(IStream *pstm,HVBHVER ver,PVOID psaveInfo, DWORD cb,PDIGSIGBLOB pdigsigBlob );
MSOAPI_(HRESULT) MsoHrCopyRgvarg(VARIANT *rgvargSrc,VARIANT *rgvargDst,int cvarg);
MSOAPIX_(void) MsoRelandClearRgvarg(VARIANT *rgvarg,int cvarg);
MSOAPI_(void) MsoClearRgvarg(VARIANT *rgvarg,int cvarg);
MSOAPIX_(void) MsoRelandClearvarg(VARIANT *pvar);
MSOAPI_(void) MsoInitRgvarg(VARIANT *rgvarg,int cvarg);

MSOAPIX_(HRESULT) MsoIStorageToIStream(IStorage *pstg,IStream *pstm);
MSOAPIX_(HRESULT)MsoIStreamToIStorage(IStream *pstm,IStorage *pstg);
MSOAPI_(HRESULT)MsoHrSaveVBAStgAsHTML(REFCLSID rAppClsid,IStorage *pDocIstg,IStream *pIstr);
MSOAPI_(HRESULT)MsoHrLoadVBAStgFromHTML(IStorage *pDocIstg,IStream *pIstr);
MSOAPIX_(HRESULT)MsoSaveVBAPrjAsHTML(REFCLSID rAppClsid,IVbaProject *pVbaProj,IStorage *pDocIstg,IStream *pIstr);
MSOAPIX_(HRESULT)MsoLoadVBAPrjFromHTML(IVbaProject *pVbaProj,IStorage *pDocIstg,IStream *pIstr);
MSOAPI_(BOOL) MsoFExportScriptAnchor(MSOHSP hsp);

MSOAPIX_(void) MsoGetLCID (LCID *plcid);
MSOAPI_(BOOL)MsoFVbpSaveProject(IVbaProject *pVbaProj);
MSOAPIX_(HRESULT) MsoHrVbpCntNonDocItems(IVbaProject *pproj, UINT *pcItems);

/*---------------------------------------------------------------------------
	IMsoAddInX

	Interface definition.
----------------------------------------------------------------- KELLYLY -*/

#undef  INTERFACE
#define INTERFACE  IMsoAddInX

DECLARE_INTERFACE(IMsoAddInX)
{
	//*** FDebugMessage method ***
	MSODEBUGMETHOD

	// IMsoAddInX methods

	/* Returns the Add-Ins 'ProgId' property. */
	MSOMETHOD_(LPCWSTR, WszGetProgId) (THIS) PURE;

	/* Returns the Add-Ins 'Description' property. */
	MSOMETHOD_(LPCWSTR, WszGetDescription) (THIS) PURE;

	/* Returns the Add-Ins 'Location' property. */
	MSOMETHOD_(LPCWSTR, WszGetLocation) (THIS) PURE;

	/* Returns the Add-Ins CLSID. */
	MSOMETHOD_(GUID *, PGetGuid) (THIS) PURE;

	/* Returns the Add-Ins 'Object' property. */
	MSOMETHOD_(IDispatch *, PGetObject) (THIS) PURE;

	/* Returns TRUE if the Add-In is connected, FALSE otherwise. */
	MSOMETHOD_(BOOL, FGetConnect) (THIS) PURE;

	/* Sets the Add-In's connect state. */
	MSOMETHOD(HrSetConnect) (THIS_ BOOL fValue) PURE;

	/* Returns TRUE if a user has requested that the connect state
	   be toggled from its current connected state. The user may
	   request for the state to be toggled through the 'Office Add-Ins'
	   dialog. */
	MSOMETHOD_(BOOL, FGetToggleConnectState) (THIS) PURE;

	/* Clears the 'ToggleConnectState' flag. */
	MSOMETHOD_(void, ClearToggleConnectState) (THIS) PURE;

	/* Returns the dual-interface that represents the Add-In. */
	MSOMETHOD(HrGetDualInterface) (THIS_ IDispatch **ppDispatch) PURE;

	/* Locks the addin for Unload from Object Model */
	MSOMETHOD(HrLock) (THIS_ BOOL *pfLock) PURE;

};


/*---------------------------------------------------------------------------
	IMsoAddInsX

	Interface definition.
----------------------------------------------------------------- KELLYLY -*/

/*- Data types used by IMsoAddInsX methods. -*/

///////////////////////////////////////
// PFNADDINSXFOREACH
//
// Type definition for callback function supplied in call to
// IMsoAddInsX::HrForEach.
//
typedef BOOL (CALLBACK *PFNADDINSXFOREACH)(IMsoAddInX *pMsoAddInX, LONG lParam);

///////////////////////////////////////
// PFNGETAPPDISPOBJECT
//
// Type definition for callback function optionally supplied
// in a call to MsoHrCreateAddInsXObject.
//
typedef IDispatch* (CALLBACK *PFNGETAPPDISPOBJECT)(void);


///////////////////////////////////////
// IMsoAddInsX
//
#undef  INTERFACE
#define INTERFACE  IMsoAddInsX

DECLARE_INTERFACE(IMsoAddInsX)
{
	//*** FDebugMessage method ***
	MSODEBUGMETHOD

	// IMsoAddInsX methods

	/* Updates the MsoAddInsX collection from the registry.
	   An update is implicitly done when the MsoAddInsX object
	   is created. */
	MSOMETHOD(HrUpdate) (THIS) PURE;

	/* Returns the count of MsoAddInX objects in the MsoAddInsX
	 collection. */
	MSOMETHOD_(LONG, LCount) (THIS) PURE;

	/* Displays the MsoAddInsX dialog allowing the user to
	   connect, disconnect, add, or remove ActiveX AddIns. */
	MSOMETHOD(HrDisplayDialog) (THIS_ HWND hwndOwner) PURE;

	/* Calls a supplied function for each MsoAddInX in the
	   MsoAddInsX collection. */
	MSOMETHOD(HrForEach) (THIS_ PFNADDINSXFOREACH pfn, LONG lParam) PURE;

	/* Returns a IEnumVARIANT interface that may be used to
	   enumerate all MsoAddInX disp objects in the MsoAddInsX
	   collection. */
	MSOMETHOD(HrGetAddInsXEnum) (THIS_ IEnumVARIANT **ppenum) PURE;

	/* Clears the collection by first disconnecting any connected
	   ActiveX AddIn objects and then removing the MsoAddIns from the
	   collection */
	MSOMETHOD(HrClear) (THIS) PURE;

	/* Returns an MsoAddInX object specified by index. */
	MSOMETHOD(HrGetItem) (THIS_ long lIndex, IMsoAddInX **ppAddIn) PURE;

	/* Returns the dual-interface of the MsoAddInsX object. */
	MSOMETHOD(HrGetDualInterface) (THIS_ IDispatch **ppDispatch) PURE;

	/* Fires the OnAddInsUpdate method of all connected ActiveX
	   AddIns. */
	MSOMETHOD_(void, OnAddInsUpdate) (THIS_ SAFEARRAY *psaCustom) PURE;

	/* Fires the OnStartupComplete method of all connected ActiveX
	   AddIns. */
	MSOMETHOD_(void, OnStartupComplete) (THIS_ SAFEARRAY *psaCustom) PURE;

	/* Deletes the MsoAddInsX object. MsoAddInsX does not use reference
	   counting to control its lifetime. The creation of the MsoAddInsX
	   object is done by MsoHrCreateAddInsXObject, to destroy the
	   object call Delete off of the pointer to the MsoAddInsX returned
	   by MsoHrCreateAddInsXObject. */
	MSOMETHOD_(void, Delete) (THIS) PURE;

	/* Split the code to create the addinsx, and then Do the boot connections */
	MSOMETHOD(HrDoBootConnections) (THIS) PURE;

	MSOMETHOD(HrUnInit) (THIS) PURE;
	MSOMETHOD(HrSetFireEvents) (THIS_ BOOL fFireEvents) PURE;

};

/*---------------------------------------------------------------------------
	Macros simplifying calling methods off of IMsoAddInsX and IMsoAddInX
	easier for C programmers.
----------------------------------------------------------------- KELLYLY -*/

/*- IMsoAddInX -*/
#define IMsoAddInX_WszGetProgId(This) \
	This->lpVtbl->WszGetProgId(This)
#define IMsoAddInX_WszGetDescription(This) \
	This->lpVtbl->WszGetDescription(This)
#define IMsoAddInX_WszGetLocation(This) \
	This->lpVtbl->WszGetLocation(This)
#define IMsoAddInX_PGetGuid(This) \
	This->lpVtbl->PGetGuid(This)
#define IMsoAddInX_PGetDispObject(This) \
	This->lpVtbl->PGetDispObject(This)
#define IMsoAddInX_FGetConnect(This) \
	This->lpVtbl->FGetConnect(This)
#define IMsoAddInX_HrSetConnect(This, fValue) \
	This->lpVtbl->HrSetConnect(This, fValue)
#define IMsoAddInX_FGetToggleConnectState(This) \
	This->lpVtbl->FGetToggleConnectState(This)
#define IMsoAddInX_ClearToggleConnectState(This) \
	This->lpVtbl->ClearToggleConnectState(This)
#define IMsoAddInX_HrGetDualInterface(This, ppDispatch) \
	This->lpVtbl->HrGetDualInterface(This, ppDispatch)

/*- IMsoAddInsX -*/
#define IMsoAddInsX_HrUpdate(This) \
	This->lpVtbl->HrUpdate(This)
#define IMsoAddInsX_LCount(This) \
	This->lpVtbl->LCount(This)
#define IMsoAddInsX_HrDisplayDialog(This, hwndOwner) \
	This->lpVtbl->HrDisplayDialog(This, hwndOwner)
#define IMsoAddInsX_HrForEach(This, pfn, lParam) \
	This->lpVtbl->HrForEach(This, pfn, lParam)
#define IMsoAddInsX_HrGetAddInsXEnum(This, ppenum) \
	This->lpVtbl->HrGetAddInsXEnum(This, ppenum)
#define IMsoAddInsX_HrClear(This) \
	This->lpVtbl->HrClear(This)
#define IMsoAddInsX_HrGetItem(This, lIndex, ppAddIn) \
	This->lpVtbl->HrGetItem(This, lIndex, ppAddIn)
#define IMsoAddInsX_HrGetDualInterface(This, ppDispatch) \
	This->lpVtbl->HrGetDualInterface(This, ppDispatch)
#define IMsoAddInsX_OnAddInsUpdate(This, psaCustom) \
	This->lpVtbl->OnAddInsUpdate(This, psaCustom)
#define IMsoAddInsX_OnStartupComplete(This, psaCustom) \
	This->lpVtbl->OnStartupComplete(This, psaCustom)
#define IMsoAddInsX_Delete(This) \
	This->lpVtbl->Delete(This)
#define IMsoAddInsX_HrDoBootConnections(This) \
	This->lpVtbl->HrDoBootConnections(This)
#define IMsoAddInsX_HrUnInit(This) \
	This->lpVtbl->HrUnInit(This)
#define IMsoAddInsX_HrSetFireEvents(This, fEvents)\
	This->lpVtbl->HrSetFireEvents(This, fEvents)


// The application is recognized by these rids.
struct msoAddinXIDs {
	int	msoridLM;
	int	msoridCU;
	int	nAppBitNumber;
};

typedef struct msoAddinXIDs* LPADDINRIDS;

/*---------------------------------------------------------------------------
	MsoHrCreateAddInsXObject

	Creates an IMsoAddInsX object and returns a reference to it through
	the out-parameter ppAddInsX. pAppRids identifies the host application.

	When the host's application's IDispatch object is needed, the addins
	object will call the host to obtain the IDispatch object.  How does it
	call the app?  First the addins object will check to see if the app
	registered a msocbAddinGetIDispatch callback (see MsoPfnSetCallback in
	msouser.h for details).  If there is no msocbAddinGetIDispatch
	callback, then pfnGetAppDispObject will be called.  If pfnGetAppDispObject
	is NULL, then the application's IMsoUser::FGetIDispatchApp function will
	be called instead.

	This function should be called during application boot, following
	object model construction and preceding object event firing.
----------------------------------------------------------------- brianhi -*/

MSOAPI_(HRESULT) MsoHrCreateAddInsXObject (LPADDINRIDS pAppRids, HMSOINST hmsoinst,
	PFNGETAPPDISPOBJECT pfnGetAppDispObject, IMsoAddInsX **ppAddInsX);

// The iActivateMethod is assigned one of these 3 values.
#define msoAddinLaunchedUI		1
#define	msoAddinLaunchedOle		2
#define	msoAddinLaunchedAuto	3

MSOAPI_(HRESULT) MsoHrCreateAddInsXObjectEx(LPADDINRIDS pAppRids,
	HMSOINST hmsoinst, PFNGETAPPDISPOBJECT pfnGetAppDispObject,
	IMsoAddInsX **ppMsoAddInsX, int iActivateMethod);

MSOAPI_(BOOL) MsoFValidAddinsExist(LPADDINRIDS pAppRids);
MSOAPI_(HRESULT) MsoHrGetVseHTMLProjObj (HMSOINST hmsoinst, IMsoVsProjectProxy *pproxy, IDispatch **ppdisp);
MSOAPI_(BOOL) MsoFInMsoFireEvent ();
MSOAPI_(void) MsoToggleTBSUpdEvent(BOOL fEnable);

typedef void(*fnRemoveTag)(LPCWSTR szTag, VOID* pParameter);

MSOAPI_(void) MsoEnumRemoveComAddinTags(
	int msoridAddinsDir,
	fnRemoveTag fnRemoveCommandBarButton,
	fnRemoveTag fnRemoveCommandBar,
	void* pParameter);

#endif // VBE

// Should use MsoFPrivacyMacros which has a Cancel button, rather than MsoPrivacyMacros, which doesn't.
MSOAPI_(void) MsoPrivacyMacros();
MSOAPI_(BOOL) MsoFPrivacyMacros();

enum
	{
	imsoAppIdUnknown = 0,
	imsoAppIdWord,
	imsoAppIdExcel,
	imsoAppIdPowerPoint,
	imsoAppIdOutlook,
	imsoAppIdAccess,
	imsoAppIdPublisher,
	imsoAppIdFrontPage,
	imsoAppIdProject,
	};

MSOAPI_(int) MsoAppIdFromInst(HMSOINST pinst);


#endif //$[VSMSO]

#endif // MSOPM_H

