//+---------------------------------------------------------------------------
//
//  Microsoft Windows
//  Copyright (C) Microsoft Corporation, 1994 - 1997.
//
//  File:       nlcheck_.h
//
//  Contents:   Office private extensions to NLCheck API
//
//  History:
//
//----------------------------------------------------------------------------
#ifndef _FACTOID_H_
#define _FACTOID_H_

//+---------------------------------------------------------------------------
// Dependencies
//----------------------------------------------------------------------------
#include <objbase.h>
#include <ocidl.h>

//+---------------------------------------------------------------------------
// Basic Definitions
//----------------------------------------------------------------------------

//TLB version numbers
#define wVerMstagTlbMajor		1
#define wVerMstagTlbMinor		2

//Binary file format version numbers
#define wVerMajorCur			1
#define wVerMinorCur			0

			
typedef unsigned int FECODE;	// Factoid Error Code
#define fecNone                0   // Success - no error code
#define fecErrors              1   // one or more errors; see CRB
#define fecUnknown             2   // Unknown error
#define fecOOM                 3   // out of memory; skip sentence
#define fecInvalidParameter    4   // invalid parameter passed to function
#define fecInvalidId           5   // invalid nl id
#define fecEndOfDoc            6   // End of document reached on fetch
#define fecEndOfPara           7   // End of paragraph reached
#define fecStopProcessing      8   // Stop processing text and return to caller
#define fecBeginOfPara         9   // Begin of para reached on Fetch
#define fecBeginOfDoc          10  // Begin of document reached on fetch
#define fecLimitHit            11  // Don't need to fetch more end of specified range
#define fecInvalidRange        12  // The supplied CA range is invalid
#define fecInvalidVersion      13  // The supplied version is not supported
#define fecIOError             14  // General IO Error has occurred
#define fecNotImplemented      15  // Function not implemented
#define fecServerNotInited	   16
#define fecUntrustComponent	   17
#define fecRecognizerNotInited 18
#define fecNoRecognizer		   19
#define fecStopBatchUpdate     20  // Stop batch update, presumably because a runaction altered the doc.

#if defined(_M_IX86)
#define CFAPI          FECODE __declspec(dllexport) WINAPI
#else
#define CFAPI          FECODE WINAPI
#endif

typedef HANDLE  		FDOC;		// Factoid Doc
typedef void 		   *FACID;      // FActoid Checker Id type
typedef unsigned int 	FCC;    	// Factoid Check Type
typedef HANDLE 			FCH;		// Factoid Check Handle
typedef long   			FCP;		// Factoid CP

#ifndef XCHAR
typedef unsigned short XCHAR;
#endif

#pragma pack(push, 4)

typedef struct _FCA
	{
    FDOC doc;
    FCP  cpFirst;
    FCP  cpLim;
	} FCA, * PFCA;

#pragma pack(pop)

#define facidNil   ((FACID)0)
#define fchNil     ((FCH)0)
#define fdocNil    ((FDOC)0)

typedef interface IFactoidPropertyBag IFactoidPropertyBag;
typedef interface IFactoidPropertyBagStore IFactoidPropertyBagStore;
typedef interface IFactoidServer IFactoidServer;
typedef interface ISmartTagTokenList ISmartTagTokenList;
typedef interface ISmartTagProperties ISmartTagProperties;
typedef interface ISmartDocProperties ISmartDocProperties;
typedef interface ISmartTagRecognizerSite ISmartTagRecognizerSite;

#pragma pack(push, 4)

typedef struct _fvi             // filled in by CheckVersion
{
    DWORD   dwVersion;                  // api version
    DWORD   dwEngineVersion;            // engine version. HIWORD(major) LOWORD(minor)
    DWORD   fcopt;                       // bitfield for module functionality
} FVI;

#pragma pack(pop)

#define cchMaxFactoidHtmlPrefix	8

//
// 	Check Factoid Initialization structure
//
#define fInitRecognizerThread				0x00000001
#define fWaitRecogSvrInit					0x00000002
#define fSecurityCheck				0x00000004
#define fSmartTagServer				0x00000008
#define fInitRecognizerFRTE				0x00000010

#define wFactoidMsgDeleteKeyVal		0x1	//l1 = IFactoidPropertyBag; l2 = index
#define wFactoidMsgGetPrefix		0x2 //l1 = XCHAR *szUri, l2 = XCHAR szPrefix[cchMaxHtmlPrefix]
#define	wFactoidMsgValidPrefix		0x3 //l1 = XCHAR *szPrefix
//#define wFactoidMsgSecurityCheck	0x4 //unused - do not reuse code.
#define wFactoidMsgFilterPrivacy	0x5	//l1 = BOOL *pfFilterPrivacy -> get rid of download URLs for currently saving pbagstore

#define MAX_FACTOID_SERVER_NONE				-1

typedef FECODE (WINAPI *PFNFACTOIDUSERNOTIFY)(int wFactoidMsg, long l1, long l2);

#pragma pack(push, 4)

typedef struct _cfinit
	{
  	DWORD 	dwVersion;      // use dwMAKEVERSION(checkVerMajor, checkVerMinor)
  	LCID 	lcid;
  	int		grf;
	int		idApp;			// obsolete; must be 0.  szAppWndClass is now used
	int		imaxSrvr;
	char*	szSrvrReg;
	PFNFACTOIDUSERNOTIFY pfnNotify;
	char	szAppWndClass[255];
	XCHAR	xszAppName[255];
	}CFINIT;

#pragma pack(pop)

// Alert: this list needs to be consistent with IF_TYPE defined in mstag.idl
#define	ftChar 			0x00000001	// Currently not in use
#define	ftSingleWord 	0x00000002	// Currently not in use
#define	ftRegExp 		0x00000004	// Currently not in use
#define	ftPara			0x00000008	// Currently used by Word only
#define	ftCell			0x00000010	// Currently used by Excel only
#define ftRun			0x00000020	// used by IE stbho

//
// NLcheckk db query Server Information ( a stripped down version of IFSIB )
//
#define cchMaxSrvrName	256  
#define cchMaxSrvrPath	512
#define cchMaxSrvrDesr	512
#define cchMaxSrvrProgId 256
#define cchMaxSrvrSolution 256

#pragma pack(push, 4)

typedef struct _NLSI
	{
	int		idSrvr;
	int 	fEnabled;	
	int		fHasPropDlg;
	XCHAR	xszName[cchMaxSrvrName];
	XCHAR 	xszDscr[cchMaxSrvrDesr];
	XCHAR	xszPath[cchMaxSrvrPath];
	XCHAR	xszProgId[cchMaxSrvrProgId];
	XCHAR	xszSolution[cchMaxSrvrSolution];
	}NLSI;
#define cbNLSI	(sizeof(NLSI))

typedef struct _FRTE
	{
	ISmartTagRecognizerSite* pRecogSite;
	IFactoidPropertyBag* pfpb;
	const XCHAR* xszUri;
	const XCHAR* xszTag;
	int ich;
	int cch;
	int iTag;
	} FRTE;

#pragma pack(pop)

//+---------------------------------------------------------------------------
// Magic number that defines a factoid dll
//----------------------------------------------------------------------------
#define FCHECK_FACTOID_MAGICNUM	0x33093544

//+---------------------------------------------------------------------------
// Extended ctcc codes
//----------------------------------------------------------------------------
#define fccCheckFactoid        0x00001000
#define fccCheckBackground     0x20000000  // perform the checking in the

//+---------------------------------------------------------------------------
// Extended ido codes
//----------------------------------------------------------------------------
#define idoGetFactoidServer             1000
#define idoCheckInitFactoid				1001
#define idoCheckTextFactoid       		1002
#define idoCheckBatchUpdateFactoid      1003
#define idoCheckGetDBServerInfo         1004
#define idoCheckSetDBServerOption		1005

//
// Callback procedure used by NLBatchUpdate()
//
// NLEC CALLBACK NLBatchUpdateProc(NLCA *pnlca,
//								int		idSrvr,
//								FACT	fact,
//								int		ichFactoid,
//								int		cchFactoid,
//								XCHAR	rgxchTag[],
//								int		cchTag)
//
typedef FECODE (WINAPI *PFNBATCHUPDATEPROC)(FCA *pnlca,
								IFactoidPropertyBag *pifpb,
								FCP		ichFactoid,
								FCP		cchFactoid,
								const WCHAR	*rgxchRaw,
								int		cchRaw,
								void 	*pvData);

typedef FECODE (WINAPI *PFNADDHANDLEPROC)(FCA*, FCH, PVOID);

//
//	Handle status code
//
#define hscUnknown				0
#define hscPending				1
#define hscPartialCalced		2
#define hscFinished				3
#define hscError                4

//
//  Known Office Client Id
//
//  These are obsolete.
#define FTD_APPID_SPECCRASH		1
#define FTD_APPID_TESTUTIL		2
#define FTD_APPID_WORD			3
#define FTD_APPID_EXCEL			4
#define FTD_APPID_TRIDENT		5
#define FTD_APPID_IEBROWSER		6
#define FTD_APPID_MIN			FTD_APPID_TESTUTIL
#define FTD_APPID_MAX			FTD_APPID_IEBROWSER

// smart tag status types.
#define stsEnabled		0
#define stsDisabled		1
#define stsCrashed		2

#define fFactoidServerTurnOff		0x00000001
#define fFactoidServerTurnOn		0x00000002

// allow states for items in STST structure
#define asUndefined		-1
#define asDisabled		0
#define asEnabled		1
#define asDontCare		2 

// these MUST MAP to the C_TYPE enumeration in mstag.idl!!
// we call them ACTYPE's (Action Control Types) because CTYPE is defined in access.
typedef enum {
	actypeSmartTag = 1,
	actypeLink = 2,
	actypeHelp = 3,
	actypeHelpURL = 4,
	actypeSeparator = 5,
	actypeButton = 6,
	actypeLabel = 7,
	actypeImage = 8,
	actypeCheckbox = 9,
	actypeTextbox = 10,
	actypeListbox = 11,
	actypeCombo = 12,
	actypeActiveX = 13,
	actypeDocumentFragment = 14,
	actypeDocumentFragmentURL = 15,
	actypeRadioGroup = 16,
	} ACTYPE;

#pragma pack(push, 4)

typedef struct _STST
{
	DWORD asLabel;
	DWORD asIndicator;
	DWORD asButton;
	DWORD asSave;
} STST;

typedef struct _STAP
{
	char szWndClass[255];
	char szFriendlyName[255];
	STST ststAbilities;
	STST ststSettings;
	NLSI* prgnlsi;
	int cnlsi;
} STAP;

// smart document list/combo box data
typedef struct _sdlcbd
{
	SAFEARRAY *psabList;
	int iSelected;
	BOOL fEditable;
} SDLCBD;

// smart document help data
typedef struct _sdhd
{
	union
		{
		XCHAR *stzHelpContent;
		XCHAR *stzHelpUrl;
		};
} SDHD;

// smart document action control data
typedef struct _sdacd
{
	ACTYPE actype;	// the control type.
	ISmartDocProperties* psdp;		// the property bag.

	// this union will contain all the different types of data which can be found.
	// different types of data are valid for different controls.
	union
		{
		SDHD sdhd;
		BOOL fChecked;
		XCHAR *stzTextboxContent;
		XCHAR *stzImageSrc;		
		SDLCBD sdlcbd;
		XCHAR *stzDocumentFragmentContent;		
		XCHAR *stzDocumentFragmentUrl;
		ISmartDocProperties *paxp;		
		};
} SDACD;

typedef struct _actioncaption
{
	XCHAR 	*stzCaption;
	XCHAR 	*stzName;
	void	*pvActionid;
	SDACD	sdacd;
} ACTIONCAPTION;

typedef struct _keyval
{
	XCHAR *rgchKey;
	int cchKey;
	XCHAR *rgchVal;
	int cchVal;
} KEYVAL;

#pragma pack(pop)

// smart tag lifespan types
typedef enum 
	{
	fperNormal,
	fperTransitory,
	fperTemporary,
	fperMax
	} FPER;	

// definitions of expiring tag value
#define ichFacYear					0
#define cchFacYear					4
#define ichFacMonth					5
#define cchFacMonth					2
#define ichFacDay					8
#define cchFacDay					2
#define ichFacHour					11
#define cchFacHour					2
#define ichFacMin					14
#define cchFacMin					2
#define cchFacExpireKeyValFull		16				// length of yyyy:MM:dd:hh:mm
#define cchFacExpireKeyValShort		10				// length of yyyy:MM:dd

// smart tag display types
#define ffstButton					0x0000001
#define ffstIndicator				0x0000002
#define grffstNone					0x0000000
#define grffstAll					0xfffffff

#define SMTAGID_ALL_TAGS			-1
#define IDISP_DONTCARE				-1

#define SZ_SmartTagComponentUpdateMessage "SmartTagInstall"

DEFINE_GUID(IID_IFactoidPropertyBagStore, 0xc4c3e75c, 0xb73d, 0x11d3, 0xb2, 0xcf, 0x0, 0x50, 0x4, 0x89, 0xd6, 0xa3);
DEFINE_GUID(IID_IFactoidPropertyBag, 0xc4c3e75d, 0xb73d, 0x11d3, 0xb2, 0xcf, 0x0, 0x50, 0x4, 0x89, 0xd6, 0xa3);
DEFINE_GUID(IID_IFactoidServer, 0x3b744d8e, 0xb8a5, 0x11d3, 0xb2, 0xcf, 0x0, 0x50, 0x4, 0x89, 0xd6, 0xa3);
DEFINE_GUID(GUID_TypeLibMstag, 0x9B92EB61, 0xCBC1, 0x11D3, 0x8C, 0x2D, 0x0, 0xA0, 0xCC, 0x37, 0xB5, 0x91);
DEFINE_GUID(GUID_IEOOUIBehaviorFactory, 0x38481807, 0xCA0E, 0x42D2, 0xBF, 0x39, 0xB3, 0x3A, 0xF1, 0x35, 0xCC, 0x4D);

#undef INTERFACE
#define INTERFACE IFactoidPropertyBag
// {C4C3E75D-B73D-11d3-B2CF-00500489D6A3}
DECLARE_INTERFACE_ (IFactoidPropertyBag, IUnknown)
{
	//IUnknown
	STDMETHOD (QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;
	STDMETHOD_ (ULONG, AddRef) (THIS) PURE;
	STDMETHOD_ (ULONG, Release) (THIS) PURE;

	STDMETHOD_ (void, SaveRgbe) (THIS) PURE;
	STDMETHOD_ (BOOL, FSave) (THIS) PURE;
	STDMETHOD_ (void, Free) (THIS) PURE; 
	
	//ACTIONS
	//returned stz is not alloc'ed read-only 
	STDMETHOD_ (const XCHAR *, StzGetFactoidCaption) (THIS) PURE;
	STDMETHOD_ (const XCHAR *, StzGetFactoidCaptionStrict) (THIS) PURE;

	//returned captions are not alloc'ed read-only
	//use the FACTOIDCAPTION to invoke action
	STDMETHOD_ (int, CActionCaptions) (THIS) PURE;
	STDMETHOD_ (BOOL, FGetActionCaption) (THIS_ ACTIONCAPTION *pac, int iac) PURE;	
	STDMETHOD_ (BOOL, FGetActionCaption2) (THIS_ ACTIONCAPTION *pac, int iac, IDispatch* pTarget, BSTR bstrText, BSTR bstrXml) PURE;
	STDMETHOD_ (BOOL, FGetActionCaptionSmartDoc) (THIS_ ACTIONCAPTION *pac, int iac, IDispatch* pTarget, BSTR bstrText, BSTR bstrXml) PURE;
	STDMETHOD_ (ACTYPE, ActypeFromAction) (THIS_ int iac) PURE;
	STDMETHOD (InvokeAction) (THIS_ void *pvActionid, IDispatch *pTarget, BSTR bstrText, BSTR bstrXml) PURE;
	//FMerge Deletes pfpb
	STDMETHOD_ (BOOL, FMerge) (THIS_ IFactoidPropertyBag *pfpb) PURE;
	STDMETHOD_ (BOOL, FSameFactoid) (THIS_ IFactoidPropertyBag *pfpb) PURE;
	
	//The following strings are not allocated for you.  read only.
	STDMETHOD_ (const XCHAR *, SzUri) (THIS) PURE;
	STDMETHOD_ (const XCHAR *, SzTag) (THIS) PURE;
	STDMETHOD_ (const XCHAR *, SzUrl) (THIS) PURE;
	STDMETHOD_ (const XCHAR *, SzHtmlPrefix) (THIS) PURE;
	STDMETHOD_ (BSTR, BstrName) (THIS) PURE;

	STDMETHOD (WriteKeyVal) (THIS_ KEYVAL keyval) PURE;
	STDMETHOD_ (BOOL, FGetKeyVal) (THIS_ int ikeyval, KEYVAL *pkeyval) PURE;
	STDMETHOD_ (int, CKeyVal) (THIS) PURE;
	STDMETHOD_ (void, DeleteKey) (THIS_ int ikey) PURE;
	STDMETHOD (WriteKeyValStz) (THIS_ LPCOLESTR xstz) PURE;

	STDMETHOD_ (BSTR, GetVerbActionLink) (THIS_ void *pvActionid, IDispatch *pTarget, BSTR bstrText, BSTR bstrXml) PURE;
	STDMETHOD_ (void, GetClsidRecognizer) (THIS_ CLSID *pclsid) PURE;
	STDMETHOD_ (BOOL, FAllCaptionsDynamic) (THIS) PURE;
	STDMETHOD_ (BOOL, FAnyCaptionsDynamic) (THIS) PURE;
	STDMETHOD_ (BOOL, FShouldHideIndicator) (THIS) PURE;
	STDMETHOD_ (BOOL, FNeedDispatchToDrawMenu) (THIS) PURE;
	STDMETHOD_ (BOOL, FGetDBServerInfo) (THIS_ int i, NLSI *pnlsi, int *pidSrvrAdjReal) PURE;

	STDMETHOD (OnCheckBoxChange) (THIS_ void *pvActionid, IDispatch* pdispTarget, BOOL fChecked) PURE;	
	STDMETHOD (OnTextboxContentChange) (THIS_ void *pvActionid, IDispatch* pdispTarget, BSTR bstrContent) PURE;	
	STDMETHOD (OnListOrComboSelectChange) (THIS_ void *pvActionid, IDispatch* pdispTarget, int iSelected, BSTR bstrSelected) PURE;	
	STDMETHOD (ImageClick) (THIS_ void *pvActionid, BSTR bstrText, BSTR bstrXml, IDispatch* pdispTarget, int xCoord, int yCoord) PURE;
	STDMETHOD (OnRadioSelectChange) (THIS_ void *pvActionid, IDispatch* pdispTarget, int iSelected, BSTR bstrRadio) PURE;	
};

#undef INTERFACE
#define INTERFACE IFactoidPropertyBagStore
//{C4C3E75C-B73D-11d3-B2CF-00500489D6A3}
DECLARE_INTERFACE_ (IFactoidPropertyBagStore, IUnknown)
{
	//IUnknown
	STDMETHOD (QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;
	STDMETHOD_ (ULONG, AddRef) (THIS) PURE;
	STDMETHOD_ (ULONG, Release) (THIS) PURE;
	
	STDMETHOD_ (BOOL, FCommitPbag) (THIS_ IFactoidPropertyBag *ppbag) PURE;
	STDMETHOD_ (void, SaveRgbe) (THIS) PURE;
	STDMETHOD_ (BOOL, FBeginSave) (THIS_ IStream *pIStream) PURE;
	STDMETHOD_ (void, EndSave) (THIS) PURE;
	STDMETHOD_ (BOOL, FBeginLoad) (THIS_ IStream *pIStream) PURE;
	STDMETHOD_ (void, EndLoad) (THIS) PURE;
	STDMETHOD_ (BOOL, FLoadPropBag) (THIS_ IFactoidPropertyBag **ppfpb) PURE;
	STDMETHOD_ (void, Free) (THIS) PURE;
	STDMETHOD_ (BOOL, FNewFactoidPropBag) (THIS_ IFactoidPropertyBag **ppfpb, 
								LPCOLESTR xstzUrnTag, LPCOLESTR xstzProp, GUID guidAction) PURE;
	STDMETHOD_ (IFactoidPropertyBag *, NewFactoidPropBagFac) (THIS_ const XCHAR *szFac) PURE;
	STDMETHOD_ (IFactoidPropertyBag *, NewFactoidPropBagUriTag) (THIS_ const XCHAR *szUri, const XCHAR *szTag) PURE;
	STDMETHOD_ (IFactoidPropertyBag *, NewFactoidPropBagUriTagNoNewType) (THIS_ const XCHAR *szUri, const XCHAR *szTag) PURE;
	STDMETHOD_ (IFactoidPropertyBag *, ClonePropBag) (THIS_ IFactoidPropertyBag *pfpb) PURE;
	STDMETHOD_ (BOOL, FBeginSaveHTML) (THIS_ const XCHAR *szHTMLPrefix) PURE;
	STDMETHOD_ (void, EndSaveHTML) (THIS) PURE;
	
	//FBeginSaveHTML must be called prior to the following methods
	//The array is allocated by the caller.  The strings are allocated by the functions
	//The string arrays are valid until EndSaveHTML is called.
	STDMETHOD_ (int, CUriPrefix) (THIS_ const XCHAR **rgszUri, const XCHAR **rgszPrefix, int cUriPrefixMax) PURE;
	STDMETHOD_ (int, CUriTagUrl) (THIS_ const XCHAR **rgszUri, const XCHAR **rgszTag, const XCHAR **rgszUrl, int cUriTagUrlMax) PURE;
	STDMETHOD_ (int, CTagPrefix) (THIS_ const XCHAR **rgszTag, const XCHAR **rgszPrefix, int cTagPrefixMax) PURE;

	STDMETHOD_ (int, GetPropertyBags) (THIS_ IFactoidPropertyBag **rgpfpb, int cMax) PURE;
	STDMETHOD_ (void, SetSaveType) (THIS_ int fBackground) PURE;

	// property bag helper functions
	STDMETHOD_ (BOOL, FKeyIsBuiltInProp) (THIS_ const XCHAR* xszQualProp, int cchQualProp, const XCHAR** pxchProp, int* pcchProp);
	STDMETHOD_ (FPER, FperFromPbag) (THIS_ IFactoidPropertyBag *pfpb) PURE;
	STDMETHOD_ (int, FExpireTimeFromPbag) (THIS_ IFactoidPropertyBag *pfpb, SYSTEMTIME* psystime) PURE;
	STDMETHOD_ (int, FExpiredFromPbag) (THIS_ IFactoidPropertyBag *pfpb) PURE;
	STDMETHOD_ (int, FGetRunActionCaptionID) (THIS_ IFactoidPropertyBag *pfpb, void** pppv) PURE;
	STDMETHOD_ (int, FstGetFactoidShowType) (THIS_ IFactoidPropertyBag *pfpb) PURE;
};


typedef int (*PFNCOMMITSUPLTAG)(void* pv, BSTR bstrFac, int ichOffset, int cch, 
				   			    ISmartTagProperties* pfp);


#undef INTERFACE
#define INTERFACE IFactoidServer
// {3B744D8E-B8A5-11d3-B2CF-00500489D6A3}
DECLARE_INTERFACE_ (IFactoidServer, IUnknown)
{
	//IUnknown
	STDMETHOD (QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;	
	STDMETHOD_ (ULONG, AddRef) (THIS) PURE;
	STDMETHOD_ (ULONG, Release) (THIS) PURE;

	//FCHECK API
	STDMETHOD_ (FECODE, CheckInitFactoid) (THIS_ FACID *pcid, CFINIT *pcfinit) PURE;
	STDMETHOD_ (FECODE, CheckTerminate) (THIS_ FACID facid, DWORD *pdwTerm) PURE;

	STDMETHOD_ (FECODE, CheckBatchUpdateFactoid) (THIS_ FACID facid, FCA *pfca, FCH fch,
				PFNBATCHUPDATEPROC pfnBatchUpdProc, int *phsc, int *pifac, PVOID pvData) PURE;

    STDMETHOD_ (FECODE, CheckTextFactoid) (THIS_ FACID facid, 
											FCC 	fcc,
											FCA		*pcaInput,
											PFNADDHANDLEPROC pfnAddHandleProc,
											const XCHAR	*rgxchPara,
											FCP		cchPara,
											ISmartTagTokenList *psttl,
											LCID	lcidText,
											DWORD	ft,
											PVOID 	pvData) PURE;

   	STDMETHOD_ (FECODE, CheckGetDBServerInfo) (THIS_ FACID facid, int idSrvrAdj, NLSI *pnlsi) PURE;
   	STDMETHOD_ (FECODE, CheckSetDBServerOption) (THIS_ FACID facid, int idSrvrAdj, int grfSrvrOpt) PURE;
				   
	STDMETHOD_ (FECODE, CheckFreeHandle) (THIS_ FACID facid, FDOC fdoc, FCH fch) PURE;
	STDMETHOD_ (FECODE, CheckVersion)(THIS_ FVI *pfvi) PURE;
	
	STDMETHOD_ (IFactoidPropertyBagStore *, NewPropertyBagStore) (THIS) PURE;

	STDMETHOD_ (void, SaveRgbeHandle)(THIS_ FACID facid, FDOC fdoc, FCH fch) PURE;
	STDMETHOD_ (void, SaveRgbe) (THIS_ int fMarkAll) PURE;
	STDMETHOD_ (BOOL, FDebugChk) (THIS) PURE;
	STDMETHOD_ (void, SuspendRecognizer) (THIS_ int fSync) PURE;
	STDMETHOD_ (void, ResumeRecognizer) (THIS_ int fSync) PURE;
	STDMETHOD_ (BOOL, FEnsureFactoid)(THIS_ const XCHAR *szUri, const XCHAR *szTag, const XCHAR *szUrl) PURE;
	STDMETHOD_ (BOOL, FExceptionErrMsg)(THIS_ XCHAR *xszBuf, int cchBuf, int *pfStatus) PURE;
	STDMETHOD_ (FCH,  FchGetNextHandle) (THIS_ FACID facid, FDOC fdoc, FCA *pfca) PURE;

	STDMETHOD_ (void, SetThreadPriority) (THIS_ int nPriority) PURE;
	STDMETHOD_ (void, ResetThreadPriority) (THIS);
	STDMETHOD_ (FECODE, RestartServers) (THIS_ int fWaitRecogSvr);
	STDMETHOD_ (BOOL, FGetFactoidType) (THIS_ int idFactoid, const XCHAR **ppszUri, const XCHAR **ppszTag) PURE;
	STDMETHOD_ (IFactoidPropertyBag*, BindPropertyBagToFactoid) (THIS_ ISmartTagProperties *pf, BSTR bstrFac) PURE;

	STDMETHOD_ (int, FGetSmartTagAppPrefs) (THIS_ const char* szAppWndClass, STAP* pstap) PURE;
	STDMETHOD_ (int, CGetSmartTagAppPrefs) (THIS_ STAP** pprgstap) PURE;
	STDMETHOD_ (void, FreeSmartTagAppPrefs) (THIS_ STAP* rgstap, int crgstap, int fOnlyFreeNlsi) PURE;
	STDMETHOD_ (void, ApplySmartTagAppPrefs) (THIS_ FACID facid, STAP* pstap) PURE;
	STDMETHOD_ (int, CGetAllRecognizerInfo) (THIS_ FACID facid, NLSI** pprgnlsi) PURE;

	STDMETHOD_ (void, LaunchSmartTagSite) (THIS_ int lcid);	
	STDMETHOD_ (BOOL, FHasPropertyPage) (THIS_ FACID facid, int idSrvrAdj) PURE;
	STDMETHOD_ (void, InvokePropertyPage) (THIS_ FACID facid, int idSrvrAdj) PURE;	
	STDMETHOD_ (void, AddToIgnoreList) (THIS_ BSTR bstrText, BSTR bstrUri) PURE; 
	STDMETHOD_ (int, FOnIgnoreList) (THIS_ BSTR bstrFac, BSTR bstrText) PURE;	
	STDMETHOD_ (void, CheckSupplementalListForGrammar) (THIS_ PFNCOMMITSUPLTAG pfncst, void* pv, 
														BSTR bstrFacName, BSTR bstrRawText,
													 	ISmartTagTokenList* psttl) PURE;

	STDMETHOD_ (void, TokenListFromBstr) (THIS_ BSTR bstrRawText, int lcid, ISmartTagTokenList** ppsttl) PURE;
	STDMETHOD_ (void, SetCurrentSolution) (THIS_ WCHAR* wzSolution) PURE;
	STDMETHOD_ (const XCHAR*, XstzGetCurrentSolution) (THIS) PURE;
	STDMETHOD_ (BOOL, FInitSmartDocWithUri) (THIS_ const XCHAR* szUri, const XCHAR* szFilePath, const XCHAR* szRegKeyRoot, IDispatch* pdispDocument) PURE;	
	STDMETHOD_ (void, PaneUpdateComplete) (THIS_ IDispatch* pdispDoc);
	STDMETHOD_ (BOOL, FEnumSmartDocProps) (THIS_ ISmartDocProperties* psdp, int iProp, 
									      BSTR* pbstrKey, BSTR* pbstrVal) PURE;	

	STDMETHOD_ (BOOL, FInitFRTE) (THIS_ BSTR bstrText, int lcidText, int ft, FRTE* pfrte) PURE;
	STDMETHOD_ (BOOL, FEnumFRTE) (THIS_ FRTE* pfrte) PURE;	
};

//definition must match mstag.idl
#ifndef __ISmartTagProperties_INTERFACE_DEFINED__
#define __ISmartTagProperties_INTERFACE_DEFINED__

#undef INTERFACE
#define INTERFACE ISmartTagProperties
//{54F37842-CDD7-11d3-B2D4-00500489D6A3}
DECLARE_INTERFACE_ (ISmartTagProperties, IDispatch)
{
	STDMETHOD (QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;
	STDMETHOD_ (ULONG, AddRef) (THIS) PURE;
	STDMETHOD_ (ULONG, Release) (THIS) PURE;
             
	STDMETHOD (GetTypeInfoCount) (THIS_ UINT *pctinfo) PURE;
	STDMETHOD (GetTypeInfo) (THIS_ UINT itinfo, LCID lcid, ITypeInfo** pptinfo) PURE;
	STDMETHOD (GetIDsOfNames) (THIS_ REFIID riid, _In_count_(cNames) LPOLESTR *rgszNames, UINT cnames, LCID lcid, _Out_cap_(cNames) DISPID *rgdispid) PURE;
	STDMETHOD (Invoke) (THIS_ DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams,
						VARIANT *pvar, EXCEPINFO *pexcepinfo, UINT *puArgErr) PURE;
	
	STDMETHOD (get_Read) (THIS_ BSTR bstrKey, BSTR *pbstrValue) PURE;
	STDMETHOD (Write) (THIS_ BSTR bstrKey, BSTR bstrValue) PURE;
	STDMETHOD (get_Count) (THIS_ int *pwcount);
	STDMETHOD (get_KeyFromIndex) (THIS_ int index, BSTR *pbstrKey) PURE;
	STDMETHOD (get_ValueFromIndex) (THIS_ int index, BSTR *pbstrValue) PURE;
};

#endif

typedef IFactoidServer * (WINAPI *PFNFACTOIDSERVER)(void);

#ifdef __cplusplus
extern "C" IFactoidServer* __cdecl InitFactoidServer(HINSTANCE hinst);
extern "C" void __cdecl FreeFactoidServer(IFactoidServer*);
#else
IFactoidServer* __cdecl InitFactoidServer(HINSTANCE hinst);
void __cdecl FreeFactoidServer(IFactoidServer*);
#endif

#ifdef __cplusplus
extern "C" void __cdecl MSTAGLibCheckMem(int fIgnoreSpecial);
extern "C" void __cdecl MSTAGLibInit(void);
extern "C" void __cdecl MSTAGLibTerm(void);
extern "C" CLSID __cdecl MSTAGLibIIDMofl(void);
#else
void __cdecl MSTAGLibCheckMem(int fIgnoreSpecial);
void __cdecl MSTAGLibInit(void);
void __cdecl MSTAGLibTerm(void);
CLSID __cdecl MSTAGLibIIDMofl(void);
#endif //__cplusplus

// use these to distinguish between smart tag and smart doc action captions.
#define FIsSmartTagActionCaption(ac) (ac.sdacd.actype == actypeSmartTag)
#define FIsSmartDocActionCaption(ac) (ac.sdacd.actype != actypeSmartTag)
#define FIsSmartDocHelpMenuActionCaption(ac) (ac.sdacd.actype == actypeHelp || ac.sdacd.actype == actypeHelpURL)
#define FIsSmartDocDocFragMenuActionCaption(ac) (ac.sdacd.actype == actypeDocumentFragment || ac.sdacd.actype == actypeDocumentFragmentURL)

#endif //_FACTOID_H_

