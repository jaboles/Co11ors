#pragma once

/****************************************************************************
	MsoWOpt.h

	Owner: ilyav
 	Copyright (c) 1997 Microsoft Corporation

	This file contains the exported interfaces and declarations for
	the Web Options dialog and associated functionality.
****************************************************************************/

#ifndef MSOWOPT_H
#define MSOWOPT_H

#if !defined(MSOSTD_H)
#include <msostd.h>
#endif

//-----------------------------------------------------------------------------
//Miniwebs for Word was cut for Office 10 but may be back.  
//#ifdef MINIWEB will be used to hide the code.
//-----------------------------------------------------------------------------
//htmlopt.des additions:
//	SUB_DIALOG sabGeneralWord = fGeneralWord AT (0, 0, 276, 184) ID _fn4cys_g4_0
//		IF fGeneralWord 
//			TEXT "Navigation" AT (7, 24, 47, 10) ID _fn4d1z_1e_1
//				TMC tmcWebOptNavigationLabel
//			TEXT "" AT (49, 27, 215, 1) ID _fn4d2l_p0_2
//				TMC tmbWebOptBorder16 BORDER
//			CHECK_BOX "&Add navigation controls" AT (15, 33, 248, 10) ID _fn4d59_g4_0
//				TMC tmcWebOptAddNavigation ACTION ARG fAddNavigation
//			TEXT "Navigate by:" AT (21, 43, 54, 10) ID _fn4d72_k0_0
//				TMC tmcWebOptNavigateByLabel
//			ENDIF 
//		RADIO_GROUP
//			ID _fn4d8v_eq_1
//			TMC tmcWebOptNavigateBy ARG iNavigateBy
//			{
//			IF fGeneralWord 
//				RADIO_BUTTON "Printed Page" AT (32, 53, 68, 10) ID _fn4d8v_eq_2
//					TMC tmcWebOptPrintedPage ACTION
//				RADIO_BUTTON "Heading, down to level:" AT (32, 63, 88, 10) ID _fn4db3_q4_0
//					TMC tmcWebOptHeading ACTION
//				RADIO_BUTTON "&Section" AT (32, 73, 97, 10) ID _fn4ddj_c8_1
//					TMC tmcWebOptSection ACTION
//				RADIO_BUTTON "S&creenful of data based on screen size:" AT (32, 83, 135, 10) ID _fn4de5_fu_2
//					TMC tmcWebOptScreenful ACTION
//				RADIO_BUTTON "&Existing manual web breaks" AT (32, 93, 115, 10) ID _fn4dij_4g_0
//					TMC tmcWebOptExisting ACTION
//				ENDIF 
//			}
//		IF fGeneralWord 
//			EDIT AT (127, 59, 40, 12) ID _fn4dbp_hs_1
//				TMC tmcWebOptHeadingLevel ACTION ARG iHeadingLevel
//				CHAR_VALIDATED SPIN 
//				PARSE_PROC WCWOPT_HeadingLevelParseProc
//				SIZE 1 
//			LIST_BOX AT (172, 79, 52, 48) ID _fn4dh4_p0_0
//				TMC tmcWebOptScreenfulList DROP_DOWN_SIBLING ARG iScreenFulSize
//				LIST_BOX_PROC WCWOPT_ScreenSizeListProc
//			TEXT "Navigation links:" AT (21, 108, 70, 10) ID _fn4djc_go_0
//				TMC tmcWebOptNavigationLinksLabel
//			LIST_BOX AT (89, 107, 173, 70) ID _fn4dn9_hs_1
//				TMC tmcWebOptNavigationLinksList DROP_DOWN_SIBLING ARG iNavigationLinks
//				LIST_BOX_PROC WListEntblStt WPARAM `msosttWebOptNavigationLinks`
//				NO_SCROLL_BAR 
//			TEXT "Links' location on web pages:" AT (22, 121, 116, 10) ID _fn4dmi_iw_0
//				TMC tmcWebOptLinksLocationLabel
//			LIST_BOX AT (140, 121, 80, 40) ID _fn4dl2_dw_0
//				TMC tmcWebOptLinksLocationList DROP_DOWN_SIBLING ARG iLinksLocation
//				LIST_BOX_PROC WListEntblStt WPARAM `msosttWebOptLinksLocation`
//				NO_SCROLL_BAR 
//			ENDIF 
//ostrman.pp additions
//			DefineStt(msosttWebOptNavigationLinks, Allocated, OINTL)
//			DefineIds("Previous|Next"!"Web Option - Main Dialogs"                    , msoidsNavigationLinksPreviousNext               , msosttWebOptNavigationLinks)
//			DefineIds("Previous|Home|End|Next"!"Web Option - Main Dialogs"           , msoidsNavigationLinksPreviousHomeEndNext        , msosttWebOptNavigationLinks)
//			DefineIds("Previous|Web Page #|Next"!"Web Option - Main Dialogs"         , msoidsNavigationLinksPreviousWebPageXNext       , msosttWebOptNavigationLinks)
//			DefineIds("Previous|Web Page X of Y|Next"!"Web Option - Main Dialogs"    , msoidsNavigationLinksPreviousWebPageXofYNext    , msosttWebOptNavigationLinks)
//			DefineIds("Previous|Home|Web Page #|End|Next"!"Web Option - Main Dialogs", msoidsNavigationLinksPreviousHomeWebPageXEndNext, msosttWebOptNavigationLinks)
//			DefineIds("Previous|Web Page X of Y|End|Next"!"Web Option - Main Dialogs", msoidsNavigationLinksPreviousWebPageXofYEndNext , msosttWebOptNavigationLinks)
//			DefineIds("Keep current links"!"Web Option - Main Dialogs"               , msoidsNavigationLinksKeepCurrentLinks           , msosttWebOptNavigationLinks)
//
//			DefineStt(msosttWebOptLinksLocation, Allocated, OINTL)
//			DefineIds("Top and Bottom"!"Web Option - Main Dialogs" , msoidsLinksLocationTopBottom  , msosttWebOptLinksLocation)
//			DefineIds("Top"!"Web Option - Main Dialogs"            , msoidsLinksLocationTop        , msosttWebOptLinksLocation)
//			DefineIds("Bottom"!"Web Option - Main Dialogs"         , msoidsLinksLocationBottom     , msosttWebOptLinksLocation)
//			DefineIds("Keep current"!"Web Option - Main Dialogs"   , msoidsLinksLocationKeepCurrent, msosttWebOptLinksLocation)
//#define MINIWEB 1
//-----------------------------------------------------------------------------

//-------------------------------------------------------------------
// Web Options dlg functions
//-------------------------------------------------------------------
interface IMsoWebOptionsDlg;

//-------------------------------------------------------------------
// Web Options Dlg Command codes.
// These codes are returned from IMsoWebOptionsDlg::FRun.
//-------------------------------------------------------------------
enum
{
	msowoptcmdCancel =  0,         // User cancelled dlg.
	msowoptcmdOK     =  1,         // User OK'ed dlg.
	msowoptcmdErr    = -1,         // Some error occurred.
};

typedef enum
{
	msowoptTargetBrowserV3,
	msowoptTargetBrowserV4,
	msowoptTargetBrowserIE4,
	msowoptTargetBrowserIE5,
	msowoptTargetBrowserIE6,
	msowoptTargetBrowserERROR,
} MSOWOPT_TARGETBROWSERS;

// We recognize exactly six apps.
// MUST MATCH WCWOPT_TABLE_FLAGS for the apps
enum MSOWOPT_APP
{
	msowoptAppAccess,
	msowoptAppExcel,
	msowoptAppPPT,
	msowoptAppPub,
	msowoptAppWord,
};

// Some options belong to enumerated types.
enum MSOWOPT_PPT_FRAME_COLORS
{
	msowoptFrameColorsBrowserColors,
	msowoptFrameColorsPresentationSchemeTextColor,
	msowoptFrameColorsPresentationSchemeAccentColor,
	msowoptFrameColorsWhiteTextOnBlack,
	msowoptFrameColorsBlackTextOnWhite
};
#define MSOWOPT_FrameColorsToVBAFrameColors(fc) (EPPFrameColors)((fc) + 1)
#define MSOWOPT_VBAFrameColorsToFrameColors(fc) (MSOWOPT_PPT_FRAME_COLORS)((fc) - 1)

typedef enum 
{
	msowoptWordBrowserLevelV4  = 0, 
	msowoptWordBrowserLevelIE5 = 1,
	msowoptWordBrowserLevelIE6 = 2,
} MSOWOPT_WORD_BROWSER_LEVEL;

enum MSOWOPT_SCREEN_SIZE
{
	msowoptScreenSize544x376,
	msowoptScreenSizeWebTV=msowoptScreenSize544x376,
	msowoptScreenSize640x480,
	msowoptScreenSize720x512,
	msowoptScreenSize800x600,
	msowoptScreenSize1024x768,
	msowoptScreenSize1152x882,
	msowoptScreenSize1152x900,
	msowoptScreenSize1280x1024,
	msowoptScreenSize1600x1200,
	msowoptScreenSize1800x1440,
	msowoptScreenSize1920x1200,
	msowoptScreenSizeMax,
};
#define MSOWOPT_ScreenSizeToVBAScreenSize(ss) (MsoScreenSize)(ss)
#define MSOWOPT_VBAScreenSizeToScreenSize(ss) (enum MSOWOPT_SCREEN_SIZE)(ss)

#ifdef MINIWEB
enum WCWOPT_LINKS_LOCATION
{
	msowoptLinksLocationTopBottom  ,
	msowoptLinksLocationTop        ,
	msowoptLinksLocationBottom     ,
	msowoptLinksLocationKeepCurrent,
};
#endif

#define MAX_PixelsPerInch 480
#define MIN_PixelsPerInch  19

// The scripts we recognize:
enum MSOWOPT_SCRIPT
{
	msoScriptArabic,
	msoScriptCyrillic,
	msoScriptLatin,
	msoScriptGreek,
	msoScriptHebrew,
	msoScriptJapan,
	msoScriptKorea,
	msoScriptOther,
	msoScriptChina,
	msoScriptThai,
	msoScriptTaiwan,
	msoScriptVietNam,
	msoScriptHindi,
	msoScriptMax,
	msoScriptNil = -1,
};
#define MSOWOPT_ScriptToVBAScript(s) (MsoCharacterSet)((s) + 1)
#define MSOWOPT_VBAScriptToScript(s) (enum MSOWOPT_SCRIPT)((s) - 1)

#define MSOWOPT_EncodingToVBAEncoding(enc) (MsoEncoding)(enc)
#define MSOWOPT_VBAEncodingToEncoding(enc) (DWORD)(enc)

// The global options.  They customize the application or Office,
// and not the document.
struct MSOWOPT_ALL_GLOBAL
{
	BYTE     iHyperlinkColor_Acc;
	BYTE     iFollowedHyperlinkColor_Acc;
	union
		{
		unsigned grf;
		struct
			{
			unsigned fUnderlineHyperlinks_Acc        : 1;
			unsigned fSaveHiddenData_XL              : 1;
			unsigned fLoadPictures_XL                : 1;
			unsigned fUpdateLinksOnSave              : 1;
			unsigned fCheckIfOfficeIsHTMLEditor      : 1;
			unsigned fCheckIfWordIsDefaultHTMLEditor : 1;
			unsigned fAlwaysSaveInDefaultEncoding    : 1;
			unsigned fUseMHTML                       : 1;
			unsigned unused                          :24;
			};
		};
};

// The doc-specific options.  They customize the document, although the default
// value is saved in the registry.
struct MSOWOPT_ALL_DOCSPEC
{
	BYTE     frameColors_PPT;                     /* enum MSOWOPT_PPT_FRAME_COLORS */
	BYTE     screenSize;                          /* enum MSOWOPT_SCREEN_SIZE */
#ifdef MINIWEB
	BYTE     iScreenfulSize;                      /* enum MSOWOPT_SCREEN_SIZE */
	BYTE     iNavigationLinks;                    /* enum MSOWOPT_NAVIGATION_LINKS */
	BYTE     iLinksLocation;                      /* enum WCWOPT_LINKS_LOCATION */
#endif
	union
		{
		unsigned grf;
		struct
			{
			unsigned fRelyOnCSS_WordXL               : 1;
			unsigned fIncludeNavigation_PPT          : 1;
			unsigned fResizeGraphics_PPT             : 1;
			unsigned fOptimizeForBrowser_Word        : 1;
			unsigned fOrganizeInFolder               : 1;
			unsigned fUseLongFileNames               : 1;
			unsigned fDownloadComponents_AccXL       : 1;
			unsigned fRelyOnVML                      : 1;
			unsigned fAllowPNG                       : 1;
			unsigned fShowSlideAnimation_PPT         : 1;
			unsigned unused2                         : 1; //Was MHTML but was moved to Global options
			unsigned fPPTFormat                      : 1; //True == Dual.  False == V4+.
			unsigned fAddHlinkToWebNavBar            : 1;
			unsigned fLoopForever                    : 1;
#ifdef MINIWEB
			unsigned fAddNavigation                  : 1;
			unsigned iNavigateBy                     : 3; //5 choices.
			unsigned unused                          :13;
#endif
			unsigned unused                          :18;
			};
		};

	DWORD dwPixelsPerInch;
	DWORD uiCodePage;
	DWORD dwLoopTime;
#ifdef MINIWEB
	DWORD dwHeadingLevel;
#endif

	WCHAR *pwzLocationOfComponents_AccXL;
	WCHAR *pwzPageTitle;
	WCHAR *pwzKeywords;
	WCHAR *pwzDescription;
	WCHAR *pwzBgndSound;
};


struct MSOWOPT_ENCODING
{
    DWORD dwFlags;
    UINT  uiCodePage;
    UINT  uiFamilyCodePage;
    WCHAR wzName[MAX_MIMECP_NAME];
};

MSOAPI_(HRESULT)                    MsoHrGenerateEncodingsList(struct MSOWOPT_ENCODING** prgenc, int* piencMac);
MSOAPI_(MSOWOPT_WORD_BROWSER_LEVEL) MsoGetBrowserLevel(struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt);
MSOAPI_(MSOWOPT_WORD_BROWSER_LEVEL) MsoGetBrowserLevelFromTargetBrowser(MSOWOPT_TARGETBROWSERS tb);
MSOAPI_(MSOWOPT_TARGETBROWSERS)     MsoGetTargetBrowserFromBrowserLevel(MSOWOPT_WORD_BROWSER_LEVEL bl);
MSOAPI_(MSOWOPT_TARGETBROWSERS)     MsoGetDefaultTargetBrowser();
MSOAPI_(void)                       MsoSetDefaultTargetBrowser(MSOWOPT_TARGETBROWSERS);
MSOAPI_(MSOWOPT_TARGETBROWSERS)     MsoGetTargetBrowser(struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt);
MSOAPI_(void)                       MsoSetTargetBrowser(struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt, MSOWOPT_TARGETBROWSERS);

// All Boolean values.
enum MSOWOPT_WEBOPTBOOL
{
	msowoptbool_RelyOnCSS_WordXL,               //0
	msowoptbool_RelyOnVML,                      //1
	msowoptbool_OptimizeForBrowser_Word,        //2
	msowoptbool_PPTFormat,                      //3
	msowoptbool_AllowPNG,                       //4
	msowoptbool_MHTML,                          //5
	msowoptbool_UnderlineHyperlinks_Acc,        //6
	msowoptbool_SaveHiddenData_XL,              //7
	msowoptbool_LoadPictures_XL,                //8
	msowoptbool_UpdateLinksOnSave,              //9
	msowoptbool_CheckIfOfficeIsHTMLEditor,      //10
	msowoptbool_CheckIfWordIsDefaultHTMLEditor, //11
	msowoptbool_AlwaysSaveInDefaultEncoding,    //12
	msowoptbool_IncludeNavigation_PPT,          //13
	msowoptbool_ResizeGraphics_PPT,             //14
	msowoptbool_OrganizeInFolder,               //15
	msowoptbool_UseLongFileNames,               //16
	msowoptbool_DownloadComponents_AccXL,       //17
	msowoptbool_ShowSlideAnimation_PPT,         //18
	msowoptbool_AddHlinkToWebNavBar,            //19
	msowoptbool_LoopForever,                    //20
#ifdef MINIWEB
	msowoptbool_AddNavigation,
	msowoptbool_NavigateBy,
#endif
	msowoptbool_cMax
};

// All enumerated values.
enum MSOWOPT_WEBOPTENUM
{
	msowoptenum_FrameColors_PPT           ,
	msowoptenum_ScreenSize                ,
	msowoptenum_HyperlinkColor_Acc        ,
	msowoptenum_FollowedHyperlinkColor_Acc,
#ifdef MINIWEB
	msowoptenum_ScreenfulSize             ,
	msowoptenum_NavigationLinks           ,
	msowoptenum_LinksLocation             ,
#endif
};

// All integer values.
enum MSOWOPT_WEBOPTINT
{
	msowoptint_PixelsPerInch              ,
	msowoptint_LoopTime                   ,
	msowoptint_CodePage                   ,
#ifdef MINIWEB
	msowoptint_HeadingLevel               ,
#endif
};

// All string values.
enum MSOWOPT_WEBOPTSTR
{
	msowoptstr_LocationOfComponents_AccXL     ,
	msowoptstr_PageTitle                      ,
	msowoptstr_Keywords                       ,
	msowoptstr_Description                    ,
	msowoptstr_BackgndSound                   ,
	//Items below here aren't in the structure, they're derived from
	//other settings.
	msowoptstr_LocationOfComponentsSetup_AccXL, //msowoptstr_LocationOfComponents_AccXL + "files/PFILES/COMMON/MSSHARED/WEBCOMPS/10/setup.exe"
};

// Font information is special because the user can change it for all scripts,
// while the app only needs it for the current one.
struct MSOWOPT_FONTINFO
{
	WCHAR wzFixedFace[LF_FACESIZE];
	WCHAR wzPropFace[LF_FACESIZE];
	USHORT uihpFixedSize;	// in half-points.
	USHORT uihpPropSize;	// in half-points.
};


// Function DrawColorComboItem from %access%\util\winutil.c
typedef BOOL (
#ifdef DEBUG
				  __cdecl
#else
				  __fastcall
#endif
				  * MSOPFNFACCDRAWCOLORCOMBOITEM)(DRAWITEMSTRUCT *);
// Function DrawColorComboItem from %access%\util\winutil.c
typedef BOOL (
#ifdef DEBUG
				  __cdecl
#else
				  __fastcall
#endif
				  * MSOPFNFACCGETCOLORCOMBOITEMTEXT)(int nColor, WCHAR *wzColor, int cchMax);

// Reload the document in the given encoding.
typedef BOOL (MSOSTDAPICALLTYPE *MSOPFNFRELOADENCPROC) (UINT *puiCodePage, DWORD dwParam);

#define MSOWOPT_MAX_PANE 6

interface IMsoWebOptionsDlg;

//Enabled             = 0000
//Diabled             = 0001
//Hidden              = 0010
//Hidden and Disabled = 0011  Valid but weird state
//Hidden and Enabled  = 0010  Valid but weird state

#define RELOAD_ENABLED  0x0000
#define RELOAD_DISABLED 0x0001
#define RELOAD_HIDDEN	0x0002
#define RELOAD_DATAPAGE 0x0010

//Enabled means the text and control are visible and enabled
//Disabled means the text and control are visible and disabled
//Hidden means the text and control are hidden and hint text is visible
//Datapage is like Hidden except there's substitute text specifically for datapages

#undef  INTERFACE
#define INTERFACE  IMsoWebOptionsDlg
DECLARE_INTERFACE_(IMsoWebOptionsDlg, ISimpleUnknown)
{
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;
	MSOMETHOD_(void, Delete) (THIS) PURE;

	// Show the dialog.
	MSOMETHOD_(BOOL, FRun) (THIS_ HWND hwndOwner, LONG *plRetCode) PURE;

	// Set the values before the dialog is up.
	MSOMETHOD_(void, SetOptions) (THIS_ struct MSOWOPT_ALL_GLOBAL *pGlobalOpt, struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt) PURE;

	// Get the values after the dialog is down.
	MSOMETHOD_(void, GetOptions) (THIS_ struct MSOWOPT_ALL_GLOBAL *pGlobalOpt, struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt) PURE;

	// Function DrawColorComboItem from %access%\util\winutil.c
	MSOMETHOD_(void, SetAccDrawColorComboItem) (THIS_ MSOPFNFACCDRAWCOLORCOMBOITEM pfnFAccDrawColorComboItem) PURE;

	// Function GetColorComboItemText from %access%\util\winutil.c
	MSOMETHOD_(void, SetAccGetColorComboItemText) (THIS_ MSOPFNFACCGETCOLORCOMBOITEMTEXT pfnFAccGetColorComboItemText) PURE;

	// In which pane to start?
	MSOMETHOD_(void, SetStartPane) (THIS_ int iPane) PURE;

	// Enable reload - as which page initially, and using what procedure?
	MSOMETHOD_(void, SetReload) (THIS_ DWORD uiCodePageReload, MSOPFNFRELOADENCPROC pfnFReloadEncProc, DWORD dwReloadParam) PURE;

	// Returns the status of the last Encoding Reload
	MSOMETHOD_(BOOL, FLastReloadState) (THIS) PURE;

	// Set PowerPoint's color schemes.
	MSOMETHOD_(void, SetPPTColors) (THIS_ COLORREF frameTextColor, COLORREF frameAccent1Color, COLORREF frameBGColor) PURE;

	// In which pane did we finish?
	MSOMETHOD_(int, GetStartPane) (THIS) PURE;
	
	// Have the document-specific options changed?
	MSOMETHOD_(BOOL, FGetChangedDocSpecOpt) (THIS) PURE;

	// Pass an enum for the background sound MRU
	MSOMETHOD_(void, SetBgndSoundMRU) (THIS_ WCHAR **prgwz, int iCount) PURE;

	// Pass an IMsoFindFile for the Browse button to use
	MSOMETHOD_(void, SetIFindFile) (THIS_ IMsoFindFile *piff) PURE;

	// Allow the invalid charset message for MHTML to be shown
	MSOMETHOD_(void, SetShowMhtmlCharsetMsg) (THIS_ BOOL fShowMsg) PURE;

	// Allow the client to specify the mode of Reload
	MSOMETHOD_(void, SetReloadOptions) (THIS_ int grfReload) PURE;
};

// Create and return in *ppiwoptdlg an IMsoWebOptionsDlg instance that can be
// used to invoke the Web Options dialog. hinst is the HMSOINST of
// the host.  When host is finishied with the object, it must call 
// IMsoWebOptionsDlg::Delete to delete the object.
// Return TRUE if successful.
enum MSOWOPT_MODE {msowoptmodeNormal, msowoptmodeSave, msowoptmodeAccessSpecial};
MSOAPI_(BOOL) MsoFCreateWebOptionsDlg(HMSOINST hinst, enum MSOWOPT_APP app, enum MSOWOPT_MODE mode, IMsoWebOptionsDlg **ppiwoptdlg);


//	Read global web options from the registry.
MSOAPI_(void) MsoRegGetGlobalWebOptions(struct MSOWOPT_ALL_GLOBAL *pGlobalOpt, enum MSOWOPT_APP app);

// Read document-specific web options from the registry.
MSOAPI_(void) MsoRegGetDefaultDocSpecWebOptions(struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt, enum MSOWOPT_APP app);

// Get the "factory settings" of document-specific web options.
MSOAPI_(void) MsoFactoryGetDefaultDocSpecWebOptions(struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt, enum MSOWOPT_APP app);

// Get the default font info for web pages in the given encoding.
MSOAPI_(enum MSOWOPT_SCRIPT) MsoScriptOfCodePage(UINT uiCodePage);
MSOAPI_(void) MsoRegGetDefaultHTMLFontInfo(enum MSOWOPT_SCRIPT script, struct MSOWOPT_FONTINFO *pFontInfo, enum MSOWOPT_APP app);

// Set the default font info for web pages in the given encoding.
MSOAPIX_(void) MsoRegSetDefaultHTMLFontInfo(enum MSOWOPT_SCRIPT script, struct MSOWOPT_FONTINFO *pFontInfo, enum MSOWOPT_APP app);

// Convert the enum options to/from strings.
MSOAPI_(const WCHAR *) MsoGetWebOptEnumString(enum MSOWOPT_WEBOPTENUM we, unsigned uiValue);
MSOAPI_(int) MsoGetWebOptEnumValue(enum MSOWOPT_WEBOPTENUM we, const WCHAR *pwzValue, int cchValue);

// Call this when the structure is no longer needed.
MSOAPI_(void) MsoReleaseDocSpecWebOptions(struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt, enum MSOWOPT_APP app);

// Map the screen size enum to a point value.
MSOAPI_(POINT) MsoGetWebOptScreenSize(enum MSOWOPT_SCREEN_SIZE ss);

// Get/Set a single option.  Used in VBA.
MSOAPI_(BOOL) MsoRegGetSingleBoolWebOption(enum MSOWOPT_WEBOPTBOOL wb);
MSOAPI_(void) MsoRegSetSingleBoolWebOption(enum MSOWOPT_WEBOPTBOOL wb, BOOL f);
MSOAPI_(DWORD) MsoRegGetSingleIntWebOption(enum MSOWOPT_WEBOPTINT wi);
MSOAPI_(DWORD) MsoRegGetSingleDefaultIntWebOption(enum MSOWOPT_WEBOPTINT wi);
MSOAPI_(void) MsoRegSetSingleIntWebOption(enum MSOWOPT_WEBOPTINT wi, DWORD dw);
MSOAPI_(void) MsoRegGetSingleStrWebOption(enum MSOWOPT_WEBOPTSTR ws, WCHAR *pwz, DWORD cwch);
MSOAPI_(void) MsoRegSetSingleStrWebOption(enum MSOWOPT_WEBOPTSTR ws, const WCHAR *pwz);
MSOAPI_(int) MsoRegGetSingleEnumWebOption(enum MSOWOPT_WEBOPTENUM we);
MSOAPI_(int) MsoRegGetSingleDefaultEnumWebOption(enum MSOWOPT_WEBOPTENUM we);
MSOAPI_(BOOL) MsoFRegSetSingleEnumWebOption(enum MSOWOPT_WEBOPTENUM we, unsigned uiValue);

// Is this an autodetect pseudo-codepage?
MSOAPI_(BOOL) MsoFCodePageIsAutoDetect(UINT uiCodePage);

// Office's localized name of a codepage.
MSOAPIX_(BOOL) MsoFWtzGetNameOfCodePage(UINT uiCodePage, WCHAR *wtzName, int cchName);

// WebOptions macro recording
MSOAPI_(BOOL) MsoWebOptionsRecording(const struct MSOWOPT_ALL_GLOBAL *pGlobalOpt, struct MSOWOPT_ALL_DOCSPEC *pDocSpecOpt, enum MSOWOPT_APP msowoptApp);

// Return an Ole Automation WebPageFonts object.
MSOAPI_(BOOL) MsoFGetWebPageFontsObject(HMSOINST hinst, enum MSOWOPT_APP msowoptApp, interface WebPageFonts **ppWebPageFonts);


#endif //MSOWOPT_H
