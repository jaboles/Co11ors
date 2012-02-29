#pragma once

/*****************************************************************************

    MSOFC.H

    Owner:    BenW
    Copyright (c) 1994 Microsoft Corporation

    This file contains header information for the Office Assistant.

    Historical Note: The Hungarian for an Office Assistant Object is a
    TFC, which stands for something that BillG said once that
    shouldn't really go into a header file.  So, just think of it as
    "The Friendly Character."

    There are 4 sections here:
    a. Assistant related things
    b. Balloon services things
    c. The ITFC interface
    d. Exported C functions

************************************************* BenW ******************/
#ifndef MSOFC_H
#define MSOFC_H

#if 0 //$[VSMSO]
#include "aw.h"

#define TFC_NOTIFICATION 1
#define MODELESS_ALERTS 1

#define cchPhraseMax 256

// An FCN is a Notification sent to the office assistant (FC)
#define msofcnActivate 1
#define msofcnDeactivate 2


// An FCA is an Animation that the Office Assistant (FC) can do
// These are integer representations of animation names
// Note that this enumeration maps to an array of string animation names
// in fc.cpp and the msoAnimation constants in mso9.odl are mapped to this
// enumeration in the functions MsoFcaFromMsoAnimationType and MsoAnimationTypeFromMsoFca
// in fcoa.cpp. If you change any of this here, propagate to these locations as well.
enum {
    msofcaIdle = 0,             // Idle needs to be the first animation (so it isn't included in random animation selection)
    msofcaFakeIdle,
    msofcaGreeting,
    msofcaBeginSpeaking,
    msofcaCharacterSuccessMajor,
    msofcaGetAttentionMajor,
    msofcaGetAttentionMinor,
    msofcaSearching,
    msofcaPrinting,
    msofcaGestureRight,
    msofcaWritingNotingSomething,
    msofcaWorkingAtSomething,
    msofcaThinking,
    msofcaSendingMail,
    msofcaListensToComputer,
    msofcaAppear,
    msofcaGetArtsy,
    msofcaGetTechy,
    msofcaGetWizardy,
    msofcaCheckingSomething,
    msofcaLookDown,
    msofcaLookDownLeft,
    msofcaLookDownRight,
    msofcaLookLeft,
    msofcaLookRight,
    msofcaLookUp,
    msofcaLookUpLeft,
    msofcaLookUpRight,
    msofcaSaving,
    msofcaGestureDown,
    msofcaGestureLeft,
    msofcaGestureUp,
    msofcaEmptyTrash,
    msofcaRestPose,
    msofcaGoodbye,              // Rest Pose, Goodbye, and Disappear need to be the last 3 animations
    msofcaDisappear,            // before msofcaMax (this way we won't include them in the randome animation selection).
                                // If you add more such animations, change DoRandomAnimation.
    msofcaMax,                  // Always leave as last
    msofcaNil = -1
    };

// An ANF is an ANimation Flag passed to ITFC::FDoAnimation
#define msoanfLoop      1   // loop the animation until it's stopped or a new one comes in
#define msoanfNoSound   2   // turn off sound temporarily just for this animation   (NOTE This flag is ignored.  Sounds are either globally ON or OFF)

#if 0 //$[VSMSO]
#include "msobal.h" /* This is the simple header for the Office Balloon DLL. */
#endif //$[VSMSO]

/* Tips */

typedef unsigned long HTIP;

/* Tip Types */
#define msottpGeneral           0
#define msottpFeature           1
#define msottpKeyShortcut       2
#define msottpMouseShortcut 3
#define msottpFun                   4
#define msottpTipOfDay          5
#define msottpMax                   6

/* Tip Priorities */
#define msotprUrgent   0
#define msotprHigh     1
#define msotprMedium   2
#define msotprLow      3
#define msotprMax      4

#define msocwchShort    302

typedef struct
    {
    WCHAR wz[msocwchShort];
    int tcid;
    int tbct;
    void *pvClient;
    IMsoControlUser *picu;
    } MSOTCN;

#endif //$[VSMSO]
// Balloon button notification function
typedef unsigned long HBALLOON;
typedef VOID (__stdcall * MSOPFNDOIBTN)(HBALLOON, int, LONG);
#define MSODOIBTN VOID __stdcall
#if 0 //$[VSMSO]
typedef BOOL (__stdcall * MSOPFNTFCFFILTERMSG)(LPMSG);
#define MSOTFCFFILTERMSG BOOL __stdcall
typedef BOOL (__stdcall * MSOPFNTFCFDOIDLE)(BOOL);
#define MSOTFCFDOIDLE BOOL __stdcall
typedef VOID (__stdcall * MSOPFNTFCWIZ)(int mdWizChange, long lPrivate);
#define MSOTFCWIZ VOID __stdcall
typedef VOID (__stdcall * MSOPFNTFCBLNNOTIFY)(BOOL fEnter, int md, long hbal);
#define MSOTFCBLNNOTIFY VOID __stdcall

typedef VOID (__stdcall * PFNSHOWTIP)(HTIP);


//
// Wizards
//

#define msomdWizActInactive 0
#define msomdWizActActive       1
#define msomdWizActSuspend      2
#define msomdWizActResume       3

#define msomdWizLocalStateOn    1
#define msomdWizLocalStateOff   2
#define msomdWizShowHelp        3
#define msomdWizSuspending      4
#define msomdWizResuming        5

/* Wizard Style Bits */
#define msowsCustomTeaser       1
#define msowsNoAutoDownTeaser 2

// Bit masks for each office wizard that has an assistant for determining whether the assistant comes up with the wizard by default
typedef enum
    {
    msofcwBitDcw = 0x01 // The ODSO Data Connection Wizard
    } MSOFCW;

//
// OLE Activation
#define msomdTFCOleClientStart  0
#define msomdTFCOleClientEnd        1
#define msomdTFCOleServerStart  2
#define msomdTFCOleServerEnd        3

#endif //$[VSMSO]

//
// Alerts
//

//$[VSMSO] Moved here from msobal.h
/* Balloon Modes */
#define msomdModal    0
#define msomdAutoDown 1
#define msomdModeless 2
#define msomdUIModal  3	// WARNING! Do not use this mode unless you know
								// what the hell you're doing!
#define msomdMax		 4   // unused one, for error checking only

// ALertButtons (alb)
#define msoalbOK               0
#define msoalbOKCancel         1
#define msoalbAbortRetryIgnore 2
#define msoalbYesNoCancel      3
#define msoalbYesNo            4
#define msoalbRetryCancel      5
#define msoalbYesAllNoCancel   6
#define msoalbNone               7

#define msoalbCustomOne 8
#define msoalbCustomTwo 9
#define msoalbCustomThree 10
#define msoalbCustomFour 11
#define msoalbCustomFive 12

#define FCustomAlb(alb) ((alb)>=msoalbCustomOne && (alb)<=msoalbCustomFive)

#define calbMax				 13

// Bitwise or'd into MsoAlertWtz arg to produce a "don't show this again" checkbox
#define MB_DONTSHOWAGAIN    0x10000000

// Bitwise or'd into MsoAlertWtz return value to indicate "don't show this again" was checked
#define IDDONTSHOWAGAIN		0x1000

// ALertiCons	(alc)
#define	msoalcNoIcon			0
#define	msoalcCritical			1
#define	msoalcQuery				2
#define	msoalcWarning			3
#define	msoalcInfo				4
#define	msoalcMax				5

#define msoalcStop      msoalcWarning
#define msoalcNote      msoalcInfo
#define msoalcCaution   msoalcQuery

// ALertDefaults (ald)
#define msoaldFirst 0
#define msoaldSecond    1
#define msoaldThird 2
#define msoaldFourth 3
#define msoaldFifth     4

// ALertCancelButtons (alq)
#define msoalqDefault -1    // just "do the right thing"
#define msoalqFirst 0
#define msoalqSecond    1
#define msoalqThird 2
#define msoalqFourth 3
#define msoalqFifth     4

#define msoalidOK           IDOK
#define msoalidCancel   IDCANCEL
#define msoalidAbort        IDABORT
#define msoalidRetry        IDRETRY
#define msoalidIgnore   IDIGNORE
#define msoalidYes      IDYES
#define msoalidNo           IDNO
#define msoalidYesAll   8

#define msoalidCustomOne    101
#define msoalidCustomTwo    102
#define msoalidCustomThree  103
#define msoalidCustomFour   104
#define msoalidCustomFive   105

//$[VSMSO_BEGIN] Moved here from msobal.h
/* Don't change these (at least from OK to Ignore unless you fix mpibtnID in fcalert.cpp */
#define msoibtnMin		  msoibtnCBoxFirst	// Minimum button ID.
#define msoibtnCBoxFirst  (msoibtnCBoxLast - msoicbxMax + 1)
#define msoibtnCBoxLast	  (msoibtnUserFirst - 1)
#define msoibtnUserFirst  (msoibtnUserLast - msoisbSBSMax + 1)
#define msoibtnUserLast	  -21

/* The following three ibtn values are not actually used by the Windows   */
/* balloon code. They are included here for compatibility with MacOffice. */
#define msoibtnSaveAll	  -20	
#define msoibtnSave		  -19
#define msoibtnDontSave	  -18
/* End of MacOffice ibtn Values */

#define msoibtnPrev		  -17
#define msoibtnMore		  -16			  
#define msoibtnYesToAll	  -15
#define msoibtnOptions	  -14
#define msoibtnTips		  -13
#define msoibtnClose		  -12
#define msoibtnSnooze     -11
#define msoibtnSearch	  -10
#define msoibtnIgnore		-9
#define msoibtnAbort			-8
#define msoibtnRetry			-7
#define msoibtnNext			-6
#define msoibtnBack			-5
#define msoibtnNo				-4
#define msoibtnYes			-3
#define msoibtnCancel		-2
#define msoibtnOK				-1
#define msoibtnNull			 0
#define msoibtnChoiceFirst	 1	// Minimum ID for a "choice" button.
#define msoibtnCBMPFirst    1000
//$[VSMSO_END]

#if 0 //$[VSMSO]
/*--------------------------------------------------------------------------*/
//
// Watsonized Alerts
//

// structures used when calling watsonized alerts

// watsonized alert bucketing parameters
typedef struct
{
    DWORD iWAlertID;
    DWORD dwLastShipAssertTag;
    HRESULT hrError;
    int grfAbt;
} WABP;


// watsonized alerts response
typedef struct
{
    BOOL fCorpResponse;     // true if this is a corporate response
    union
        {
        struct
            {
            BYTE *pbResponse;
            int cbResponse;
            };
        WCHAR *wzErrorText;
        };
    BOOL fShowErrorMsg;     // NOTE:  If fShowErrorMsg is TRUE, then wzErrorText is used, otherwise pbResponse/cbResponse
    BOOL fOpenExpanded;	// true if we should open the infotext expanded
} WARESP;

// watsonized alert response info
typedef struct
{
    WARESP warespMS;    // MS response
    WARESP warespCorp;  // corporate response
} WARI;

// watsonized alert dialog data
typedef struct _wadd
{
    WABP wabp;  // bucketing params
    WARI wari;  // response info

    int grfAbtReg;  // grfAbt we picked up from watsonrc.dll
    DWORD iInfoText;    // response ID from watsonrc.dll, -1 if none
    DWORD dwOptions;        // the high order bits we pulled out of watsonrc.dll

    WCHAR *wzMoreInfo;	// text for the More Information button
    WCHAR *wzLessInfo;	// text for the Less Information button
    WCHAR *wzHelpWindow;	// text for the Open in Help Window button
} WADD;
#endif //$[VSMSO]

// bit flags for grfAbt
#define fabtAlertID     1
#define fabtShipAssert  2
#define fabtHResult     4

#if 0 //$[VSMSO]
// maximum number of IDs per application
#define msowaidMax      (100000L)

// this constant is used when no alert ID is passed into LDoAlert*
#define msowaidNoID		(99999L)

// starting Watsonized Alert ID for each application (no overlap)
#define msowaidExcel        (100000L)
#define msowaidWord         (200000L)
#define msowaidOutlook      (300000L)
#define msowaidPowerpoint   (400000L)
#define msowaidAccess       (500000L)
#define msowaidGraph        (600000L)
#define msowaidMso			(700000L)
#define msowaidConferencing (800000L)
#define msowaidProject (900000L)
#define msowaidOneNote		(1000000L)
#define msowaidPublisher		(1100000L)

#define msowaidFirst			msowaidExcel
// when new applications are added, change this
#define msowaidLast			msowaidPublisher		

// exported APIs used with Watsonized Alerts

MSOAPI_(void) MsoSendWatsonAlertSqmStream(WABP *);
MSOAPI_(BOOL) MsoFRetrieveWatsonAlertData(WADD *pwadd, DWORD iWAlertID, DWORD dwLastShipAssertTag, HRESULT hrError, int grfAbt);
MSOAPI_(DWORD) MsoDwLastShipAssertTag();
MSOAPI_(void) MsoLaunchDWForWatsonAlert(WADD *pwadd, const WCHAR *pwzAlertText);
MSOAPI_(void) MsoFreePwadd(WADD *pwadd);
MSOAPI_(BOOL) MsoFGetLastWAlertHR(HRESULT *phrError);
MSOAPI_(void) MsoSetLastWAlertHR(HRESULT hrError);
MSOAPI_(void) MsoResetLastWAlertHR();

typedef BOOL (MSOPRIVCALLTYPE *MSOGETALERTUSERDOCS)(WCHAR *wzUserDocs, int cchMax);
MSOAPI_(void) MsoSetWzAlertUserDocsCallback(MSOGETALERTUSERDOCS pfn);

/*--------------------------------------------------------------------------*/
//
// Assistant Options
//


typedef struct
    {
    BOOL fAssistWithHelp;   // bring up the asst when F1 or Cmd-? is pressed
    BOOL fAssistWithWizards;    // master setting for whether asst helps with wizards
    BOOL fAssistWithAlerts; // send alerts through the assistant when it's up

    BOOL fMoveAway;         // should asst move when in the way
    BOOL fAWContextHelp;        // should the Answer Wizard do context help
    BOOL fSound;                // should the assistant make noise

    BOOL fUnused;               // BUG: remove this on interface change day
    BOOL fAllHelpInVBE;     // search product and programmability help while in VBE

    BOOL fFeatureTips;      // show tips about using features more effectively
    BOOL fMouseTips;            // show tips about using the mouse more effectively
    BOOL fKeyboardTips;     // show tips about keyboard shortcuts
    BOOL fHighPriTipsOnly;  // only show high priority tips
    BOOL fShowTipOfTheDay;  // show tip of the day at startup

    } AOP;


// This is the interface to the Office Assistantfor designer.
#undef  INTERFACE
#define INTERFACE  IMsoAsstButtonUserHelper

DECLARE_INTERFACE(IMsoAsstButtonUserHelper)
{

    MSOMETHOD_(BOOL, FEnabled) (THIS_ int tcid) PURE;
    MSOMETHOD_(BOOL, FChangeLabel) (THIS_ HMSOINST hmsoinst,int tcid,
                                    const WCHAR *wtzOld, _Out_cap_(257) WCHAR *wtzNew) PURE;
    MSOMETHOD_(BOOL, FGetQuickTip) ( THIS_ int tcid,
                                    _Out_cap_(MAX_PATH+1) WCHAR *wtzHelpFile, int *piHelpContext) PURE;
    MSOMETHOD_(BOOL, FClick) ( THIS_ int tcid,int *ptbbs, int grf) PURE;
    MSOMETHOD_(UINT,  Tbbs) ( THIS_ int tcid) PURE;
    MSOMETHOD_(BOOL, FGetString) ( THIS_ HMSOINST hmsoinst,int tcid,int tbstr,_Out_cap_(257) WCHAR *wtz) PURE;
    MSOMETHOD_(BOOL, FWtzGetAccelerator)(THIS_  int tcid, _Out_ WCHAR *wtz) PURE;
    MSOMETHOD_(int,TbcuOleUsage)(THIS_  int tcid) PURE;

};



/*--------------------------------------------------------------------------*/

// This is the interface to the Office Assistant.

#undef  INTERFACE
#define INTERFACE  IMsoTFC

DECLARE_INTERFACE(IMsoTFC)
{

   //*** FDebugMessage method ***
   MSODEBUGMETHOD

    // ITFC methods
    MSOMETHOD_(BOOL, FInitialize) (THIS_ SYSTEMTIME *ptm) PURE;
    MSOMETHOD_(BOOL, FInstalled)(THIS) PURE;
    MSOMETHOD_(BOOL, FIsAssistantMsg) (THIS_ MSG *pmsg) PURE;

    MSOMETHOD_(void, SetCurrentTFC)(THIS) PURE;

    MSOMETHOD_(void, Show) (THIS_ int fcs) PURE;
    MSOMETHOD_(void, Move) (THIS_ int xLeft, int yTop) PURE;
    MSOMETHOD_(void, Obscure) (THIS_ int x, int y) PURE;
    MSOMETHOD_(void, Unobscure)(THIS) PURE;
    MSOMETHOD_(BOOL, FMoveAway) (THIS_ RECT *prcAvoid, RECT *prcContainer) PURE;
    MSOMETHOD_(BOOL, FGetRect) (THIS_ RECT *prc) PURE;
    MSOMETHOD_(int, FcsGet) (THIS) PURE;

    MSOMETHOD_(void, Help) (THIS) PURE;

    MSOMETHOD_(long, LStartWizard) (THIS_ int fca, int mdWizAct, int ws,
                                              MSOPFNTFCWIZ pfnWiz, long lPrivate,
                                              RECT *prcAvoid) PURE;
    MSOMETHOD_(void, EndWizard) (THIS_ long lWizID, BOOL fSuccess, int fca) PURE;
    MSOMETHOD_(void, ActivateWizard) (THIS_ long lWizID, int fca, int mdWizAct) PURE;

    MSOMETHOD_(BOOL, FDoAnimation) (THIS_ int fca, WORD wFlags) PURE;
    MSOMETHOD_(void, StopAnimation) (THIS_ int fca) PURE;
    MSOMETHOD_(BOOL, FGetAnimation) (THIS_ int *pfca) PURE;
    MSOMETHOD_(BOOL, FDoAmbientAnim) (THIS_ int x, int y, int num, int den)
                                                PURE;
    MSOMETHOD_(VOID *, PibuGet) (THIS) PURE;

    MSOMETHOD_(VOID, DoWelcomeBalloon) (THIS_ WCHAR *wz, int hid) PURE;

    MSOMETHOD_(VOID, DoAssistantBalloon) (THIS) PURE;
    MSOMETHOD_(long, LDoBalloon)    (THIS_ MSOBC *pbc, RECT *prcAvoid, int md,
                    MSOPFNDOIBTN pfnDoIbtn, int fca) PURE;
    MSOMETHOD_(void, RefillBalloon) (THIS_ long hbal, MSOBC *pbc, MSOPFNDOIBTN pfnDoIbtn,
                                int fca) PURE;
    MSOMETHOD_(void, CloseBalloon) (THIS_ long hbal) PURE;
    /* BneGetLast - Returns the last balloon error (bne). This last error
        value describes the result of the last call to LDoBalloon. It is
        unaffected by other TFC calls even if they result in the creation and/or
        showing of balloons. */
    MSOMETHOD_(int, BneGetLast) (THIS) PURE;

    MSOMETHOD_(int, LDoAlert) (THIS_ const WCHAR *wzTitle, const WCHAR *wzText,
                                    int alb, int alc, int ald, int alq,
                                    int md, BOOL fSysAlert,
                                    const WCHAR *wzHelpFile, long lwHelpID,
                                    MSOPFNDOIBTN pfnModelessAlertProc, long lPrivate,
                                    WCHAR **rgwtzCustomButtons) PURE;

#if MODELESS_ALERTS
    MSOMETHOD_(void, CloseAlert)(THIS_ long hal) PURE;
    MSOMETHOD_(BOOL, FIsAlertMsg)(THIS_ long hal, MSG *pmsg) PURE;
#endif

    MSOMETHOD_(BOOL, FAddDefaultAWDatabase) (THIS) PURE;
    MSOMETHOD_(BOOL, FAddAWDatabase) (THIS_ WCHAR *wtzFile, BOOL fRemote) PURE;
    MSOMETHOD_(VOID, DeleteAWDatabase) (THIS_ WCHAR *wtzFile) PURE;
    MSOMETHOD_(VOID, ClearAWDatabaseList) (THIS) PURE;
    MSOMETHOD_(VOID, ResetAWDatabaseList) (THIS_ BOOL fDelayInit) PURE;
    MSOMETHOD_(void, SetAWDatabaseType) (THIS_ BOOL fProgram) PURE;

//   MSOMETHOD_(PXUTOPIC *, PpxutopicGetAW) ( THIS_ WCHAR * wzQuery, BOOL fAddApp,
//                              int cAWTopicMax, int iDbMask ) PURE;

    MSOMETHOD_(void, NotifyOleActivate) (THIS_ BOOL fActive, interface IOleInPlaceFrame * pipf ) PURE;

    MSOMETHOD_(BOOL, FDoNotification)(THIS_ MSOBC *pbc, BOOL fQueue, int fca) PURE;
    MSOMETHOD_(void, SendTip)(THIS_ HTIP htip, int ttp, int tpr) PURE;
    MSOMETHOD_(BOOL, FGetIDispatchAssistant)(THIS_ IDispatch **ppidisp) PURE;
    MSOMETHOD_(VOID, GetAssistantOptions)(THIS_ AOP *paop) PURE;
    MSOMETHOD_(BOOL, FSetAssistantFileName)(THIS_ WCHAR *wzAsstFile, int fca) PURE;

    /* MdTopBalloon and PbcTopBalloon are intended purely for debugging */
    /* purposes. They are not sanctioned for use in shipping code.      */
    MSOMETHOD_(int, MdTopBalloon) (THIS) PURE;
    MSOMETHOD_(MSOBC *, PbcTopBalloon)(THIS) PURE;

    MSOMETHOD_(VOID, EnterModalState) (THIS_ BOOL fEnter) PURE;
    MSOMETHOD_(VOID, SetBalloonNotifyProc) (THIS_ MSOPFNTFCBLNNOTIFY pfn) PURE;

    MSOMETHOD_(UTOPIC *, RgutopicGetAW) ( THIS_ WCHAR * wzQuery, BOOL fAddApp,
                                int cAWTopicMax, int iDbMask, int * pcutopic ) PURE;
    MSOMETHOD_(VOID, FreeRgutopic) (THIS_ UTOPIC * rgutopic, int cutopic) PURE;
    MSOMETHOD_(void, SetAWDatabaseSet) (THIS_ int iDatabaseSet) PURE;

    MSOMETHOD_(int, LDoAlertEx) (THIS_ const WCHAR *wzTitle, const WCHAR *wzText,
                                    int alb, int alc, int ald, int alq,
                                    int md, BOOL fSysAlert,
                                    const WCHAR *wzHelpFile, long lwHelpID,
                                    MSOPFNDOIBTN pfnModelessAlertProc, long lPrivate,
                                    WCHAR **rgwtzCustomButtons, int iAlertID, int msorid) PURE;
    MSOMETHOD_(VOID, SetHwnd) (THIS_ HWND hwnd) PURE;

    // Tip lightbulb
    MSOMETHOD_(void, SendTipEx)(THIS_ HTIP htip, int ttp, int tpr,
                                   POINT * ppt, PFNSHOWTIP pfnShowTip) PURE;
    MSOMETHOD_(void, MoveTipBulb)(THIS_ int x, int y) PURE;
    MSOMETHOD_(void, KillTipBulb)(THIS) PURE;
    MSOMETHOD_(BOOL, FGetTipBulbPosn)(THIS_ int * px, int * py) PURE;
    MSOMETHOD_(BOOL, FFullyInited)(THIS) PURE;
    MSOMETHOD_(BOOL, FInHomeBalloonFromClick) (THIS) PURE;
    MSOMETHOD_(void, EnableBalloon) (THIS_ long hbal, BOOL fEnable) PURE;

    MSOMETHOD_(IMsoAsstButtonUserHelper *, PiabuhGet) (THIS) PURE;
    MSOMETHOD_(BOOL, FWizardAssistantOn)(THIS_ MSOFCW msofcw) PURE;
    MSOMETHOD_(void, SetWizardAssistant) (THIS_ MSOFCW msofcw, BOOL fOn) PURE;
    MSOMETHOD_(BOOL, FHasFocus) (THIS) PURE;

    MSOMETHOD_(int, LDoAlertWAHr) (THIS_ const WCHAR *wzTitle, const WCHAR *wzText,
                                    int alb, int alc, int ald, int alq,
                                    int md, BOOL fSysAlert,
                                    const WCHAR *wzHelpFile, long lwHelpID,
                                    MSOPFNDOIBTN pfnModelessAlertProc, long lPrivate,
                                    WCHAR **rgwtzCustomButtons, int iAlertID, int msorid, DWORD iWAlertID, HRESULT hrError, BOOL fHResult) PURE;

};


/****************************************************************************
    The IMsoTFCUser interface provides methods required by the assistant.
****************************************************************** CMoore **/

#undef  INTERFACE
#define INTERFACE  IMsoTFCUser



DECLARE_INTERFACE(IMsoTFCUser)
{
   //*** FDebugMessage method ***
   MSODEBUGMETHOD

   // General Assistant things
    MSOMETHOD_(VOID, AssistWithWizards) (THIS_ BOOL fAssist) PURE;

    // Balloon things
    MSOMETHOD_(BOOL, FDoIdle) (THIS_ BOOL fFirstIdle) PURE;
    MSOMETHOD_(BOOL, FFilterMsg) (THIS_ LPMSG) PURE;
    MSOMETHOD_(HBITMAP, HbitmapFromRef) (THIS_ WCHAR *wz, int id) PURE;

    // Tip things
    MSOMETHOD_(BOOL, FGetTcn) (THIS_ HTIP htip, MSOTCN *ptcn) PURE;
    MSOMETHOD_(VOID, TipDiscarded) (THIS_ HTIP htip) PURE;
    MSOMETHOD_(VOID, TipShown) (THIS_ HTIP htip) PURE;
    MSOMETHOD_(VOID, ResetTips) (THIS) PURE;

    // AnswerWizard things
    MSOMETHOD_(MSOCTT *, PcttGetContext) (THIS) PURE;
    MSOMETHOD_(VOID, FreePctt) (THIS_ MSOCTT * pctt) PURE;
    MSOMETHOD_(BOOL, FDoAWHelpAction) (THIS_ int act, const WCHAR *wzHelpFile,
                                       int hid) PURE;

    // Notification things
    MSOMETHOD_(BOOL, FDoQNotifIbtn) (THIS_ MSOBC *pbc, int ibtn) PURE;

    // AnswerWizard databases location (for moving init out of boot)
    MSOMETHOD_(BOOL, FInitAnswerWizard) (THIS_ DWORD future) PURE;
};



//  Exported functions
MSOAPI_(BOOL) MsoFCreateITFC(HMSOINST hmsoinst, IMsoTFC **ppitfc,
                                IMsoTFCUser *pifcu, int idb);
MSOAPI_(BOOL) MsoFCreateITFCNoAssistant(HMSOINST hmsoinst, IMsoTFC **ppitfc,
                                IMsoTFCUser *pifcu, int idb);
MSOAPI_(BOOL) MsoFCreateITFCHwnd(HMSOINST hmsoinst, IMsoTFC **ppitfc,
                                IMsoTFCUser *pifcu, int idb, HWND hwndMain);
MSOAPI_(void) AppTurnAsstOn();

// This API is used by the Assist.DLL to find out if the current app intends to use the mso assistant
MSOAPIX_(BOOL) MsoFUsingAssistant();

#if DEBUG
MSOAPI_(void) MsoSetAWDebugOptions( int major, int minor, BOOL fOn );
MSOAPI_(void) MsoSetAWDebugFileName( const char * szFile );
MSOAPI_(BOOL) MsoFAddAWDb(const WCHAR *wtzFile, BOOL fRemote);
MSOAPI_(VOID) MsoClearAWDbList(VOID);
MSOAPI_(UTOPIC *) MsoRgutopicISearch(const WCHAR * wzTarget, const WCHAR * wzQuery,
                MSOCTT * pctt, int iDbMask, int iTopicFilter, int cTopicMax,
                BOOL fAddApp, BOOL fUserLanguage,
                BOOL * pfPhraseHelpAvail, BOOL * pfFileOk, int * pcutopic);
MSOAPI_(VOID) MsoFreeRgutopic(UTOPIC * rgutopic, int cutopic);
MSOAPI_(VOID) MsoSetShutdownTimeout(DWORD);
#endif

MSOAPI_(void) MsoDestroyITFC(IMsoTFC *pitfc);

/*  Displays an Alert through the Office assistant pitfc with title
    wzTitle and text wzText.  alb determines the button to display,
    ald determines the default button, alq tells which button is
    the cancel button, md is the modality of the alert, fSysAlert
    is true for system alert boxes, wzHelpFile and lwHelpID specify
    help file reference for the alert, rgwtzCustomButtons is an array
    of custom button names for the custom button types. */
MSOAPI_(int) LDoAlertTFC(IMsoTFC *ptfc, const WCHAR *wzTitle, const WCHAR *wzText,
    int alb, int alc, int ald,  int alq, int md, BOOL fSysAlert,
    const WCHAR *wzHelpFile, long lwHelpID, MSOPFNDOIBTN pfnModelessAlertProc,
    long lPrivate,  WCHAR **rgwtzCustomButtons);

/* Extends the above function to display an extendable alert if the
    administrator has set a registry key exists that matches iAlertID */
MSOAPI_(int) LDoAlertTFCEx(IMsoTFC *pitfc, const WCHAR *wzTitle, const WCHAR *wzText,
    int alb, int alc, int ald,  int alq, int md, BOOL fSysAlert,
    const WCHAR *wzHelpFile, long idh, MSOPFNDOIBTN pfnModelessAlertProc,
    long lPrivate, WCHAR ** rgwtzCustomButtons, int iAlertID, int msorid);

MSOAPI_(int) LDoAlertTFCWA(IMsoTFC *pitfc, const WCHAR *wzTitle, const WCHAR *wzText,
    int alb, int alc, int ald,  int alq, int md, BOOL fSysAlert,
    const WCHAR *wzHelpFile, long idh, MSOPFNDOIBTN pfnModelessAlertProc,
    long lPrivate, WCHAR ** rgwtzCustomButtons, int iAlertID, int msorid, DWORD iWAlertID);

MSOAPI_(int) LDoAlertTFCWAHr(IMsoTFC *pitfc, const WCHAR *wzTitle, const WCHAR *wzText,
    int alb, int alc, int ald,  int alq, int md, BOOL fSysAlert,
    const WCHAR *wzHelpFile, long idh, MSOPFNDOIBTN pfnModelessAlertProc,
    long lPrivate, WCHAR ** rgwtzCustomButtons, int iAlertID, int msorid, DWORD iWAlertID, HRESULT hrError, BOOL fHResult);


// mapping from alert ids to watson alert id
typedef struct
{
	int ids;	// alert IDS
	int idwa;	// watsonized alert ID
}
WALERTMAPPING;

MSOAPI_(int) MsoIdWAFromIds(int ids, const WALERTMAPPING rgWAlertMapping[], int cAlerts);

#endif //$[VSMSO]

int LDoAlertTFCWAHrEx(void *pitfc, const WCHAR *wzTitle, const WCHAR *wzText,
        int alb, int alc, int ald,  int alq, int md, BOOL fSysAlert,
        const WCHAR *wzHelpFile, long idh, MSOPFNDOIBTN pfnModelessAlertProc,
        long lPrivate, _Out_ WCHAR ** rgwtzCustomButtons, int iAlertID, int msorid, DWORD iWAlertID, HRESULT hrError, int grfAbt);

int LDoAlertTFCIds(void *pitfc, const WCHAR *wzTitle, const WCHAR *wzText,
	int alb, int alc, int ald,	int alq, int md, BOOL fSysAlert,
	const WCHAR *wzHelpFile, long idh, MSOPFNDOIBTN pfnModelessAlertProc,
	long lPrivate, _Out_ WCHAR ** rgwtzCustomButtons, int idsAlert, int msorid);

int LDoAlertTFCIdsEx(void *pitfc, const WCHAR *wzTitle, const WCHAR *wzText,
	int alb, int alc, int ald,	int alq, int md, BOOL fSysAlert,
	const WCHAR *wzHelpFile, long idh, MSOPFNDOIBTN pfnModelessAlertProc,
	long lPrivate, _Out_ WCHAR ** rgwtzCustomButtons, int idsAlert, int msorid, int iWAlertID, BOOL fFromMSO);


#if 0 //$[VSMSO]
/* ///////////////////////////////////////////////////////////////////////////
// Load sets of AW databases
// MsoFGimmeAWDatabases loads localized ones through Darwin
// MsoFCustomAWDatabases loads a list from the registry
/////////////////////////////////////////////////////////////////JJames//// */

#if 1
// old API, replace with MsoFGimmeAWDatabases
 typedef struct {
    int qfid; // language qualified component for this database, msoqcidNil to end
    int ftid; // feature that must be installed to use it, msoftidNil otherwise
    } MSOAWDATA;
MSOAPIX_(BOOL) MsoFStandardAWDatabases(const MSOAWDATA *aAWdata, BOOL fRemote);
#endif
MSOAPI_(BOOL) MsoFGimmeAWDatabases(msocidT qcid, LCID lcid, DWORD dwGimmeFlags, BOOL fRemote);
MSOAPI_(BOOL) MsoFCustomAWDatabases(int ridDatabaseHook, int ridPolicy, BOOL fRemote);

#endif //$[VSMSO]

/* ///////////////////////////////////////////////////////////////////////////
// Get help file filenames
// MsoFGetHelpFile gets it through Gimme ID
// MsoFGetHelpFileFromWz does registry or Darwin lookup to find a help file
/////////////////////////////////////////////////////////////////JJames//// */
MSOAPI_(BOOL) MsoFGetHelpFile(int qfid, _Out_z_cap_(cchMax) WCHAR *wzPath, DWORD cchMax);
MSOAPI_(BOOL) MsoFGetHelpFileFromWz(const WCHAR *wzHelp, _Out_z_cap_(cchMax) WCHAR *wzPath, DWORD cchMax);

#endif
