#pragma once

/****************************************************************************
	MsoUser.h

	Owner: DavePa
 	Copyright (c) 1994 Microsoft Corporation

	Declarations for common functions and interfaces required for apps
	to use the Office DLL.
****************************************************************************/

#ifndef MSOUSER_H
#define MSOUSER_H

#include "msodebug.h"

#include "msoiv.h" // Instrumented Version for Office9 #defines and typedef

#ifndef MSO_NO_INTERFACES
interface IMsoControlContainer;
#endif // MSO_NO_INTERFACES

/****************************************************************************
	The ISimpleUnknown interface is a variant on IUnknown which supports
	QueryInterface but not reference counts.  All objects of this type
	are owned by their primary user and freed in an object-specific way.
	Objects are allowed to extend themselves by supporting other interfaces
	(or other versions of the primary interface), but these interfaces
	cannot be freed without the knowledge and cooperation of the object's
	owner.  Hey, it's just like a good old fashioned data structure except
	now you can extend the interfaces.
****************************************************************** DAVEPA **/

#undef  INTERFACE
#define INTERFACE  ISimpleUnknown

DECLARE_INTERFACE(ISimpleUnknown)
{
	/* ISimpleUnknown's QueryInterface has the same semantics as the one in
		IUnknown, except that QI(IUnknown) succeeds if and only if the object
		also supports any real IUnknown interfaces, QI(ISimpleUnknown) always
		succeeds, and there is no implicit AddRef when an non-IUnknown-derived
		interface is requested.  If an object supports both IUnknown-derived
		and ISimpleUnknown-derived interfaces, then it must implement a
		reference count, but all active ISimpleUnknown-derived interfaces count
		as a single reference count. */
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;
};


/****************************************************************************
	HMSOINST is an opaque reference to an Office instance record.  Each
	thread of each EXE or DLL that uses Office must call MsoFInitOffice
	to init Office and get an HMSOINST.
****************************************************************** DAVEPA **/
#ifndef HMSOINST
typedef struct MSOINST *HMSOINST;  // MSOINST is defined only within Office
#endif

#include "msotc.h"

/****************************************************************************
	The IMsoUser interface has methods for Office to call back to the
	app for general information that is common across Office features.
****************************************************************** DAVEPA **/

#undef  INTERFACE
#define INTERFACE  IMsoUser

enum {
	msofmGrowZone = 1,
};

#define MSOCCHMAXSHORTAPPID 15
enum {
	msocchMaxShortAppId = MSOCCHMAXSHORTAPPID
};


/*	dlgType sent to IMsoUser::FPrepareForDialog. Modal dialogs have LSB 0.*/
#define msodlgWindowsModal			0x00000000
#define msodlgWindowsModeless		0x00000001
#define msodlgSdmModal				0x00000010
#define msodlgSdmModeless			0x00000011
#define msodlgUIModalWinModeless	0x00000101
#define msodlgUIModalSdmModeless	0x00000111
#define msodlgSdmModalNWaitAct	0x00001000

// Help Pane options for MsoFStartHelpPane's HELPPANE struct
#define msoHelpPaneDefaults			0x00000000 // accept all defaults
#define msoHelpPaneNoAsstCenter		0x00000001 // no Assistance Center link
#define msoHelpPaneNoTraining		0x00000002 // no Training link
#define msoHelpPaneNoUpdates		0x00000008 // no Product Updates link
#define msoHelpPaneNoContactUs		0x00000010 // no Contact Us link
#define msoHelpPaneNoWhatsNew		0x00000020 // no Accessibility Help option
#define msoHelpPaneNoSpotLight		0x00000040 // no Spotlight feature
#define msoHelpPaneNoStartHelpPane	0x00000080 // skip the start help pane altogether and go right to the TOC pane
#define msoHelpPaneNoNativeTBS		0x00000100 // WARNING: DO NOT USE THIS FLAG UNLESS YOU KNOW WHAT YOU ARE DOING! The calling app has no natice toobar set ptr so this flag indicates that the help pane should "borrow" one from the owner app
#define msoHelpPaneShowCategory		0x00000200 // If set, the wzCaegory buffer in HELPPANEDATA must also be set indicating to show the TOC pane with a specific category already showing
#define msoHelpPaneVBE				0x00000400 // If set, Load only the VB TOCs from the Manifest (use only for VBE application)
#define msoHelpPaneRefreshTOC		0x00000800 // If set, flush the current contents of the TOC pane when shown, and reload it
#define msoHelpPaneDenyOnlineTOC	0x00001000 // If set, do not initiate an Online TOC query even if the online content setting is turned on (Visio11 needs this)

// Note, always memset init the HELPPANE struct before using it
// Also, the wzCategory string should be a '/' delimited category asset id string
// like this: CH000000011033/CH000000231033/CH000000171033, etc., representing the
// assets ids for all parent categories of the target category on down to the
// target category itself, in order from root to leaf.
typedef struct _HELPPANEDATA
{
	DWORD qfidTOCManifest; // gimme id for xml file containing all offline tocs that could be loaded into TOC help pane along with data about they should be loaded
	DWORD dwFlags; // if not set, then msoHelpPaneDefaults is implied
	void* pitbs;	// toolbarset pointer, if not set, it is obtained via the hwnd passed into MsoFStartHelpPane	
	WCHAR wzCategory[MAX_PATH]; // if set, the dwFlags should also have the msoHelpPaneShowCategory flag set. Used only for a SRP breadcrumb expansion.
	WCHAR wzQueryID[MAX_PATH]; // query id of a breadcrumb expansion from the Search Results Pane (SRP). Used only by the SRP when the msoHelpPaneShowCategory is set.  
}HELPPANEDATA;

typedef HELPPANEDATA* LPHELPPANEDATA;

// Notification codes for FNotifyAction methods
// Names containing 'Query' indicate app's return value is sought.  Other
// values are strictly to notify app.
enum
	{
	msonaStartHelpMode = 0,         // User entered Quick tip mode (Shift-F1).  App should update any internal state
	msonaEndHelpMode,               // Quick tip was displayed.  App should restore cursor.
	msonaBeforePaletteRealize,      // Office is going to realize one or more palettes, see comment below
	msonaQueryDisablePip,           // Gives the app the chance to refuse a .PIP file, should it not want one.
	msonaQueryInsertPicture,        // Asks the app if we can insert a picture
	msonaQueryAcbAware,             // Asks the app if it's fully aware of the Active ClipBoard
	msonaAddinBeforeConnect,
	msonaAddinAfterConnect,
	msonaAddinBeforeDisconnect,
	msonaAddinAfterDisconnect,
	msonaRegisterServices,          // Register NETUI accessible services
	msonaBlockWMPaintPumping,       // Block WM_PAINT pumping furing COM OOP calls
	msonaUnblockWMPaintPumping,     // Unblock WM_PAINT pumping furing COM OOP calls
	msonaHideDRMUI,					// Hides (and effectively disables) all DRM UI
	};

// Subsystem classifications for FEmNotifyAction methods
enum
	{
	msoemssToolbar = 0,			// Command bars
	msoemssAppEm,					// App Event Monitor
	msoemssTip,						// Possible future feature: tip interface
	msoemssTimer,					// Possible future feature: timer notify
	};

/* About msonaBeforePaletteRealize:

	Office will call FNotifyAction(msonaBeforePaletteRealize) to let the app
	it's going to realize a palette. The app should start palette management
	if it has delayed doing so until it absolutely needs to.

	The app should select and realize a palette, and from now on, should
	respond to palette messages WM_QUERYNEWPALETTE and WM_PALETTECHANGED.
*/


DECLARE_INTERFACE(IMsoUser)
{
   /* Debuging interfacing for this interface */
   MSODEBUGMETHOD

	/* Return an IDispatch object for the Application object in 'ppidisp'
		Return fSuccess. */
	MSOMETHOD_(BOOL, FGetIDispatchApp) (THIS_ IDispatch **ppidisp) PURE;

	/* Return the long representing the application, as required by the
		"Creator" method of VBA objects. */
	MSOMETHOD_(LONG, LAppCreatorCode) (THIS) PURE;		//  REVIEW:  PETERO:  Is this MAC only?

	/* If the host does not support running macros then return FALSE,
		else check the macro reference in wtzMacro, which is in a 257 char buffer,
		for validity, modify it in-place if desired, and return TRUE if valid.
		The object trying to attach the macro, if any, is given by 'pisu'.
		The format of macro references is defined by the host, but the typical
		simple case would be the name of a VBA Sub.  The host may delay
		expensive validation checks until FRunMacro as desired. */
	MSOMETHOD_(BOOL, FCheckMacro) (_In_z_ THIS_ WCHAR *wtzMacro, _Out_ ISimpleUnknown *pisu) PURE;

	/* Run the macro given by the reference wtz (which has been checked for
		validity by FCheckMacro).  The object to which the macro is attached,
		if any, is given by 'pisu'.  Return TRUE if successful (FALSE if the
		host does not support running macros). */
	MSOMETHOD_(BOOL, FRunMacro) (THIS_ const WCHAR *wtzMacro, ISimpleUnknown *pisu,
										 VARIANT *pvarResult, VARIANT *rgvar,
										 int cvar) PURE;

	/* When a low memory condition occurs this callback method will be invoked.  The
		Application should free up cbBytesNeeded or more if it can.  Return back the
		actual number of bytes that were freed. */
	MSOMETHOD_(int, CbFreeMem) (THIS_ int cbBytesNeeded, int msofm) PURE;

	/* Office will call this in deciding whether or not to do certain actions
		that require OLE. */
	MSOMETHOD_(BOOL, FIsOleStarted) (THIS) PURE;

	/* Office will call this in deciding whether or not to do certain actions
		that require OLE. If the Application supports delayed OLE initialization
		and OLE has not been started, try to start OLE now.  Office makes no
		guarantee that it will cache the value returned here, so this may be
		called even after OLE has been started. */
	MSOMETHOD_(BOOL, FStartOle) (THIS) PURE;
	/* If a Picture Container is being created Office will call back to the IMsoUser
		to fill the Picture Container with control(s). */
	// TODO: TCoon unsigned int should be UCBK_SDM
	MSOMETHOD_(BOOL, FFillPictureContainer) (THIS_ interface IMsoControlContainer *picc,
															unsigned int tmc, unsigned int wBtn,
															BOOL *pfStop, int *pdx, int *pdy) PURE;
	/* The app should pass thru the parameters to WinHelp or the equivalent
		on the Mac */
	MSOMETHOD_(void, CallHelp)(THIS_ HWND hwnd, const WCHAR *wzHelpFile,
			UINT uCommand, DWORD dwData) PURE;
	// WHAT IS THIS?
	/* The init call to initialize sdm. Get called when first sdm
	   dialog needs to come up. */
	MSOMETHOD_(BOOL, FInitDialog)(THIS) PURE;

	/* AutoCorrect functions. Used to inegrate this feature with the apps
		undo functionality and extended AC functionality in Word. */
	MSOMETHOD_(void, ACRecordVars)(THIS_ DWORD dwVars) PURE;
	MSOMETHOD_(BOOL, ACFFullService)(THIS) PURE;
	MSOMETHOD_(void, ACRecordRepl)(THIS_ int, _In_ WCHAR *wzFrom, _Out_ WCHAR *wzTo) PURE;
	MSOMETHOD_(void, ACAdjustAC)(THIS_ int iwz, int idiwz) PURE;

	/* Return the CLSID of the application */
	MSOMETHOD_(void, GetAppClsid) (THIS_ LPCLSID *) PURE;

	/* Before and After doing a sdm dialog, call back to the application for
		them to do their own init and cleanup.
		The dlg parameter is a bitmap flags defined here as msodlgXXXX
		*/
 	MSOMETHOD_(BOOL, FPrepareForDialog) (THIS_ void **ppvDlg, int dlgType) PURE;
 	MSOMETHOD_(void, CleanupFromDialog) (THIS_ void *pvDlg) PURE;

	// Applications must provide a short (max 15 char + '\0') string which
	// identifies the application.  This string is used as the application ID
	// with ODMA.  This string may be displayed to the user, so it should be
	// localized.  But strings should be chosen so that localized versions
	// can often use the same string.  (For example, "MS Excel" would be a
	// good string for Excel to use with most Western-language versions.)  If
	// the file format changes for a localized version (eg. for Far East or
	// bi-di versions), a different string should be used for the localized
	// versions whose file format is different.  (It is assumed that all
	// versions with the same localized string can read each other's files.)
	// The application should copy the string into the buffer provided.
	// This string cannot begin with a digit.  The application can assume
	// that wzShortAppId points to a buffer which can hold msocchMaxShortAppId
	// Unicode characters plus a terminating '\0' character.
	// If you have questions, contact erikhan.
	MSOMETHOD_(void, GetWzShortAppId) (THIS_ _Out_z_cap_(MSOCCHMAXSHORTAPPID) WCHAR *wzShortAppId) PURE;

	MSOMETHOD_(void, GetStickyDialogInfo) (THIS_ int hidDlg, POINT *ppt) PURE;
	MSOMETHOD_(void, SetPointStickyDialog) (THIS_ int hidDlg, POINT *ppt) PURE;

	/* Called before command bars start tracking, and after they stop. Note
		that this will be called even in the HMenu cases, and on the Mac.
		Also, when real command bars start tracking, you are called on
		OnComponentActivate by the Component Manager. Make sure you know which
		callback you want to use.
		This callback is used by Excel to remove/put back a keyboard change they
		have on the Mac. */
	MSOMETHOD_(void, OnToolbarTrack) (THIS_ BOOL fStart) PURE;

	/* Notification that the action given by 'na' occurred.
		Return TRUE if the
		notification was processed.
	*/
	MSOMETHOD_(BOOL, FNotifyAction) (THIS_ int na) PURE;

	// TODO(JBelt): this callback is obsolete
	/* Called back by the Office Darwin layer to let the app hook up its additional
		Darwin tables. The MSOTCFCF structure is explained in detail in msotc.h.
		Fill pfcf with the structure.
		You will be called on this API the very first time Office encounters a
		file id which is outside its scope. This can't happen unless you called
		MsoFGimmeFile or one of its friends with such an id in the first place.
		Non-Darwinized apps can just do nothing here. */
	MSOMETHOD_(void, HookDarwinTables) (THIS_ MSOTCFCF *pfcf) PURE;

	/*  Handle all event monitor notifications.
		There was an action of interest to the event monitor in Office, such as
		toolbar	activity.  The 'subsystem' in which the action occurs is
		indicated by emss.  na is a subsystem-specific notify action
		identifier.  A negative valued na indicates a pre-action notification
		to the event monitor.  Not all events generate a pre-action, but all do
		generate a post-action.  Arguments are packed into the structure at
		pvArgs, in a subsystem and na-specific fashion.  pvArgs is maintained
		(i.e. allocated and freed, if necessary) entirely on the Office side.
		ppvEmNotify is provided for app-side communication between pre- and
		post-action notifications.  ppvEmNotify is maintained entirely by the
		application.
		Return TRUE from FEmNotifyAction if the notification was processed.
		Currently, the return value is only relevant in the case where TRUE
		is returned from a pre-action notification, in which case no
		post-action notification is sent for that event.
	*/
	MSOMETHOD_(BOOL, FEmNotifyAction) (THIS_ int emss, int na,
										void **ppvEmNotify, void *pvArgs) PURE;

	/* Called by an office for a button customized to be hyperlink passing the
	   mode in which the App should open the hyperlink provided in pwzSource
	   Return TRUE if app opened Hyperlink for the given mode
	          FALSE if hyperlink could not be opened/app doesn't care about it */
	MSOMETHOD_(BOOL, FOpenHyperlink) (THIS_ LPCWSTR pwzSource,LPCWSTR pwzLocation,
									  DWORD grfwtbnt,int mode) PURE;
};


/****************************************************************************
	IMsoUser10 is an Office10 extension of IMsoUser interface

	Office code cannot assume that clients implement IMsoUser10 interface.

****************************************************************** IgorZ **/

#undef  INTERFACE
#define INTERFACE  IMsoUser10

DECLARE_INTERFACE(IMsoUser10)
{
	MSOMETHOD_(HRESULT, HrOnMsoFInitOffice)(
		HWND hwndMain,
		HINSTANCE hinstClient,
		IMsoUser *piuser,
		const WCHAR *wzHostName,
		HMSOINST *phinst,
		BOOL *pfHandled) PURE;

	MSOMETHOD_(void, OnMsoUninitOffice)(HMSOINST hinst, BOOL *pfHandled) PURE;
};


MSOAPI_(BOOL) MsoFSetInstIMsoUser10(HMSOINST hinst, IMsoUser10 *pUser10);


// NOTE: Another copy of this definition is in msosdm.h
#ifndef PFNFFillPictureContainer
typedef BOOL (MSOSTDAPICALLTYPE *PFNFFillPictureContainer) (interface IMsoControlContainer *picc,
														unsigned int tmc, unsigned int wBtn,
														BOOL *pfStop, int *pdx, int *pdy);
#endif
#ifndef PFNFFillPictureContainerEx
typedef BOOL (MSOSTDAPICALLTYPE *PFNFFillPictureContainerEx) (interface IMsoControlContainer *picc,
														unsigned int tmc, unsigned int wBtn,
														BOOL *pfStop, int *pdx, int *pdy, UINT *pufFlags);
#endif



// What does an application do when it needs mso to call it back sometime?
// It registers a callback, of course.  We have callbacks all over the place
// and it's about time they started coming together.  Here's a mechanism
// for registering a callback in a common way.
//
// First, the callback of interest is identified by an msocb constant.  The app
// determines which callback(s) it wants to register, and calls MsoFSetCallback.
// This will return the previously registered callback for that msocb.
//
// The callback signature should be defined for each msocb.
//
// --brianhi

typedef void (MSOSTDAPICALLTYPE * PFNGENERICCALLBACK)(void);

enum
{
	// msocbAddinGetIDispatch:  The addins object uses this callback to obtain
	// an IDispatch object from a host application.

	msocbAddinGetIDispatch = 0,		// use PFNADDINGETIDISPATCH

	// Add new callback types here

	msocbCallbackCount
};

typedef IDispatch * (MSOSTDAPICALLTYPE * PFNADDINGETIDISPATCH)(const WCHAR * pwzAddinPath);


MSOAPI_(PFNGENERICCALLBACK) MsoPfnSetCallback(UINT msocb, PFNGENERICCALLBACK pfn);	// returns previous callback
MSOAPI_(PFNGENERICCALLBACK) MsoPfnGetCallback(UINT msocb);



/*****************************************************************************
	Registry structure for initing MSO for MsoFLangChanged
*****************************************************************************/
typedef struct _MSOREGLANG
{
	int msoridAppRegistryLang;	// For ORAPI apps
	WCHAR* pwzAppRegistryLang;	// If above 0, used to get registry entry (FP)
}MSOREGLANG;

#if DEBUG

/*****************************************************************************
	Block Entry structure for Memory Checking
*****************************************************************************/
typedef struct _MSOBE
{
	void* hp;
	int bt;
	unsigned cb;
	BOOL fAllocHasSize;
	HMSOINST pinst;
}MSOBE;


/****************************************************************************
	The IMsoDebugUser interface has Debug methods for Office to call back
   to the app for debugging information that is common across Office features.
****************************************************************** JIMMUR **/

#undef  INTERFACE
#define INTERFACE  IMsoDebugUser

DECLARE_INTERFACE(IMsoDebugUser)
{
   /* Call the MsoFSaveBe API for all of the structures in this application
		so that leak detection can be preformed.  If this function returns
		FALSE the memory check will be aborted. The lparam parameter if the
		same lparam value passed to the MsoFChkMem API.  This parameter should
		in turn be passed to the MsoFSaveBe API which this method should call
		to write out its stuctures. */
   MSOMETHOD_(BOOL, FWriteBe) (THIS_ LPARAM) PURE;

   /* This callback allows the application to abort an on going memory check.
	   If this function return TRUE the memory check will be aborted.
		If FALSE then the memory check will continue.  The application should
		check its message queue to determine if the memory check should
		continue.  The lparam paramater if the same lparam value passed to the
		MsoFChkMem API.  This allows the application to supply some context if
		it is required. */
   MSOMETHOD_(BOOL, FCheckAbort) (THIS_ LPARAM) PURE;

   /* This callback is called when duplicate items are  found in the heap.
      This provides a way for an applications to manage its referenced counted
		items.  The prgbe parameter is a pointer to the array of MSOBE records. The
		ibe parameter is the current index into that array.  The cbe parameter
		is the count of BEs in the array.  This method should look at the MSOBE in
		question and return back the next index that should checked.  A value of
		0 for the return value will designate that an error has occured.*/
   MSOMETHOD_(int, IbeCheckItem) (THIS_ LPARAM lParam, MSOBE *prgbe, int ibe, int cbe) PURE;

	/* This call back is used to aquire the strigstring name of a Bt. This is used
		when an error occurs during a memory integrity check.  Returning FALSE means
		that there is no string.*/
	MSOMETHOD_(BOOL, FGetSzForBt) (THIS_ LPARAM lParam, MSOBE *pbe, int *pcbsz,
												char **ppszbt) PURE;

	/* This callback is used to signal to the application that an assert is
		about to come up.  szTitle is the title of the assert, and szMsg is the
		message to be displayed in the assert, pmb contains the messagebox
		flags that will be used for the assert.  Return a MessageBox return code
		(IDABORT, IDRETRY, IDIGNORE) to stop the current assert processing and
		simulate the given return behavior.  Returns 0 to proceed with default
		assert processing.  The messagebox type can be changed by modifying
		the MB at *pmb.  iaso contains the type of assert being performed */
	MSOMETHOD_(int, PreAssert) (THIS_ int iaso, char* szTitle, char* szMsg, UINT* pmb) PURE;

	/* This callback is used to signal to the application that an assert has
		gone away.  id is the MessageBox return code for the assert.  The return
		value is used to modify the MessageBox return code behavior of the
		assert handler */
	MSOMETHOD_(int, PostAssert) (THIS_ int id) PURE;
};

MSOAPI_(BOOL) MsoFWriteHMSOINSTBe(LPARAM lParam, HMSOINST hinst);
#endif // DEBUG


/****************************************************************************
	Initialization of the Office DLL
****************************************************************************/

/* Initialize the Office DLL.  Each thread of each EXE or DLL using the
	Office DLL must call this function.  On Windows, 'hwndMain' is the hwnd of
	the app's main window, and is used to detect context switches to other
	Office apps, and to send RPC-styles messages from one office dll to another.
	On the Mac, this used to establish window ownership (for WLM apps), and can
	be NULL for non-WLM apps.  The 'hinst' is the HINSTANCE of
	the EXE or DLL.  The interface 'piuser' must implement the IMsoUser
	interface for this use of Office.  wzHostName is a pointer to the short name
	of the host to be used in menu item text. It must be no longer than 32
	characters including the null terminator.
	The HMSOINST instance reference
	for this use of Office is returned in 'phinst'.  Return fSuccess. */
MSOAPI_(BOOL) MsoFInitOffice(HWND hwndMain, HINSTANCE hinstClient,
									  IMsoUser *piuser, const WCHAR *wzHostName,
									  HMSOINST *phinst);

/* As above, but establishes a app specific registry entry to check an apps last
	UI Language.  This is compared to the current UI lang and can then correctly
	tell the app and COM addins when the lang has changed under it. */
MSOAPI_(BOOL) MsoFInitOfficeEx(HWND hwndMain, HINSTANCE hinstClient,
									  IMsoUser *piuser, const WCHAR *wzHostName,
									  HMSOINST *phinst, MSOREGLANG* pMLRApp);

/*  MsoFOfficeInitialized
	Returns if the loaded mso.dll is in an initialized state.  Allows an mso client to enable
	functionality only when someone has fully inited mso.  */
MSOAPI_(BOOL) MsoFOfficeInitialized();

/* Uninitialize the Office DLL given the HMSOINST as returned by
	MsoFInitOffice.  The 'hinst' is no longer valid after this call. */
MSOAPI_(void) MsoUninitOffice(HMSOINST hinst);

/* Load and register the Office OLE Automation Type Library by searching
	for the appropriate resource or file (don't use existing registry entries).
	Return typelib in ppitl or just register and release if ppitl is NULL.
	Return HRESULT returned	from LoadTypeLib/RegisterTypeLib. */
MSOAPI_(HRESULT) MsoHrLoadTypeLib(ITypeLib **ppitl);

/* Register everything that Office needs in the registry for a normal user
	setup (e.g. typelib, proxy interfaces).  Return NOERROR or an HRESULT
	error code. */
MSOAPI_(HRESULT) MsoHrRegisterAll();

/* Unregister anything that is safe and easy to unregister.
	Return NOERROR or an HRESULT error code. */
MSOAPIX_(HRESULT) MsoHrUnregisterAll();

/* Reset the hwndMain of the hinst to the passed in hwndMain.  -- Word::Stevera */
MSOAPI_(BOOL) MsoFSetInstHwndMain(HMSOINST hinst, HWND hwndMain);

/* Apps can call this to get an IDispatch interface to the Answer Wizard object. */
MSOAPI_(BOOL) MsoFGetIDispatchAnswerWizard(HMSOINST hinst, IDispatch **ppidisp);

#if DEBUG
	/* Add the IMsoDebugUser interface to the HMSOINST instance reference.
	   Return fSuccess. */
	MSOAPI_(BOOL) MsoFSetDebugInterface(HMSOINST hinst, IMsoDebugUser *piodu);
	MSOAPI_(BOOL) MsoFGetDebugInterface(HMSOINST hinst, IMsoDebugUser **ppiodu);

#endif

/****************************************************************************
	Other APIs of global interest
****************************************************************************/

/* A generic implementation of QueryInterface for an object given by pisu
	with a single ISimpleUnknown-derived interface given by riidObj.
	Succeeds only if riidQuery == riidObj or ISimpleUnknown.
	Returns NOERROR and pisu in *ppvObj if success, else E_NOINTERFACE. */
MSOAPI_(HRESULT) MsoHrSimpleQueryInterface(ISimpleUnknown *pisu,
							REFIID riidObj, REFIID riidQuery, void **ppvObj);

/* Like MsoHrSimpleQueryInterface except succeeds for either riidObj1
	or riidObj2, returning pisu in both cases and therefore useful for
	inherited interfaces. */
MSOAPI_(HRESULT) MsoHrSimpleQueryInterface2(ISimpleUnknown *pisu,
							REFIID riidObj1, REFIID riidObj2, REFIID riidQuery,
							void **ppvObj);

/* This message filter is called for EVERY message the host app receives.
	If the procedure processes it should return TRUE otw FALSE. */
MSOAPI_(BOOL) FHandledLimeMsg(MSG *pmsg);


/*************************************************************************
	MSOGV -- Generic Value

	Currently we have a bunch of fields in Office-defined structures
	with names like pvClient, pvDgs, etc.  These are all declared as
	void *'s, but really they're just for the user of Office to stick
	some data in an Office structure.

	The problem with using void * and calling these fields pvFoo is that
	people keep assuming that you could legitimately compare them against
	NULL and draw some conclusion (like that you didn't need to call the
	host back to free	stuff).  This tended to break hosts who were storing
	indices in these fields.

	So I invented "generic value" (great name, huh?)  Variables of this
	type are named gvFoo.  Almost by definition, there is NO gvNil.

	This type will always be unsigned and always big enough to contain
	either a uint or a pointer.  We don't promise that this stays the
	same length forever, so don't go saving them in files.
************************************************************ PeterEn ****/
typedef void *MSOGV;
#define msocbMSOGV (sizeof(MSOGV))


#define msoclrNil   (0xFFFFFFFF)
#define msoclrBlack (0x00000000)
#define msoclrWhite (0x00FFFFFF)
#define msoclrNinch (0x80000001)


/*************************************************************************
	Stream I/O Support Functions

  	MsoFByteLoad, MsoFByteSave, MsoFWordLoad, MsoFWordSave, etc.
	The following functions are helper functions to be used when loading or
	saving toolbar data using an OLE 2 Stream.  They take care of the stream
	I/O, byte swapping for consistency between Mac and Windows, and error
	checking.  They should be used in all FLoad/FSave callback functions.
	MsoFWtzLoad expects wtz to point at an array of 257 WCHARs.  MsoFWtzSave
	will save an empty string if wtz is passed as NULL.

	SetLastError:  can be set to values from IStream's Read and Write methods
************************************************************ WAHHABB ****/
MSOAPIX_(BOOL) MsoFByteLoad(LPSTREAM pistm, BYTE *pb);
MSOAPIX_(BOOL) MsoFByteSave(LPSTREAM pistm, const BYTE b);
MSOAPI_(BOOL) MsoFWordLoad(LPSTREAM pistm, WORD *pw);
MSOAPI_(BOOL) MsoFWordSave(LPSTREAM pistm, const WORD w);
MSOAPI_(BOOL) MsoFLongLoad(LPSTREAM pistm, LONG *pl);
MSOAPI_(BOOL) MsoFLongSave(LPSTREAM pistm, const LONG l);
MSOAPIX_(BOOL) MsoFWtzLoad(LPSTREAM pistm, _Out_z_cap_(cchMax) WCHAR *wtz, int cchMax);
MSOAPIX_(BOOL) MsoFWtzSave(LPSTREAM pistm, const WCHAR *wtz);


MSOAPIMX_(int) MsoCchGetUsersFilesFolder(WCHAR *wzFilename, int cchFilename);

#ifdef MAPIVIM
/*	Returns the a full pathname to the MAPIVIM DLL in wzPath.   */
MSOAPI_(int) MsoFGetMapiPath(WCHAR* wzPath, BOOL fInstall);
#endif // MAPIVIM

MSOAPIMX_(WCHAR *) MsoWzGetKey(const WCHAR *wzApp, const WCHAR *wzSection, WCHAR *wzKey);


/****************************************************************************
	Stuff about File IO
************************************************************** PeterEn *****/

/* MSOFO = File Offset.  This is the type in which Office stores seek
	positions in files/streams.  I kinda wanted to use FP but that's already
	a floating point quantity. Note that the IStream interfaces uses
	64-bit quantities to store these; for now we're just using 32.  These
	are exactly the same thing as FCs in Word. */
typedef ULONG MSOFO;
#define msofoFirst ((MSOFO)0x00000000)
#define msofoLast  ((MSOFO)0xFFFFFFFC)
#define msofoMax   ((MSOFO)0xFFFFFFFD)
#define msofoNil   ((MSOFO)0xFFFFFFFF)

/* MSODFO = Delta File Offset.  A difference between two MSOFOs. */
typedef MSOFO MSODFO;
#define msodfoFirst ((MSODFO)0x00000000)
#define msodfoLast  ((MSODFO)0xFFFFFFFC)
#define msodfoMax   ((MSODFO)0xFFFFFFFD)
#define msodfoNil   ((MSODFO)0xFFFFFFFF)


/****************************************************************************
	Office ZoomRect animation code
****************************************************************************/
MSOAPI_(void) MsoZoomRect(RECT *prcFrom, RECT *prcTo, BOOL fAccelerate, HRGN hrgnClip);
MSOAPI_(void) MsoZoomRectEx(RECT *prcFrom, RECT *prcTo, BOOL fAccelerate, HRGN hrgnClip, int delay);

/*	On the Windows side we don't call OleInitialize at boot time - only
	CoInitialize. On the Mac side this is currently not being done because
	the Running Object Table is tied in with OleInitialize - so we can't
	call RegisterActiveObject if OleInitialize is not called - may
	want to revisit this issue. */

/*	Should be called before every call that requires OleInitialize to have
	been called previously. This function calls OleInitialize if it hasn't
	already been called. */
MSOAPI_(BOOL) MsoFEnsureOleInited();
/*	If OleInitialize has been called then calls OleUninitialize */
MSOAPI_(void) MsoOleUninitialize();

// Delayed Drag Drop Registration
/*	These routines are unnecessary on the Mac since Mac OLE doesn't require OLE
    to be initialized prior to using the drag/drop routines */
/*	All calls to RegisterDragDrop should be replaced by
	MsoHrRegisterDragDrop. RegisterDragDrop requires OleInitialize so
	during boot RegisterDragDrop should not be called. This function
	adds the drop target to a queue if OleInitialize hasn't already been
	called. If it has then it just calls RegisterDragDrop. */
MSOAPI_(HRESULT) MsoHrRegisterDragDrop(HWND hwnd, IDropTarget *pDropTarget);

/*	All calls to RevokeDragDrop should be replaced by
	MsoHrRevokeDragDrop. If a delayed queue of drop targets exists
	then this checks the queue first for the target. */
MSOAPI_(HRESULT) MsoHrRevokeDragDrop(HWND hwnd);

/*	Since all drop targets previously registered at boot time are now
	stored in a queue, we need to make sure we register them sometime.
	These can become drop targets
	a. if we are initiating a drag and drop - in which case we call this
	function before calling DoDragDrop (inside MsoHrDoDragDrop).
	b. while losing activation - so we might become the drop target of
	another app. So this function is called from the WM_ACTIVATEAPP
	message handler. */
MSOAPI_(BOOL) MsoFRegisterDragDropList();

/*	This function should be called instead of DoDragDrop - it first
	registers any drop targets that may be in the lazy init queue. */
MSOAPI_(HRESULT) MsoHrDoDragDrop(IDataObject *pDataObject,
	IDropSource *pDropSource, DWORD dwOKEffect, DWORD *pdwEffect);


/*	Module names MsoLoadModule supports */
/*  IF ANY THING IS CHANGED HERE - CHANGE GLOBALS.CPP! */

enum
{
	msoimodWinnls,		// System International utilities
	#define msoimodGetMax (msoimodWinnls+1)

	msoimodShell,		// System Shell
	msoimodHlink,		// Hyperlink APIs
	msoimodDarwin,		// Darwin
	msoimodRichEdit,    // Richedit dll
	msoimodMsoHev,		// Msohev.dll
	msoimodOlepro32,    // olepro32.dll - OleCreateFontIndirect & OldCreatePictureIndirect
	msoimodWFF,         // ippwff.dll - IWebFolderForms dll
	msoimodWtsapi32,	// WTS Api's
	msoimodWsmEng, // wsmeng.dll
#ifdef DEBUG
	msoimodDbgHelp, // dbghelp.dll
#endif
	msoimodMax,
};

/* ifn enums for Modules loaded by MsoLoadModule */
/* THE ORDER MUST MATCH THAT IN GLOBALS.CPP! -- MRuhlen */

enum
{
	ifnDllGetClassObject,
	ifnInvalidateDriveType,
	ifnCIDLData_CreateFromIDArray,
	ifnSHCreateLinks,
	cfnShell
};

enum
{
	ifnHlinkCreateFromMoniker,
	ifnHlinkCreateFromString,
	ifnHlinkCreateFromData,
	ifnHlinkCreateBrowseContext,
	ifnHlinkClone,
	ifnHlinkNavigateToStringReference,
	ifnHlinkOnNavigate,
	ifnHlinkUpdateStackItem,
	ifnHlinkOnRenameDocument,
	ifnHlinkNavigate,
	ifnHlinkResolveMonikerForData,
	ifnHlinkResolveStringForData,
	ifnHlinkParseDisplayName,
	ifnHlinkQueryCreateFromData,
	ifnHlinkSetSpecialReference,
	ifnHlinkGetSpecialReference,
	ifnHlinkCreateShortcut,
	ifnHlinkResolveShortcut,
	ifnHlinkIsShortcut,
	ifnHlinkResolveShortcutToString,
	ifnHlinkCreateShortcutFromString,
	ifnHlinkGetValueFromParams,
	ifnHlinkCreateShortcutFromMoniker,
	ifnHlinkResolveShortcutToMoniker,
	ifnHlinkTranslateURL,
	ifnHlinkCreateExtensionServices,
	ifnHlinkPreprocessMoniker,
	cfnHlink
};

enum
{
	ifnMsiGetProductCodeW,
	ifnMsiGetComponentPathW,
	ifnMsiReinstallFeatureW,
	ifnMsiReinstallProductW,
	ifnMsiQueryFeatureStateW,
	ifnMsiQueryProductStateW,
	ifnMsiUseFeatureW,
	ifnMsiGetUserInfoW,
	ifnMsiInstallMissingFileW,
	ifnMsiSetInternalUI,
	ifnMsiInstallProductW,
	ifnMsiEnumComponentQualifiersW,
	ifnMsiProvideQualifiedComponentW,
	ifnMsiVerifyPackageW,
	ifnMsiConfigureFeatureW,
	ifnMsiConfigureProductW,
	ifnMsiConfigureProductExW,
	ifnMsiProvideComponentW,
	ifnMsiInstallMissingComponentW,
	ifnMsiEnableLogW,
	ifnMsiCollectUserInfoW,
	ifnMsiGetProductInfoW,
	ifnMsiSetExternalUIW,
	ifnMsiUseFeatureExW,
	ifnMsiProvideQualifiedComponentExW,
	ifnMsiLocateComponentW,
	ifnMsiEnumComponentQualifiersA,
	ifnMsiEnumClientsW,
	ifnMsiEnumFeaturesW,
	ifnMsiGetFeatureUsageW,
	ifnMsiViewExecute,
	ifnMsiDatabaseOpenViewW,
	ifnMsiOpenDatabaseW,
	ifnMsiCloseHandle,
	ifnMsiRecordGetStringW,
	ifnMsiViewFetch,
	ifnMsiRecordIsNull,
	ifnMsiEnumComponentsW,
	cfnDarwin
};

enum
{
	ifnPHevCreateFileInfo,
	ifnWHevParseFile,
	ifnFHevActivateApp,
	ifnHevDestroyFileInfo,
	ifnFHevRegister,
	ifnFHevSetDefaultEditor,
	ifnWHeviAppFromProgId,
	ifnFHevAddToFileInfo,
	ifnHevFGetProgIDFromFile,
	ifnHevFGetCreatorAppIcon,
	ifnHevFGetCreatorAppName,
	ifnHevFQueryDefaultEditor,
	ifnHevFSetExtraData,
	ifnHevFCheckNonMSApp,
	ifnHevFGetFileIcons,
        ifnHevFIsHTMLOrMHTML,
	cfnMsoHev
};

enum
{
    ifnOleCreateFontIndirect,
    ifnOleCreatePictureIndirect,
    cfnOlepro32
};

enum
{
	ifnGetIWFFPtr,
	cfnWFF
};

enum
{
	ifnWTSQuerySessionInformationW,
	ifnWTSFreeMemory,
	cfnWtsapi32
};

enum
{
	ifnWsmEngineCreate,
	cfnWsmEng
};

enum
{
	ifnSymInitialize,
	ifnSymCleanup,
	ifnSymGetOptions,
	ifnSymSetOptions,
	ifnSymGetSymFromAddr,
	ifnSymGetLineFromAddr,
	cfnDbgHelp
};

// we don't bother loading any functions out of these modules
#define cfnUser 0
#define cfnGdi 0
#define cfnWinnls 0
#define cfnWinmm 0
#define cfnRichEdit 0


/*	Returns the module handle of the given module imod. Loads it if it is
	not loaded already.  fForceLoad will force a LoadLibrary on the DLL
	even if it is already in memory. */
MSOAPI_(HINSTANCE) MsoLoadModule(int imod, BOOL fForceLoad);

MSOAPIX_(void) MsoFreeModule(int imod);

MSOAPI_(BOOL) MsoFModuleLoaded(int imod);

/*	Returns the proc address in the module imod of the function
	szName.  Returns NULL if the module is not found or if the entry
	point does not exist in the module. */
MSOAPI_(FARPROC) MsoGetProcAddress(int imod, const char* szName);


/*	This API should be called by the client before MsoFInitOffice to set
	our locale id so that we can load the correct international dll.
	Defaults to the user default locale if app doesn't call this API before. */
MSOAPI_(void) MsoSetLocale(LCID dwLCID);


/* Cover for standard GetTextExtentPointW that:
	1. Uses GetTextExtentPoint32W on Win32 (more accurate)
	2. Fixes Windows bug	when cch is 0.  If cch is 0 then the correct dy
		is returned and dx will be 0.  Also, if cch is 0 then wz can be NULL. */
MSOAPI_(BOOL) MsoFGetTextExtentPointW(HDC hdc, const WCHAR *wz, int cch, LPSIZE lpSize);

/* Covers for Windows APIs that need to call the W versions if on a
	Unicode system, else the A version. */
MSOAPI_(LRESULT) MsoDispatchMessage(const MSG *pmsg);
MSOAPI_(LRESULT) MsoSendMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
MSOAPI_(LONG) MsoPostMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
MSOAPI_(LRESULT) MsoCallWindowProc(WNDPROC pPrevWndFunc, HWND hwnd, UINT msg,
		WPARAM wParam, LPARAM lParam);

MSOAPIX_(LONG) MsoGetWindowLong(HWND hwnd, int nIndex);
MSOAPI_(LONG) MsoSetWindowLong(HWND hwnd, int nIndex, LONG dwNewLong);

#ifdef _WIN64
MSOAPIX_(LONG_PTR) MsoGetWindowLongPtr(HWND hwnd, int nIndex);
MSOAPIX_(LONG_PTR) MsoSetWindowLongPtr(HWND hwnd, int nIndex, LONG_PTR dwNewLong);
#else
#define MsoGetWindowLongPtr(hwnd, nIndex) MsoGetWindowLong(hwnd, nIndex)
#define MsoSetWindowLongPtr(hwnd, nIndex, dwNewLong) MsoSetWindowLong(hwnd, nIndex, dwNewLong)
#endif // _WIN64

#define ETO_MSO_IME_UL_WORKAROUND 0x0800000
#define ETO_MSO_NO_GLYPH 0x1000000
#define ETO_MSO_DISPLAY_HOTKEY 0x2000000
#define ETO_MSO_NO_FONTLINK	0x20000000
#define	ETO_MSO_DONT_CALL_UCSCRIBE	0x40000000
#define ETO_MSO_FORCE_ENHMETAFILE 0x80000000
MSOAPI_(int) MsoGetWindowTextWtz(HWND hwnd, _Out_z_cap_(cchMax) WCHAR *wtz, int cchMax);
MSOAPIX_(BOOL) MsoSetWindowTextWtz(HWND hwnd, const WCHAR *wtz);
MSOAPI_(BOOL) MsoAppendMenuW(HMENU hMenu, UINT uFlags, UINT uIDNewItem,
								LPCWSTR lpNewItem);
MSOAPI_(BOOL) MsoInsertMenuW(HMENU hMenu, UINT uPosition, UINT uFlags,
								UINT uIDNewItem, LPCWSTR lpNewItem);

/*	Return the facename of the system UI (dialog) font in wtzFaceName. */
MSOAPI_(VOID) MsoGetSystemUIFont(_Out_z_cap_(cchMax) WCHAR *wtzFaceName, int cchMax);

/*	Return TRUE if the user settings indicate that we should not use Tahoma
	and instead use the system UI font.  If we should use the system font
	and plf is non-NULL and points at a font with facename currently equal
	to "Tahoma", then overwrite it with the the appropriate system UI
	font (e.g. "MS Sans Serif" or "MS Dialog") with the same attributes. */
MSOAPI_(BOOL) MsoFOverrideOfficeUIFont(LOGFONT *plf);

/*	Return TRUE if the user settings indicate that we should not use Tahoma
	and instead use the system UI font.  If we should use the system font
	and the Windows dialog at hwndDlg is using the Tahoma font, then set
	the font for all controls in the dialog to the default system dialog font.
	Call this in WM_INITDIALOG for Windows dialogs.
		If this function creates a new font, it returns that font in *phfont.
	In this case, the caller is responsible for deleting the font when the
	dialog is closed. */
MSOAPI_(BOOL) MsoFOverrideOfficeUIWinDlgFont(HWND hwndDlg, HFONT *phfont);

/* If the LOGFONT at plf is a default system UI font then change *plf
	to substitute the Tahoma font as appropriate for the system.
	Return TRUE if *plf was changed. */
MSOAPI_(BOOL) MsoFSubstituteTahomaLogfont(LOGFONT *plf);

/*	Some Far East languages need a minimum 9 point UI font size because
	otherwise the glyphs are unreadable.  On other languages, we want to go
	no lower than 8 pt (even if buggy system settings return a smaller value).
	This is controlled by a resource.  If the the font at plf describes a
	smaller size, increase it to the minimum. */
MSOAPI_(void) MsoEnsureMinUIFontSize(LOGFONT *plf);

// Fonts supported by MsoFGetFontSettings
enum
	{
	msofntMenu,
	msofntTooltip,
	};

/* Return font and color info for the font given by 'fnt' (see msofntXXX).
	If fVertical, then the font is rotated 90 degrees if this fnt type
	supports rotation in Office.  If phfont is non-NULL, return the HFONT used
	for this item.  This font is owned and cached by Office and should not
	be deleted.  If phbrBk is non-NULL, return a brush used for the
	background of this item	(owned by Office and should not be deleted).
	If pcrText is non-NULL,	return the COLOREF used for the text color for
	this item. Return TRUE if all requested info was returned. */
MSOAPI_(BOOL) MsoFGetFontSettings(int fnt, BOOL fVertical, HFONT *phfont,
		HBRUSH *phbrBk, COLORREF *pcrText);

/* If the system suppports NotifyWinEvent, then call it with the given
	parameters (see \otools\inc\win\winable.h). */
MSOAPI_(void) MsoNotifyWinEvent(DWORD dwEvent, HWND hwnd, LONG idObject, LONG idChild);

/* Return FALSE iff we don't need to call MsoNotifyWinEvent.  This is only an
	optimization to avoid prep work in the caller since calling MsoNotifyWinEvent
	is always safe and fast if nobody is listening. */
MSOAPIX_(BOOL) MsoFNotifyWinEvents();

/* Return TRUE if an Accessibility screen reader is running. */
MSOAPI_(BOOL) MsoFScreenReaderPresent();

/* Call LResultFromObject in oleacc.dll to thunk an IUnknown object into
   an LRESULT to allow for cross-process access.  The wParam is the parameter
   as passed to WM_GETOBJECT. */
MSOAPI_(LRESULT) MsoLThunkIUnknown(IUnknown *punk, WPARAM wParam);

/*	Return TRUE if build version of OleAcc.Dll is greater than or equal to
	the version number passed in (in the form A.B.C.D). */
MSOAPIX_(BOOL) MsoFOleAccDllVersion(short A, short B, short C, short D);

/****************************************************************************
   MsoNotifyIMEWindowChange

	This function should be called whenever the set of visible IME windows
	changes.	Usually, calling this function in response to IMN_OPENCANDIDATE
	and IMN_CHANGECANDIDATE messages is sufficient. Unfortunately, Office is
	unable to catch these messages itself because they go to the window that
	has focus, not to the top-level window of the application.
		By calling this function, the application allows Office to do useful
	things like move the assistant off of the IME windows.
****************************************************************************/
MSOAPI_(void) MsoNotifyIMEWindowChange(void);


/****************************************************************************
	MsoFHelpIdHasHelp

	This function reports on whether a given Help ID has been registered with
	MSO to have help turned on.  In particular, this is used in dialog
	managers needing to know whether to show (and handle) the "?" button.
****************************************************************************/
MSOAPI_(BOOL) MsoFHelpIdHasHelp(int helpId);


/****************************************************************************
	MsoFShowHelpForHelpId

	This function brings up help for a given Help ID has been registered with
	MSO to have help turned on.
****************************************************************************/
MSOAPI_(BOOL) MsoFShowHelpForHelpId(HWND hwndApp, int helpId);



/*-------------------------------------------------------------------------*/

#define HELP_VBA_COMMAND   0x10000

typedef enum
{
	msoargtSwitch,
	msoargtFile,
	msoargtString,
	msoargtProfile,
	msoargtAutomation,
	msoargtRegserver,
	msoargtUnregserver,
	msoargtSwitchData,
	msoargtEmbedding,
	msoargtSafe,
	msoargtDDE,
} ARGT;

typedef struct
{
	ARGT argt;
	int ch;  // NOTE: don't change this to a CHAR since ARGC is overloaded to handle WCHAR cmdline
	union
		{
		CHAR *szData;
		int fFound;
		WCHAR *wzData;
		};
} ARGC;

MSOAPI_(int) MsoParseCommandLine(ARGC *pargc, unsigned int carg, _Inout_ _Deref_prepost_z_ CHAR **pszCmdLine, int fDestructive);
MSOAPI_(int) MsoParseCommandLineW(ARGC *pargc, unsigned int carg, _Inout_ _Deref_prepost_z_ WCHAR **pszCmdLine, int fDestructive);

/****************************************************************************
	MsoGetIntlSysSettings

	This API is important to non-US apps, namely MidEast and FarEast.  It takes
	a BOOL param, fRefresh, which will be FALSE most of the time--which is at the
	init time.  Apps will only set it to TRUE when they think that System settings
	have been changed.

	BIDI_TODO: It would be a good idea to remove the fRefresh param, and refresh it
	automatically when WM_SETTINGCHANGE is sent.  Where should we trap it?
****************************************************************************/
MSOAPI_(DWORD) MsoGetIntlSysSettings(BOOL fRefresh);
MSOAPI_(BOOL) MsoFEditLangSupported(WORD lid);

#define MSOI_ARABIC_SYSTEM_INSTALLED		0x00000001	// Arabic APIs
#define MSOI_ARABIC_FONTS_INSTALLED			0x00000002	// Arabic Fonts
#define MSOI_ARABIC_KBD_INSTALLED			0x00000004	// Arabic Keyboards



//SOUTHASIA
// These uses the higher nibble of the second byte to define for SOUTHASIA.
#define MSOI_HINDI_FULLY_INSTALLED			0x00001000
#define MSOI_THAI_FULLY_INSTALLED			0x00002000
#define MSOI_VIETNAMESE_FULLY_INSTALLED		0x00004000
#define MSOI_MSO9SA_RUNNING                 0x00008000 // use to mark this MSO version as SA enabled
//SOUTHASIA

#define MSOI_ARABIC_FULLY_INSTALLED			(MSOI_ARABIC_SYSTEM_INSTALLED \
											| MSOI_ARABIC_FONTS_INSTALLED \
											| MSOI_ARABIC_KBD_INSTALLED)
											// Full Arabic system (0x00000007)

#define MSOI_HEBREW_SYSTEM_INSTALLED		0x00000010	// Hebrew APIs
#define MSOI_HEBREW_FONTS_INSTALLED			0x00000020	// Hebrew Fonts
#define MSOI_HEBREW_KBD_INSTALLED			0x00000040	// Hebrew Keyboard

#define MSOI_HEBREW_FULLY_INSTALLED			(MSOI_HEBREW_SYSTEM_INSTALLED \
											| MSOI_HEBREW_FONTS_INSTALLED \
											| MSOI_HEBREW_KBD_INSTALLED)
											// Full Hebrew system (0x00000070)

#define MSOI_URDU_SYSTEM_INSTALLED			0x00000100	// Urdu APIs
#define MSOI_URDU_FONTS_INSTALLED			0x00000200	// Urdu Fonts
#define MSOI_URDU_KBD_INSTALLED				0x00000400	// Urdu Keyboard

#define MSOI_URDU_FULLY_INSTALLED			(MSOI_URDU_SYSTEM_INSTALLED \
											| MSOI_URDU_FONTS_INSTALLED \
											| MSOI_URDU_KBD_INSTALLED)
											// Full Urdu system (0x00000700)

#define MSOI_FARSI_SYSTEM_INSTALLED			0x00001000	// Farsi APIs
#define MSOI_FARSI_FONTS_INSTALLED			0x00002000	// Farsi Fonts
#define MSOI_FARSI_KBD_INSTALLED			0x00004000	// Farsi Keyboard

#define MSOI_FARSI_FULLY_INSTALLED			(MSOI_FARSI_SYSTEM_INSTALLED \
											| MSOI_FARSI_FONTS_INSTALLED \
											| MSOI_FARSI_KBD_INSTALLED)
											// Full Farsi system (0x00007000)

#define MSOI_SINDHI_SYSTEM_INSTALLED		0x00001000	// Sindhi APIs
#define MSOI_SINDHI_FONTS_INSTALLED			0x00002000	// Sindhi Fonts
#define MSOI_SINDHI_KBD_INSTALLED			0x00004000	// Sindhi Keyboard

#define MSOI_SINDHI_FULLY_INSTALLED			(MSOI_SINDHI_SYSTEM_INSTALLED \
											| MSOI_SINDHI_FONTS_INSTALLED \
											| MSOI_SINDHI_KBD_INSTALLED)
											// Full Sindhi system (0x00070000)


#define MSOI_NT4_RUNNING					0x00000100	// Either Windows NT 4
#define MSOI_NT5_RUNNING					0x00000200	// Or Windows NT 5
#define MSOI_WIN95_RUNNING					0x00000400	// Or Windows 95
#define MSOI_WIN98_RUNNING					0x00000800	// Or Windows 98

// Does it support Unicode?
MSOAPI_(BOOL) MsoFUnicodeCommCtrl();

#endif // MSOUSER_H
