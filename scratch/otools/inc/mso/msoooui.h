/****************************************************************************
	MsoOOUI.h

	Owner: NitinG
	Copyright (c) 1999 Microsoft Corporation

	This file contains the exported interfaces and declarations for
	Office On-Object Control.
****************************************************************************/

#ifndef MSOOOUI_H
#define MSOOOUI_H


/****************************************************************************
	On-Obj String types - For IMsoOnObjControlUser::FGetString
****************************************************************** NITING **/
enum
	{
	msooostrTooltip = 0,
	msooostrMSAADescription,
	};


/****************************************************************************
	On-Obj Styles

	The lower byte of the style are continous numbers (not to be bit-wise ored)
	used to index into the rgoopDefault) array. The higher three bytes are
	other properties that can be or'ed including the first byte. To add more
	default styles, add a new msooosDefaultXXX and add a new record to
	rgoopDefault array.
****************************************************************** NITING **/
// Use the default tcid, position and sound for various types. To look at the
// various default values look at the rgoopDefault array.
#define msooosDefaultNil          0x00000000
#define msooosDefaultAutoText     0x00000001
#define msooosDefaultAutoList     0x00000002
#define msooosDefaultPaste        0x00000003
#define msooosDefaultErrorCheck   0x00000004
#define msooosDefaultFactoid	  0x00000005
#define msooosDefaultFactoidExcel 0x00000006
#define msooosDefaultDragFill	  0x00000007
#define msooosDefaultInsertCells  0x00000008
#define msooosDefaultAutoFit	  0x00000009
#define msooosDefaultAutoFitFE	  0x0000000a
#define msooosDefaultHyperlink	  0x0000000b
#define msooosDefaultAutoFlow	  0x0000000c
#define msooosDefaultPictResize	  0x0000000d
#define msooosDefaultInsertObject 0x0000000e
#define msooosDefaultPersona      0x0000000f
#define msooosDefaultXdt          0x00000010
#define msooosDefaultPF           0x00000011

// Add new defaults here - don't forget to bump up msooosDefaultMax, and
// update msoosDefaultMask appropriately.
#define msooosDefaultMax          0x00000012
#define msooosDefaultMask         0x000000FF

// Hide/Show after hide/show delay instead of immediately
#define msooosHideTimer           0x00000100
#define msooosShowTimer           0x00000200
#define msooosTimers              (msooosHideTimer | msooosShowTimer)

// Send a click notification (msooonClick) when clicked
#define msooosNotifyClick         0x00000400

// Use current mouse/ip positions to show/hide the controls. This style is
// used by the On-Obj Manager to figure out if it should hide/show the
// control when NotifyCursorPos is called. If a control does NOT have this
// style set, then the manager will show the control irrespective of cursor
// position. The control itself does not use this style. If you are planning
// to manage the show/hide of control yourself and not append it to the
// On-Obj Manager, then this style is irrelevant.
#define msooosShowOnCursor        0x00000800

/* Set this flag if you wish the position of the OOUI to remain fixed rather */
/* than be recalculated on each call to IMsoOnObjControl::Show. Note that,   */
/* even with this flag set, a call to IMsoOnObjControl::FUpdateHitRegion will*/
/* trigger a recalculation of position.                                      */
#define msooosFixedPosition		 0x00001000

// Send a hover notification (msooonHover) when hovered over
#define msooosNotifyHover         0x00002000

// Send a resize notification (msooonResize) when OOC is resized
#define msooosNotifyResize        0x00004000

// Disable hiding the ooui under the mouse when the caret moves around,
// but the mouse does not move.  This basically enables the showing of two
// oouis, one under the mouse and one under the caret.
#define msooosShowDual			0x00008000

// Disable the ooui animation that would otherwise be played when showing
// an ooui for the first x times.  Having this flag set on an ooui will not
// reduce the value of x by one.
#define msooosDontAnimate		0x00010000


// Combo of above styles
#define msooosAutoText            (msooosDefaultAutoText | msooosTimers | \
								   msooosShowOnCursor | msooosNotifyHover)
#define msooosAutoList            msooosDefaultAutoList
#define msooosPaste               (msooosDefaultPaste | msooosFixedPosition)
#define msooosFactoid			  (msooosDefaultFactoid | msooosTimers | \
								   msooosShowOnCursor | msooosNotifyHover)
#define msooosInsertObject        (msooosDefaultInsertObject | msooosFixedPosition)
#define msooosPersona			  (msooosDefaultPersona | msooosTimers | msooosDontAnimate | \
								   msooosShowOnCursor | msooosNotifyHover | msooosNotifyClick)
#define msooosXdt				(msooosDefaultXdt | msooosDontAnimate)


/****************************************************************************
	On-Obj Notifications - return FALSE if you do not handle the notification.
****************************************************************** NITING **/
typedef enum
	{
	msooonDelete = 0,	// The OOC is about to be deleted. The client will
						// usually delete the pvClient data on this
						// notification. Lparam is NULL.
	msooonClick,		// The OOC has been clicked. To get this notification
						// pass the msooosNotifyClick style when setting up
						// the control. Returning FALSE will make office handle
						// click as usual. Returning TRUE will stop further
						// office processing. Lparam is NULL.
	msooonNeedCtls,		// The OOC needs the controls for drop down menu
						// Lparam is NULL. The client should call FSetMenuCtls
						// in return to set the controls.
	msooonDropMenu,		// The OOC is about to drop the drop down menu
						// Lparam is NULL.
	msooonMenuItemClick,// One of the menu items was clicked. LPARAM is the
						// index of the item clicked.
	msooonHidden,		// The OOC has just been hidden (as a result of timer
						// or otherwise). Lparam is NULL.
	msooonHover,		// On mouse move, LParam == TRUE if the mouse is over the OOUI
	msooonODelete,	// Office wishes to delete the control. Reason is in lParam.
						// Return TRUE to prevent deletion.
	msooonResize,		// Resizing the control.  LParam is 0 for light and 1 
						// for heavy mode
	msooonMove,         // a WM_MOVE message has been passed to the control
	} MSOOON;

enum
	{
	msoodelDefault,			// default
	msoodelClipboard,		// clipboard changed
	};
	

/****************************************************************************
	On-Object Properties: Used for IMsoOnObjControl::SetProperty.
****************************************************************** NITING **/
enum
	{
	msooopMin = 0,
	msooopTcidLight = msooopMin,
	msooopTcidHeavy,
	msooopPositionDefault,
	msooopPositionFallback,
	msooopSound,
	msooopGap,
	msooopTcidHighContrastLight,
	msooopLightIconWidth,
	msooopLightIconHeight,
	msooopLargeLightIconWidth,
	msooopLargeLightIconHeight,
	msooopAniWidth,
	msooopAniHeight,
	msooopLightHitWidth,
	msooopLightHitHeight,
	msooopAlwaysActive, // if TRUE, this OOUI will activate even if the parent window isnt active, else it's FALSE (default)
	msooopFRefreshCtls,	// refresh menu controls each time ooui is clicked (see 11.131011)
	// Add properties above this
	msooopMax
	};


/****************************************************************************
	On-Object Positions: Used for IMsoOnObjControl::SetProperty with
	msooopPositionDefault, msooopPositionFallback.

    LTLT   LTRT          RTLT   RTRT
          -------------------- 
    LTLB | LTRB          RTLB | RTRB
         |                    |
         |                    |
         |                    |
    LBLT | LBRT          RBLT | RBRT
          -------------------- 
    LBLB   LBRB          RBLB   RBRB

****************************************************************** NITING **/
enum
	{
	msoooPosNil = -1,	// Users shouldn't pass this value
	msoooPosMin = 0,
	msoooPosLTLT = msoooPosMin,
	msoooPosLTLB,
	msoooPosLTRT,
	msoooPosLTRB,	// Avoid using this: It will put OOC on top of Hit-Region.
	msoooPosLBLT,
	msoooPosLBLB,
	msoooPosLBRT,	// Avoid using this: It will put OOC on top of Hit-Region.
	msoooPosLBRB,
	msoooPosRTLT,
	msoooPosRTLB,	// Avoid using this: It will put OOC on top of Hit-Region.
	msoooPosRTRT,
	msoooPosRTRB,
	msoooPosRBLT,	// Avoid using this: It will put OOC on top of Hit-Region.
	msoooPosRBLB,
	msoooPosRBRT,
	msoooPosRBRB,
	msoooPosMax
	};


/****************************************************************************
	Defines the IMsoOnObjControl interface

****************************************************************** NITING **/
#undef  INTERFACE
#define INTERFACE  IMsoOnObjControl

DECLARE_INTERFACE_(IMsoOnObjControl, ISimpleUnknown)
{
	// ISimpleUnknown methods
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;

	/* Standard FDebugMessage method */
	MSODEBUGMETHOD

	/* IMsoOnObjControl methods */
	/* Delete this On-Obj control */
	MSOMETHOD_(void, Delete) (THIS_ BOOL fSkipOon) PURE;

	/* Show this control */
	MSOMETHOD_(void, Show) (THIS_ BOOL fImmediate) PURE;

	/* Hide this control */
	MSOMETHOD_(void, Hide) (THIS_ BOOL fImmediate) PURE;

	/* Move the control and its hit region. */
	MSOMETHOD_(void, Move) (THIS_ int dx, int dy) PURE;

	/* Move the control and its hit region. */
	MSOMETHOD_(void, MoveEx) (THIS_ int dx, int dy, BOOL fRecalc) PURE;

	/* Update the hit region */
	MSOMETHOD_(BOOL, FUpdateHitRegion) (THIS_ int crc, RECT *rgrc) PURE;

	/* Update the hit region Extended version with fForceCalcPos flag */
	MSOMETHOD_(BOOL, FUpdateHitRegionEx) (THIS_ int crc, RECT *rgrc, BOOL fForceCalcPos) PURE;
	
	/* Get the current hit region. */
	MSOMETHOD_(int, CGetHitRegion) (_Out_cap_(crcMax) THIS_ RECT *rgrc, int crcMax) PURE;

	/* Set the parent hwnd of the control. */
	MSOMETHOD_(void, SetParentHwnd) (THIS_ HWND hwnd) PURE;

	/* Get the parent hwnd of the control. */
	MSOMETHOD_(HWND, HwndParentGet) (THIS) PURE;

	/* Tell the control whether the cursor or caret is in or out */
	/* of its hit region.                                        */
	MSOMETHOD_(void, SetCursorPos) (THIS_ BOOL fInHitRgn, BOOL fCaret) PURE; 

	/* Notify the control of the current cursor or caret position so that
	   it can show or hide itself appropriately. */
	MSOMETHOD_(void, NotifyCursorPos) (THIS_ POINT pt,	BOOL fCaret) PURE;
	
	/* Update the rcContainer for the control. fAllowShow lets */
	/* caller determine if show should be done.                */ 
	MSOMETHOD_(void, UpdateRcContainer) (THIS_ RECT *prc, BOOL fAllowShow) PURE;

	/* Set the tcid for the i'th menu item. */
	MSOMETHOD_(void, SetTcid) (THIS_ int i, int tcid) PURE;

	/* Set the fDisabled flag for the i'th menu item. */
	MSOMETHOD_(void, SetDisabled) (THIS_ int i, BOOL fDisabled) PURE;

	/* Set the tbbs for the i'th menu item. */
	MSOMETHOD_(void, SetTbbs) (THIS_ int i, int tbbs) PURE;

	/* Set the string for the i'th menu item. */
	MSOMETHOD_(void, SetWtz) (THIS_ int i, const WCHAR *wtz) PURE;

	/* Set the drop down menu's controls */
	MSOMETHOD_(BOOL, FSetMenuCtls) (THIS_ int cctcr, MSOCTCR *rgctcr) PURE;

	/* Remove the i'th menu's control. */
	MSOMETHOD_(BOOL, FRemoveMenuCtl) (THIS_ int ictcr) PURE;

	/* Get the MSOCTCR for the i'th menu item. */
	MSOMETHOD_(BOOL, FGetMenuCtl) (THIS_ int i, MSOCTCR *pctcr) PURE;

	/* Copy the menu from this control to piooc. */
	MSOMETHOD_(void, CopyMenuCtls) (THIS_ IMsoOnObjControl *piooc) PURE;
	
	/* Returns the number of menu items */
	MSOMETHOD_(int, CctcrGet) (THIS) PURE;
	
	
	/* Get the ppvclient data */
	MSOMETHOD_(void **, PpvClientGet) (THIS) PURE;

	/* Set a property for the OOC. lParam depends on the oop value. Use
	   msooopXXX values for oop. */
	MSOMETHOD_(void, SetProperty) (THIS_ int oop, LPARAM lParam) PURE;
	
	/* Passes notification on to the control user */
	MSOMETHOD_(BOOL, FNotifyAction) (THIS_ MSOOON oon, LPARAM lParam) PURE;
	
	/* Set the sound for the OOC.  iSound must be >= 0 and < msosndMax */
	MSOMETHOD_(void, SetSound) (THIS_ int iSound) PURE;
	
	/* Returns the sound for the OOC. */
	MSOMETHOD_(int, ISoundGet) (THIS) PURE;
	
	/* Sets the string used for the MSAA name property. */
	MSOMETHOD_(BOOL, FSetMSAAName) (THIS_ const WCHAR *wzName) PURE;
	
	/* Gets the string used for the MSAA name property. */
	MSOMETHOD_(const WCHAR *, WzGetMSAAName) (THIS) PURE;
	
	/* Sets the string used for the MSAA description property. */
	MSOMETHOD_(BOOL, FSetMSAADescription) (THIS_ const WCHAR *wzDescription) PURE;

	/* Gets the string used for the MSAA description property. */
	MSOMETHOD_(const WCHAR *, WzGetMSAADescription) (THIS) PURE;

	/* Returns the visible state of this OOC. */
	MSOMETHOD_(BOOL, FVisible) (THIS) PURE;
	
	/* Returns the style of this OOC. */
	MSOMETHOD_(int, OosGet) (THIS) PURE;
	
	/* Sets the style of this OOC. */
	MSOMETHOD_(BOOL, FSetOos) (THIS_ int oos) PURE;

	// GetWindowRect for the OOUI window
	MSOMETHOD_(BOOL, FGetOOUIWindowRect) (THIS_ RECT *prc) PURE;
	
	/* Get a property for the OOC. Use msooopXXX values for oop. */
	MSOMETHOD_(int, GetProperty) (THIS_ int oop) PURE;
	
	/* Returns true if the OOC intersects it's pareent window (not scrolled off) */
	MSOMETHOD_(BOOL, FIntersectsParent) (THIS) PURE;
	
	/* Returns the deleted state of this OOC. */
	MSOMETHOD_(BOOL, FDeleted) (THIS) PURE;
	
	/* Returns true if the OOC will accept acceleration (visible, not deleted, etc) */
	MSOMETHOD_(BOOL, FCanAccelerate) (THIS) PURE;

	/* Set the hwnd which can have the focus in addtion to the */
	/* parent in order to show the control. */
	MSOMETHOD_(void, SetFocusHwnd) (THIS_ HWND hwnd) PURE;

	/* Get the hwnd which can have the focus in addition to the */
	/* parent in order to show the control. */
	MSOMETHOD_(HWND, HwndFocusGet) (THIS) PURE;

	/* Get the X and Y offsets for light weight mode */
	MSOMETHOD_(void, GetLtXYOffsets) (THIS_ int *pdx, int *pdy) PURE;

	/* Sets the string used for the MSAA Keyboard Shortcut property. */
	MSOMETHOD_(BOOL, FSetMSAAKeyboardShortcut) (THIS_ const WCHAR *wzKeyboardShortcut) PURE;

	/* Gets the string used for the MSAA Keyboard Shortcut property. */
	MSOMETHOD_(const WCHAR *, WzGetMSAAKeyboardShortcut) (THIS) PURE;
	
	/* Sets the string used for the MSAA Default Action property. */
	MSOMETHOD_(BOOL, FSetMSAADefaultAction) (THIS_ const WCHAR *wzDefaultAction) PURE;

	/* Gets the string used for the MSAA Default Action property. */
	MSOMETHOD_(const WCHAR *, WzGetMSAADefaultAction) (THIS) PURE;
};


/****************************************************************************
	Defines the IMsoOnObjControlUser interface

****************************************************************** NITING **/
#undef  INTERFACE
#define INTERFACE  IMsoOnObjControlUser

DECLARE_INTERFACE_(IMsoOnObjControlUser, ISimpleUnknown)
{
	// ISimpleUnknown methods
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;

	/* Standard FDebugMessage method */
	MSODEBUGMETHOD

	/* IMsoOnObjControlUser methods */
	/* Asks the control for a string based on the value of oostr
		(see msooostrXXX).  The control should return the string in the wtz
		buffer (limited to 255 chars).  If the control returns FALSE, no string 
		is assumed to have been returned. If the control returns TRUE, but the
		length of the string is 0, then no tooltip will be generated or
		displayed. */
	MSOMETHOD_(BOOL, FGetString) (THIS_ IMsoOnObjControl *pooc, int oostr,
							_Out_z_cap_(cch) WCHAR *wtz, int cch) PURE;

	/* Notify the user of some control action. Value of lParam depends on
	   the notifcation (oon). See msooonXXX for notifications. Return FALSE
	   if you do not handle the notification. */
	MSOMETHOD_(BOOL, FNotifyAction) (THIS_ IMsoOnObjControl *pooc, MSOOON oon,
							LPARAM lParam) PURE;
};


/****************************************************************************
	Defines the IMsoOnObjManager interface

****************************************************************** NITING **/
#undef  INTERFACE
#define INTERFACE  IMsoOnObjManager

DECLARE_INTERFACE_(IMsoOnObjManager, ISimpleUnknown)
{
	// ISimpleUnknown methods
	MSOMETHOD(QueryInterface) (THIS_ REFIID riid, void **ppvObj) PURE;

	/* Standard FDebugMessage method */
	MSODEBUGMETHOD

	/* IMsoOnObjManager methods */
	/* Deletes IMsoOnObjManager class */
	MSOMETHOD_(void, Delete) (THIS) PURE;

	/* Append the OOC to the list */
	MSOMETHOD_(BOOL, FAppend) (THIS_ IMsoOnObjControl *pooc) PURE;

	/* Remove the OOC from the list. Doesn't delete the control */
	MSOMETHOD_(BOOL, FRemove) (THIS_ IMsoOnObjControl *pooc) PURE;

	/* Show the OOC controls */
	MSOMETHOD_(void, ShowAll) (THIS) PURE;

	/* Hide the OOC controls */
	MSOMETHOD_(void, HideAll) (THIS) PURE;

	/* Notify the manager of the current cursor or caret position so that
	   it can show/hide OOCs appropriately */
	MSOMETHOD_(void, NotifyCursorPos) (THIS_ HWND hwnd, POINT pt,
							BOOL fCaret) PURE;
	
	/* see if the hwnd belongs to one of the OOC controls */
	MSOMETHOD_(BOOL, FIsOocHwnd) (THIS_ HWND hwnd, IMsoOnObjControl ** ppOoc) PURE;
};


// Accelerate to the next OOUI menu 
MSOAPI_(BOOL) MsoFAccelerateOOUI ();

MSOAPI_(BOOL) MsoFCreateOnObjManager (IMsoOnObjManager **ppoom);


MSOAPI_(BOOL) MsoFCreateOnObjControl (HMSOINST hinst, int oos, HWND hwndParent,
							RECT *prcContainer,
							int crc, RECT *rgrc, void *pvClient, 
							IMsoOnObjControlUser *pOOCU,
							IMsoOnObjControl **ppOOC);

//
// Get/Set OOUI prefs
//

// Return the global state of the Show Auto OOUI preference.
MSOAPI_(BOOL) MsoFGetShowAutoOOUIPref();

// Set the global state of the Show Auto OOUI preference.
MSOAPI_(BOOL) MsoSetShowAutoOOUIPref(BOOL fShowAutoOoui);

// Return the global state of the Show Paste Recovery OOUI preference.
MSOAPI_(BOOL) MsoFGetShowPasteRecoveryOOUIPref();

// Return the global state of the Show OOUI preference for the given style
MSOAPI_(BOOL) MsoFGetShowOOUIPrefOos(int oos);

// Set the global state of the Show Paste Recovery OOUI preference.
MSOAPI_(BOOL) MsoSetShowPasteRecoveryOOUIPref(BOOL fShowPasteOoui);

// Set the global state of the Show OOUI preference for a style.
MSOAPI_(BOOL) MsoSetShowOOUIPref(BOOL fShow, int rid, int oos1, int oos2, int oos3);

// Accelerate to OOC if appropriate
MSOAPI_(BOOL) MsoFOOCAccelerate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

// Return TRUE if the given screen point is in ANY visible OOC window
MSOAPI_(BOOL) MsoFPointInOOCWindow(POINT pt);

// Return the default value for the given property in the given style
MSOAPI_(int) MsoDefaultOOUIPropertyGet(int oos, int oop);

#endif // MSOOOUI_H

