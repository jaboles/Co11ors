// This include file defines the MSOCEO structure
// it is included from both NetUI (nuiValue.h) and MSO (msotb.h)

#ifndef MSOCEO_INCLUDED
#define MSOCEO_INCLUDED

#include "msostd.h"


typedef struct _msoceo
	{
	union
		{
		ULONG uceo;     // for bulk init, compares..
		struct
			{
			ULONG wStyle:8,         // bits 0 - 7: 0 == sys standard, 1-127 reserved
			      maskUnused:1,     // bits 8: unused and reserved
			      f3Ddisplay:1,     // bit  9: should the controls be drawn 3 D or flat
			      fInContextMenu:1, // bit 10: in context menu (or dropping from context menu)
			      fInCommandWell:1, // bit 11: in command well in the Customize dialog
			      fInHMenu:1,       // bit 12: The control is currently on an Hmenu
			      fInSDM:1,         // bit 13: The control is in a Gcc inside an SDM dialog
			      fMetricChange:1,  // bit 14: QuerySize should assume metric changes
			      fUnderlineMode:1, // bit 15: draw underline under first &
			      fForceFocus:1,    // bit 16: force selected look (opposite of fForceNoFocus) - only used in command well (?) 
			      fInMainMenu:1,    // bit 17: in a main menu bar
			      fInGcc:1,         // bit 18: is the control in a Gcc (should strip "_" from menu
			      fDropLeft:1,      // bit 19: cascading dropdown should drop to the left
			      fNotTopLevel:1,   // bit 20: control is in bar popped up by another
			      fInMenu:1,        // bit 21: control resides in a menu bar
			      fInCustomize:1,   // bit 22: is the command bar in customize mode?
			      fForceNoFocus:1,  // bit 23: do NOT draw selected, even if it is (opposite of fForceFocus)
				  fUseTextForBlack:1, // bit 24: TRUE if black should be replaced by 3D Text color
			      fMirror:1,        // bit 25: mirrored orientation (rotate 180), see note below
			      fVertical:1,      // bit 26: vertical orientation (rotate left 90)
			      fUNUSED2:1, //fExtraSize:1,     /* UNUSED?? */ // bit 27: extra large if TRUE, else extra small
			      fLarge:1,         // bit 28: standard small or large format
			      fDocked:1,        // bit 29: use standard-size docked state
			      fInDialogGcc:1, // bit 30: is this a GCC in an SDM dialog?
			      fGccOnTabSheet:1; // bit 31: Does our GCC have a gradient tab sheet background when themed?  (needed for Win32 GCC on a tab sheet)
			};
		};
	union
		{
		ULONG uceo2;     // for bulk init, compares..
		struct
			{
			ULONG fAltColors:1,
			fSDMLarge:1,
			fNoToolTip:1,
			fMenuReflow:1,          // analogous to setting msotbrMenuReflow
			fAltColors2:1,
			fNoCenterH:1,           // Only valid for Wrapped text buttons
			fCenterBtnText:1,       // Will center the text portion of a splitbutton menu
			fFocusDepressed:1,      // SDMGcc 1st depressed button will take focus on tmmFocus
			fGccKeepUnderline:1,    // Keep the apmersand on the Gcc (for SDM or other reason)
			fReturnKeys:1,          // PreTranslate won't eat keys
			fHyperlinkLook:1,       // Draw button with blue text and underline
			fInWorkPane:1,          // in a workpane command bar
			// NOTE: InCaption, InWPActiveCaption and InWPInactiveCaption represent 4 possible
			// states, so really we only need 2 bits, not 3 for these flags, and check
			// their combinations.  When the day comes we are running low on bits, we can
			// make this change.
			fInCaption:1,           // in a caption bar (like a close button)
			fInWPActiveCaption:1,   // in an active WP caption bar
			fInWPInactiveCaption:1, // in an inactiave WP caption bar
			fNestedGcc:1,           // Non-SDM GCCs nested in panes: changes focus drawing
			fWideItems:1,           // Used for quick-customize menu items
			fSecondPass:1,          // Used for quick-customize menu items
			fDroppedMenuControl:1,  // Control is a menu control, and has its menu dropped
			fGccUseNewLook:1,       // Use the new look for gcc's
			fNothingFancy:1,        // Override desaturation, shadows & popping up, and enables bitmap scaling
			fSwatch:1,              // So drawing knows which background color to use
			fFadeControl:1,         // Fade the control out using info in vai if possible
			fDrawGradient:1,		// Draw the control 'transparently' by drawing proper background gradient
			fInWPNavBar:1,          // Control is in the WP navigation bar
			fInQuickCustomize:1,    // Control is in the quick-customize menu
			fDrawTransparent:1,  // Draw with true trasnparency, Slower than DrawGradient
			fUnused3:1,             // Was fHiResIcon
			fUseAccAnyways:1,       // Use Accelerators, even if on context menu
			fDroppedFromMainMenu:1, // Dropped from the main menu (but not fInMainMenu which means in the menubar itself)
			fUnused2:2;
			};
		};
	} MSOCEO;       // ctlenv options

MSOAPI_(void) MsoCbvDrawIcon(HDC hdc, MSOCEO *pceo, int tcid, HANDLE hdibCustom,
			HANDLE hmaskCustom, RECT *prc, COLORREF crBkgd, int tbbsx, UINT fuFlags);

#endif
