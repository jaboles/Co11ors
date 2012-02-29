// List of MSO-owned dialogs with help support.

// All additions to this list must retain the list's sorted order.

// This list describes the mapping from Help ID to CHM file (via gimme ID) and topic ID for all
// MSO-owned dialogs.  App-specific IDs should be added in their respective files.
// To add help support to an SDM, CDL or NetUI dialog, simply give the dialog an appropriate
// Help ID, and enter the mapping in this list.  Doing this will provide the following:
//    1) A "?" button in the title bar
//    2) "F1" handling
// When activated, these will bring up the given help topic in the given CHM file via MsoHtmlHelp.

#include "msouser.h"
#include "helpid.h"


MSOHELPIDSUPPORTARRAY vHelpIdSupportCommon =
{
	// Help ID                // CHM ID                    // Topic ID    // Value     // Description
	{hidDlgCBCustomizeTab0,   msoqfidOfficeMainHelp,       553713665},    // 2000      // Common -> Tools   -> Customize
	{hidDlgWebInsertHLink,    msoqfidOfficeMainHelp,       553713702},    // 5000      // Common -> Insert  -> Hyperlink
	{hidDlgAVSecurity,        msoqfidOfficeMainHelp,       553714167},    // 6001      // Common -> Tools   -> Macros -> Security
	{hidDlgDrm,               msoqfidOfficeMainHelp,       553714122},    // 28000     // Common -> File    -> Permission (several dialogs)

	#if DEBUG
	{987654,                  msoqfidOfficeMainHelp,       3083525},      // 987654    // Debug dialog entry for testing Help ID support
	#endif
};


MSOHELPIDSUPPORTLIST vHelpIdSupportListCommon =
	MAKEMSOHELPIDSUPPORTLIST(vHelpIdSupportCommon);

