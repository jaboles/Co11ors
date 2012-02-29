// msohelp.h
//
// This file contains the interface for msohelp.lib.   It mainly consists of
// the MsoHelpW entry point and a set of defines which are commands for that
// function.
//
// Parameters:
//		hwndMain: Must be the Application frame window. This is used for automatic tiling.
//		lpszHelp: Fully qualified HtmlHelp Compressed File path (.CHM), Darwin Qualifier "foo.hlp" or URL
//		usCommand: WinHelp Command
//		dwData: WinHelp Data
//
// Returns: True for success, False for failure or for unsupported command
//
#ifndef __MSHLPBAR_H_
#define __MSHLPBAR_H_


#if defined(__cplusplus)
	#define HELPEXTERN_C extern "C"
#else
	#define HELPEXTERN_C
#endif
#define HELPAPI_(t) HELPEXTERN_C t __stdcall

// Custom Command

/*-------------------------------------------------------------------------*/
// sub-commands of HELP_ASSISTANT_STATE

// make this call to notify help when the assistant visible state has changed
// MsoHelp(NULL, NULL, HELP_ASSISTANT_STATE, HELP_ASSISTANT_STATE_VISIBLE);

// make this call to notify help when the assistant dock state has changed
// MsoHelp(NULL, NULL, HELP_ASSISTANT_STATE, HELP_ASSISTANT_STATE_UNDOCKED);

// make this call to notify help when the tabs should not extend regardless of assistant state
// MsoHelp(NULL, NULL, HELP_ASSISTANT_STATE, HELP_ASSISTANT_STATE_EXTENDED_ACTIVE);

// make this call to notify help when the tabs should extend regardless of assistant state
// MsoHelp(NULL, NULL, HELP_ASSISTANT_STATE, HELP_ASSISTANT_STATE_EXTENDED_ACTIVE|HELP_ASSISTANT_STATE_EXTENDED);

// NOTE: the above flags can be combined in a single call to MsoHelp.

#define HELP_ASSISTANT_STATE_VISIBLE         (1 << 0)      // OFFICE ONLY
#define HELP_ASSISTANT_STATE_UNDOCKED        (1 << 1)      // OFFICE ONLY
#define HELP_ASSISTANT_STATE_EXTENDED_ACTIVE (1 << 2)      // OFFICE ONLY
#define HELP_ASSISTANT_STATE_EXTENDED        (1 << 3)      // OFFICE ONLY

/*-------------------------------------------------------------------------*/
// sub-commands of HELP_APPSTATE

#define HELP_APPSIZECHANGING    0x5510L
#define HELP_APPEXITSIZEMOVE    0x5511L
#define HELP_ACTIVATE           0x5512L
#define HELP_NOTILE             0x5513L

/*-------------------------------------------------------------------------*/
// main commands

// make this call to send the HelpUser interface if you want notifications from the helppane.
//	MsoHelp(hwndMain, NULL, HELP_INIT, (DWORD)g_HelpUser))

// make this call to hide or show help
//	MsoHelp(hwnd, NULL, HELP_APPSTATE, <sub-command>)

#define HELP_ASSISTANT_STATE    0x5555L      // OFFICE ONLY
#define HELP_INIT               0x5556L
#define HELP_APPSTATE           0x5557L
#define HELP_GETOBJECT          0x5558L
#define HELP_GETOBJECTSIZE      0x5559L
//                              0x555AL      // NOT USED
//                              0x555BL      // NOT USED
#define HELP_DEACTIVATE_TID     0x555CL      // OFFICE ONLY
#define HELP_READY              0x555DL
#define HELP_SEARCHONWEBENABLED 0x555EL      // OFFICE ONLY
#define HELP_SEARCHONWEB        0x555FL      // OFFICE ONLY
#define HELP_BEGIN              0x5560L
#define HELP_END                0x5561L
#define HELP_CLOSE_HOME_BALLOON 0x5562L      // OFFICE ONLY
#define HELP_AW_FILES           0x5563L
#define HELP_GETHELPAPP         0x5564L
//                              0x5565L      // NOT USED
#define HELP_INFORMOFQUERY      0x5566L
#define HELP_INIT_APPPATH		0x5567L		 // OFFICE ONLY


/*-------------------------------------------------------------------------*/

#define WM_APPSTATE WM_USER + 1

#ifndef NOHTMLHELPAPI

HELPAPI_(BOOL) MsoHelpW(HWND hwndMain, LPCWSTR lpszHelp, UINT usCommand, DWORD dwData);
HELPAPI_(BOOL) MsoHelpA(HWND hwndMain, LPCSTR lpszHelp, UINT usCommand, DWORD dwData);
HELPAPI_(void) MsoHelpSetUserData(long lUserData);
HELPAPI_(long) MsoHelpLGetUserData(void);

#endif //NOHTMLHELPAPI

#endif // __MSHLPBAR_H_
