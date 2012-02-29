
#pragma once

/****************************************************************************
	MsoUI.h

	Owner: JeffJo
	Copyright (c) 1997 Microsoft Corporation

	This file contains the exported interfaces and declarations for
	Office UI and related components.
****************************************************************************/

#ifndef MSOUI_H
#define MSOUI_H


// Fake SDI Commands
enum
	{
	msouisdiCreate = 0,       // Create a new document
	msouisdiRename,           // An existing document has had a name change
	msouisdiActivate,         // A document has been activated and brought to front
	msouisdiDestroy,          // Destroy the document window
	msouisdiQuitting,         // Exiting the app, more commands may follow, but stop updating the task bar
	msouisdiShow,             // The windows is being Shown or Hidden
	msouisdiActivateApp,      // The app is being activated, not a specific document
	msouisdiDisableSDI,       // Destroys all FakeSDI tabs, subsequent calls ignored
	msouisdiReenableSDI,      // Renables FakeSDI, does not magically recreate tabs
	                          //    returns TRUE iff SDI was previously disabled
	msouisdiDisable1stStart,  // Prevents FakeSDI from starting when the 2nd doc
	                          // window is created.  Queues up the tab info until
	                          // Renable1stStart is called.  Useful when loading a doc
	                          // that will replace a default doc, but two docs may
	                          // temporarily be created. CANNOT be the first call into
	                          // MsoFDoFakeSDI, 1st structure must already be created.
	msouisdiReenable1stStart, // Reenables FakeSDI and forces creation if 2 or more
	                          // docs are currently cached.  If this is the case,
	                          // hwnd must be passed to specify the tab to be activated.
	msouisdiShowAppSlab,      // Do NOT attempt to hide the app icon when in FakeSDI
	                          // mode.  hwnd must be a top level app hwnd.  At present
	                          // there is no way to disable this once turned on.  This
	                          // setting affects all FakeSDI windows in the instance.
	};

enum
	{
	msouicaChildActivated = 3, // Child was Activated, do no further work
	msouicaContinue,           // Message was acknowledged, but please continue
	msouicaDontActivate,       // Child can't be activated
	};

enum
	{
	msouiecExecuteTCID = 1,    // Execute TCID
	msouiecExecuteDlg,    	   // Execute TCID that brings up a Dialog
	msouiecSDMCtl,	     	   	// Executes an SDM dialog control
	msouiecWinCtl,             // Executes a Windows Control
	msouiecSelectSDMTab,
	msouiecSelectWinTab,
	};

enum
	{
	msouierrSuccess = 0,
	msouierrFail,
	msouierrNotValidId,
	msouierrNoDialog,
	msouierrWrongDialog,
	msouierrAdminDisabled,
	msouierrDisabled,
	msouierrOn,
	msouierrOff,
	msouierrUnknown,
	msouierrAppModal,
	};


#if 0 //$[VSMSO]
MSOAPI_(BOOL) MsoFDoFakeSDI(int uisdi, HWND hwnd, HMSOINST hinst, const WCHAR *wtzCaption);
#endif //$[VSMSO]
#endif // MSOUI_H
