/****************************************************************************
MSOMULTIMON.H

Owner: MikeBor
Copyright (c) 2002 Microsoft Corporation

This contains declarations for the Multi Monitor MSO API.
****************************************************************************/

#pragma once


/*------------------------------------------------------------------------
	MsoMoveRectToMonitor

	This function is called before a window is displayed by an app.
	It takes the parent window of the new window and the desired position
	and size of the new window.  The position and size will be modified
	so that it is positioned on the correct monitor.

	fResizableWindow should be set to TRUE if this function can modify
	the size of prcPos along with the position.

	http://officenet/Teams/UIS/Office11%20Enhanced%20Multi%20Mon%20%20Experience/Forms/AllItems.htm

	Returns true if the position was modified, false otherwise.
---------------------------------------------------------------- MikeBor -*/
MSOAPI_(BOOL) MsoMoveRectToMonitor(HWND hwndParent, RECT * prcPos, BOOL fResizableWindow);


/*------------------------------------------------------------------------
	MsoEnsureRectVisible

	This function is called before a window is displayed by an app.
	It takes the desired position and size of the window.  
	The position and size will be modified so that it is 
	entirely visible on some combination of monitors.

	Calling this is unnecessary if MsoMoveRectToMonitor is used.

	fResizableWindow should be set to TRUE if this function can modify
	the size of prcPos along with the position.

	Returns true if the position was modified, false otherwise.
---------------------------------------------------------------- MikeBor -*/
MSOAPI_(BOOL) MsoEnsureRectVisible(HWND hwndParent, RECT * prcPos, BOOL fResizableWindow);


/*------------------------------------------------------------------------
	MsoGetMonitorNumber

	Gets the monitor number (1-9) that the given window is displayed on.
	This is mainly here for the common file dialogs which store their
	position information on a per app per monitor basis.
---------------------------------------------------------------- MikeBor -*/
MSOAPI_(int) MsoGetMonitorNumber(HWND hwndParent);


