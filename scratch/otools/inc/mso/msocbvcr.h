#pragma once

/****************************************************************************
	Msocbvcr.h

	Owner: MCirello, DMorton, MattiasA, KurtJ
	Copyright (c) 2002 Microsoft Corporation

	This file contains the exported colors used in the
	Office Toolbars and related components and is included in msotb.h, but 
	is kept separate so that projects that need to know about these
	colors, but cannot include generated headers from MSO can include this 
	file.

	NOTE: If you _delete_ a color from this table, you must update NetUI's
	nuicolors.lst to prevent breaks to the NetUI build.  Both deletes and
	adds should edit nuicolors.lst and %NETUI%\test\core\parser.cpp to
	ensure that NetUI continues to use proper colors in NetUI.dll.

	NOTE: If you _delete_ a color that's already shipped we'll come kick
	your ass since you'll break old apps that use a new build of MSO but
	they themselves have not been recomplied.  Messing with the order of
	shipped colors is also an ass kicking offense.
****************************************************************************/

#ifndef MSOCBVCR_H
#define MSOCBVCR_H

typedef enum _msocbvcr
	{
	/////////////////////////////////////////////////////////////////////////////
	// Office 10 entries.  These must not be changed.  See below.
	//
	
	msocbvcrCBBkgd = 0,
	msocbvcrCBDragHandle,
	msocbvcrCBSplitterLine,
	msocbvcrCBTitleBkgd,
	msocbvcrCBTitleText,
	msocbvcrCBBdrOuterFloating,
	msocbvcrCBBdrOuterDocked,

	msocbvcrCBCtlBkgd,
	msocbvcrCBCtlText,
	msocbvcrCBCtlTextDisabled,
	msocbvcrCBCtlBkgdMouseOver,
	msocbvcrCBCtlBdrMouseOver,
	msocbvcrCBCtlTextMouseOver,
	msocbvcrCBCtlBkgdMouseDown,
	msocbvcrCBCtlBdrMouseDown,
	msocbvcrCBCtlTextMouseDown,
	msocbvcrCBCtlBkgdSelected,
	msocbvcrCBCtlBdrSelected,
	msocbvcrCBCtlBkgdSelectedMouseOver,
	msocbvcrCBCtlBdrSelectedMouseOver,
	msocbvcrCBCtlBkgdLight,
	msocbvcrCBCtlTextLight,

	msocbvcrCBMainMenuBkgd,
	msocbvcrCBMainMenuCtlBdrGlow,
	msocbvcrCBMenuBkgd,
	msocbvcrCBMenuCtlText,
	msocbvcrCBMenuCtlTextDisabled,
	msocbvcrCBMenuBdrOuter,
	msocbvcrCBMenuIconBkgd,
	msocbvcrCBMenuIconBkgdDropped,
	msocbvcrCBMenuSplitArrow,
	msocbvcrCBMenuCtlBkgdMouseOver,
	msocbvcrCBMenuCtlBdrMouseOver,
	msocbvcrCBMenuCtlTextMouseOver,

	/////////////////////////////////////////////////////////////////////////////
	// New entries must be added below this comment so as not to change any
	// entries from Office 10 (or you'll break backward-compatibility for
	// Office apps that are not getting recompiled).
	//

	msocbvcrCBSplitterLineLight,
	msocbvcrCBShadow,
	msocbvcrCBLabelBkgnd,
	msocbvcrCBIconDisabledLight,
	msocbvcrCBIconDisabledDark,
	msocbvcrCBLowColorIconDisabled,

	msocbvcrCBGradMainMenuHorzBegin,
	msocbvcrCBGradMainMenuHorzEnd,
	msocbvcrCBGradVertBegin,
	msocbvcrCBGradVertMiddle,
	msocbvcrCBGradVertEnd,
	msocbvcrCBGradOptionsBegin,
	msocbvcrCBGradOptionsMiddle,
	msocbvcrCBGradOptionsEnd,
	msocbvcrCBGradMenuTitleBkgdBegin,
	msocbvcrCBGradMenuTitleBkgdEnd,
	msocbvcrCBGradMenuIconBkgdBegin,
	msocbvcrCBGradMenuIconBkgdMiddle,
	msocbvcrCBGradMenuIconBkgdEnd,
	msocbvcrCBGradMenuCtlMouseOverBegin,
	msocbvcrCBGradMenuCtlMouseOverMiddle,
	msocbvcrCBGradMenuCtlMouseOverEnd,
	msocbvcrCBGradOptionsSelectedBegin,
	msocbvcrCBGradOptionsSelectedMiddle,
	msocbvcrCBGradOptionsSelectedEnd,
	msocbvcrCBGradOptionsMouseOverBegin,
	msocbvcrCBGradOptionsMouseOverMiddle,
	msocbvcrCBGradOptionsMouseOverEnd,
	msocbvcrCBGradSelectedBegin,
	msocbvcrCBGradSelectedMiddle,
	msocbvcrCBGradSelectedEnd,
	msocbvcrCBGradMouseOverBegin,
	msocbvcrCBGradMouseOverMiddle,
	msocbvcrCBGradMouseOverEnd,
	msocbvcrCBGradMouseDownBegin,
	msocbvcrCBGradMouseDownMiddle,
	msocbvcrCBGradMouseDownEnd,

	msocbvcrNetLookBkgnd,
	msocbvcrCBDockSeparatorLine,
	msocbvcrCBDropDownArrow,
	msocbvcrCBDragHandleShadow,

	msocbvcrPlacesBarBkgd,

	msocbvcrVSToolWindowBorder,

	msocbvcrMax, // NOT a valid color, just a place holder

	/////////////////////////////////////////////////////////////////////////////
	// Only the colors that must always match with other colors should be placed here....
	//
	msocbvcrCBBdrInnerFloating = msocbvcrCBBkgd,
	msocbvcrCBBdrInnerDocked = msocbvcrCBBkgd,
	msocbvcrCBCtlTextSelected = msocbvcrCBCtlTextMouseOver,
	msocbvcrCBCtlTextSelectedMouseOver = msocbvcrCBCtlTextMouseDown,
	msocbvcrCBMenuCtlBkgd = msocbvcrCBMenuBkgd,
	msocbvcrCBMenuBdrInner = msocbvcrCBMenuBkgd,
	msocbvcrCBGradHorzBegin = msocbvcrCBGradVertBegin,
	msocbvcrCBGradHorzMiddle = msocbvcrCBGradVertMiddle,
	msocbvcrCBGradHorzEnd = msocbvcrCBGradVertEnd,
	msocbvcrSDFeedback = msocbvcrCBCtlBkgdMouseDown,
	msocbvcrCBCtlDisabled = msocbvcrCBCtlTextDisabled,
	msocbvcrCBGradSelectedMouseOverBegin = msocbvcrCBGradMouseDownBegin,
	msocbvcrCBGradSelectedMouseOverMiddle = msocbvcrCBGradMouseDownMiddle,
	msocbvcrCBGradSelectedMouseOverEnd = msocbvcrCBGradMouseDownEnd,
	
	} MSOCBVCR;

#endif // MSOCBVCR_H
