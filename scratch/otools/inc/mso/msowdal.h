#pragma once
/*****************************************************************************

   Module  : MSOWDAL.h
   Prefix  : None
   Owner   : SteveBre
   Security Owner: ctran

******************************************************************************

   Win32-specific Dialog AutoLayout callbacks and frame classes

*****************************************************************************/

#include "msodal.h"

using namespace DlgAutoLayout;


/***************************************************************************
   Some exported resource ids
***************************************************************************/

#ifndef IDAPPLY 
#define IDAPPLY 0x3021
#endif

#ifndef ID_BOGUS
#define ID_BOGUS 0xFFFE
#endif


/***************************************************************************/
// Constants
/***************************************************************************/

// Text to be used in the control which disables/enables AutoLayout for a
// given dialog.  The control must be of the form:
//
//	LTEXT ENABLE_OR_DISABLE_FLAG, CONTROL_ID, 0, 0, 0, 0, NOT WS_VISIBLE,
//
// where ENABLE_OR_DISABLE_FLAG is one of WZ_DAL_ON or WZ_DAL_OFF and
// CONTROL_ID is an ID given to the control.
#define WZ_DAL_ON    L"DAL=On"
#define WZ_DAL_OFF   L"DAL=Off"
#define MAX_DUMMY_DAL_TEXT \
	max(sizeof(WZ_DAL_ON) / sizeof(WCHAR), sizeof(WZ_DAL_OFF) / sizeof(WCHAR))


/***************************************************************************/

class IMsoWAutoLayoutDialog;

/***************************************************************************/
#define MSOWFRAME_FRIENDS \
   friend class MsoWFrame; \
      friend class MsoWCompositeFrame; \
         friend class MsoWContainerFrame; \
            friend class MsoWCtrlContainerFrame; \
            friend class MsoWGroupBoxFrame; \
            friend class MsoWOverlapFrame; \
            friend class MsoWTabCtrlFrame; \
         friend class MsoWFlowFrame; \
         friend class MsoWVertFlowFrame; \
         friend class MsoWHorizFrame; \
         friend class MsoWTableFrame; \
         friend class MsoWVertFrame; \
      friend class MsoWLeafFrame; \
         friend class MsoWCtrlFrame; \
         friend class MsoWListEditFrame; \
         friend class MsoWNumInputFrame; \
   friend class MsoWFrameEqualizer;

/***************************************************************************/
class MsoWFrame
{
public:
   FRAME_FORWARDS(m_rFrame, .)

protected:
   MsoWFrame( Frame& rFrame ) : m_rFrame( rFrame ) {}
   // rFrame has not been constructed yet, so do not use it except for its address (to initialize m_rFrame)

private:
   Frame& m_rFrame;
   DebugOnly( int** m_autoExpandHelper; )
   MSOWFRAME_FRIENDS
}; // MsoWFrame

/***************************************************************************/
class MsoWCompositeFrame : public MsoWFrame
{
public:
   COMPOSITEFRAME_FORWARDS_NEW(static_cast<CompositeFrame&>( m_rFrame ), .)

   // Involve WinDAL-specific classes:
   MsoWCompositeFrame& operator<<( MsoWFrame& frame ) { static_cast<CompositeFrame&>( m_rFrame ).operator<<( frame.m_rFrame ); return *this; }

protected:
   MsoWCompositeFrame( CompositeFrame& rFrame ) : MsoWFrame( rFrame ) {}
   // rFrame has not been constructed yet, so do not use it except for its address (to initialize m_rFrame)
}; // MsoWCompositeFrame

/***************************************************************************/
class MsoWHorizFrame : public MsoWCompositeFrame
{
public:
   MSOCTOR MsoWHorizFrame( IMsoWAutoLayoutDialog& dlg, MsoWCompositeFrame* pParentFrame );
   HORIZFRAME_FORWARDS_NEW(static_cast<HorizFrame&>( m_rFrame ), .)

private:
   HorizFrame m_frame;
}; // MsoWHorizFrame

/***************************************************************************/
class MsoWVertFrame : public MsoWCompositeFrame
{
public:
   MSOCTOR MsoWVertFrame( IMsoWAutoLayoutDialog& dlg, MsoWCompositeFrame* pParentFrame );
   VERTFRAME_FORWARDS_NEW(static_cast<VertFrame&>( m_rFrame ), .)

private:
   VertFrame m_frame;
}; // MsoWVertFrame

/***************************************************************************/
class MsoWFlowFrame : public MsoWCompositeFrame
{
public:
   MSOCTOR MsoWFlowFrame( IMsoWAutoLayoutDialog& dlg, MsoWCompositeFrame* pParentFrame );
   FLOWFRAME_FORWARDS_NEW(static_cast<FlowFrame&>( m_rFrame ), .)

   void BindChildPair( MsoWFrame& firstChild )
      { static_cast<FlowFrame&>( m_rFrame ).BindChildPair( firstChild.m_rFrame ); }
   int GetRowOccupiedByChild( MsoWFrame& child )
      { static_cast<FlowFrame&>( m_rFrame ).GetRowOccupiedByChild( child.m_rFrame ); }
   void ExpandChildToFillRowVert( MsoWFrame& child )
      { static_cast<FlowFrame&>( m_rFrame ).ExpandChildToFillRowVert( child.m_rFrame ); }

private:
   FlowFrame m_frame;
}; // MsoWFlowFrame

/***************************************************************************/
class MsoWVertFlowFrame : public MsoWCompositeFrame
{
public:
   MSOCTOR MsoWVertFlowFrame( IMsoWAutoLayoutDialog& dlg, MsoWCompositeFrame* pParentFrame );
   VERTFLOWFRAME_FORWARDS_NEW(static_cast<VertFlowFrame&>( m_rFrame ), .)

   void BindChildPair( MsoWFrame& firstChild )
      { static_cast<VertFlowFrame&>( m_rFrame ).BindChildPair( firstChild.m_rFrame ); }
   int GetColOccupiedByChild( MsoWFrame& child )
      { static_cast<VertFlowFrame&>( m_rFrame ).GetColOccupiedByChild( child.m_rFrame ); }
   void ExpandChildToFillColHoriz( MsoWFrame& child )
      { static_cast<VertFlowFrame&>( m_rFrame ).ExpandChildToFillColHoriz( child.m_rFrame ); }

private:
   VertFlowFrame m_frame;
}; // MsoWVertFlowFrame

/***************************************************************************/
class MsoWTableFrame : public MsoWCompositeFrame
{
public:
   MSOCTOR MsoWTableFrame( IMsoWAutoLayoutDialog& dlg, int columns, int rows, MsoWCompositeFrame* pParentFrame );
   TABLEFRAME_FORWARDS_NEW(static_cast<TableFrame&>( m_rFrame ), .)

   // Involve WinDAL-specific classes:
   MsoWTableFrame& operator<<( MsoWFrame& frame ) { static_cast<TableFrame&>( m_rFrame ).operator<<( frame.m_rFrame ); return *this; }
   // Must return TableFrame, not CompositeFrame, so both overloads can be used in a chain.
   MsoWTableFrame& operator<<( int zero ) { static_cast<TableFrame&>( m_rFrame ).operator<<( zero ); return *this; }
   void InsertChildAtCell ( MsoWFrame& child, int iCol, int iRow, int cCols = 1, int cRows = 1)
      { static_cast<TableFrame&>( m_rFrame ).InsertChildAtCell( child.m_rFrame, iCol, iRow, cCols, cRows ); }
   void ExpandChildToFillHoriz( MsoWFrame& child )
      { static_cast<TableFrame&>( m_rFrame ).ExpandChildToFillHoriz( child.m_rFrame ); }
   void ExpandChildToFillVert( MsoWFrame& child )
      { static_cast<TableFrame&>( m_rFrame ).ExpandChildToFillVert( child.m_rFrame ); }
   void ExpandChildToFillCellHoriz( MsoWFrame& child )
      { static_cast<TableFrame&>( m_rFrame ).ExpandChildToFillCellHoriz( child.m_rFrame ); }
   void ExpandChildToFillCellVert( MsoWFrame& child )
      { static_cast<TableFrame&>( m_rFrame ).ExpandChildToFillCellVert( child.m_rFrame ); }
   void AddGroupBox( WORD groupCtrlData, int left, int top, int right, int bottom, bool fNeverHasText = false )
      { static_cast<TableFrame&>( m_rFrame ).AddGroupBox( (void*)groupCtrlData, left, top, right, bottom, fNeverHasText ); }
   void AddGroupBox2Ctrls( WORD textCtrlData, WORD lineCtrlData, int left, int top, int right, int bottom, bool fNeverHasText /* = false */ )
      { static_cast<TableFrame&>( m_rFrame ).AddGroupBox2Ctrls( (void*)textCtrlData, (void*)lineCtrlData, left, top, right, bottom, fNeverHasText ); }
   
private:
   TableFrame m_frame;
}; // MsoWTableFrame

/***************************************************************************/
class MsoWContainerFrame : public MsoWCompositeFrame
{
public:
   CONTAINERFRAME_FORWARDS_NEW(static_cast<ContainerFrame&>( m_rFrame ), .)

protected:
   MsoWContainerFrame( ContainerFrame& rFrame ) : MsoWCompositeFrame( rFrame ) {}
   // rFrame has not been constructed yet, so do not use it except for its address (to initialize m_rFrame)
}; // MsoWContainerFrame

/***************************************************************************/
class MsoWCtrlContainerFrame : public MsoWContainerFrame
{
public:
   MSOCTOR MsoWCtrlContainerFrame( IMsoWAutoLayoutDialog& dlg, WORD ctrlID, MsoWCompositeFrame* pParentFrame );
   CTRLCONTAINERFRAME_FORWARDS_NEW(static_cast<CtrlContainerFrame&>( m_rFrame ), .)

private:
   CtrlContainerFrame m_frame;
}; // MsoWCtrlContainerFrame

/***************************************************************************/
class MsoWGroupBoxFrame : public MsoWContainerFrame
{
public:
   MSOCTOR MsoWGroupBoxFrame( IMsoWAutoLayoutDialog& dlg, WORD groupBoxID, MsoWCompositeFrame* pParentFrame );
   MSOCTOR MsoWGroupBoxFrame( IMsoWAutoLayoutDialog& dlg, WORD textCtrlID, WORD lineCtrlID, MsoWCompositeFrame* pParentFrame );
   GROUPBOXFRAME_FORWARDS_NEW(static_cast<GroupBoxFrame&>( m_rFrame ), .)

private:
   GroupBoxFrame m_frame;
}; // MsoWGroupBoxFrame

/***************************************************************************/
class MsoWOverlapFrame : public MsoWContainerFrame
{
public:
   MSOCTOR MsoWOverlapFrame( IMsoWAutoLayoutDialog& dlg, MsoWCompositeFrame* pParentFrame );
   OVERLAPFRAME_FORWARDS_NEW(static_cast<OverlapFrame&>( m_rFrame ), .)

private:
   OverlapFrame m_frame;
}; // MsoWOverlapFrame

/***************************************************************************/
class MsoWTabCtrlFrame : public MsoWContainerFrame
{
public:
   MSOCTOR MsoWTabCtrlFrame( IMsoWAutoLayoutDialog& dlg, WORD tabCtrlID, MsoWCompositeFrame* pParentFrame );
   TABCTRLFRAME_FORWARDS_NEW(static_cast<TabCtrlFrame&>( m_rFrame ), .)

private:
   TabCtrlFrame m_frame;
}; // MsoWTabCtrlFrame

/***************************************************************************/
class MsoWLeafFrame : public MsoWFrame
{
public:
   LEAFFRAME_FORWARDS_NEW(static_cast<LeafFrame&>( m_rFrame ), .)

   // Involve WinDAL-specific types:
   void AddLeftString( WORD strID ) { static_cast<LeafFrame&>( m_rFrame ).AddLeftString( (void*)strID ); }
   void AddTopString( WORD strID ) { static_cast<LeafFrame&>( m_rFrame ).AddTopString( (void*)strID ); }
   void AddRightString( WORD strID ) { static_cast<LeafFrame&>( m_rFrame ).AddRightString( (void*)strID ); }

protected:
   MsoWLeafFrame( LeafFrame& rFrame ) : MsoWFrame( rFrame ) {}
   // rFrame has not been constructed yet, so do not use it except for its address (to initialize m_rFrame)
}; // MsoWLeafFrame

/***************************************************************************/
class MsoWCtrlFrame : public MsoWLeafFrame
{
public:
   MSOCTOR MsoWCtrlFrame( IMsoWAutoLayoutDialog& dlg, WORD ctrlID, MsoWCompositeFrame* pParentFrame );
   CTRLFRAME_FORWARDS_NEW(static_cast<CtrlFrame&>( m_rFrame ), .)

private:
   CtrlFrame m_frame;
}; // MsoWCtrlFrame

/***************************************************************************/
class MsoWNumInputFrame : public MsoWLeafFrame
{
public:
   //MSOCTOR MsoWNumInputFrame( IMsoWAutoLayoutDialog& dlg, WORD ctrlID, MsoWCompositeFrame* pParentFrame ); // Not needed for Win32 dialogs
   MSOCTOR MsoWNumInputFrame( IMsoWAutoLayoutDialog& dlg, WORD editStrID, WORD spinnerID, MsoWCompositeFrame* pParentFrame );
   NUMINPUTFRAME_FORWARDS_NEW(static_cast<NumInputFrame&>( m_rFrame ), .)

private:
   NumInputFrame m_frame;
}; // MsoWNumInputFrame

/***************************************************************************/
class MsoWListEditFrame : public MsoWLeafFrame
{
public:
   MSOCTOR MsoWListEditFrame( IMsoWAutoLayoutDialog& dlg, WORD editStrID, WORD listBoxID, MsoWCompositeFrame* pParentFrame );
   LISTEDITFRAME_FORWARDS_NEW(static_cast<ListEditFrame&>( m_rFrame ), .)

private:
   ListEditFrame m_frame;
}; // MsoWListEditFrame

/***************************************************************************/
class MsoWFrameEqualizer
{
public:
   MSOCTOR MsoWFrameEqualizer( IMsoWAutoLayoutDialog& dlg );
   FRAMEEQUALIZER_FORWARDS(m_rEqualizer, .)

   // Involve WinDAL-specific classes:
   MsoWFrameEqualizer& operator<<( MsoWFrame& frame ) { m_rEqualizer.operator<<( static_cast<Frame&>( frame.m_rFrame ) ); return *this; }

private:
   FrameEqualizer m_equalizer;
   FrameEqualizer& m_rEqualizer;
}; // MsoWFrameEqualizer


/***************************************************************************/
class IMsoWAutoLayoutDialogUser
/***************************************************************************/
{
public:
   MSOMETHOD_(BOOL,FDlgIsModeless)() { return FALSE; }
   MSOMETHOD_(HFONT,GetDlgFontOverride)() { return NULL; }
   MSOMETHOD_(BOOL,FGetCtrlTypeOverride)( WORD /*ctrlID*/, MsoDALCtrlType& /*type*/ )
      { return FALSE; }
   MSOMETHOD_(BOOL,FGetCustomDefaultCtrlSize)( WORD /*ctrlID*/, int& /*width*/, int& /*height*/ )
      { return FALSE; }

#if( MSODAL_DEBUG_TOOLS )
   MSOMETHOD_(int,GetDlgID)() { return 1; }
#endif // MSODAL_DEBUG_TOOLS
};


/***************************************************************************/
// Initializers and stuff
/***************************************************************************/
MSOAPI_(BOOL) MsoFCreateIMsoWAutoLayoutDialog( IMsoWAutoLayoutDialog **ppDlg, HWND hwndDlg,
                                               IMsoWAutoLayoutDialogUser* pDlgUser = NULL );
MSOAPI_(void) MsoDestroyIMsoWAutoLayoutDialog( IMsoWAutoLayoutDialog *pDlg );
MSOAPI_(BOOL) MsoWIsAutoLayoutEnabled(HWND hwndDlg, int dummyID);
MSOAPI_(HWND) MsoWTabCtrlIndexToHwnd(HWND hwndTabCtrl, int index);


#if 0 // Do we need to support C clients?  Do C apps have Win32 (i.e., non-SDM) dialogs?
MSOAPI_(void) MsoWHorizFrameInit(MsoWHorizFrame *This, IMsoWAutoLayoutDialog *pDlg,
	MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWVertFrameInit(MsoWVertFrame *This, IMsoWAutoLayoutDialog *pDlg,
	MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWFlowFrameInit(MsoWFlowFrame *This, IMsoWAutoLayoutDialog *pDlg,
	MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWVertFlowFrameInit(MsoWVertFlowFrame *This, IMsoWAutoLayoutDialog *pDlg,
	MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWTableFrameInit(MsoWTableFrame *This, IMsoWAutoLayoutDialog *pDlg,
	int columns, int rows, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWCtrlContainerFrame(MsoWCtrlContainerFrame *This,
	IMsoWAutoLayoutDialog *pDlg, WORD ctrlID, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWGroupBoxFrameInit(MsoWGroupBoxFrame *This, 
	IMsoWAutoLayoutDialog *pDlg, WORD groupBoxID, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWGroupBoxFrameInit2Ctrls(MsoWGroupBoxFrame *This, 
	IMsoWAutoLayoutDialog *pDlg, WORD textCtrlID, WORD lineCtrlID, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWOverlapFrameInit(MsoWOverlapFrame *This,
	IMsoWAutoLayoutDialog *pDlg, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWTabCtrlFrameInit(MsoWTabCtrlFrame *This,
	IMsoWAutoLayoutDialog *pDlg, WORD tabCtrlID, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWCtrlFrameInit(MsoWCtrlFrame *This, IMsoWAutoLayoutDialog *pDlg,
	WORD ctrlID, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWNumInputFrameInit(MsoWNumInputFrame *This,
	IMsoWAutoLayoutDialog *pDlg, WORD editStrID, WORD spinnerID, MsoWCompositeFrame *pParentFrame);
MSOAPI_(void) MsoWListEditFrameInit(MsoWListEditFrame *This,
	IMsoWAutoLayoutDialog *pDlg, WORD editStrID, WORD listBoxID, MsoWCompositeFrame *pParentFrame);

MSOAPI_(void) MsoWFrameEqualizerInit(MsoWFrameEqualizer *This,
	IMsoWAutoLayoutDialog *pDlg);
#endif

#if( DEBUG )
// Move & resize all control windows. This catches AutoLayout errors with controls that don't specify a size.
// Should be called on each tabsheet of a tab dialog, since it currently doesn't recurse. 
// Debug-only for performance reasons.
MSOAPI_(void) MsoWDALResetCtrlRects(HWND hwndDlg, int idDALSwitch);
#endif // DEBUG


/***************************************************************************/
class IMsoWAutoLayoutDialog : public IMsoAutoLayoutDialog
/***************************************************************************/
// IMsoWAutoLayoutDialog is defined publicly so that it can be passed to functions
// taking IMsoAutoLayoutDialog.
{
#define IMsoWAutoLayoutDialog_Inherited IMsoAutoLayoutDialog

private: // Yes, private.  Use MsoFCreateIMsoWAutoLayoutDialog() and MsoDestroyIMsoWAutoLayoutDialog().
   IMsoWAutoLayoutDialog( HWND hwnd, IMsoWAutoLayoutDialogUser* pDlgUser = NULL );
   ~IMsoWAutoLayoutDialog();

public:
   MSODALOVERRIDEDEBUGMETHOD

   // Dialog:
   MSOOVERRIDEMETHOD_(bool) IsModeless() const; // override
   MSOOVERRIDEMETHOD_(bool) IsRightToLeft() const; // override
   MSOOVERRIDEMETHOD_(LID ) LanguageDisplayed() const; // override
   MSOOVERRIDEMETHOD_(void) GetBaseUnits( int& unitsHoriz, int& unitsVert, void* ctrlData, bool fIsDialog ) const; // override
   MSOOVERRIDEMETHOD_(RECT) GetCtrlParentClientRect( void* ctrlData ) const; // override
   MSOOVERRIDEMETHOD_(int ) WidthOfCaptionPlusSysBtns() const; // override
   MSOOVERRIDEMETHOD_(void) SetDlgClientSize( int width, int height ); // override
   MSOOVERRIDEMETHOD_(void) SetDlgPosition(); // override
   MSOOVERRIDEMETHOD_(HWND) GetDlgHwnd(); // override

   // Controls:
   MSOOVERRIDEMETHOD_(MsoDALCtrlType) GetCtrlType( void* ctrlData ) const; // override
   MSOOVERRIDEMETHOD_(void) GetBitmapSize( void* bitmapCtrlData, int& width, int& height ) const; // override
   MSOOVERRIDEMETHOD_(void) ComputeMinSizeForBitmapBtn( void* bitmapBtnData, int& width, int& height ) const; // override
   MSOOVERRIDEMETHOD_(bool) FGetCustomDefaultCtrlSize( void* /*ctrlData*/, int& /*width*/, int& /*height*/ ) const; // override
   MSOOVERRIDEMETHOD_(RECT) GetCtrlRect( void* ctrlData ) const; // override
   MSOOVERRIDEMETHOD_(void) SetCtrlRect( void* ctrlData, const RECT& r ); // override

   // List controls:
   MSOOVERRIDEMETHOD_(int ) MaxItemWidth( void* ctrlData ) const; // override
   MSOOVERRIDEMETHOD_(int ) FixedClosedDropDownHeight( void* ctrlData ); // override
   MSOOVERRIDEMETHOD_(int ) TotalOpenListCtrlHeightForNLines( void* ctrlData, int lines, int closedDropDownHeight ); // override
   MSOOVERRIDEMETHOD_(void) Set2PartCtrlRects( void* ctrlData, const RECT& firstPartRect, const RECT& totalRect ); // override

   // Tab controls:
   MSOOVERRIDEMETHOD_(POINT) GetWidestTabItemSize( void* tabCtrlData ) const; // override
   MSOOVERRIDEMETHOD_(int  ) GetTabRowCount( void* tabCtrlData ) const; // override
   MSOOVERRIDEMETHOD_(bool ) FTabSheetsAreSeparateWindows( void* tabCtrlData ) const; // override
   MSOOVERRIDEMETHOD_(void ) SetRectAllSheetsOfTabctrl( void* tabCtrlData, const RECT& sheetRect ); // override

   // Control text:
   MSOOVERRIDEMETHOD_(void) GetCtrlText( void* ctrlData, _Out_cap_(nMaxCount) WCHAR* wzText, int nMaxCount ) const; // override
   MSOOVERRIDEMETHOD_(void) MeasureCtrlText( void* ctrlData, int& width, int& height ) const; // override
   MSOOVERRIDEMETHOD_(void) MeasureCtrlTextWithTargetWidth(void* ctrlData, int& width,
		int& height, bool fWidenForLongWord) const; // override
   MSOOVERRIDEMETHOD_(void) MeasureText( const WCHAR* wzText, void* ctrlData, int& width, int& height ) const; // override
   MSOOVERRIDEMETHOD_(void) ShrinkCtrlTextToWidth( void* ctrlData, int maxWidth ); // override

   // Miscellaneous:
#if( DEMOAPP ) // For testing purposes only
   MSOMETHODIMP_(bool) CompositeFramesExpandByDefault() const; // override
#endif // DEMOAPP

// Debug-only:

#if( MSODAL_DEBUG_TOOLS )
   // Frame outlines and Logo:
   MSOOVERRIDEMETHOD_(int ) GetCurTabIndex( void* tabCtrlData ) const; // override
   MSOOVERRIDEMETHOD_(bool) SupportsFrameOutlines() const; // override
   MSOOVERRIDEMETHOD_(void) GetShortDlgMgrName( _Out_z_cap_(cchDlgMgr) WCHAR* wzDlgMgr, int cchDlgMgr ) const; // override

   // String length changing:
   MSOOVERRIDEMETHOD_(void) SetCtrlText( void* ctrlData, const WCHAR* wzText ); // override
   MSOOVERRIDEMETHOD_(void) GetDlgCaption(_Out_cap_(nMaxCount)  WCHAR* wzText, int nMaxCount ) const; // override
   MSOOVERRIDEMETHOD_(void) SetDlgCaption( const WCHAR* wzText ); // override
   MSOOVERRIDEMETHOD_(bool) SupportsStrLenChanging() const; // override
   MSOOVERRIDEMETHOD_(INT_PTR) GetCtrlID( void* ctrlData ) const; // override
   MSOOVERRIDEMETHOD_(int ) GetDlgID() const; // override

#endif // MSODAL_DEBUG_TOOLS

private:
	static void SetListBoxCtrlRect(HWND hwndListBox, const RECT& r);
	void SetComboBoxCtrl(HWND hwndComboBox, const RECT& r);
	static void SetSpinnerCtrlRects(HWND hwndSpinner, const RECT& r);
	
   void* m_pPrivateData;

friend MSOPUB BOOL MSOAPICALLTYPE MsoFCreateIMsoWAutoLayoutDialog(IMsoWAutoLayoutDialog**, HWND, IMsoWAutoLayoutDialogUser*);
friend MSOPUB void MSOAPICALLTYPE MsoDestroyIMsoWAutoLayoutDialog(IMsoWAutoLayoutDialog*);
}; // IMsoWAutoLayoutDialog

