#pragma once

/***************************************************************************/
// MSOSDAL.H
//
// Public header for the SDMDAL (Standard Dialog Manager Dialog AutoLayout)
// glue module.
//
// Contact:  CatMorr
// Security Owner: ctran
/***************************************************************************/

#ifndef MSOSDAL_H
#define MSOSDAL_H


#include "msosdm.h"
#include "msodal.h"


// This is the main dialog class, intended to be used by the AutoLayout frames.
// This class is abstract; there is no need for the client to know anything
// about its internals.  Clients can only declare pointers to this function,
// and must initialize any instance of this class using the initializer below
// (MsoFGetIMsoSAutoLayoutDialog).
#if DAL_CPLUSPLUS
	using namespace DlgAutoLayout;
	class IMsoSAutoLayoutDialog;
#else
	struct IMsoSAutoLayoutDialog;
	typedef struct IMsoSAutoLayoutDialog IMsoSAutoLayoutDialog;
#endif // DAL_CPLUSPLUS


/***************************************************************************/
// Generic SDM AutoLayout API's.
/***************************************************************************/

// Determines if a dialog is running (or should run) its AutoLayout code.
// Should always be called before running AutoLayout- (or non-AutoLayout-)
// specific code.
MSOAPI_(BOOL) MsoFRunningSDMDAL(void);

// This should be called any time you want to call AutoLayout outside of a
// dlmAutoLayout.  Never call your AutoLayoutMyDialog function directly,
// except within a dlmAutoLayout.
MSOAPI_(BOOL) MsoFAutoLayoutCurDlg(POINT ptDlgDimensions);

// Retrieve a dialog's IMsoSAutoLayoutDialog.  Should be called before
// creating a dialog's Frame tree.
MSOAPI_(BOOL) MsoFGetIMsoSAutoLayoutDialog(IMsoSAutoLayoutDialog **ppDlg);

// Clean-up.  Should be called on a dialog's dlmExit.
MSOAPI_(void) MsoSDMDALAutoLayoutCleanup(void);

// Convert pixels to DU's or DU's to pixels based on SDM's dialog base units.
MSOAPI_(int) MsoSDMPixToDU(int nPix, BOOL fHoriz);
MSOAPI_(int) MsoSDMDUToPix(int nDU, BOOL fHoriz);

// Calculate the text extent of a string on an SDM control.  Use
// MsoSDMGetTextDimensionsTmc() whenever the tmc for the control is available;
// only use MsoSDMGetTextDimensions() when the tmc is NOT available.
MSOAPI_(void) MsoSDMGetTextDimensionsTmc(TMC tmcCtrl, int *pWidth, 
	int *pHeight, DWORD grfSDMTextMeasure);
MSOAPI_(void) MsoSDMGetTextDimensions(const WCHAR *wzText, int *pWidth,
	int *pHeight, DWORD grfSDMTextMeasure);

// FLAGS to be used with MsoSDMGetTextDimensionsTmc()/MsoSDMGetTextDimensions():
// * grfUseTargetWidth:  If this parameter is set, the target width of the text
// should be passed in to the function.
#define grfUseTargetWidth    0x00000001
// * grfWrapLongWords, grfTruncateLongWords:  Ignored if grfUseTargetWidth is
// not set.  When a target width is provided, one of these flags (setting both
// makes no sense) specifies how to behave if the text contains one or more long
// words that have no spaces at which to wrap and that cannot fit within the target
// width.  If grfWrapLongWords is set, the text will be rigidly confined within the
// target width and will be measured assuming the long words wrap to mulitiple lines.
// If grfTruncateLongWords is set, the text will again be rigidly confined within
// the target width, instead measured as though the long words will be truncated
// by trailing ellipses.  If neither flag is set, the returned width will be wider
// than the target width, as much so as necessary to fit the longest unbroken word.
#define grfWrapLongWords     0x00000010
#define grfTruncateLongWords 0x00000100
// * grfBold:  Set if the text of the string is bold.
#define grfBold              0x00001000


/***************************************************************************/
// SDM Frame Classes
//
// * MsoSFrame
//		* MsoSCompositeFrame
//			* MsoSHorizFrame
//			* MsoSVertFrame
//			* MsoSFlowFrame
//			* MsoSVertFlowFrame
//			* MsoSTableFrame
//			* MsoSContainerFrame
//				* MsoSCtrlContainerFrame
//				* MsoSGroupBoxFrame
//				* MsoSOverlapFrame
//				* MsoSTabCtrlFrame
// * MsoSLeafFrame
//		* MsoSCtrlFrame
//		* MsoSNumInputFrame
//		* MsoSListEditFrame
/***************************************************************************/

// Due to a compiler optimization, we have to do some wacky things to establish
// C++ frame classes which can be allocated on the stack outside of the MSO.dll.
// (If the C++ frame classes were defined here in the typical way and a client
// outside of the MSO.dll were to declare an instance of one of these frame 
// classes on the stack, any attempts to call member functions of this class
// would result in link errors for the client.)  The solution used here is for
// the SDM frame classes to contain instances of what would be their parent
// classes (the generic frame classes in MSODAL.h) rather than subclassing off
// of those generic frame classes.  Hence all the extra goo in the classes
// defined below...

#if DAL_CPLUSPLUS
	#define DECLARE_ABSTRACT_WRAPPER_CLASS(cls) \
		DECLARE_CLASS(cls)
	#define DECLARE_ABSTRACT_WRAPPER_CLASS_(cls, basecls) \
		DECLARE_CLASS_(cls, basecls)
#else
	#define DECLARE_ABSTRACT_WRAPPER_CLASS(cls) \
		typedef struct cls { \
			void *pFrame; \
		} cls;
	#define DECLARE_ABSTRACT_WRAPPER_CLASS_(cls, basecls) \
		DECLARE_ABSTRACT_WRAPPER_CLASS(cls)
#endif // DAL_CPLUSPLUS


#if DAL_CPLUSPLUS
	#define DECLARE_CONCRETE_WRAPPER_CLASS(cls, containedcls) \
		DECLARE_CLASS(cls)
	#define DECLARE_CONCRETE_WRAPPER_CLASS_(cls, basecls, containedcls) \
		DECLARE_CLASS_(cls, basecls)
#else
	#ifdef CONST_VTABLE
		#undef CONST_VTBL
		#define CONST_VTBL const
		#define DECLARE_CONCRETE_WRAPPER_CLASS(cls, containedcls) \
			typedef struct cls { \
				void *pFrame; \
				const struct containedcls##Vtbl FAR* lpVtbl; \
				void *pImp; \
			} cls;
	#else
   		#undef CONST_VTBL
   		#define CONST_VTBL
   		#define DECLARE_CONCRETE_WRAPPER_CLASS(cls, containedcls) \
   			typedef struct cls { \
   				void *pFrame; \
   				struct containedcls##Vtbl FAR* lpVtbl; \
   				void *pImp; \
   			} cls;
   	#endif // CONST_VTABLE
	#define DECLARE_CONCRETE_WRAPPER_CLASS_(cls, basecls, containedcls) \
		DECLARE_CONCRETE_WRAPPER_CLASS(cls, containedcls)
#endif // DAL_CPLUSPLUS


// MsoSFrame
#undef  CLASSNAME
#define CLASSNAME MsoSFrame

DECLARE_ABSTRACT_WRAPPER_CLASS(MsoSFrame)
#if DAL_CPLUSPLUS
{
public:

	FRAME_FORWARDS(m_pFrame, ->)

	Frame *m_pFrame;

protected:
	MsoSFrame() {}
	void InitFramePtr(Frame *pFrame) { m_pFrame = pFrame; }
}; // MsoSFrame
#endif // DAL_CPLUSPLUS


// MsoSCompositeFrame
#undef  CLASSNAME
#define CLASSNAME MsoSCompositeFrame

DECLARE_ABSTRACT_WRAPPER_CLASS_(MsoSCompositeFrame, MsoSFrame)
#if DAL_CPLUSPLUS
{
public:
	COMPOSITEFRAME_FORWARDS_NEW(((CompositeFrame *) m_pFrame), ->)

	MsoSCompositeFrame& operator<<(MsoSFrame& frame)
		{ ((CompositeFrame *) m_pFrame)->operator<<(*(frame.m_pFrame));
		  return *this; }

protected:
	MsoSCompositeFrame() : MsoSFrame() {}
}; // MsoSCompositeFrame
#endif // DAL_CPLUSPLUS	


// MsoSHorizFrame
#undef  CLASSNAME
#define CLASSNAME MsoSHorizFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSHorizFrame, MsoSCompositeFrame, HorizFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSHorizFrame(IMsoSAutoLayoutDialog& dlg, MsoSCompositeFrame *pParentFrame);
	HORIZFRAME_FORWARDS_NEW(((HorizFrame *) m_pFrame), ->)

private:
	HorizFrame m_frame;
}; // MsoSHorizFrame
#endif // DAL_CPLUSPLUS	


// MsoSVertFrame
#undef  CLASSNAME
#define CLASSNAME MsoSVertFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSVertFrame, MsoSCompositeFrame, VertFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSVertFrame(IMsoSAutoLayoutDialog& dlg, MsoSCompositeFrame *pParentFrame);
	VERTFRAME_FORWARDS_NEW(((VertFrame *) m_pFrame), ->)
	
private:
	VertFrame m_frame;
}; // MsoSVertFrame
#endif // DAL_CPLUSPLUS	


// MsoSFlowFrame
#undef  CLASSNAME
#define CLASSNAME MsoSFlowFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSFlowFrame, MsoSCompositeFrame, FlowFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSFlowFrame(IMsoSAutoLayoutDialog& dlg, MsoSCompositeFrame *pParentFrame);
	FLOWFRAME_FORWARDS_NEW(((FlowFrame *) m_pFrame), ->)
	
	void BindChildPair(MsoSFrame& firstChild)
		{ ((FlowFrame *) m_pFrame)->BindChildPair(*(firstChild.m_pFrame)); }
	int GetRowOccupiedByChild(MsoSFrame& child)
		{ ((FlowFrame *) m_pFrame)->GetRowOccupiedByChild(*(child.m_pFrame)); }
	void ExpandChildToFillRowVert( MsoSFrame& child )
		{ ((FlowFrame *) m_pFrame)->ExpandChildToFillRowVert(*(child.m_pFrame)); }

private:
	FlowFrame m_frame;
}; // MsoSFlowFrame
#endif // DAL_CPLUSPLUS	


// MsoSVertFlowFrame
#undef  CLASSNAME
#define CLASSNAME MsoSVertFlowFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSVertFlowFrame, MsoSCompositeFrame, VertFlowFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSVertFlowFrame(IMsoSAutoLayoutDialog& dlg, MsoSCompositeFrame *pParentFrame);
	VERTFLOWFRAME_FORWARDS_NEW(((VertFlowFrame *) m_pFrame), ->)
	
	void BindChildPair(MsoSFrame& firstChild)
		{ ((VertFlowFrame *) m_pFrame)->BindChildPair(*(firstChild.m_pFrame)); }
	int GetColOccupiedByChild(MsoSFrame& child)
		{ ((VertFlowFrame *) m_pFrame)->GetColOccupiedByChild(*(child.m_pFrame)); }
	void ExpandChildToFillColHoriz( MsoSFrame& child )
		{ ((VertFlowFrame *) m_pFrame)->ExpandChildToFillColHoriz(*(child.m_pFrame)); }

private:
	VertFlowFrame m_frame;
}; // MsoSVertFlowFrame
#endif // DAL_CPLUSPLUS	


// MsoSTableFrame
#undef  CLASSNAME
#define CLASSNAME MsoSTableFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSTableFrame, MsoSCompositeFrame, TableFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSTableFrame(IMsoSAutoLayoutDialog& dlg, int columns, int rows,
		MsoSCompositeFrame *pParentFrame);
	TABLEFRAME_FORWARDS_NEW(((TableFrame *) m_pFrame), ->)

   void InsertChildAtCell (MsoSFrame& child, int iCol, int iRow, int cCols = 1, int cRows = 1)
		{ ((TableFrame *) m_pFrame)->InsertChildAtCell(*(child.m_pFrame), iCol, iRow, cCols, cRows); }

   void ExpandChildToFillHoriz(MsoSFrame& child)
		{ ((TableFrame *) m_pFrame)->ExpandChildToFillHoriz(*(child.m_pFrame)); }

	void ExpandChildToFillVert(MsoSFrame& child)
		{ ((TableFrame *) m_pFrame)->ExpandChildToFillVert(*(child.m_pFrame)); }

	void ExpandChildToFillCellHoriz(MsoSFrame& child)
		{ ((TableFrame *) m_pFrame)->ExpandChildToFillCellHoriz(*(child.m_pFrame)); }

	void ExpandChildToFillCellVert(MsoSFrame& child)
		{ ((TableFrame *) m_pFrame)->ExpandChildToFillCellVert(*(child.m_pFrame)); }

	void AddGroupBox(TMC tmcGroupCtrlData, int left, int top, int right, int bottom,
		bool fNeverHasText)
		{ ((TableFrame *) m_pFrame)->AddGroupBox((void *) tmcGroupCtrlData, left, top, right,
		  bottom, fNeverHasText); }

	void AddGroupBox2Ctrls(TMC tmcTextCtrlData, TMC tmcLineCtrlData, int left, int top,
		int right, int bottom, bool fNeverHasText)
		{ ((TableFrame *) m_pFrame)->AddGroupBox2Ctrls((void *) tmcTextCtrlData,
		  (void *) tmcLineCtrlData, left, top, right, bottom, fNeverHasText); }
		
	MsoSCompositeFrame& operator<<(MsoSFrame& frame)
		{ ((CompositeFrame *) m_pFrame)->operator<<(*(frame.m_pFrame));
		  return *this; }

	MsoSTableFrame& operator<<(int zero)
		{ ((TableFrame *) m_pFrame)->operator<<(zero); return *this; }
	
private:
	TableFrame m_frame;
}; // MsoSTableFrame
#endif // DAL_CPLUSPLUS


// MsoSContainerFrame
#undef  CLASSNAME
#define CLASSNAME MsoSContainerFrame

DECLARE_ABSTRACT_WRAPPER_CLASS_(MsoSContainerFrame, MsoSCompositeFrame)
#if DAL_CPLUSPLUS
{
public:
	CONTAINERFRAME_FORWARDS_NEW(((ContainerFrame *) m_pFrame), ->)

protected:
	MsoSContainerFrame() : MsoSCompositeFrame() {}
}; // MsoSContainerFrame
#endif // DAL_CPLUSPLUS


// MsoSCtrlContainerFrame
#undef  CLASSNAME
#define CLASSNAME MsoSCtrlContainerFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSCtrlContainerFrame, MsoSContainerFrame,
	CtrlContainerFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSCtrlContainerFrame(IMsoSAutoLayoutDialog& dlg, TMC tmc,
		MsoSCompositeFrame *pParentFrame);
	CTRLCONTAINERFRAME_FORWARDS_NEW(((CtrlContainerFrame *) m_pFrame), ->)
	
private:
	CtrlContainerFrame m_frame;
}; // MsoSCtrlContainerFrame
#endif // DAL_CPLUSPLUS


// MsoSGroupBoxFrame
#undef  CLASSNAME
#define CLASSNAME MsoSGroupBoxFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSGroupBoxFrame, MsoSContainerFrame,
	GroupBoxFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSGroupBoxFrame(IMsoSAutoLayoutDialog& dlg, TMC tmc,
		MsoSCompositeFrame *pParentFrame);
	GROUPBOXFRAME_FORWARDS_NEW(((GroupBoxFrame *) m_pFrame), ->)
	
private:
	GroupBoxFrame m_frame;
}; // MsoSGroupBoxFrame
#endif // DAL_CPLUSPLUS


// MsoSOverlapFrame
#undef  CLASSNAME
#define CLASSNAME MsoSOverlapFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSOverlapFrame, MsoSContainerFrame,
	OverlapFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSOverlapFrame(IMsoSAutoLayoutDialog& dlg,
		MsoSCompositeFrame *pParentFrame);
	OVERLAPFRAME_FORWARDS_NEW(((OverlapFrame *) m_pFrame), ->)
	
private:
	OverlapFrame m_frame;
}; // MsoSOverlapFrame
#endif // DAL_CPLUSPLUS


// MsoSTabCtrlFrame
#undef  CLASSNAME
#define CLASSNAME MsoSTabCtrlFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSTabCtrlFrame, MsoSContainerFrame,
	TabCtrlFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSTabCtrlFrame(IMsoSAutoLayoutDialog& dlg, TMC tmcTabs, TMC tmcBorder,
		MsoSCompositeFrame *pParentFrame);
	TABCTRLFRAME_FORWARDS_NEW(((TabCtrlFrame *) m_pFrame), ->)

private:
	TabCtrlFrame m_frame;
}; // MsoSTabCtrlFrame
#endif // DAL_CPLUSPLUS


// MsoSLeafFrame
#undef  CLASSNAME
#define CLASSNAME MsoSLeafFrame

DECLARE_ABSTRACT_WRAPPER_CLASS_(MsoSLeafFrame, MsoSFrame)
#if DAL_CPLUSPLUS
{
public:
	LEAFFRAME_FORWARDS_NEW(((LeafFrame *) m_pFrame), ->)

	void AddLeftString(TMC tmc)
		{ ((LeafFrame *) m_pFrame)->AddLeftString((void *) tmc); }
	void AddTopString(TMC tmc)
		{ ((LeafFrame *) m_pFrame)->AddTopString((void *) tmc); }
	void AddRightString(TMC tmc)
		{ ((LeafFrame *) m_pFrame)->AddRightString((void *) tmc); }

protected:
	MsoSLeafFrame() : MsoSFrame() {}
}; // MsoSLeafFrame
#endif // DAL_CPLUSPLUS


// MsoSCtrlFrame
#undef  CLASSNAME
#define CLASSNAME MsoSCtrlFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSCtrlFrame, MsoSLeafFrame, CtrlFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSCtrlFrame(IMsoSAutoLayoutDialog& dlg, TMC tmc,
		MsoSCompositeFrame *pParentFrame);
	CTRLFRAME_FORWARDS_NEW(((CtrlFrame *) m_pFrame), ->)
	
private:
	CtrlFrame m_frame;
}; // MsoSCtrlFrame
#endif // DAL_CPLUSPLUS


// MsoSNumInputFrame
#undef  CLASSNAME
#define CLASSNAME MsoSNumInputFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSNumInputFrame, MsoSLeafFrame, NumInputFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSNumInputFrame(IMsoSAutoLayoutDialog& dlg, TMC tmc,
		MsoSCompositeFrame *pParentFrame);
	NUMINPUTFRAME_FORWARDS_NEW(((NumInputFrame *) m_pFrame), ->)
	
private:
	NumInputFrame m_frame;
}; // MsoSNumInputFrame
#endif // DAL_CPLUSPLUS


// MsoSListEditFrame
#undef  CLASSNAME
#define CLASSNAME MsoSListEditFrame

DECLARE_CONCRETE_WRAPPER_CLASS_(MsoSListEditFrame, MsoSLeafFrame, ListEditFrame)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSListEditFrame(IMsoSAutoLayoutDialog& dlg, TMC tmcEditStr, TMC tmcListBox,
		MsoSCompositeFrame *pParentFrame);
	LISTEDITFRAME_FORWARDS_NEW(((ListEditFrame *) m_pFrame), ->)
	
private:
	ListEditFrame m_frame;
}; // MsoSListEditFrame
#endif // DAL_CPLUSPLUS


/***************************************************************************/
// Supplemental Frame Class:  Frame Equalizer
/***************************************************************************/

// MsoSFrameEqualizer
#undef  CLASSNAME
#define CLASSNAME MsoSFrameEqualizer

DECLARE_CONCRETE_WRAPPER_CLASS(MsoSFrameEqualizer, FrameEqualizer)
#if DAL_CPLUSPLUS
{
public:
	MSOCTOR MsoSFrameEqualizer(IMsoSAutoLayoutDialog& dlg);
	FRAMEEQUALIZER_FORWARDS(((FrameEqualizer *) m_pEqualizer), ->)
	
	MsoSFrameEqualizer& operator<<(MsoSFrame& frame)
		{ m_equalizer.operator<<(*((Frame *) frame.m_pFrame)); return *this; }

private:
	FrameEqualizer *m_pEqualizer;
	FrameEqualizer m_equalizer;
}; // MsoSFrameEqualizer
#endif // DAL_CPLUSPLUS


/***************************************************************************/
// Frame Initializers For C Clients
/***************************************************************************/
#if !DAL_CPLUSPLUS

MSOAPI_(void) MsoSHorizFrameInit(MsoSHorizFrame *This, IMsoSAutoLayoutDialog *pDlg,
	MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSVertFrameInit(MsoSVertFrame *This, IMsoSAutoLayoutDialog *pDlg,
	MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSFlowFrameInit(MsoSFlowFrame *This, IMsoSAutoLayoutDialog *pDlg,
	MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSVertFlowFrameInit(MsoSVertFlowFrame *This, IMsoSAutoLayoutDialog *pDlg,
	MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSTableFrameInit(MsoSTableFrame *This, IMsoSAutoLayoutDialog *pDlg,
	int columns, int rows, MsoSCompositeFrame *pParentFrame);
MSOAPIX_(void) MsoSCtrlContainerFrameInit(MsoSCtrlContainerFrame *This,
	IMsoSAutoLayoutDialog *pDlg, TMC tmc, MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSGroupBoxFrameInit(MsoSGroupBoxFrame *This, 
	IMsoSAutoLayoutDialog *pDlg, TMC tmc, MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSOverlapFrameInit(MsoSOverlapFrame *This,
	IMsoSAutoLayoutDialog *pDlg, MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSTabCtrlFrameInit(MsoSTabCtrlFrame *This,
	IMsoSAutoLayoutDialog *pDlg, TMC tmcTabs, TMC tmcBorder,
	MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSCtrlFrameInit(MsoSCtrlFrame *This, IMsoSAutoLayoutDialog *pDlg,
	TMC tmc, MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSNumInputFrameInit(MsoSNumInputFrame *This,
	IMsoSAutoLayoutDialog *pDlg, TMC tmc, MsoSCompositeFrame *pParentFrame);
MSOAPI_(void) MsoSListEditFrameInit(MsoSListEditFrame *This,
	IMsoSAutoLayoutDialog *pDlg, TMC tmcEditStr, TMC tmcListBox,
	MsoSCompositeFrame *pParentFrame);

MSOAPIX_(void) MsoSFrameEqualizerInit(MsoSFrameEqualizer *This,
	IMsoSAutoLayoutDialog *pDlg);

#endif // !DAL_CPLUSPLUS

#if DAL_CPLUSPLUS
MSOAPI_(BOOL) MsoDALFOrderItemsPerLocString(void **rgItems, int *pcItems, const WCHAR *wzOrder);
MSOAPI_(BOOL) MsoDALFOrderItemsPerLocStringTmc(MsoSFrame **rgItems, int *pcItems, TMC tmcOrder, MsoSCompositeFrame *pParentH);
#endif // DAL_CPLUSPLUS

#endif // MSOSDAL_H
