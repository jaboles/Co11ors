#pragma once
/*****************************************************************************

   Module  : MSODAL.h
   Owner   : SteveBre

******************************************************************************

   Public header for generic Dialog AutoLayout code

*****************************************************************************/

#if defined(__cplusplus) && !defined(CINTERFACE)
   #define DAL_CPLUSPLUS 1
#else
   #define DAL_CPLUSPLUS 0
#endif

#if( DEMOAPP ) // defined in the precompiled header of the demo app
   #if DAL_CPLUSPLUS
      #include "afxpriv.h" // for DEBUG, WM_KICKIDLE
   #else
      #define _X86_ 1
      #define WM_KICKIDLE         0x036A  // (params unused) causes idles to kick in
   #endif
   #undef DEBUG
   #define DEBUG _DEBUG
   #define NO_MSO_DEPENDENCIES 1
#else
   //#define NO_MSO_DEPENDENCIES 1 // If you're not using MSO, you should provide real implementations for AssertMsgInline, Assert, and AssertSz.
#endif

#if( DEBUG || DEMOAPP )
   #define MSODAL_DEBUG_TOOLS 1
#endif


#if( !NO_MSO_DEPENDENCIES )
   #include "msodebug.h" // for AssertMsgInline
   #include "msointl.h"  // for LID
   #define MSODALPUBDEBUGMETHODIMP  virtual MSOPUB MSODEBUGMETHODIMP
   #define MSODALDEBUGMETHODIMP     virtual        MSONONVIRTDEBUGMETHODIMP
   #define MSODALOVERRIDEDEBUGMETHOD __override    MSODEBUGMETHODIMP
#else
   // the following warnings are disabled for the standard header files to allow Warning Level 4 clean compile
   #pragma warning( push )                // save default warnings
   #pragma warning( disable : 4201 )      // nonstandard extension used : nameless struct/union
   #pragma warning( disable : 4214 )      // nonstandard extension used : bit field types other than int
   #include <windef.h>
   #include <stdarg.h>
   #include <winbase.h>
   #include <wingdi.h>
   #define OEMRESOURCE
   #include <winuser.h>
   #pragma warning( pop )                 // restore default warnings
   #if( DEMOAPP )
      #define OBM_CHECKBOXES      32759 // Why wasn't this defined in winuser.h?
      #define UINT_PTR            UINT
//      #define GetWindowLongPtr    GetWindowLong
//      #define GWLP_HINSTANCE      GWL_HINSTANCE
      #define DWORD_PTR           DWORD
      #if DAL_CPLUSPLUS
         #include "stdafx.h" // for ASSERT
      #else
         #define ASSERT(f)
      #endif
      #define DxOfRect(rc)        ((rc).right - (rc).left)
      #define DyOfRect(rc)        ((rc).bottom - (rc).top)
      #define AssertMsgInline(f, szMsg) ASSERT(((szMsg),f))
      #define Assert(f)                 ASSERT(f)
      #define AssertSz(f, szMsg)        ASSERT(((szMsg),f))
   #else
      #if( DEBUG )
         #define AssertMsgInline(f, szMsg) do { if( !(f) ) { /* Should alert */ } } while (0)
         #define Assert(f)                 do { if( !(f) ) { /* Should alert */ } } while (0)
         #define AssertSz(f, szMsg)        do { if( !(f) ) { /* Should alert */ } } while (0)
      #else
         #define AssertMsgInline(f, szMsg)
         #define Assert(f)
         #define AssertSz(f, szMsg)
      #endif
   #endif
   typedef WORD LID;
   #if( DEMOAPP )
      #define lidAmerican 0x0409
   #endif
   #define lidHebrew 0x040d

   #if( DEBUG )
	   #define DebugOnly(e) e
	   #define DebugElse(s, t)	s
      #define AssertDo(f) Assert(f)
   #else
	   #define DebugOnly(e)
	   #define DebugElse(s, t) t
      #define AssertDo(f) f
   #endif
   #define MsoAssertTag(f, num)		   Assert(f)
   #define MsoAssertSzTag(f, sz, num)	AssertSz(f, sz)
   #define MsoAssertDoTag(f, num)      AssertDo(f)
   #define MsoNotReachedTag(tag)       MsoAssertTag(FALSE, tag)
   #define MsoNotReachedSzTag(sz, tag) MsoAssertSzTag(FALSE, sz, tag)

   // String funcs
   #define MsoWzIndex      DlgAutoLayout_MsoWzIndex
   #if( MSODAL_DEBUG_TOOLS )
      #define WzTruncCopy     DlgAutoLayout_WzTruncCopy
      #define MsoWzCchCopy    DlgAutoLayout_MsoWzCchCopy
      #define MsoWzCopy       DlgAutoLayout_MsoWzCopy
      #define MsoWzCchAppend  DlgAutoLayout_MsoWzCchAppend
      #define PwchDecodeUint  DlgAutoLayout_PwchDecodeUint
      #define MsoWzDecodeUint DlgAutoLayout_MsoWzDecodeUint
      #define MsoWzDecodeInt  DlgAutoLayout_MsoWzDecodeInt
      #define MsoSzDecodeInt  DlgAutoLayout_MsoSzDecodeInt
      #define MsoSzDecodeUint DlgAutoLayout_MsoSzDecodeUint
      #define PchDecodeUint   DlgAutoLayout_PchDecodeUint
   #endif

   int MsoCchWzLen(const WCHAR* wz);
   WCHAR* MsoWzIndex(const WCHAR* wz, int wch);
   #if( MSODAL_DEBUG_TOOLS )
      WCHAR* WzTruncCopy(_Out_z_cap_(cchDst) WCHAR *wzDst, const WCHAR *wzSrc, unsigned cchDst);
      WCHAR* MsoWzCchCopy(const WCHAR* wzFrom, _Out_z_cap_(cch) WCHAR* wzTo, int cch);
      WCHAR* MsoWzCopy(const WCHAR* wzFrom, _Out_ WCHAR* wzTo);
      WCHAR* MsoWzCchAppend(const WCHAR* wzFrom, _Out_z_cap_(cch) WCHAR* wzTo, int cch);
      WCHAR* PwchDecodeUint(_Out_cap_(1) WCHAR* rgwch, unsigned u, int wBase);
      int MsoWzDecodeUint(_Out_z_cap_(cch) WCHAR* rgwch, int cch, unsigned u, int wBase);
      int MsoWzDecodeInt(_Out_z_cap_(cch) WCHAR* rgwch, int cch, int w, int wBase);
      int MsoSzDecodeInt(_Out_z_cap_(cch) CHAR* rgch, int cch, int w, int wBase);
      int MsoSzDecodeUint(_Out_z_cap_(cch) CHAR* rgch, int cch, unsigned u, int wBase);
      CHAR* PchDecodeUint(_Out_cap_(1) CHAR* rgch, unsigned u, int wBase);
   #endif

   #if DAL_CPLUSPLUS
      #define MSOEXTERN_C extern "C"
   #else
      #define MSOEXTERN_C
   #endif
   //#if !defined(OFFICE_BUILD)
	   //#define MSOPUB     __declspec(dllimport)
	   //#define MSOPUBDATA __declspec(dllimport)
   //#else
#if 0 //$[VSMSO]
	   #define MSOPUB     __declspec(dllexport)
	   #define MSOPUBDATA __declspec(dllexport)
#else //$[VSMSO]
	   #define MSOPUB
	   #define MSOPUBDATA
#endif //$[VSMSO]
   //#endif
   #define MSOAPICALLTYPE __stdcall
   #define MSOAPI_(t) MSOEXTERN_C MSOPUB t MSOAPICALLTYPE
   #define MSODALPUBDEBUGMETHODIMP
   #define MSODALDEBUGMETHODIMP
   #define MSODALOVERRIDEDEBUGMETHOD
   #if DAL_CPLUSPLUS
      #define MSOMETHOD_(t,m)   virtual t STDMETHODCALLTYPE m
   #else
      #define MSOMETHOD_(t,m)   t (STDMETHODCALLTYPE * m)
   #endif
   #define MSOMETHODIMP_(t)  t STDMETHODCALLTYPE
   #define MSOOVERRIDEMETHOD_  MSOMETHODIMP_
   #define __ecount(var)
   #define __zcount(var)
   #define __fallthrough

   #ifndef cElements
      #if !DAL_CPLUSPLUS || defined _PREFAST_
         #define cElements(ary) (sizeof(ary) / sizeof(ary[0]))
      #else
         template<typename T> static char cElementsVerify(void const *, T) throw() { return 0; }
         template<typename T> static void cElementsVerify(T *const, T *const *) throw() {};
         #define cElements(arr) (sizeof(cElementsVerify(arr,&(arr))) * sizeof(arr)/sizeof(*(arr)))
      #endif
   #endif

   #ifndef RgC
      #define RgC(ary) (ary),cElements(ary)
   #endif
#endif

#if !_PREFIX_
#define MSOCTOR                    MSOPUB MSOAPICALLTYPE
#define MSODTOR                    MSOPUB MSOAPICALLTYPE
#else
#define MSOCTOR                
#define MSODTOR              
#endif

#define MSODTORVIRT_DECL   virtual MSOPUB MSOAPICALLTYPE
#define MSODTORVIRT_IMP            MSOPUB MSOAPICALLTYPE
#define MSODALAPI_(t) MSOAPI_(t)
#define MSODALAPIX_(t) t
#define MSODALDATA MSOPUBDATA


// FUTURE: Include these macros from elsewhere?
#undef CLASSNAME
#undef THISPTR_
#if defined(CLASSNAME) || defined(THISPTR_)
   break build: Change to unused macros.
#endif

#if DAL_CPLUSPLUS
   #define C_ONLY(stuff)
   #define CPP_ONLY(stuff) stuff
	#define CPP_ELSE(cpp, c)	cpp
   #define CLASS_KEYWORD class
   #define THISPTR_
   #define THISPTR                    void
   #define DEFAULTS_TO(value)    = value
   #define DECLARE_CLASS(cls)    CLASS_KEYWORD cls
   #define DECLARE_CLASS_(cls, basecls)    CLASS_KEYWORD cls : public basecls
   #define METHOD_CONST const
   #define REFERENCE                       &
#else
   #define C_ONLY(stuff) stuff
   #define CPP_ONLY(stuff)
	#define CPP_ELSE(cpp, c)	c
   #define THISPTR_                   CLASSNAME FAR* This,
   #define THISPTR                    CLASSNAME FAR* This
   #define DEFAULTS_TO(value)
   #ifdef CONST_VTABLE
      #undef CONST_VTBL
      #define CONST_VTBL const
      #define DECLARE_CLASS(cls)      typedef struct cls { \
                                          const struct cls##Vtbl FAR* lpVtbl; \
                                          void* pImp; \
                                      } cls; \
                                      typedef const struct cls##Vtbl cls##Vtbl; \
                                      const struct cls##Vtbl
   #else
      #undef CONST_VTBL
      #define CONST_VTBL
      #define DECLARE_CLASS(cls)      typedef struct cls { \
                                          struct cls##Vtbl FAR* lpVtbl; \
                                          void* pImp; \
                                      } cls; \
                                      typedef struct cls##Vtbl cls##Vtbl; \
                                      struct cls##Vtbl
   #endif
   #define DECLARE_CLASS_(cls, basecls)    DECLARE_CLASS(cls)
   #define METHOD_CONST
   #define REFERENCE                       *
   #define bool                            unsigned char
#endif


#if DAL_CPLUSPLUS
/*********************/
namespace DlgAutoLayout
/*********************/
{
   class AutoLayoutDialogData;
   class IMsoAutoLayoutDialog;
   class FrameEqualizerImp;
   class FrameImp;
#else
   struct IMsoAutoLayoutDialog;
   typedef struct IMsoAutoLayoutDialog IMsoAutoLayoutDialog;
#endif // DAL_CPLUSPLUS


//-------------------------------------------------------------------------------------------------
// Office-wide layout metrics, mostly in dialog units (DU).
//
// Clients should almost always use these, rather than hard-coded numbers or locally-defined
// constants, as parameters to control- and margin-sizing APIs.  This way the metrics can be
// changed centrally, and all dialogs will automatically reflect the changes.
//
// (Many of these values were taken from PowerPoint's "GPU Dialog Layout Metrics" document,
// and some may still reflect Win3.x metrics instead of more current standards.  These should
// be reviewed, and will likely change, making it all the more important that all apps use
// them so they will get the changes.)
//-------------------------------------------------------------------------------------------------
// Convert pixels to DU's or DU's to pixels based on the passed in dialog base units.
MSOAPI_(int) MsoPixToDU(int nPix, int nPixPerChar, BOOL fHoriz);
MSOAPI_(int) MsoDUToPix(int nDU, int nPixPerChar, BOOL fHoriz);

#define MAX_CTRL_TEXT_LEN 1500 // Does not include null terminator.  Can be bigger than 255.
#define MAX_CAPTION_LEN    500 // Does not include null terminator.

// Dialog border margins, WITHIN THE CLIENT AREA -- these do not include any non-client borders or caption.
#define DIALOG_TOP_MARGIN_DU                 4
#define DIALOG_BOTTOM_MARGIN_DU              4
#define DIALOG_LEFT_MARGIN_DU                4
#define DIALOG_RIGHT_MARGIN_DU               4
// The following values are likely to change -- do not rely on them.  Also, note that they
// do not include the insetting of the dialog from the edges of the WorkPane toolbar.
#define WORKPANE_TOP_MARGIN_DU               4
#define WORKPANE_BOTTOM_MARGIN_DU            4
#define WORKPANE_LEFT_MARGIN_DU              4
#define WORKPANE_RIGHT_MARGIN_DU             4

// Tabsheet border margins
#define TABSHEET_TOP_MARGIN_DU               6
#define TABSHEET_LEFT_MARGIN_DU              6
#define TABSHEET_RIGHT_MARGIN_DU             6
#define TABSHEET_BOTTOM_MARGIN_DU            6

// Spacing between independent controls
#define FRAME_H_MARGIN_DU                    6
#define FRAME_V_MARGIN_DU                    2
#define FRAME_DIFFGRP_V_MARGIN_DU            5

// Spacing between controls in groups
#define PUSHBTN_SAMEGRP_V_MARGIN_DU          4
#define PUSHBTN_DIFFGRP_V_MARGIN_DU          PUSHBTN_SAMEGRP_V_MARGIN_DU // Identical margin, per O10 bug 147948
#define PUSHBTN_SAMEGRP_H_MARGIN_DU          6
#define CHECKBOX_V_MARGIN_DU                 2
#define RADIOBTN_V_MARGIN_DU                 2
#define HYPERLINK_V_MARGIN_DU                2

// Spacing between controls and their label text
#define LABEL_V_MARGIN_DU                    2
#define LABEL_LEFT_MARGIN_DU                 2 // Space between a control and its left-side label
#define LABEL_RIGHT_MARGIN_DU                2 // Space between a control and its right-side label

// Spacing inside multi-element controls
#define LISTEDIT_V_MARGIN_DU                 0
#define LISTEDIT_LEFT_INDENT_DU              0
#define LISTEDIT_RIGHT_INDENT_DU             0

// Sizes of specific controls
#define PUSHBTN_HEIGHT_DU                   14
#define PUSHBTN_MIN_WIDTH_DU                50
#define WORKPANE_PUSHBTN_MIN_WIDTH_DU        0
#define PUSHBTN_MIN_WIDTH_AROUND_TEXT_DU    12 // Min width in excess of the text width that must be added to the width
                                               // of a pushbutton (total space for left and right sides)
#define STATICSTR_HEIGHT_DU                  8
#define PREFERRED_CLOSED_DROPDOWN_HEIGHT_DU 13 // Used when the closed-state height of dropdowns is settable.
                                                    // (Depending on the dialog manager, it may not be.)
#define CHECKBOX_TEXT_HEIGHT_DU             10 // Must be > 8 else focus rect obscures text
#define RADIOBTN_TEXT_HEIGHT_DU             10 // Must be > 8 else focus rect obscures text
#define HYPERLINK_TEXT_HEIGHT_DU            10 // Must be > 8 else focus rect obscures text
#define HYPERLINK_ICON_TO_TEXT_DIST_DU       3 // Margin between icon and text - taken from ppt\appframe\dlnative.cpp
#define HYPERLINK_H_FOCUS_RECT_PIX           2 // Left + right margin (for focus rect) - ""
#define HYPERLINK_V_FOCUS_RECT_PIX           2 // Top + bottom margin (for focus rect) if total height > HYPERLINK_TEXT_HEIGHT_DU
                                               // 3 matches height for powerpoint hyperlinks
#define MULTILINE_CHECKBOX_EXTRA_HEIGHT_PIX  3 // Necessary for multiline checkboxes and radio buttons especially with BS_TOP style
#define VTRACKBAR_WIDTH_DU                  19 // This value is based on the size of the trackbar in the Audio Tuning
                                               // Wizard of Microsoft NetMeeting 2.11, and should probably be reviewed.
                                               // Perhaps it should be related (proportional) to some system metric, like scrollbar width?
#define HTRACKBAR_HEIGHT_DU                 VTRACKBAR_WIDTH_DU // This equality should probably be reviewed.
MSODALAPIX_(int) MsoDALNumInputHeightPix( const IMsoAutoLayoutDialog REFERENCE dlg ); // *** Returns PIXELS -- not DUs!  Don't use DUToPixV(). ***  Returns 0 if failed.
MSODALAPIX_(int) MsoDALNumInputSpinnerWidthPix();                   // *** Returns PIXELS -- not DUs!  Don't use DUToPixV(). ***  Returns 0 if failed.
MSODALAPI_(int) MsoDALEditStrHeightPix( const IMsoAutoLayoutDialog REFERENCE dlg ); // *** Returns PIXELS -- not DUs!  Don't use DUToPixV(). ***  Returns 0 if failed.
MSODALAPI_(int) MsoDALEditStrTextHeightPix( int pixPerChar ); // Returns the height in PIXELS that the edit box should be (it subtracts off border and internal padding.


// Vertical text offsets of specific controls (for baseline alignment)
#define EDITSTR_HEIGHT_ABOVE_TEXT_PIX        3 // The number of pixels at the top of an edit-string that cannot be used
                                               // to display text.  To achieve baseline alignment between a static string
                                               // and an edit-string next to it, the top of the static string should be
                                               // positioned this many pixels lower than the top of the edit-string.
                                               // Includes 2-pixel border plus 1-pixel internal margin.
#define TEXTLISTBOX_HEIGHT_ABOVE_TEXT_PIX    2 // The number of pixels at the top of a text listbox that cannot be used
                                               // to display text.  To achieve baseline alignment between a static string
                                               // and the first item of text listbox next to it, the top of the static
                                               // string should be positioned this many pixels lower than the top of the
                                               // listbox.  Includes just the 2-pixel border of the listbox.

// Group-boxes and group-lines
#define GROUPBOX_TOP_MARGIN_DU              ( STATICSTR_HEIGHT_DU + 4 ) // Includes 8
                                               // DUs of label text (half above and half below box crossbar) plus 
                                               // 4 DUs blank space between text and interior controls
#define GROUPBOX_NO_LABEL_TOP_MARGIN_DU      5 // Used when the group-box has empty (i.e., no) label:  This represents
                                               // the 2-pixel thickness of the group-box crossbar as well as the margin
                                               // beneath the crossbar but above the interior controls. So the actual blank
                                               // space is 2 pixels less than GROUPBOX_NO_LABEL_TOP_MARGIN_DU DUs.
#define GROUPBOX_BOTTOM_MARGIN_DU            4
#define GROUPBOX_LEFT_MARGIN_DU             FRAME_H_MARGIN_DU
#define GROUPBOX_RIGHT_MARGIN_DU            FRAME_H_MARGIN_DU
#define GROUPBOX_MIN_WIDTH_AROUND_TEXT_DU   12 // Min width in excess of the label width that must be added to the width
                                               // of a groupbox, including both margins surrounding the text, as well as
                                               // both end segments of the line
#define GROUPBOX_OUTLINE_GIRTH_PIX           2 // Thickness of the outline, in PIXELS -- not DUs!  Don't use DUToPixV().
MSODALAPI_(int) MsoDALGroupBoxNoLabelTopAdjustPix( const IMsoAutoLayoutDialog REFERENCE dlg ); // *** Returns PIXELS -- not DUs!  Don't use DUToPixV(). ***
#define GROUPLINE_TOP_MARGIN_DU             GROUPBOX_TOP_MARGIN_DU // See comment at GROUPBOX_TOP_MARGIN_DU
#define GROUPLINE_NO_LABEL_TOP_MARGIN_DU    GROUPBOX_NO_LABEL_TOP_MARGIN_DU // See comment at GROUPBOX_NO_LABEL_TOP_MARGIN_DU
#define GROUPLINE_BOTTOM_MARGIN_DU           3
#define GROUPLINE_LEFT_MARGIN_DU            GROUPBOX_LEFT_MARGIN_DU
#define GROUPLINE_RIGHT_MARGIN_DU            0
#define GROUPLINE_GAP_AFTER_TEXT_DU         LABEL_LEFT_MARGIN_DU // Space between text and line
#define GROUPLINE_MIN_WIDTH_BEYOND_TEXT_DU  16 // Min length of the line part of the GroupLine, INCLUDING space between text and line (GROUPLINE_GAP_AFTER_TEXT_DU)
#define GROUPLINE_LINE_HEIGHT_PIX           GROUPBOX_OUTLINE_GIRTH_PIX // Thickness of the line, in pixels
MSODALAPIX_(int) MsoDALGroupLineNoLabelTopAdjustPix( IMsoAutoLayoutDialog REFERENCE dlg ); // *** Returns PIXELS -- not DUs!  Don't use DUToPixV(). ***
#define UNDERLINEDHEADING_TOP_MARGIN_DU             GROUPLINE_TOP_MARGIN_DU
#define UNDERLINEDHEADING_NO_LABEL_TOP_MARGIN_DU    GROUPLINE_NO_LABEL_TOP_MARGIN_DU
#define UNDERLINEDHEADING_BOTTOM_MARGIN_DU          GROUPLINE_BOTTOM_MARGIN_DU
#define UNDERLINEDHEADING_LEFT_MARGIN_DU            GROUPLINE_LEFT_MARGIN_DU
#define UNDERLINEDHEADING_RIGHT_MARGIN_DU           GROUPLINE_RIGHT_MARGIN_DU
#define UNDERLINEDHEADING_MIN_WIDTH_BEYOND_TEXT_DU  GROUPLINE_MIN_WIDTH_BEYOND_TEXT_DU
#define UNDERLINEDHEADING_LINE_HEIGHT_PIX            1 // Thickness of the line, in pixels

MSODALAPI_(int) MsoDALCheckBoxBMWidthPlusMargin( const IMsoAutoLayoutDialog REFERENCE dlg );
// Width of checkbox bitmap PLUS the space between the bitmap and the text -- NOT JUST between the
// bitmap and the left side of the focus rectangle that surrounds the text.  The focus rectangle,
// when present, juts into this margin.  It also extends, on its right side, past the end of the
// text -- but *this* extra width is not included by this function.  Returns 0 if failed.

MSODALAPI_(int) MsoDALCheckBoxBMWidth();
// Width of checkbox bitmap, NOT including the space between the bitmap and the text.  Returns 0 if failed.

MSODALAPI_(int) MsoDALCheckBoxBMHeight();
// Height of checkbox bitmap.  Returns 0 if failed.


typedef enum MsoDALCtrlType
{
   MSODAL_OUTOFCONTROL,
   MSODAL_DUMMYCONTROL,
   MSODAL_STATICSTRING,
   MSODAL_EDITSTRING,
   MSODAL_DROPDOWNLIST,
   MSODAL_PUSHBUTTON,
   MSODAL_BITMAPBUTTON,
   MSODAL_RADIOBUTTON,
   MSODAL_CHECKBOX,
   MSODAL_SCROLLBAR,
   MSODAL_LISTBOX,
   MSODAL_COMBOBOX,
   MSODAL_USERCONTROL,
   MSODAL_CUSTOMDROPDOWN,
   MSODAL_NUMBERINPUT,
   MSODAL_TRACKBAR,
   MSODAL_GROUPBOX,
   MSODAL_GROUPLINE,
   MSODAL_UNDERLINEDHEADING,
   MSODAL_TABCONTROL,
   MSODAL_BITMAPCONTROL,
   MSODAL_NUMCTRLTYPES,
   MSODAL_HYPERLINK,
} MsoDALCtrlType;


typedef enum MsoDALDlgType
{
   MSODAL_NORMAL,
   MSODAL_WORKPANE, // Has special border margins, may differ in other metrics too
   MSODAL_WIZARD, // Has no dialog border margins, and does not show tabs of tabcontrol
} MsoDALDlgType;


typedef struct MsoDALLayoutParams
{
   CPP_ONLY( MsoDALLayoutParams() : type( MSODAL_NORMAL ) {} )

   MsoDALDlgType type;

   // New params may be added here:

} MsoDALLayoutParams;


typedef enum MsoDALFrameData
{
	MSODAL_PARENT_FRAME,
	MSODAL_FIRST_CHILD_FRAME,
	MSODAL_NEXT_SIBLING_FRAME
} MsoDALFrameData;


//------------------------------------------------------------------------------------
// AutoLayout API's
//------------------------------------------------------------------------------------

MSODALAPI_(BOOL) MsoDALFDlgIsValid( const IMsoAutoLayoutDialog REFERENCE dlg );
// Is the IMsoAutoLayoutDialog valid, or was there some kind of failure during its creation or use,
// such as memory allocation?

MSODALAPI_(void) MsoDALMarkDlgInvalid( const IMsoAutoLayoutDialog REFERENCE dlg );
// Call this if there is any kind of failure or exception while executing one of the overridden
// IMsoAutoLayoutDialog callback methods, or if anything weird occurs that should cause us to abort
// the entire AutoLayout.

MSODALAPI_(void) MsoDALSuspendFrameDestruction( IMsoAutoLayoutDialog REFERENCE dlg );
// Call this to prevent frames from being destructed during Layout(), if you need to check their
// dimensions after Layout().  Remember to call MsoDALDestructFrames() when you don't need the
// frames anymore, or you will leak memory!  Or just use the class FrameLifetime (defined below).

MSODALAPI_(void) MsoDALDestructFrames( IMsoAutoLayoutDialog REFERENCE dlg );
// If you called MsoDALSuspendFrameDestruction() then you must call this afterwards, when you no
// longer need the frames.


#if DAL_CPLUSPLUS

class FrameLifetime
{
public:
   FrameLifetime( IMsoAutoLayoutDialog& dlg ) : m_dlg( dlg )
      { MsoDALSuspendFrameDestruction( dlg ); }
   ~FrameLifetime()
      { MsoDALDestructFrames( m_dlg ); }
private:
   IMsoAutoLayoutDialog& m_dlg;
};

#endif // DAL_CPLUSPLUS


#if( MSODAL_DEBUG_TOOLS )
   MSODALAPI_(void) MsoDALRunDebugDialog( HWND hwndAppWindow );
   MSODALAPI_(void) MsoDALDrawDebugMarkings( IMsoAutoLayoutDialog REFERENCE dlg );
   MSODALAPI_(BOOL) MsoDALFGetFontHeightOverride( int REFERENCE lfHeight );

   MSODALAPI_(BOOL) MsoDALFRtlDlgForcingEnabled();
   // Is the "Force right-to-left layout" tool turned on?

   MSODALAPI_(BOOL) MsoDALFAutoFailOn(); // Cryptic name so you won't be tempted to call this -- for use by glues only.

   MSODALAPI_(BOOL) MsoDALStrLenToolEnabled();
#endif // MSODAL_DEBUG_TOOLS


#undef  CLASSNAME
#define CLASSNAME Frame

// Flags used (via SetSetWidthFlags()) to alter the default behavior of SetWidth():
typedef enum MsoDALSetWidthFlag
{
   MSODAL_SW_INCLUDESTRIM        = 1,
      // Width for non-text elements of the control (such as dropdown buttons,
      // listbox scrollbars, edit-string borders and margins, checkbox bitmaps
      // and focus rectangles, etc.) is already included in the width passed
      // by the caller, and should not be added automatically.
   MSODAL_SW_NOSHRINKWRAP        = 2,
      // Even if the control's (single- or multi-line) text does not fill the
      // complete width specified, do not reduce the control's width to tightly
      // fit the text.  (In other words, override the default behavior, whereby
      // SetWidth() will "shrink-wrap" a text control to the minimum width
      // necessary to accommodate its text.)
   MSODAL_SW_NOWIDERFORLONGWORD  = 4,
      // By default, if there is a single word (commonly, filename) that is wider
      // than the specified width, SetWidth() will increase the width by up to 90%
      // to try to accommodate it.  If MSODAL_SW_NOWIDERFORLONGWORD is specified,
      // this flexibility is disabled, so long words will either be truncated or
      // wrapped in mid-word. (Alternatively, SetMaxPcntToExceedStrTargetWidth()
      // can be used to change this limit from the default 90%.)
   MSODAL_SW_NOSETHEIGHT         = 8,
      // Prevent SetWidth() from also increasing the height of a text control
      // that must wrap to multiple lines.
   MSODAL_SW_NOOPIFSTRHASNEWLINE = 16,
      // Make SetWidth() do nothing (be a no-op) if the text control contains
      // a newline character.  This allows localizers complete control over how
      // the text wraps.  Should not be used except in very special cases.
   MSODAL_SW_SIMPLE              = MSODAL_SW_INCLUDESTRIM | MSODAL_SW_NOSHRINKWRAP |
                                   MSODAL_SW_NOWIDERFORLONGWORD | MSODAL_SW_NOSETHEIGHT,
} MsoDALSetWidthFlag;

DECLARE_CLASS(Frame)
{
CPP_ONLY(public:)

#define FRAME_VIRTUALS \
   DebugOnly( CPP_ELSE( MSODTORVIRT_DECL ~Frame(), \
                        MSOMETHOD_(void,PlaceholderForFrameDtorDoNotCall) (THISPTR) ); ) \
      /* No destructor in ship in order to avoid 15 bytes of destructor-calling code      */ \
      /* getting generated for every frame in the world.  This saves over 22KB in         */ \
      /* PowerPoint client code.  Since there's no destructor, all Frames for the current */ \
      /* dialog get "destructed" at the end of CompositeFrame::Layout().                  */ \
      \
   MSOMETHOD_(void,ExpandToFillHoriz) (THISPTR); \
   MSOMETHOD_(void,ExpandToFillVert) (THISPTR); \
      /* Expand this one child frame to fill space available within its parent.  To expand        */ \
      /* multiple sibling frames, consider calling ExpandChildrenToFillHoriz/Vert() on the parent */ \
      \
   MSOMETHOD_(void,DontExpandHoriz) (THISPTR); \
   MSOMETHOD_(void,DontExpandVert) (THISPTR); \
      /* Used in conjunction with ExpandChildrenToFill[Horiz/Vert]() or */ \
      /* Same[Width/Height]Children() called on the parent frame        */ \
      \
   MSOMETHOD_(void,SetWidth) (THISPTR_ int width); \
   MSOMETHOD_(void,SetWidthDU) (THISPTR_ int width); \
   MSOMETHOD_(void,SetWidthChars) (THISPTR_ int numChars); \
      /* Sets the width of the frame (immediately -- not during Layout()).  For text controls, the       */ \
      /* behavior can be quite complicated (as explained in the subsequent paragraphs), and may be       */ \
      /* customized by calling SetSetWidthFlags() with the appropriate flags prior to the call to        */ \
      /* SetWidth[XX]().  The default behavior is pretty good though, and SetSetWidthFlags() should      */ \
      /* generally not be needed.                                                                        */ \
      /*                                                                                                 */ \
      /* For control types with borders, interior margins, and/or potential scrollbars, extra width      */ \
      /* is automatically added to the width specified, to accommodate these elements.  It is rarely     */ \
      /* advisable to disable this width flexibility, but this can be done if necessary by calling       */ \
      /* SetSetWidthFlags(MSODAL_SW_INCLUDESTRIM) before SetWidth[XX](), and then the specified width    */ \
      /* will be set on the overall control, without adding any extra width for the control's "trim".    */ \
      /*                                                                                                 */ \
      /* If the frame corresponds to a control that supports multi-line wrapping text (e.g., static      */ \
      /* string, checkbox, or radio button), then SetWidth[XX]() will also set the height, based on the  */ \
      /* number of lines required to fit all the text within the specified width.  However, the height   */ \
      /* will never be decreased from the prior height -- only increased.  To prevent SetWidth[XX]()     */ \
      /* from also setting the height of such a control, call SetSetWidthFlags(MSODAL_SW_NOSETHEIGHT)    */ \
      /* before SetWidth[XX]().                                                                          */ \
      /*                                                                                                 */ \
      /* If the frame corresponds to a control that supports multi-line wrapping text and there is a     */ \
      /* single word (commonly, filename) that is wider than the specified width, the width will         */ \
      /* increase by up to 90% to try to accommodate it.  If it still can't fit on one line, the extra   */ \
      /* characters will be clipped by static string controls, or wrapped to subsequent lines by         */ \
      /* read-only edit (incl. RichEdit) controls.  (To disable this automatic widening, call            */ \
      /* SetSetWidthFlags(MSODAL_SW_NOWIDERFORLONGWORD).  To merely change the width flexibility from    */ \
      /* 90% to some other amount, call SetMaxPcntToExceedStrTargetWidth().  Either call must be made    */ \
      /* prior to the call to SetWidth[XX]().)                                                           */ \
      /*                                                                                                 */ \
      /* If the frame corresponds to a control that supports multi-line wrapping text, but the specified */ \
      /* width is greater than needed to show all the text (either because the text is so short that it  */ \
      /* doesn't fill even a single line at that width, or because at the specified wrapping width it    */ \
      /* wraps in such a way that no line extends precisely to the end of the specified width), then     */ \
      /* SetWidth[XX]() will "shrink-wrap" the frame to the minimum width necessary to accommodate its   */ \
      /* text.  To override this behavior and ensure that the frame does not get less width than         */ \
      /* specified, call SetSetWidthFlags(MSODAL_SW_NOSHRINKWRAP) before SetWidth[XX]().                 */ \
      /*                                                                                                 */ \
      /* For CompositeFrames, SetWidth[XX]() will only increase the width from the minimum width         */ \
      /* required to accommodate its children -- never decrease it.                                      */ \
      \
   MSOMETHOD_(void,SetHeight) (THISPTR_ int height); \
   MSOMETHOD_(void,SetHeightDU) (THISPTR_ int height); \
      /* Sets the height of the frame (immediately -- not during Layout().  For CompositeFrames,     */ \
      /* SetHeight[XX]() will only increase the height from the minimum height required to           */ \
      /* accommodate its children -- never decrease it.                                              */ \
      \
   MSOMETHOD_(void,Indent) (THISPTR_ int numIndents DEFAULTS_TO(1) ); \
      /* Indent (from left) by the standard indent width (MsoDALCheckBoxBMWidthPlusMargin()) */ \
      /* Multiple calls to this and the IndentBy[XX]() methods are cumulative.               */ \
      \
   MSOMETHOD_(void,IndentBy) (THISPTR_ int indentWidth ); \
   MSOMETHOD_(void,IndentByDU) (THISPTR_ int indentWidth ); \
      /* Indent (from left) by an arbitrary amount, in pixels or dialog units, respectively */ \
      /* Multiple calls to these and the Indent() method are cumulative.                    */ \
      \
   MSOMETHOD_(void,LowerBy) (THISPTR_ int height ); \
   MSOMETHOD_(void,LowerByDU) (THISPTR_ int height ); \
      /* Just like IndentBy() and IndentByDU(), but for the vertical dimension. */ \
      \
   MSOMETHOD_(void,SetMargin) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetMarginDU) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetMarginChars) (THISPTR_ int numChars); \
      /* Sets the margin that precedes this frame (i.e., appears between this frame and the previous */ \
      /* sibling) iff the parent frame is a VertFrame or HorizFrame.  A no-op otherwise.             */ \
      \
   MSOMETHOD_(void,ExpandToFitStr) (THISPTR_ const WCHAR* wzStr); \
      /* For a control that can have different strings dynamically set into it, call this method           */ \
      /* once for each possible string value, to ensure that the frame is large enough to accommodate      */ \
      /* any of the possible values.  Any necessary space for the non-text parts of the control (for       */ \
      /* groupboxes, checkboxes, etc.) is automatically added.  This method expands the frame              */ \
      /* immediately -- not during Layout().  The string is treated as single-line, even if SetWidth[XX]() */ \
      /* was previously called on the frame, unless the string contains actual newline characters, in      */ \
      /* which case the frame may expand vertically as well as horizontally.                               */ \
      \
   MSOMETHOD_(void,AlignWithGroupBox) (THISPTR); \
      /* Rarely used.  If the frame's parent is a TableFrame, and AddGroupBox[XX]() has been called, the */ \
      /* appropriate row and column margins were increased to fit the 4 segments of the groupbox control */ \
      /* in them, and no frames -- just other TableFrame-owned, margin-dwelling groupbox controls -- are */ \
      /* able to precisely align with them ... unless this method is called.  Makes the frame jut into a */ \
      /* row or column margin (which it otherwise will never do) to align with the appropriate groupbox. */ \
      \
   MSOMETHOD_(void,SetSetWidthFlags) (THISPTR_ int flags); \
      /* Set MsoDALSetWidthFlag combinations to alter the behavior of subsequent calls to SetWidthXX() */ \
      \
   MSOMETHOD_(int,DUToPixH) (THISPTR_ int x) METHOD_CONST; \
   MSOMETHOD_(int,DUToPixV) (THISPTR_ int y) METHOD_CONST; \
      /* Convert from DUs to pixels.                                              */ \
      /* To call these after Layout(), a FrameLifetime (or similar) must be used. */ \
      \
   MSOMETHOD_(int,Width) (THISPTR) METHOD_CONST; \
   MSOMETHOD_(int,Height) (THISPTR) METHOD_CONST; \
      /* Return the current size of the frame.                                    */ \
      /* To call these after Layout(), a FrameLifetime (or similar) must be used. */ \
      \
   MSOMETHOD_(Frame*,GetFrame) (THISPTR_ MsoDALFrameData msodalfd) METHOD_CONST; \
      /* return parent, first child, or next sibling frame data */

   FRAME_VIRTUALS

#if DAL_CPLUSPLUS
protected:
   Frame();

   FrameImp* m_pImp;

   friend class FrameEqualizerImp;
   friend class FrameImp;
#endif // DAL_CPLUSPLUS
};

#undef  CLASSNAME
#define CLASSNAME CompositeFrame

typedef enum ExpandChildrenAlgorithm
{
   NO_EXPANSION = 0,
   SAME_SIZE, // Expand children to completely fill parent, forcing all children 
              // to be the same size in the specified dimension, EVEN IF THE
              // PARENT FRAME MUST GET LARGER TO ACCOMMODATE THEM
   PROPORTIONATE, // Expand children proportionately to their pre-expanded sizes
                  // to completely fill parent
   TOWARD_EQUAL_SIZES // Expand children to completely fill parent, making all children
                      // the same size if there is already sufficient space in parent
                      // frame to do so, but never force the parent to grow in order
                      // to achieve this equality.
} ExpandChildrenAlgorithm;

typedef enum AlignDir
{
   ALIGN_TOPLEFT,
   ALIGN_TOP, // means TOPCENTER
   ALIGN_TOPRIGHT,
   ALIGN_LEFT, // means CENTERLEFT
   ALIGN_CENTER,
   ALIGN_RIGHT, // means CENTERRIGHT
   ALIGN_BOTTOMLEFT,
   ALIGN_BOTTOM, // means BOTTOMCENTER
   ALIGN_BOTTOMRIGHT
} AlignDir;

DECLARE_CLASS_(CompositeFrame, Frame)
{
CPP_ONLY(public:)

#define COMPOSITEFRAME_VIRTUALS_NEW \
   MSOMETHOD_(CompositeFrame REFERENCE, CPP_ELSE(operator<<, InsertFrame)) (THISPTR_ Frame REFERENCE frame); \
      /* Insert this frame into the specified parent.  Rarely needed, since it's better to pass */ \
      /* the parent as a param at the time of frame construction.                               */ \
      \
   MSOMETHOD_(BOOL,Layout) (THISPTR_ MsoDALLayoutParams* pParams DEFAULTS_TO(NULL) ); \
      /* Perform the layout.  Also destructs all frames for the current IAutoLayoutDialog (even frames   */ \
      /* not inserted into the tree), unless a FrameLifetime (or similar) is preventing the destruction. */ \
      /* Returns TRUE iff succeeded.                                                                     */ \
      \
   MSOMETHOD_(void,AlignChildren) (THISPTR_ AlignDir dir); \
      /* Set the direction in which all children are to be aligned. */ \
      \
   MSOMETHOD_(void,ExpandChildrenToFillHoriz) (THISPTR_ ExpandChildrenAlgorithm expand DEFAULTS_TO(TOWARD_EQUAL_SIZES) ); \
   MSOMETHOD_(void,ExpandChildrenToFillVert) (THISPTR_ ExpandChildrenAlgorithm expand DEFAULTS_TO(TOWARD_EQUAL_SIZES) ); \
      /* Specify that this frame's children are to expand (during Layout()) to fill the available space,  */ \
      /* using the specified algorithm.  DontExpand[Horiz/Vert]() may be called on individual children to */ \
      /* prevent them from expanding.                                                                     */ \
      \
   MSOMETHOD_(void,SameWidthChildren) (THISPTR); \
   MSOMETHOD_(void,SameHeightChildren) (THISPTR); \
      /* Specify that during Layout(), all children are to be made the same size as the largest one in */ \
      /* the specified dimension.  DontExpand[Horiz/Vert]() may be called on individual children to    */ \
      /* prevent them from expanding.                                                                  */ \
      \
   MSOMETHOD_(void,SetWidthsChildren) (THISPTR_ int width); \
   MSOMETHOD_(void,SetWidthFlagsChildren) (THISPTR_ int flags); \
   MSOMETHOD_(void,RemoveDlgBorderMargins) (THISPTR); \
      /* Remove the margins that appear around the perimeter of the overall dialog by default.          */ \
      /* Causes the 'AutoLayout' logo to be drawn transparently, to prevent it from obscuring controls. */ \
      \
   MSOMETHOD_(void,SetDlgClientWidth) (THISPTR_ int width); \
   MSOMETHOD_(void,SetDlgClientHeight) (THISPTR_ int height); \
      /* Enlarges the dialog to the specified client size, but will never shrink the dialog */ \
      /* from the minimum size required to fit all the frames.                              */ \
      \
   MSOMETHOD_(void,SetDialogUsesMultiFonts) (THISPTR_ bool fUsesMultiFonts ); \
      /* Specifies that the current dialog may use fonts other than the dialog font. */ \
      /* Setting this to true will calculate control sizes correctly for each font,  */ \
      /* but will result in more computation time.                                   */ \
      \
   MSOMETHOD_(void,SetPerimeterMargin) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetPerimeterMarginDU) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetLeftMargin) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetRightMargin) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetTopMargin) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetBottomMargin) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetLeftMarginDU) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetRightMarginDU) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetTopMarginDU) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetBottomMarginDU) (THISPTR_ int margin);
#define COMPOSITEFRAME_VIRTUALS \
            C_ONLY( FRAME_VIRTUALS ) \
            COMPOSITEFRAME_VIRTUALS_NEW

   COMPOSITEFRAME_VIRTUALS

#if DAL_CPLUSPLUS
protected:
   CompositeFrame();
#endif // DAL_CPLUSPLUS
};


#undef  CLASSNAME
#define CLASSNAME HorizFrame

DECLARE_CLASS_(HorizFrame, CompositeFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR HorizFrame( IMsoAutoLayoutDialog& dlg, CompositeFrame* pParentFrame );
   MSOCTOR HorizFrame( IMsoAutoLayoutDialog& dlg, void* horizData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define HORIZFRAME_VIRTUALS_NEW \
	MSOMETHOD_(void,DistributeChildren) (THISPTR); \
      /* Makes the children repel each other, and spread throughout the available width */ \
      /* with equal distances between them all.                                         */ \
      \
	MSOMETHOD_(void,SetAllMargins) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetAllMarginsDU) (THISPTR_ int margin); \
      /* Equivalent to calling SetMargin[XX]() with the given value on every child, except for */ \
      /* the first child, whose margin is set to 0 (which it already is by default).           */

#define HORIZFRAME_VIRTUALS \
	C_ONLY( COMPOSITEFRAME_VIRTUALS ) \
	HORIZFRAME_VIRTUALS_NEW

   HORIZFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME VertFrame

DECLARE_CLASS_(VertFrame, CompositeFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR VertFrame( IMsoAutoLayoutDialog& dlg, CompositeFrame* pParentFrame );
   MSOCTOR VertFrame( IMsoAutoLayoutDialog& dlg, void* vertData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define VERTFRAME_VIRTUALS_NEW \
	MSOMETHOD_(void,DistributeChildren) (THISPTR); \
      /* Makes the children repel each other, and spread throughout the available height */ \
      /* with equal distances between them all.                                          */ \
      \
	MSOMETHOD_(void,SetAllMargins) (THISPTR_ int margin); \
	MSOMETHOD_(void,SetAllMarginsDU) (THISPTR_ int margin); \
      /* Equivalent to calling SetMargin[XX]() with the given value on every child, except for the */ \
      /* first child, whose margin is set to 0 (which it already is by default).                 */

#define VERTFRAME_VIRTUALS \
	C_ONLY( COMPOSITEFRAME_VIRTUALS ) \
	VERTFRAME_VIRTUALS_NEW

	VERTFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME FlowFrame

DECLARE_CLASS_(FlowFrame, CompositeFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR FlowFrame( IMsoAutoLayoutDialog& dlg, CompositeFrame* pParentFrame );
   MSOCTOR FlowFrame( IMsoAutoLayoutDialog& dlg, void* flowData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define FLOWFRAME_VIRTUALS_NEW \
   MSOMETHOD_(void,SetVertMargin) (THISPTR_ int pixelMargin); \
   MSOMETHOD_(void,SetVertMarginDU) (THISPTR_ int marginDU); \
      /* Set the vertical margin between rows.  (All rows share the same margin,      */ \
      /* since it is unknown which children will end up in which rows anyway).        */ \
      /* (Horizontal margins between children are set by calling Frame::SetMarginXX() */ \
      /* on the individual children.)                                                 */ \
      \
   MSOMETHOD_(void,ShrinkWidthBy) (THISPTR_ int delta); \
   MSOMETHOD_(void,ShrinkWidthByDU) (THISPTR_ int deltaDU); \
      /* Reduce flow width by the specified amount.  By default, the flow      */ \
      /* width is equal to the dialog's visible client width,                  */ \
      /* minus any left and right border margins.  So, for example, if the     */ \
      /* FlowFrame is inside a GroupBoxFrame of type UNDERLINEDHEADING (the    */ \
      /* type used in workpanes), you should call                              */ \
      /* ShrinkWidthByDU( UNDERLINEDHEADING_LEFT_MARGIN_DU ) to account for    */ \
      /* the extra indenting of the FlowFrame from the left border of the      */ \
      /* dialog, so that the right side of the FlowFrame doesn't extend beyond */ \
      /* the right side of the dialog.  Alternatively, SetWidthXX() may be     */ \
      /* called on the FlowFrame to set a different flow width altogether.     */ \
      /* If Indent() or IndentByXX() is called on the FlowFrame, a             */ \
      /* corresponding call to ShrinkWidthByXX() is not necessary -- the       */ \
      /* flow width will shrink automatically by the amount of the indent.     */ \
      \
   MSOMETHOD_(void,IndentLatterRows) (THISPTR); \
      /* Indent all rows after the first, by standard indent width. */ \
      \
   MSOMETHOD_(void,IndentLatterRowsBy) (THISPTR_ int width); \
   MSOMETHOD_(void,IndentLatterRowsByDU) (THISPTR_ int widthDU); \
      /* Indent all rows after the first, by the specified width. */ \
      \
   MSOMETHOD_(void,IndentLatterRowsByFirstChild) (THISPTR); \
      /* Indent all rows after the first, such that the left edge of the  */ \
      /* first child in each of them aligns with the left edge of the     */ \
      /* second child of the first row.  If ANY latter row child is too   */ \
      /* wide to fit (even in its own row) after being indented by that   */ \
      /* amount, then ALL the latter rows are indented only by the        */ \
      /* standard indent width (assuming it is less), and the FlowFrame's */ \
      /* second child will move down to start on a new row, also indented */ \
      /* by the standard indent width.                                    */ \
      \
   MSOMETHOD_(void,DistributeRows) (THISPTR); \
      /* Makes the rows repel each other, and spread throughout the */ \
      /* available space with equal distances between them all.     */ \
      \
   MSOMETHOD_(void,DistributeChildrenHoriz) (THISPTR); \
      /* Makes the children in a row repel each other, and spread the */ \
      /* available space with equal horizontalmargins between them.   */ \
      \
	MSOMETHOD_(void,SetAllMargins) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetAllMarginsDU) (THISPTR_ int margin); \
      /* Equivalent to calling SetMargin[XX]() with the given value on every child, except for */ \
      /* the first child, whose margin is set to 0 (which it already is by default).           */ \
      \
   MSOMETHOD_(void,BindChildPair) (THISPTR_ Frame REFERENCE firstChild); \
      /* Bind firstChild with the child immediately following it.  (Can also */ \
      /* be used to bind a set of more than two children.)  The bound set of */ \
      /* children will all be kept in the same row whenever the flow         */ \
      /* width is great enough to make that possible, even if that results   */ \
      /* in less efficient use of the preceding and succeeding rows.  But    */ \
      /* if the flow width is too narrow to hold them all in the same row,   */ \
      /* then the bound set will be broken up into multiple rows (without    */ \
      /* any special indentation of the latter rows of the bound set, aside  */ \
      /* from existing indentation due to IndentLatterRowsXX()).  If a bound */ \
      /* set of children must be broken into multiple rows, then no child    */ \
      /* from outside of that bound set will appear on the same row as any   */ \
      /* child from that bound set. */ \
      \
   MSOMETHOD_(int,GetRowOccupiedByChild) (THISPTR_ Frame REFERENCE child); \
      /* Returns the row number that the child is in.  Can only be called        */ \
      /* after Layout(), which means a FrameLifetime (or equivalent) must        */ \
      /* be used to prevent the FlowFrame from being destructed inside Layout(). */ \
      \
   MSOMETHOD_(void,ExpandChildToFillRowVert) (THISPTR_ Frame REFERENCE child); \
      /* Expand the child to fill the vertical space in the row, WITHOUT expanding */ \
      /* the row height.  To also expand the row to fill all available vertical    */ \
      /* space in the FlowFrame, call ExpandToFillVert() on the child.             */

#define FLOWFRAME_VIRTUALS \
	C_ONLY( COMPOSITEFRAME_VIRTUALS ) \
	FLOWFRAME_VIRTUALS_NEW

	FLOWFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME VertFlowFrame

DECLARE_CLASS_(VertFlowFrame, CompositeFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR VertFlowFrame( IMsoAutoLayoutDialog& dlg, CompositeFrame* pParentFrame );
   MSOCTOR VertFlowFrame( IMsoAutoLayoutDialog& dlg, void* flowData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define VERTFLOWFRAME_VIRTUALS_NEW \
   MSOMETHOD_(void,SetHorizMargin) (THISPTR_ int pixelMargin); \
   MSOMETHOD_(void,SetHorizMarginDU) (THISPTR_ int marginDU); \
      /* Set the horizontal margin between rows.  (All columns share the same margin, */ \
      /* since it is unknown which children will end up in which columns anyway).     */ \
      /* (Vertical margins between children are set by calling Frame::SetMarginXX()   */ \
      /* on the individual children.)                                                 */ \
      \
   MSOMETHOD_(void,ShrinkHeightBy) (THISPTR_ int delta); \
   MSOMETHOD_(void,ShrinkHeightByDU) (THISPTR_ int deltaDU); \
      /* Reduce flow height by the specified amount.  By default, the flow      */ \
      /* height is equal to the dialog's visible client height,                 */ \
      /* minus any top and bottom border margins.  So, for example, if the      */ \
      /* FlowFrame is inside a GroupBoxFrame of type UNDERLINEDHEADING (the     */ \
      /* type used in workpanes), you should call                               */ \
      /* ShrinkHeightByDU( UNDERLINEDHEADING_LEFT_MARGIN_DU ) to account for    */ \
      /* the extra indenting of the FlowFrame from the top border of the        */ \
      /* dialog, so that the bottom of the FlowFrame doesn't extend beyond      */ \
      /* the bottom of the dialog.  Alternatively, SetHeightXX() may be         */ \
      /* called on the VertFlowFrame to set a different flow height altogether. */ \
      \
   MSOMETHOD_(void,DistributeCols) (THISPTR); \
      /* Makes the columns repel each other, and spread throughout  */ \
      /* the available space with equal distances between them all. */ \
      \
   MSOMETHOD_(void,DistributeChildrenVert) (THISPTR); \
      /* Makes the children in a column repel each other, and spread   */ \
      /* the available space with equal vertical margins between them. */ \
      \
   MSOMETHOD_(void,SameWidthColumns) (THISPTR); \
      /* Makes all columns have the same width as the widest column. */ \
      \
	MSOMETHOD_(void,SetAllMargins) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetAllMarginsDU) (THISPTR_ int margin); \
      /* Equivalent to calling SetMargin[XX]() with the given value on every child, except for */ \
      /* the first child, whose margin is set to 0 (which it already is by default).           */ \
      \
   MSOMETHOD_(void,BindChildPair) (THISPTR_ Frame REFERENCE firstChild); \
      /* Bind firstChild with the child immediately following it.  (Can also      */ \
      /* be used to bind a set of more than two children.)  The bound set of      */ \
      /* children will all be kept in the same column whenever the flow           */ \
      /* height is great enough to make that possible, even if that results       */ \
      /* in less efficient use of the preceding and succeeding columns.  But      */ \
      /* if the flow height is too short to hold them all in the same column,     */ \
      /* then the bound set will be broken up into multiple columns .  If a bound */ \
      /* set of children must be broken into multiple columns, then no child      */ \
      /* from outside of that bound set will appear on the same column as any     */ \
      /* child from that bound set. */ \
      \
   MSOMETHOD_(int,GetColOccupiedByChild) (THISPTR_ Frame REFERENCE child); \
      /* Returns the column number that the child is in.  Can only be called      */ \
      /* after Layout(), which means a FrameLifetime (or equivalent) must be      */ \
      /* used to prevent the VertFlowFrame from being destructed inside Layout(). */ \
      \
   MSOMETHOD_(void,ExpandChildToFillColHoriz) (THISPTR_ Frame REFERENCE child); \
      /* Expand the child to fill the horizontal space in the column, WITHOUT expanding */ \
      /* the column width.  To also expand the column to fill all available horizontal  */ \
      /* space in the VertFlowFrame, call ExpandToFillHoriz() on the child.             */

#define VERTFLOWFRAME_VIRTUALS \
	C_ONLY( COMPOSITEFRAME_VIRTUALS ) \
	VERTFLOWFRAME_VIRTUALS_NEW

	VERTFLOWFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME TableFrame

DECLARE_CLASS_(TableFrame, CompositeFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR TableFrame( IMsoAutoLayoutDialog& dlg, int cColumns, int cRows, CompositeFrame* pParentFrame );
   MSOCTOR TableFrame( IMsoAutoLayoutDialog& dlg, int cColumns, int cRows, void* tableData,
                       CompositeFrame* pParentFrame );
	MSOOVERRIDEMETHOD_(CompositeFrame&) operator<< (THISPTR_ Frame& frame); // override
      // Duplicated from CompositeFrame::operator<< so not hidden by overload.  Does not occupy
      // an extra v-table entry since it's an override.  Hence no C placeholder necessary.
#endif // DAL_CPLUSPLUS

#define TABLEFRAME_VIRTUALS_NEW \
	MSOMETHOD_(TableFrame REFERENCE, CPP_ELSE(operator<<, InsertEmptyCell)) (THISPTR_ int zero); \
      /* Pass 0 to reserve the next cell in the table as empty.  The next actual child to be */ \
      /* inserted will skip this cell and go into the following one.                         */ \
      \
   MSOMETHOD_(void,InsertChildAtCell) (THISPTR_ Frame REFERENCE child, int iCol, int iRow, \
                                      int cCols DEFAULTS_TO( 1 ), \
                                      int cRows DEFAULTS_TO( 1 )); \
      /* Insert a child into a specific cell or cells, with cCols meaning column span and  */ \
      /* cRows meaning row span.  */ \
      \
   MSOMETHOD_(void,ExpandChildToFillHoriz) (THISPTR_ Frame REFERENCE child); \
   MSOMETHOD_(void,ExpandChildToFillVert) (THISPTR_ Frame REFERENCE child); \
      /* Expand the row/col where the child is in to occupy the empty space in the frame, and */ \
      /* then expand the child vertically/horizontally to fill the space in the expanded cell */ \
      \
   MSOMETHOD_(void,ExpandChildToFillCellHoriz) (THISPTR_ Frame REFERENCE child); \
   MSOMETHOD_(void,ExpandChildToFillCellVert) (THISPTR_ Frame REFERENCE child); \
      /* Expand the child vertically/horizontally to fill the space of the cell. The size of */ \
      /* the row/col the child is in is NOT changed. */ \
      \
   MSOMETHOD_(void,SetColMargin) (THISPTR_ int iCol, int margin); \
   MSOMETHOD_(void,SetColMarginDU) (THISPTR_ int iCol, int margin); \
   MSOMETHOD_(void,SetRowMargin) (THISPTR_ int iRow, int margin); \
   MSOMETHOD_(void,SetRowMarginDU) (THISPTR_ int iRow, int margin); \
      /* Set the margin preceding (above or to the left of) the specified (0-based) column or row.  There */ \
      /* is one more row margin than there are rows, so to set the size of the row margin along the very  */ \
      /* bottom of the TableFrame (beneath the last row), pass iRow = cRows.  Similarly for columns.      */ \
      \
   MSOMETHOD_(void,SetAllColMargins) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetAllColMarginsDU) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetAllRowMargins) (THISPTR_ int margin); \
   MSOMETHOD_(void,SetAllRowMarginsDU) (THISPTR_ int margin); \
      /* Equivalent to calling SetColMargin[XX]() or SetRowMargin[XX]() with the given value on every column */ \
      /* or row, respectively, except for the first (index=0) and the imaginary "post-last" (index=cColumns  */ \
      /* or index=cRows) columns/rows, whose margins are set to 0 (which they already are by default).       */ \
      \
   MSOMETHOD_(void,DistributeCols) (THISPTR); \
   MSOMETHOD_(void,DistributeRows) (THISPTR); \
      /* Makes the columns or rows (respectively) repel each other, and spread throughout the */ \
      /* available space with equal distances between them all.                               */ \
      \
   MSOMETHOD_(void,AddGroupBox) (THISPTR_ void* groupCtrlData, int left, int top, \
                                int right, int bottom, BOOL fNeverHasText DEFAULTS_TO(FALSE)); \
   MSOMETHOD_(void,AddGroupBox2Ctrls) \
                                (THISPTR_ void* textCtrlData, void* lineCtrlData, int left, int top, \
                                int right, int bottom, BOOL fNeverHasText DEFAULTS_TO(FALSE)); \
      /* Adds a groupbox control (including groupline) in the row and column margins, specifying the 4    */ \
      /* "vertices" on the grid of margin intersections that the groupbox is to connect.  The grid coords */ \
      /* are 0-based, with (0,0) being the top-left corner of the TableFrame (above and to the left of    */ \
      /* the first child, and (cCols,cRows) being the bottom-right corner of the TableFrame (below and    */ \
      /* to the right of the last child).                                                                 */ \
      /* (E.g. left=0, top=0, right=3, bottom=2 would create a groupbox starting at the extreme top-left  */ \
      /* and extending 3 cells across and 2 cells down, thus enclosing 6 cells.)                          */ \
      \
   MSOMETHOD_(void,ExpandColToFillHoriz) (THISPTR_ int iCol); \
   MSOMETHOD_(void,ExpandRowToFillVert) (THISPTR_ int iRow); \
      /* Call ExpandChildToFillVert/Horiz on all children in the row/col to expand all of them */ \
      \
   MSOMETHOD_(void,DontExpandColHoriz) (THISPTR_ int iCol, bool fIncludeMultiColChildren DEFAULTS_TO(true)); \
   MSOMETHOD_(void,DontExpandRowVert) (THISPTR_ int iRow, bool fIncludeMultiRowChildren DEFAULTS_TO(true)); \
      /* Set flags in each children in the row/col so they will not expand */ \
      \
   MSOMETHOD_(void,SetWidthColChildren) (THISPTR_ int iCol, int width, bool fIncludeMultiColChildren DEFAULTS_TO(true)); \
   MSOMETHOD_(void,SetWidthColChildrenDU) (THISPTR_ int iCol, int width, bool fIncludeMultiColChildren DEFAULTS_TO(true)); \
   MSOMETHOD_(void,SetHeightRowChildren) (THISPTR_ int iRow, int height, bool fIncludeMultiRowChildren DEFAULTS_TO(true)); \
   MSOMETHOD_(void,SetHeightRowChildrenDU) (THISPTR_ int iRow, int height, bool fIncludeMultiRowChildren DEFAULTS_TO(true)); \
      /* Set the height/width of all children in the row/col specified */

#define TABLEFRAME_VIRTUALS \
	C_ONLY( COMPOSITEFRAME_VIRTUALS ) \
	TABLEFRAME_VIRTUALS_NEW

	TABLEFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME ContainerFrame

DECLARE_CLASS_(ContainerFrame, CompositeFrame)
{
CPP_ONLY(public:)

#define CONTAINERFRAME_VIRTUALS \
	C_ONLY( COMPOSITEFRAME_VIRTUALS )
	CONTAINERFRAME_VIRTUALS

#if DAL_CPLUSPLUS
protected:
   ContainerFrame();
#endif // DAL_CPLUSPLUS
};


#undef  CLASSNAME
#define CLASSNAME CtrlContainerFrame

DECLARE_CLASS_(CtrlContainerFrame, ContainerFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR CtrlContainerFrame( IMsoAutoLayoutDialog& dlg, void* ctrlData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define CTRLCONTAINERFRAME_VIRTUALS  C_ONLY( CONTAINERFRAME_VIRTUALS )

	CTRLCONTAINERFRAME_VIRTUALS 
};


#undef  CLASSNAME
#define CLASSNAME GroupBoxFrame

DECLARE_CLASS_(GroupBoxFrame, ContainerFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR GroupBoxFrame( IMsoAutoLayoutDialog& dlg, void* groupCtrlData, CompositeFrame* pParentFrame );
   MSOCTOR GroupBoxFrame( IMsoAutoLayoutDialog& dlg, void* textCtrlData, void* lineCtrlData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define GROUPBOXFRAME_VIRTUALS_NEW \
   MSOMETHOD_(void,SetNeverHasText) (THISPTR);
   // Specifies that this groupbox does not have (and will never have) text, so no extra space should
   // be allocated above the top crossbar, and less space should be allocated beneath the crossbar.

#define GROUPBOXFRAME_VIRTUALS \
	C_ONLY( CONTAINERFRAME_VIRTUALS ) \
	GROUPBOXFRAME_VIRTUALS_NEW

	GROUPBOXFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME OverlapFrame

DECLARE_CLASS_(OverlapFrame, ContainerFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR OverlapFrame( IMsoAutoLayoutDialog& dlg, CompositeFrame* pParentFrame );
   MSOCTOR OverlapFrame( IMsoAutoLayoutDialog& dlg, void* overlapData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define OVERLAPFRAME_VIRTUALS  C_ONLY( CONTAINERFRAME_VIRTUALS )

	OVERLAPFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME TabCtrlFrame

DECLARE_CLASS_(TabCtrlFrame, ContainerFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR TabCtrlFrame( IMsoAutoLayoutDialog& dlg, void* tabCtrlData, CompositeFrame* pParentFrame );
   MSOCTOR TabCtrlFrame( IMsoAutoLayoutDialog& dlg, void* tabsData, void* sheetBorderData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define TABCTRLFRAME_VIRTUALS  C_ONLY( CONTAINERFRAME_VIRTUALS )

	TABCTRLFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME LeafFrame

DECLARE_CLASS_(LeafFrame, Frame)
{
CPP_ONLY(public:)

#if DAL_CPLUSPLUS
   enum // sample param values for SetMaxPcntToExceedStrTargetWidth():
   {
      UNLIMITED_WIDTH = -1, // Make string wide enough so that any contained filename 
                            // or other really long word that is not broken up by spaces 
                            // will fit on its own line, NO MATTER HOW WIDE.
      FORCE_FIXED_WIDTH = 0, // String may not deviate from the specified width.  Same as calling
                             // SetSetWidthFlags(MSODAL_SW_NOWIDERFORLONGWORD), which is preferred.
      DEFAULT_MAX_PERCENT_WIDER = 90, // Allow string to widen to fit a big word on its
                                      // own line, up to n% LONGER THAN (not "of") the
                                      // specified width.
   }; // TODO: This enum should be moved outside of class for C support
#endif // DAL_CPLUSPLUS

#define LEAFFRAME_VIRTUALS_NEW \
   MSOMETHOD_(void,AddLeftString) (THISPTR_ void* strCtrlData); \
   MSOMETHOD_(void,AddTopString) (THISPTR_ void* strCtrlData); \
   MSOMETHOD_(void,AddRightString) (THISPTR_ void* strCtrlData); \
   MSOMETHOD_(void,SetHeightLines) (THISPTR_ int numLines); \
   MSOMETHOD_(void,SetMaxPcntToExceedStrTargetWidth) (THISPTR_ int percent); \
      /* See comments at Frame::SetWidth().  Use of SetSetWidthFlags(MSODAL_SW_NOWIDERFORLONGWORD) */ \
      /* is preferred to SetMaxPcntToExceedStrTargetWidth(0).                                      */ \
      \
   MSOMETHOD_(void,IgnoreSetWidthIfStrHasNewline) (THISPTR); /* OBSOLETE -- Use SetSetWidthFlags(MSODAL_SW_NOOPIFSTRHASNEWLINE) instead */ \
   MSOMETHOD_(void,DisableSetWidthFlexibility) (THISPTR); /* OBSOLETE -- Use SetSetWidthFlags(MSODAL_SW_INCLUDESTRIM | MSODAL_SW_NOWIDERFORLONGWORD) instead */ \
   MSOMETHOD_(void,IncreaseWidthBy) (THISPTR_ int increase); \
   MSOMETHOD_(void,IncreaseHeightBy) (THISPTR_ int increase); \
   MSOMETHOD_(void,IncreaseWidthByDU) (THISPTR_ int increase); \
   MSOMETHOD_(void,IncreaseHeightByDU) (THISPTR_ int increase);

#define LEAFFRAME_VIRTUALS \
            C_ONLY( FRAME_VIRTUALS ) \
            LEAFFRAME_VIRTUALS_NEW

	LEAFFRAME_VIRTUALS

#if DAL_CPLUSPLUS
protected:
   LeafFrame();
#endif // DAL_CPLUSPLUS
};


#undef  CLASSNAME
#define CLASSNAME CtrlFrame

DECLARE_CLASS_(CtrlFrame, LeafFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR CtrlFrame( IMsoAutoLayoutDialog& dlg, void* ctrlData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define CTRLFRAME_VIRTUALS_NEW \
   MSOMETHOD_(void,SetWidthFromCtrl) (THISPTR); \
   MSOMETHOD_(void,SetHeightFromCtrl) (THISPTR); \
   MSOMETHOD_(void,SetSizeFromCtrl) (THISPTR); \
   MSOMETHOD_(void,SetWidthToFitPreexistingItems) (THISPTR_ int minWidthChars DEFAULTS_TO(5), \
                                                            int maxWidthChars DEFAULTS_TO(80), \
                                                            int extraWidthChars DEFAULTS_TO(0)); \
   MSOMETHOD_(void,SetLeftAdjust) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetRightAdjust) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetTopAdjust) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetBottomAdjust) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetLeftAdjustDU) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetRightAdjustDU) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetTopAdjustDU) (THISPTR_ int adjust); \
   MSOMETHOD_(void,SetBottomAdjustDU) (THISPTR_ int adjust); \
   MSOMETHOD_(void,ShrinkWrapToText) (THISPTR_ bool fShrinkWrap); \
      /* If fShrinkWrap is true, reduce control width to tightly fit its text.  If false, do not reduce    */ \
      /* control width.  This API only applies to CHECKBOX, RADIOBUTTON, HYPERLINK, and STATIC_STRING.     */ \
      /* By default, STATIC_STRING does not shrink wrap, all others do.                                    */ \
      /* Note that this shrink wrap only changes the control width, not frame width.                       */

#define CTRLFRAME_VIRTUALS \
	C_ONLY( LEAFFRAME_VIRTUALS ) \
	CTRLFRAME_VIRTUALS_NEW

	CTRLFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME NumInputFrame

DECLARE_CLASS_(NumInputFrame, LeafFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR NumInputFrame( IMsoAutoLayoutDialog& dlg, void* numInputData, CompositeFrame* pParentFrame );
   MSOCTOR NumInputFrame( IMsoAutoLayoutDialog& dlg, void* editStrData, void* spinnerData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define NUMINPUTFRAME_VIRTUALS_NEW \
	MSOMETHOD_(void,SetWidthDigits) (THISPTR_ int numDigits);
	
#define NUMINPUTFRAME_VIRTUALS \
	C_ONLY( LEAFFRAME_VIRTUALS ) \
	NUMINPUTFRAME_VIRTUALS_NEW

	NUMINPUTFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME ListEditFrame

DECLARE_CLASS_(ListEditFrame, LeafFrame)
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR ListEditFrame( IMsoAutoLayoutDialog& dlg, void* editStrData, void* listBoxData, CompositeFrame* pParentFrame );
#endif // DAL_CPLUSPLUS

#define LISTEDITFRAME_VIRTUALS_NEW \
	MSOMETHOD_(void,SetWidthToFitPreexistingItems) (THISPTR_ int minWidthChars DEFAULTS_TO(5), \
		                                                      int maxWidthChars DEFAULTS_TO(80), \
		                                                      int extraWidthChars DEFAULTS_TO(0));
	
#define LISTEDITFRAME_VIRTUALS \
	C_ONLY( LEAFFRAME_VIRTUALS ) \
	LISTEDITFRAME_VIRTUALS_NEW

	LISTEDITFRAME_VIRTUALS
};


#undef  CLASSNAME
#define CLASSNAME FrameEqualizer

DECLARE_CLASS(FrameEqualizer)
//
// A FrameEqualizer takes a set of LeafFrames, finds the greatest width
// ("and/or height", from here on in) among them, and sets the widths of all 
// others in the set to the that greatest width.  The widening takes place
// at the beginning of CompositeFrame::Layout(), BEFORE any requested
// expansion-to-fill takes effect, so the final expanded widths may not be
// equal.  To cause equalization to take place prior to the call to Layout(), 
// clients may optionally call EqualizeNow().  (In this case it will NOT
// equalize again inside Layout().) 
//
// Since CompositeFrame::SameWidthChildren(), on the other hand, DOES
// guarantee equal widths even after expansion-to-fill, its use is generally
// preferable to that of FrameEqualizer.  FrameEqualizer 
// is provided for those cases when SameWidthChildren() cannot be used, due
// to the positions of the frames to be equalized in the frame hierarchy.
//
// Usage:
//    FrameEqualizer btnsEqualizer(dlg);
//    btnsEqualizer << okBtnFrame << cancelBtnFrame << someOtherFrame;
//    //btnsEqualizer.EqualizeNow(); // This call is optional
//
// Currently only supports width equalization -- not height.  As an interim workaround, use
// Frame::Height() to find the tallest frame height, then call SetHeight() on all frames.
{
#if DAL_CPLUSPLUS
public:
   MSOCTOR FrameEqualizer( IMsoAutoLayoutDialog& dlg );
#endif // DAL_CPLUSPLUS

#define FRAMEEQUALIZER_VIRTUALS \
   DebugOnly( CPP_ELSE( MSODTORVIRT_DECL ~FrameEqualizer(), \
                        MSOMETHOD_(void,PlaceholderForFrameEqDtorDoNotCall) (THISPTR) ); ) \
      /* No destructor in ship.  See comment at ~Frame for explanation. */ \
   MSOMETHOD_(FrameEqualizer REFERENCE, CPP_ELSE(operator<<, AddFrame)) (THISPTR_ Frame REFERENCE frame); \
	MSOMETHOD_(void,EqualizeNow) (THISPTR); \
      /* Call this to force equalization to take place immediately, instead of inside */ \
      /* the call to Layout().  Most clients should not call this.                    */

	FRAMEEQUALIZER_VIRTUALS
		
#if DAL_CPLUSPLUS
protected:
   FrameEqualizerImp* m_pImp;

   friend class FrameEqualizerImp;
#endif // DAL_CPLUSPLUS
};

#if !DAL_CPLUSPLUS
   // C style constructors for concrete classes:
   void STDMETHODCALLTYPE HorizFrameCtor( HorizFrame* This, IMsoAutoLayoutDialog* dlg, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE HorizFrameCtorWithData( HorizFrame* This, IMsoAutoLayoutDialog* dlg, void* horizData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE VertFrameCtor( VertFrame* This, IMsoAutoLayoutDialog* dlg, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE VertFrameCtorWithData( VertFrame* This, IMsoAutoLayoutDialog* dlg, void* vertData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE FlowFrameCtor( FlowFrame* This, IMsoAutoLayoutDialog* dlg, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE FlowFrameCtorWithData( FlowFrame* This, IMsoAutoLayoutDialog* dlg, void* flowData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE VertFlowFrameCtor( VertFlowFrame* This, IMsoAutoLayoutDialog* dlg, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE VertFlowFrameCtorWithData( VertFlowFrame* This, IMsoAutoLayoutDialog* dlg, void* flowData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE TableFrameCtor( TableFrame* This, IMsoAutoLayoutDialog* dlg, int cColumns, int cRows, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE TableFrameCtorWithData( TableFrame* This, IMsoAutoLayoutDialog* dlg, int cColumns, int cRows, void* tableData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE CtrlContainerFrameCtor( CtrlContainerFrame* This, IMsoAutoLayoutDialog* dlg, void* ctrlData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE GroupBoxFrameCtor( GroupBoxFrame* This, IMsoAutoLayoutDialog* dlg, void* groupCtrlData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE GroupBoxFrameCtor2( GroupBoxFrame* This, IMsoAutoLayoutDialog* dlg, void* textCtrlData, void* lineCtrlData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE OverlapFrameCtor( OverlapFrame* This, IMsoAutoLayoutDialog* dlg, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE OverlapFrameCtorWithData( OverlapFrame* This, IMsoAutoLayoutDialog* dlg, void* overlapData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE TabCtrlFrameCtor( TabCtrlFrame* This, IMsoAutoLayoutDialog* dlg, void* tabCtrlData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE CtrlFrameCtor( CtrlFrame* This, IMsoAutoLayoutDialog* dlg, void* ctrlData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE NumInputFrameCtor( NumInputFrame* This, IMsoAutoLayoutDialog* dlg, void* numInputData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE NumInputFrameCtor2( NumInputFrame* This, IMsoAutoLayoutDialog* dlg, void* editStrData, void* spinnerData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE ListEditFrameCtor( ListEditFrame* This, IMsoAutoLayoutDialog* dlg, void* editStrData, void* listBoxData, CompositeFrame* pParentFrame );
   void STDMETHODCALLTYPE FrameEqualizerCtor( FrameEqualizer* This, IMsoAutoLayoutDialog* dlg );
#endif // !DAL_CPLUSPLUS

typedef enum DalLayoutFlag
{
   LAYOUTFLAG_NONE = 0x00,
   LAYOUTFLAG_NEEDSSECONDPASS = 0x01,
   LAYOUTFLAG_NEEDSSETVISIBLEWIDTH = 0x02
} DalLayoutFlag;

typedef enum DalLayoutData
{
   LAYOUTDATA_CONSTRAINT_PASS1 = 1,
   LAYOUTDATA_CONSTRAINT_PASS2 = 2,
   LAYOUTDATA_POSITION_PASS1 = 3,
   LAYOUTDATA_FLAGS = 4
} DalLayoutData;

#if DAL_CPLUSPLUS
class IMsoAutoLayoutDialog // Must endure until dialog is destroyed, or at least as long as frame outlines may be drawn
{
protected:
   MSOCTOR IMsoAutoLayoutDialog();
   MSODTORVIRT_DECL ~IMsoAutoLayoutDialog();

public:
   MSODALPUBDEBUGMETHODIMP

   //---------------
   // Dialog:
   //---------------

   // DO NOT IMPLEMENT -- WILL BE REMOVED
   MSOMETHOD_(bool,IsModeless)() const = 0;

   // Is the dialog right-to-left?  Typically, it is iff the WS_EX_RIGHT extended style flag is set
   // on it.  If this function returns true, then the control positions will be flipped right-to-left.
   // Note that Middle Eastern languages may have individual dialogs that are NOT right-to-left, so
   // it is generally incorrect to just check the UI language.
   MSOMETHOD_(bool,IsRightToLeft)() const = 0;

   // !!! This function in all dialog managers should examine the resource flags of the control
   // to determine whether the control can wrap text or not
   // However, it works perfectly fine for wrapping label lines if the function
   // only returns true
   // The default implementation returns false
   MSOMETHOD_(bool,FControlWrapsText)( void* /* ctrlData */ ) const { return false; }

   // Return the LID of the current UI language.
   MSOMETHOD_(LID,LanguageDisplayed)() const = 0;

   // Return the dialog or control base units for the current dialog or control font. 
   // ctrlData is ignored if fIsDialog is set to true.
   // For more info, see documentation on the
   // Win32 APIs GetDialogBaseUnits() (http://msdn.microsoft.com/library/sdkdoc/winui/dlgboxes_7wfn.htm)
   // and MapDialogRect() (http://msdn.microsoft.com/library/sdkdoc/winui/dlgboxes_4hf8.htm).
   MSOMETHOD_(void,GetBaseUnits)( int& unitsHoriz, int& unitsVert, void* ctrlData, 
                                  bool fIsDialog ) const = 0;

   // Return the client rectangle of the window whose origin the control coordinates are expressed
   // in terms of.  This can be either the tabsheet window (if there is a tabcontrol and its sheets are
   // implemented using a separate window and the control lives on the tabsheet), or the overall dialog.
   // In Win32 dialogs, this is generally the parent HWND of the control HWND.
   // DELAYABLE: Called for RTL dlgs only.
   MSOMETHOD_(RECT,GetCtrlParentClientRect)( void* ctrlData ) const = 0;

   // Return the width of the dialog caption (which generally uses a different font than the dialog font)
   // plus any additional necessary titlebar width for the system buttons (Context Help [?] + Close [X]).
   // It is not necessary to include the width of the left and right borders of the titlebar.  Returning
   // a width that is too small may cause a dialog caption to be truncated when the caption is longer than
   // the width otherwise required for the dialog.  Including a bit of safety buffer space (e.g., assuming
   // the Context Help button will always be present even though it may not) and returning a width greater
   // than necessary is OK; it may just cause dialogs with long captions to be that much wider than necessary.
   // DELAYABLE: Only affects dialogs with long captions.
   MSOMETHOD_(int,WidthOfCaptionPlusSysBtns)() const = 0;

   // Set the client size of the dialog window.  This means extra width and height must be added for
   // the (non-client) titlebar and borders.  If host wants to reposition the dialog (e.g., to keep it
   // centered, etc.) based on its new size, this function is a good place to do it.  Alternatively, it
   // may be done after the AutoLayout has completed.
   MSOMETHOD_(void,SetDlgClientSize)( int width, int height ) = 0;

   // DO NOT IMPLEMENT -- WILL BE REMOVED
   MSOMETHOD_(void,SetDlgPosition)() = 0;

   // Return the HWND of the overall dialog.
   MSOMETHOD_(HWND,GetDlgHwnd)() = 0;

   // Returns whether or not the dialog manager starts a new coordinate system
   // with each frame. Default is false
   MSOMETHOD_(bool,FEveryFrameStartsNewCoordSystem)() const { return false; }

   //---------------
   // Controls:
   //---------------

   // For the given ctrlData, return the most closely-matching MsoDALCtrlType value.
   MSOMETHOD_(MsoDALCtrlType,GetCtrlType)( void* ctrlData ) const = 0;

   // Given a control of type MSODAL_BITMAPCONTROL or MSODAL_BITMAPBUTTON, compute the size of the bitmap.
   // Do not add any border -- just provide the bitmap size.
   MSOMETHOD_(void,GetBitmapSize)( void* bitmapCtrlData, int& width, int& height ) const = 0;

   // DO NOT IMPLEMENT -- WILL BE REMOVED
   MSOMETHOD_(void,ComputeMinSizeForBitmapBtn)( void* bitmapBtnData, int& width, int& height ) const = 0;

   // Optionally set the initial width and height to be used by the CtrlFrame of the specified control, and
   // return true.  If the engine's default sizing strategy is to be used, then return false.  Typically,
   // controls of type MSODAL_USERCONTROL will have their default initial sizes overridden by this function;
   // other controls typically will not.  However, if the desired initial size of the MSODAL_USERCONTROL has
   // already been set prior to the creation of its frame, return false and current dimensions of the control
   // will be preserved.  (It is not possible to set just one dimension and not the other in this function.)
   MSOMETHOD_(bool,FGetCustomDefaultCtrlSize)( void* /*ctrlData*/, int& /*width*/, int& /*height*/ ) const
      { return false; }

   // Return the current coordinates of the specified control.
   MSOMETHOD_(RECT,GetCtrlRect)( void* ctrlData ) const = 0;

   // Set the position and size of the specified control to the specified rectangle.
   MSOMETHOD_(void,SetCtrlRect)( void* ctrlData, const RECT& r ) = 0;

   // Set the position and size of a table row
   // No-op if this method is not overridden
   MSOMETHOD_(void,SetTableRowRect)( void* /*tableData*/, int /*iRow*/, Frame* /*pFrame*/, const RECT& /*r*/ ) { return; };

   //---------------
   // List controls:
   //---------------

   // Given a control of type MSODAL_DROPDOWNLIST, MSODAL_LISTBOX, or MSODAL_COMBOBOX, return the greatest width,
   // in pixels, of any item in the control.
   MSOMETHOD_(int,MaxItemWidth)( void* ctrlData ) const = 0;

   // If the closed-state height of the specified MSODAL_DROPDOWNLIST or MSODAL_CUSTOMDROPDOWN control or the
   // height of the edit box part of the specified MSODAL_COMBOBOX control is determined automatically by the
   // dialog manager and cannot be user-set, then return that fixed height, in pixels.  If the closed-state
   // height is variable and can be user-set, then return -1.
   MSOMETHOD_(int,FixedClosedDropDownHeight)( void* ctrlData ) = 0;

   // Return the total height, in pixels, that the specified MSODAL_DROPDOWNLIST (when in open state),
   // MSODAL_LISTBOX, or MSODAL_COMBOBOX control must occupy in order to display the specified number of lines
   // (items) in the list part of the control.  If the control is a MSODAL_DROPDOWNLIST or MSODAL_COMBOBOX,
   // then the returned height should include the edit-box part of the control, whose height is provided in
   // the param closedDropDownHeight.  (For listboxes, the param closedDropDownHeight should be ignored.)  The
   // returned height should also include any top or bottom border of the list part of the control (which, if
   // excluded, might cause the last item to be truncated or dropped).
   MSOMETHOD_(int,TotalOpenListCtrlHeightForNLines)( void* ctrlData, int lines, int closedDropDownHeight ) = 0;

   // Set the positions and sizes of both parts of a 2-part control (currently, just MSODAL_DROPDOWNLIST or
   // MSODAL_COMBOBOX control), or, if the dialog manager doesn't allow the sizes of both parts to be individually
   // set, then just of the total, combined control or whichever of the two parts can have its coordinates set.
   MSOMETHOD_(void,Set2PartCtrlRects)( void* ctrlData, const RECT& firstPartRect, const RECT& totalRect ) = 0;

   // Return true only if the list-edit combo behaves as a dropdown -- i.e., if the listBox is not visible
   // unless the user clicks on a dropdown button on the edit control.  The default implementation returns false.
   MSOMETHOD_(bool,ListEditIsDropDown)( void* /*editStrData*/, void* /*listBoxData*/ ) const { return false; }

   //---------------
   // Tab controls:
   //---------------

   // Return POINT specifying the width of the widest tab (commonly the tab with the longest text title)
   // as well as its height, which we presume is the height of all the other tabs too.
   MSOMETHOD_(POINT,GetWidestTabItemSize)( void* tabCtrlData ) const = 0;

   // Return the number of tab rows needed to fit all the tabs, given the current width of the tab-control.
   MSOMETHOD_(int,GetTabRowCount)( void* tabCtrlData ) const = 0;

   // Are the tabsheets implemented as separate windows?
   MSOMETHOD_(bool,FTabSheetsAreSeparateWindows)( void* /*tabCtrlData*/ ) const { return true; }

   // If the tabsheets are implemented as separate windows, set the position (relative to the overall, 
   // parent dialog) and size of each of them to the specified rectangle.  Do nothing if there are no
   // separate tabsheet windows.
   MSOMETHOD_(void,SetRectAllSheetsOfTabctrl)( void* tabCtrlData, const RECT& sheetRect ) = 0;

   //---------------
   // Control text:
   //---------------

   // Retrieve the text of the specified control into the provided buffer wzText.  nMaxCount specifies the
   // maximum number of characters to copy to the buffer, including the null terminator.  If the text exceeds
   // this limit, it should be truncated to avoid buffer overflow.
   MSOMETHOD_(void,GetCtrlText)( void* ctrlData, _Out_cap_(nMaxCount) WCHAR* wzText, int nMaxCount ) const = 0;

   // Return the width and height of the text of the specified control.
   MSOMETHOD_(void,MeasureCtrlText)( void* ctrlData, int& width, int& height ) const = 0;

   // Measure the text of the given control based on wrapping width specified by the parameter 'width', and
   // set the resulting height and adjusted width.  (I.e., if measuring with DrawText(), use the DT_WORDBREAK
   // flag.)  Since the actual width of the wrapping text may be less or, in the case of text with one
   // really long word with no spaces at which to wrap, greater than the target width passed in, the width
   // must be set as well as the height.
   //
   // The param fWidenForLongWord specifies how to behave if the text contains one or more long words that
   // have no spaces at which to wrap and that cannot fit within the target width.  If fWidenForLongWord is
   // false, the text should be rigidly confined within the target width, and, depending on how the control
   // displays a long word that's wider than the control's width, should be measured either as though the long
   // word wraps to multiple lines (as in read-only edit-strings and SDM static strings) or as though the long
   // word is truncated to the control's width (as in Win32 static-strings).  If it is true, the returned
   // width may be wider than the target width, as wide as necessary to fit the longest unbroken word.
   MSOMETHOD_(void,MeasureCtrlTextWithTargetWidth)( void* ctrlData, int& width, int& height,
                                                    bool fWidenForLongWord ) const = 0;

   // Return the width and height of the text, using font info from the specified control.
   MSOMETHOD_(void,MeasureText)( const WCHAR* wzText, void* ctrlData, int& width, int& height ) const = 0;

   // Set the text of the given control to an abbreviated form of its current text, perhaps employing an
   // ellipsis, that fits within the specified maxWidth.  Line breaks should not be inserted -- the text
   // should remain a single line.
   MSOMETHOD_(void,ShrinkCtrlTextToWidth)( void* ctrlData, int maxWidth ) = 0;
   MSOMETHOD_(bool,FNeedShrinkCtrlTextToWidth)( void* /*ctrlData*/ ) const { return true; }

   //----------------
   // Miscellaneous:
   //----------------

   // The default implementation returns true.  Only return false if the dialog uses an old-style
   // AutoLayout that relies on composite frames not expanding by default.
   MSOMETHOD_(bool,CompositeFramesExpandByDefault)() const { return true; }

   // return custom work pane right padding
   MSOMETHOD_(bool,GetWorkPaneRightPadding)( void* /*ctrlData*/, int& /*padding*/ ) { return false; }

   //------------------------
   // Layout constraints caching:
   //------------------------

   // Called by DAL to allow clients to cache results of the ::CalcConstraints() call on a given frame
   MSOMETHOD_(void,SetCalcConstraintSize)(DalLayoutData , void* /*ctrlData*/, SIZE /*size*/) const { return; }
   // Called by DAL on composite frames with no children to check whether client has constraints cached from 
   // previous Layout pass(s). Clients should return true if constraints are cached.
   MSOMETHOD_(bool,GetCalcConstraintSize)(DalLayoutData , void* /*ctrlData*/, SIZE* /*psize*/, DalLayoutFlag*) const { return false; }
   // Called by DAL on composite frames to allow clients to cache frame visible size.
   MSOMETHOD_(void,SetCalcPositionVisibleSize)(void* /*ctrlData*/, SIZE /*psize*/) { return; }

   //----------------
   // Allocator
   //----------------

   // This allocator is used by DAL to allocate and free frame implementations. MSO memory allocator
   // is used by default implementation.
#if( !NO_MSO_DEPENDENCIES )
#if DEBUG
   virtual MSOPUB MSOMETHODIMP_(void*) PvAllocFrameMemory(size_t size, int dg, const CHAR* szFile, int li);
#else
   virtual MSOPUB MSOMETHODIMP_(void*) PvAllocFrameMemory(size_t size, int dg);
#endif
   virtual MSOPUB MSOMETHODIMP_(void) FreeFrameMemory(void *pv) const;
#endif


   // Returns whether or not the dialog manager wants DAL to automatically add
   // extra border margins on edit strings.  Default is true.
   MSOMETHOD_(bool,FAddExtraPixForEditStrBorders)() const { return true; }

   //---------------
   // Debug-only:
   //---------------

#if( MSODAL_DEBUG_TOOLS )
   //--------------------------
   // Frame outlines and Logo:
   //--------------------------

   // Return zero-based index of the currently selected tab.  (Used for drawing frame outlines.)
   MSOMETHOD_(int,GetCurTabIndex)( void* tabCtrlData ) const = 0;

   // Return whether this dialog supports the drawing of frame outlines.  The default implementation
   // returns true; override if there are some specific dialogs for which you never want outlines drawn.
   MSOMETHOD_(bool,SupportsFrameOutlines)() const { return true; }

   // Return the name of the dialog manager, to be used in the logo painted at the top-left of each
   // dialog.  It must be no longer than cchDlgMgr characters (including the null terminator).
   MSOMETHOD_(void,GetShortDlgMgrName)( WCHAR* wzDlgMgr, int cchDlgMgr ) const { if ( cchDlgMgr > 0 ) wzDlgMgr[0] = 0; }

   //-------------------------
   // String length changing:
   //-------------------------

   // Set the control's text to the text provided.  (Used by the string length changing test tool.)
   MSOMETHOD_(void,SetCtrlText)( void* ctrlData, const WCHAR* wzText ) = 0;

   // Retrieve the dialog's caption into the provided buffer wzText.  nMaxCount specifies the maximum
   // number of characters to copy to the buffer, including the null terminator.  If the text exceeds this
   // limit, it should be truncated to avoid buffer overflow.  (Used by the string length changing test tool.)
   MSOMETHOD_(void,GetDlgCaption)(_Out_cap_(nMaxCount)  WCHAR* wzText, int nMaxCount ) const = 0;

   // Set the dialog's caption to the string provided.  (Used by the string length changing test tool.)
   MSOMETHOD_(void,SetDlgCaption)( const WCHAR* wzText ) = 0;
   
   // Return whether this dialog supports the automated changing of its strings by the string
   // length changing test tool.  The default implementation returns true; override if there
   // are some specific dialogs for which you never want the string lengths changed, even when
   // the tool is active.
   MSOMETHOD_(bool,SupportsStrLenChanging)() const { return true; }

   // Return a numeric ID for the control.  Used by the string length changing tool, when
   // choosing a target length for the control's text.
   MSOMETHOD_(INT_PTR,GetCtrlID)( void* ctrlData ) const = 0;

   // Return a non-negative ID for the dialog.  Used to index into an array in which the string length
   // changing tool stores the most recent test suite scheme used for each individual dialog.  (It's
   // useful for the tool to track each dialog's status in the test suite individually, since some
   // dialogs can only be invoked alternately with a companion dialog, and we don't want to miss
   // every second test suite scheme for the dialog being tested.)
   MSOMETHOD_(int,GetDlgID)() const = 0;

   // Provide non-negative (dialog) IDs for all tabsheets in the tabcontrol (up to cMaxSheets), in the
   // order of their appearance in the tabcontrol.  Return the number of tabsheet IDs written.
   MSOMETHOD_(int,GetTabSheetIDs)( _Out_cap_(cMaxSheets) int ids[], int cMaxSheets ) const { UNREFERENCED_PARAMETER(ids); UNREFERENCED_PARAMETER(cMaxSheets); return 0; }

   // Returns whether or not the dialog manager allows methods to be called
   // twice.  Used for dialog managers who call commands, such as SetWidth,
   // multiple times on one control.  Default is false
   MSOMETHOD_(bool,FAllowsDupMethodCalls)() const { return false; }

#endif // MSODAL_DEBUG_TOOLS

private:
   AutoLayoutDialogData* m_pData;

   friend class FrameImp;
}; // IMsoAutoLayoutDialog

} // namespace DlgAutoLayout
#endif // DAL_CPLUSPLUS

#undef THISPTR_
#undef CLASSNAME


//------------------------------------------------------------------------------------
// The following macros may be useful for defining dialog-manager-specific wrapper
// classes around Dialog AutoLayout frame classes.  They declare functions to be
// inline-forwarded from the wrapper object to the contained frame object.
//------------------------------------------------------------------------------------

#define FRAME_FORWARDS(FORWARDEE, DEREF) \
   void ExpandToFillHoriz() { FORWARDEE DEREF ExpandToFillHoriz(); } \
   void ExpandToFillVert() { FORWARDEE DEREF ExpandToFillVert(); } \
   void DontExpandHoriz() { FORWARDEE DEREF DontExpandHoriz(); } \
   void DontExpandVert() { FORWARDEE DEREF DontExpandVert(); } \
   void SetWidth( int width) { FORWARDEE DEREF SetWidth( width ); } \
   void SetWidthDU( int width) { FORWARDEE DEREF SetWidthDU( width ); } \
   void SetWidthChars( int numChars) { FORWARDEE DEREF SetWidthChars( numChars ); } \
   void SetHeight( int height) { FORWARDEE DEREF SetHeight( height ); } \
   void SetHeightDU( int height) { FORWARDEE DEREF SetHeightDU( height ); } \
   void Indent( int numIndents DEFAULTS_TO(1) ) { FORWARDEE DEREF Indent( numIndents ); } \
   void IndentBy( int indentWidth ) { FORWARDEE DEREF IndentBy( indentWidth ); } \
   void IndentByDU( int indentWidth ) { FORWARDEE DEREF IndentByDU( indentWidth ); } \
   void LowerBy( int height ) { FORWARDEE DEREF LowerBy( height ); } \
   void LowerByDU( int height ) { FORWARDEE DEREF LowerByDU( height ); } \
   void SetMargin( int margin) { FORWARDEE DEREF SetMargin( margin ); } \
   void SetMarginDU( int margin) { FORWARDEE DEREF SetMarginDU( margin ); } \
   void SetMarginChars( int numChars) { FORWARDEE DEREF SetMarginChars( numChars ); } \
   void ExpandToFitStr( const WCHAR* wzStr) { FORWARDEE DEREF ExpandToFitStr( wzStr ); } \
   void AlignWithGroupBox() { FORWARDEE DEREF AlignWithGroupBox(); } \
   void SetSetWidthFlags( int flags) { FORWARDEE DEREF SetSetWidthFlags( flags ); } \
   int DUToPixH( int x) METHOD_CONST { return FORWARDEE DEREF DUToPixH( x ); } \
   int DUToPixV( int y) METHOD_CONST { return FORWARDEE DEREF DUToPixV( y ); } \
   int Width() METHOD_CONST { return FORWARDEE DEREF Width(); } \
   int Height() METHOD_CONST { return FORWARDEE DEREF Height(); } \
   Frame* GetFrame(MsoDALFrameData msodalfd) METHOD_CONST { return FORWARDEE DEREF GetFrame(msodalfd); }


#define COMPOSITEFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   /*CompositeFrame& operator<<( Frame& frame ) { return FORWARDEE DEREF operator<<( frame ); } Wrong frame classes -- custom forwarder needed*/ \
   BOOL Layout( MsoDALLayoutParams* pParams DEFAULTS_TO(NULL) ) { return FORWARDEE DEREF Layout( pParams ); } \
   void AlignChildren( AlignDir dir) { FORWARDEE DEREF AlignChildren( dir ); } \
   void ExpandChildrenToFillHoriz( ExpandChildrenAlgorithm expand DEFAULTS_TO(TOWARD_EQUAL_SIZES) ) { FORWARDEE DEREF ExpandChildrenToFillHoriz( expand ); } \
   void ExpandChildrenToFillVert( ExpandChildrenAlgorithm expand DEFAULTS_TO(TOWARD_EQUAL_SIZES) ) { FORWARDEE DEREF ExpandChildrenToFillVert( expand ); } \
   void SameWidthChildren() { FORWARDEE DEREF SameWidthChildren(); } \
   void SameHeightChildren() { FORWARDEE DEREF SameHeightChildren(); } \
   void SetWidthsChildren( int width) { FORWARDEE DEREF SetWidthsChildren( width ); } \
   void SetWidthFlagsChildren( int flags) { FORWARDEE DEREF SetWidthFlagsChildren( flags ); } \
   void RemoveDlgBorderMargins() { FORWARDEE DEREF RemoveDlgBorderMargins(); } \
   void SetDlgClientWidth( int width) { FORWARDEE DEREF SetDlgClientWidth( width ); } \
   void SetDlgClientHeight( int height) { FORWARDEE DEREF SetDlgClientHeight( height ); }\
   void SetDialogUsesMultiFonts( bool fUsesMultiFonts) { FORWARDEE DEREF SetDialogUsesMultiFonts( fUsesMultiFonts ); } \
   void SetPerimeterMargin( int margin) { FORWARDEE DEREF SetPerimeterMargin( margin ); } \
   void SetPerimeterMarginDU( int margin) { FORWARDEE DEREF SetPerimeterMarginDU( margin ); } \
   void SetLeftMargin( int margin) { FORWARDEE DEREF SetLeftMargin( margin ); } \
   void SetRightMargin( int margin) { FORWARDEE DEREF SetRightMargin( margin ); } \
   void SetTopMargin( int margin) { FORWARDEE DEREF SetTopMargin( margin ); } \
   void SetBottomMargin( int margin) { FORWARDEE DEREF SetBottomMargin( margin ); } \
   void SetLeftMarginDU( int margin) { FORWARDEE DEREF SetLeftMarginDU( margin ); } \
   void SetRightMarginDU( int margin) { FORWARDEE DEREF SetRightMarginDU( margin ); } \
   void SetTopMarginDU( int margin) { FORWARDEE DEREF SetTopMarginDU( margin ); } \
   void SetBottomMarginDU( int margin) { FORWARDEE DEREF SetBottomMarginDU( margin ); }

#define COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            FRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define HORIZFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void DistributeChildren() { FORWARDEE DEREF DistributeChildren(); } \
   void SetAllMargins( int margin ) { FORWARDEE DEREF SetAllMargins( margin ); } \
   void SetAllMarginsDU( int margin ) { FORWARDEE DEREF SetAllMarginsDU( margin ); }

#define HORIZFRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            HORIZFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define VERTFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void DistributeChildren() { FORWARDEE DEREF DistributeChildren(); } \
   void SetAllMargins( int margin ) { FORWARDEE DEREF SetAllMargins( margin ); } \
   void SetAllMarginsDU( int margin ) { FORWARDEE DEREF SetAllMarginsDU( margin ); }

#define VERTFRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            VERTFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define FLOWFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetVertMargin( int pixelMargin ) { FORWARDEE DEREF SetVertMargin( pixelMargin ); } \
   void SetVertMarginDU( int marginDU ) { FORWARDEE DEREF SetVertMarginDU( marginDU ); } \
   void ShrinkWidthBy( int delta ) { FORWARDEE DEREF ShrinkWidthBy( delta ); } \
   void ShrinkWidthByDU( int deltaDU ) { FORWARDEE DEREF ShrinkWidthByDU( deltaDU ); } \
   void IndentLatterRows() { FORWARDEE DEREF IndentLatterRows(); } \
   void IndentLatterRowsBy( int width ) { FORWARDEE DEREF IndentLatterRowsBy( width ); } \
   void IndentLatterRowsByDU( int widthDU ) { FORWARDEE DEREF IndentLatterRowsByDU( widthDU ); } \
   void IndentLatterRowsByFirstChild() { FORWARDEE DEREF IndentLatterRowsByFirstChild(); } \
   void DistributeRows() { FORWARDEE DEREF DistributeRows(); } \
   void DistributeChildrenHoriz() { FORWARDEE DEREF DistributeChildrenHoriz(); } \
   void SetAllMargins( int margin ) { FORWARDEE DEREF SetAllMargins( margin ); } \
   void SetAllMarginsDU( int margin ) { FORWARDEE DEREF SetAllMarginsDU( margin ); }

#define FLOWFRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            FLOWFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define VERTFLOWFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetHorizMargin( int pixelMargin ) { FORWARDEE DEREF SetHorizMargin( pixelMargin ); } \
   void SetHorizMarginDU( int marginDU ) { FORWARDEE DEREF SetHorizMarginDU( marginDU ); } \
   void ShrinkHeightBy( int delta ) { FORWARDEE DEREF ShrinkHeightBy( delta ); } \
   void ShrinkHeightByDU( int deltaDU ) { FORWARDEE DEREF ShrinkHeightByDU( deltaDU ); } \
   void DistributeCols() { FORWARDEE DEREF DistributeCols(); } \
   void DistributeChildrenVert() { FORWARDEE DEREF DistributeChildrenVert(); } \
   void SameWidthColumns() { FORWARDEE DEREF SameWidthColumns(); } \
   void SetAllMargins( int margin ) { FORWARDEE DEREF SetAllMargins( margin ); } \
   void SetAllMarginsDU( int margin ) { FORWARDEE DEREF SetAllMarginsDU( margin ); }

#define VERTFLOWFRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            VERTFLOWFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define TABLEFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetColMargin( int iCol, int margin ) { FORWARDEE DEREF SetColMargin( iCol, margin ); } \
   void SetColMarginDU( int iCol, int margin ) { FORWARDEE DEREF SetColMarginDU( iCol, margin ); } \
   void SetRowMargin( int iRow, int margin ) { FORWARDEE DEREF SetRowMargin( iRow, margin ); } \
   void SetRowMarginDU( int iRow, int margin ) { FORWARDEE DEREF SetRowMarginDU( iRow, margin ); } \
   void SetAllColMargins( int margin ) { FORWARDEE DEREF SetAllColMargins( margin ); } \
   void SetAllColMarginsDU( int margin ) { FORWARDEE DEREF SetAllColMarginsDU( margin ); } \
   void SetAllRowMargins( int margin ) { FORWARDEE DEREF SetAllRowMargins( margin ); } \
   void SetAllRowMarginsDU( int margin ) { FORWARDEE DEREF SetAllRowMarginsDU( margin ); } \
   void SetWidthColChildren( int iCol, int width, bool fIncludeMultiColChildren DEFAULTS_TO(true) ) { FORWARDEE DEREF SetWidthColChildren( iCol, width, fIncludeMultiColChildren ); } \
   void SetWidthColChildrenDU( int iCol, int width, bool fIncludeMultiColChildren DEFAULTS_TO(true) ) { FORWARDEE DEREF SetWidthColChildrenDU( iCol, width, fIncludeMultiColChildren ); } \
   void SetHeightRowChildren( int iRow, int width, bool fIncludeMultiRowChildren DEFAULTS_TO(true) ) { FORWARDEE DEREF SetHeightRowChildren( iRow, width, fIncludeMultiRowChildren ); } \
   void SetHeightRowChildrenDU( int iRow, int width, bool fIncludeMultiRowChildren DEFAULTS_TO(true) ) { FORWARDEE DEREF SetHeightRowChildrenDU( iRow, width, fIncludeMultiRowChildren ); } \
   void DontExpandColHoriz( int iCol, bool fIncludeMultiColChildren DEFAULTS_TO(true) ) { FORWARDEE DEREF DontExpandColHoriz( iCol, fIncludeMultiColChildren ); } \
   void DontExpandRowVert( int iRow, bool fIncludeMultiRowChildren DEFAULTS_TO(true) ) { FORWARDEE DEREF DontExpandRowVert( iRow, fIncludeMultiRowChildren ); } \
   void ExpandColToFillHoriz( int iCol ) { FORWARDEE DEREF ExpandColToFillHoriz( iCol ); } \
   void ExpandRowToFillVert( int iRow ) { FORWARDEE DEREF ExpandRowToFillVert( iRow ); } \
   void DistributeCols() { FORWARDEE DEREF DistributeCols(); } \
   void DistributeRows() { FORWARDEE DEREF DistributeRows(); } \
   void AddGroupBox( void* groupCtrlData, int left, int top, int right, int bottom, BOOL fNeverHasText ) \
      { FORWARDEE DEREF AddGroupBox( groupCtrlData, left, top, right, bottom, fNeverHasText ); } \
   void AddGroupBox2Ctrls( void* textCtrlData, void* lineCtrlData, \
                     int left, int top, int right, int bottom, BOOL fNeverHasText ) \
      { FORWARDEE DEREF AddGroupBox2Ctrls( textCtrlData, lineCtrlData, left, top, right, bottom, fNeverHasText ); }

#define TABLEFRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            TABLEFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define CONTAINERFRAME_FORWARDS_NEW(FORWARDEE, DEREF)

#define CONTAINERFRAME_FORWARDS(FORWARDEE, DEREF) \
            COMPOSITEFRAME_FORWARDS(FORWARDEE, DEREF) \
            CONTAINERFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define CTRLCONTAINERFRAME_FORWARDS_NEW(FORWARDEE, DEREF)

#define CTRLCONTAINERFRAME_FORWARDS(FORWARDEE, DEREF) \
            CONTAINERFRAME_FORWARDS(FORWARDEE, DEREF) \
            CTRLCONTAINERFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define GROUPBOXFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetNeverHasText() { FORWARDEE DEREF SetNeverHasText(); }

#define GROUPBOXFRAME_FORWARDS(FORWARDEE, DEREF) \
            CONTAINERFRAME_FORWARDS(FORWARDEE, DEREF) \
            GROUPBOXFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define OVERLAPFRAME_FORWARDS_NEW(FORWARDEE, DEREF)

#define OVERLAPFRAME_FORWARDS(FORWARDEE, DEREF) \
            CONTAINERFRAME_FORWARDS(FORWARDEE, DEREF) \
            OVERLAPFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define TABCTRLFRAME_FORWARDS_NEW(FORWARDEE, DEREF)

#define TABCTRLFRAME_FORWARDS(FORWARDEE, DEREF) \
            CONTAINERFRAME_FORWARDS(FORWARDEE, DEREF) \
            TABCTRLFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define LEAFFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void AddLeftString( void* strCtrlData) { FORWARDEE DEREF AddLeftString( strCtrlData ); } \
   void AddTopString( void* strCtrlData) { FORWARDEE DEREF AddTopString( strCtrlData ); } \
   void AddRightString( void* strCtrlData) { FORWARDEE DEREF AddRightString( strCtrlData ); } \
   void SetHeightLines( int numLines) { FORWARDEE DEREF SetHeightLines( numLines ); } \
   void SetMaxPcntToExceedStrTargetWidth( int percent) { FORWARDEE DEREF SetMaxPcntToExceedStrTargetWidth( percent ); } \
   void IgnoreSetWidthIfStrHasNewline() { FORWARDEE DEREF IgnoreSetWidthIfStrHasNewline(); } \
   void DisableSetWidthFlexibility() { FORWARDEE DEREF DisableSetWidthFlexibility(); } \
   void IncreaseWidthBy( int increase ) { FORWARDEE DEREF IncreaseWidthBy( increase ); } \
   void IncreaseHeightBy( int increase ) { FORWARDEE DEREF IncreaseHeightBy( increase ); } \
   void IncreaseWidthByDU( int increase ) { FORWARDEE DEREF IncreaseWidthByDU( increase ); } \
   void IncreaseHeightByDU( int increase ) { FORWARDEE DEREF IncreaseHeightByDU( increase ); }

#define LEAFFRAME_FORWARDS(FORWARDEE, DEREF) \
            FRAME_FORWARDS(FORWARDEE, DEREF) \
            LEAFFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define CTRLFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetWidthFromCtrl() { FORWARDEE DEREF SetWidthFromCtrl(); } \
   void SetHeightFromCtrl() { FORWARDEE DEREF SetHeightFromCtrl(); } \
   void SetSizeFromCtrl() { FORWARDEE DEREF SetSizeFromCtrl(); } \
   void SetWidthToFitPreexistingItems( int minWidthChars DEFAULTS_TO(5), \
                                       int maxWidthChars DEFAULTS_TO(80), \
                                       int extraWidthChars DEFAULTS_TO(0) ) \
         { FORWARDEE DEREF SetWidthToFitPreexistingItems( minWidthChars, maxWidthChars, extraWidthChars ); } \
   void SetLeftAdjust( int adjust ) { FORWARDEE DEREF SetLeftAdjust( adjust ); } \
   void SetRightAdjust( int adjust ) { FORWARDEE DEREF SetRightAdjust( adjust ); } \
   void SetTopAdjust( int adjust ) { FORWARDEE DEREF SetTopAdjust( adjust ); } \
   void SetBottomAdjust( int adjust ) { FORWARDEE DEREF SetBottomAdjust( adjust ); } \
   void SetLeftAdjustDU( int adjust ) { FORWARDEE DEREF SetLeftAdjustDU( adjust ); } \
   void SetRightAdjustDU( int adjust ) { FORWARDEE DEREF SetRightAdjustDU( adjust ); } \
   void SetTopAdjustDU( int adjust ) { FORWARDEE DEREF SetTopAdjustDU( adjust ); } \
   void SetBottomAdjustDU( int adjust ) { FORWARDEE DEREF SetBottomAdjustDU( adjust ); } \
   void ShrinkWrapToText( bool fShrinkWrap ) { FORWARDEE DEREF ShrinkWrapToText( fShrinkWrap ); }

#define CTRLFRAME_FORWARDS(FORWARDEE, DEREF) \
            LEAFFRAME_FORWARDS(FORWARDEE, DEREF) \
            CTRLFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define NUMINPUTFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetWidthDigits( int numDigits ) { FORWARDEE DEREF SetWidthDigits( numDigits ); }

#define NUMINPUTFRAME_FORWARDS(FORWARDEE, DEREF) \
            LEAFFRAME_FORWARDS(FORWARDEE, DEREF) \
            NUMINPUTFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define LISTEDITFRAME_FORWARDS_NEW(FORWARDEE, DEREF) \
   void SetWidthToFitPreexistingItems( int minWidthChars DEFAULTS_TO(5), \
                                       int maxWidthChars DEFAULTS_TO(80), \
                                       int extraWidthChars DEFAULTS_TO(0) ) \
         { FORWARDEE DEREF SetWidthToFitPreexistingItems( minWidthChars, maxWidthChars, extraWidthChars ); }

#define LISTEDITFRAME_FORWARDS(FORWARDEE, DEREF) \
            LEAFFRAME_FORWARDS(FORWARDEE, DEREF) \
            LISTEDITFRAME_FORWARDS_NEW(FORWARDEE, DEREF)


#define FRAMEEQUALIZER_FORWARDS(FORWARDEE, DEREF) \
   void EqualizeNow() { FORWARDEE DEREF EqualizeNow(); }
