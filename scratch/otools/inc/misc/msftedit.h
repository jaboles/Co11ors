//
// MSFTEDIT.H
// 
// Purpose:
// 	RICHEDIT v5.0 public definitions
// 
// Copyright (c) 2001, Microsoft Corporation
//

#ifndef _MSFTEDIT_
#define _MSFTEDIT_

// Include definitions from previous versions
#include <richedit.h>

// Special RichEdit20W plan text class
#define RICHEDIT_CLASSWPT		L"RichEdit20WPT"

// These are C++ only interfaces
#ifdef __cplusplus

// Declaration for events associated with blobs and text trackers.
class RichEditEvent
{
public:
	enum EVENTID
	{
		EVENTID_FORMATCHANGED,
		EVENTID_BLOB_MOVED,
		EVENTID_BLOB_DELETED,
		EVENTID_SLBLOB_RESIZE,
		EVENTID_TRACKER_CPCHANGE,
		EVENTID_TRACKER_PRECPCHANGE
	};

	struct BasicEvent {
		EVENTID id;
	};

	// Bit value in dwFormatFlags
	enum FORMATFLAGS 
	{
		LAYOUT_CHANGE=1
	};

	struct FormatChangeEvent : public BasicEvent
	{
		DWORD		dwCFMask;
		DWORD		dwPFMask;
		DWORD		dwFormatFlags;
		LONG		cpFormatChangeMin;
		LONG		cpFormatChangeMost;
	};

	struct TrackerCPChangeEvent : public BasicEvent
	{
		LONG		cpChangeMin;
		LONG		cpChangeMost;
		LONG		cchDel;
	};

	struct TrackerPreCPChangeEvent : public BasicEvent
	{
		LONG		cpMin;
		LONG		cpMost;
		UINT		msg;
		WPARAM		wparam;
		LPARAM		lparam;
	};

	struct SLBlobResizeEvent : public BasicEvent
	{
		LONG  xWidthNew;
	};

	BasicEvent *GetBasicEvent()
	{
		return _pBasicEvent;
	};

	void SetBasicEvent(BasicEvent *pBasicEvent)
	{
		_pBasicEvent = pBasicEvent;
	};

private:
	BasicEvent	*_pBasicEvent;

};

// IID_IPesrsistTextBlobBits : {ED2D3F04-3735-429a-BA86-3AC3138CE470}
//	IPersistTextBlobBits is an alternate very low overhead interface for serializing blobs
//	Rich Edit will query blobs for this custom interface and use it instead of ITextBlob::Load and ITextBlob::Save
DEFINE_GUID(IID_IPersistTextBlobBits, 0xed2d3f04, 0x3735, 0x429a, 0xba, 0x86, 0x3a, 0xc3, 0x13, 0x8c, 0xe4, 0x70);
interface IPersistTextBlobBits : public IUnknown
{
	STDMETHOD(Load)(
		void	*pv,	// In : Pointer to buffer from which bits should be loaded
		int		cb		// In : Size of Buffer
	) PURE;

	STDMETHOD(Save)(
		void	*pv,	// In : Pointer to buffer where bits should be saved
		int		&cb		// In/Out : Size of buffer, set to required size if pv is null
	) PURE;
};

// IID_ITextBlob : 53ff24d9-6414-418f-8e8d-ddd1ba692db5
//    Blobs are light weight client drawn objects hosted by rich edit controls.
//    Parts of a blob may overlap with text in the rich edit control.
//    The ITextBlob interface provides the functionality necessary for
//    a rich edit control to be able to host these objects.
//    It provides the methods necessary for rendering, reflow, and persistence of the data.
//    Clients of rich edit must provide implementations of the interface to be used by rich edit.
DEFINE_GUID(IID_ITextBlob, 0x53ff24d9, 0x6414, 0x418f, 0x8e, 0x8d, 0xdd, 0xd1, 0xba, 0x69, 0x2d, 0xb5);
interface ITextBlob : public IUnknown
{

	// BLOB_FLAGS define bits for BLOB_PROPERTIES::dwFlags.
	enum BLOB_FLAGS
	{
		BLOB_DRAWS_BACKGROUND	  = 1,	// Blob draws background (as in normal, as in selected case)
										// NOTE: normal case never occurs for transparent RECs
		BLOB_DRAWS_EFFECTS		  = 2,	// Blob draws effects (underline, strikeout)
		BLOB_PRESERVES_LINE_HEIGHT= 4,	// When there is no text in the line (it has blobs only),
										// blobs having this flag set preserve line's ascent and
										// descent proper to the given text formatting. When none
										// of blobs of the line have this flag set, its ascent and
										// descent will be minimal for blobs to fit in.
	};

	// BLOB_DRAW_FLAGS define bits for BLOB_DRAW_INFO::dwFlags.
	enum BLOB_DRAW_FLAGS
	{
		BLOB_DRAW_BACKGROUND	=  1,	// On next Draw call, render background
		BLOB_DRAW_UNDERLINE		=  2,	// On next Draw call, render underline
		BLOB_DRAW_STRIKEOUT		=  4,	// On next Draw call, render strikeout
		BLOB_DRAW_NO_FOREGROUND	=  8,	// On next Draw call, do not render foreground
										//  (affects the next Draw call only)
		BLOB_DRAW_SELECTED		= 16,	// Duplicates fSelected argument of Draw
										//  - to stress that given colors are
										//  selection foreground/background colors
 	};

	// Special value for BLOB_PROPERTIES fields yUnderlinePos, yStrikeoutPos
	enum
	{
		EFFECT_OFFSET_UNDEFINED = 0x80000000
	};

	// All coordinates in BLOB_PROPERTIES are in HIMETRIC.

	struct BLOB_PROPERTIES
	{
		UINT  cbSize;			// Length of the structure in bytes
		LONG  xWidth;			// Overall width of the blob (including overlap)
		LONG  yHeight;			// Overall height of the blob (including overlap)
		RECT  rcNoOverlap;		// The part of a blob that can't overlap text

		DWORD dwFlags;			// Control flags (see BLOB_FLAGS)

		// The following properties have no meaning for multiline blobs
		LONG  dyDescent;		// Amount of no-overlap rect height used for line descent
								// - alwyas when TA_BASELINE alignment; for underlining
								// when other alignment (TA_TOP or TA_BOTTOM)
		LONG  dyUnderlineOffset;// Desirable underline offset below the baseline
								// (EFFECT_POSITION_UNDEFINED means: use position
								// defined by the context)
		LONG  dyStrikeoutOffset;// Strikeout offset above the baseline
								// (EFFECT_POSITION_UNDEFINED means: use position
								// defined by the context)
		SHORT dyUnderlineWidth;	// Desirable underline thickness
								// (zero means: use thickness defined by the context)
		SHORT dyStrikeoutWidth;	// Desirable strikeout thickness
								// (zero means: use thickness defined by the context)
	};

	// All coordinates in BLOB_DRAW_INFO are in device units (pixels).

	struct BLOB_DRAW_INFO
	{
		UINT		cbSize;			// Length of the structure in bytes
		DWORD		dwFlags;		// BLOB_DRAWS_BACKGROUND, BLOB_DRAWS_EFFECTS
									// and BLOB_DRAWS_NO_FOREGROUND could be set
		RECT		rcNoOverlap;	// Blob no-overlap (layout) rect
		RECT		rcBackground;	// Blob background rect (line-height)
		COLORREF	crBackground;	// Background color
		COLORREF	crForeground;	// Foreground (text) color
		RECT		rcUnderline;	// Underline rectangle (when the blob draws 
									//  effects on its own)
		RECT		rcStrikeout;	// Strikeout rectangle (when the blob draws 
									//  effects on its own)
	};

	// Gets information needed to handle layout and drawing
	// A blob has two size rectangles of interest
	// The first is the overall size rectangle.
	// The second defines the overlap allowed with text around the blob
	// We also want the 
	STDMETHOD(GetBlobProperties)(
		BLOB_PROPERTIES *pBlobProperties
	) PURE;

	// Draw the blob in the given rectangle
	// The rectangle accomodates the full extent of the blob
	// It is up to the client to apply any necessary clipping.
	STDMETHOD(Draw)(
		HDC hdc,
		BOOL fSelected,
		LPCRECT lprcPos
	) PURE;

	// Returns the number of alternate text candidates for a blob
	// Used to retrieve alternate text for a blob for operations such as searches
	// If the requested input index is out of range an error code should be returned
	STDMETHOD_(int, GetTextItem)(
		int index,
		BSTR *pbstrText
	) PURE;

	// A point inside a blob's rectangle may not be a hit.
	// Pt will be normalized to blob origin
	STDMETHOD(QueryHitTest)(
		POINT pt,
		BOOL &fHit
	) PURE;

	STDMETHOD(GetBlobClassId)(
		CLSID *pClsid
	) PURE;

	STDMETHOD(Load)(
		IStorage *pstg
	) PURE;

	STDMETHOD(Save)(
		IStorage *pstg
	) PURE;

	// If a polygon is provided, rich edit will tightwrap text around the region
	// This method will only be called for "figure" blobs.
	STDMETHOD(GetPolygon) (
		SHORT nVertices,	// Number of vertices given by InserMultiLineBlob
		LPPOINT lpPoints	// Vector of point for blob to fill out
	) PURE;

	// Notify blob about interesting events
	STDMETHOD(OnEvent)(
		RichEditEvent *pevent
	) PURE;

	STDMETHOD(SetDrawInfo) (
		BLOB_DRAW_INFO const& DrawInfo
	) PURE;
};

// IID_ITextTracker : 78e3ccbd-7d6f-4a8d-a662-88ab429d8c43
DEFINE_GUID(IID_ITextTracker, 0x78e3ccbd, 0x7d6f, 0x4a8d, 0xa6, 0x62, 0x88, 0xab, 0x42, 0x9d, 0x8c, 0x43);
interface ITextTracker : public IUnknown
{
	enum TEXTTRACKER_GRAVITY
	{
		TEXTTRACKER_GRAVITY_BACKWARD = 1,
		TEXTTRACKER_GRAVITY_FORWARD = 2,
	};

	struct TRACKER_PROPERTIES
	{
		UINT cbSize;
		BYTE bCpMinGravity;			// Value from TEXT_TRACKER_CPGRAVITY
		BYTE bCpMostGravity;		// Value from TEXT_TRACKER_CPGRAVITY
		BYTE fCling;
		BYTE fPreChangeNotify;		// Tracker wants notification before changes
	};

	STDMETHOD(GetProperties)(
		TRACKER_PROPERTIES *pTrackerProperties
	) PURE;

	// Notify tracker about interesting events
	STDMETHOD(OnEvent)(
		RichEditEvent *pevent
	) PURE;
};

typedef int HBLOB;

typedef int HTRACKER;

typedef int HBOUNDARY;

// IID_IEnumTextBlobs : 5fed305e-baaf-42c6-9ed9-f25f5f53a6ec
DEFINE_GUID(IID_IEnumTextBlobIds, 0x5fed305e, 0xbaaf, 0x42c6, 0x9e, 0xd9, 0xf2, 0x5f, 0x5f, 0x53, 0xa6, 0xec);
interface IEnumTextBlobs : public IUnknown
{
	STDMETHOD(Clone)(
		IEnumTextBlobs **ppEnumBlobs
	) PURE;

	STDMETHOD(Next)(
		ULONG ulCount,
		HBLOB *phblob,
		ULONG *pcFetched
	) PURE;

	STDMETHOD(Reset)( void ) PURE;

	STDMETHOD(Skip)( ULONG ulCount ) PURE;      
};

// IID_IEnumTextTrackers : c78232b3-a5b8-4613-b6c3-685e5ce2d259
DEFINE_GUID(IID_IEnumTextTrackers, 0xc78232b3, 0xa5b8, 0x4613, 0xb6, 0xc3, 0x68, 0x5e, 0x5c, 0xe2, 0xd2, 0x59);
interface IEnumTextTrackers : public IUnknown
{
	STDMETHOD(Clone)(
		IEnumTextTrackers **ppEnumTrackers
	) PURE;

	STDMETHOD(Next)(
		ULONG ulCount,
		HTRACKER *phtracker,
		ULONG *pcFetched
	) PURE;

	STDMETHOD(Reset)( void ) PURE;

	STDMETHOD(Skip)( ULONG ulCount ) PURE;      
};

// IID_ITextMarkContainer : 6bda2cdc-d9b3-4310-8fcb-5690308c2e7c
DEFINE_GUID(IID_ITextMarkContainer, 0x6bda2cdc, 0xd9b3, 0x4310, 0x8f, 0xcb, 0x56, 0x90, 0x30, 0x8c, 0x2e, 0x7c);
interface ITextMarkContainer : public IUnknown
{
	// ------- Blob Methods ----------------------
														  
	typedef HRESULT (WINAPI *CreateBlobProc)(ITextBlob **, REFCLSID, HBLOB, LPVOID);

	STDMETHOD(SetBlobFactory) (
		CreateBlobProc pCreateBlobFunction,
		LPVOID pvClientData
	) PURE;

	enum SLBLOB_FLAGS {
		SLBLOB_RIGHTTOLEFT = 1,
		SLBLOB_TREATASSPACE = 2,
		SLBLOB_SHRINKTOFIT = 4
	};

	STDMETHOD(InsertSingleLineBlob) (
		LONG cp,
		ITextBlob *pBlob,
		DWORD dwFlags,			// Combination of SLBLOB_FLAGS
		BYTE  bTextAlign,		// One of TA_TOP, TA_BOTTOM, TA_BASELINE
		HBLOB &hblob			// Out : Rich edit handle for the blob (for queries)
	) PURE;

	enum MLBLOB_FLAGS
	{
		// Horizontal alignment flags
		MLBLOB_ALIGN_PAGE_LEFT = 0,
		MLBLOB_ALIGN_PAGE_CENTER = 1,
		MLBLOB_ALIGN_PAGE_RIGHT = 2,

		// Vertical alignment flags
		MLBLOB_ALIGN_NEXT_LINE = 0,
		MLBLOB_ALIGN_PAGE_TOP = 4,
		MLBLOB_ALIGN_PAGE_MIDDLE = 8,
		MLBLOB_ALIGN_PAGE_BOTTOM = 12
	};

	struct MLBLOB_CREATE_DATA
	{
		UINT  cbSize;
		SHORT nVertices;		// Number of vertices in blob's polygon
		SHORT xDistTextLeft;	// Tight wrap text distance on the left side
		SHORT xDistTextRight;	// Tight wrap text distance on the right side
		LONG  xDisplacement;	// Alignment displacement in x direction
		LONG  yDisplacement;	// Alignment displacement in y direction
								// Both displacements are not taken into account
								// for MLBLOB_ALIGN_NEXT_LINE alignment.
	};

	STDMETHOD(InsertMultiLineBlob) (
		LONG cp,
		ITextBlob *pBlob,
		DWORD dwFlags,			// Combination of MLBLOB_FLAGS
		MLBLOB_CREATE_DATA const *pmlbcd,	// NULL means all zeros
		HBLOB &hblob			// Out : Rich edit handle for the blob (for queries)
	) PURE;

	STDMETHOD(PinBlob) (
		LPPOINT lpPoint,			// A value of NULL will cause the blob to become unpinned
		HBLOB hblob
	) PURE;

	STDMETHOD(BlobFromHBlob) (
		HBLOB hblob,
		ITextBlob **ppBlob
	) PURE;

	STDMETHOD(BlobFromCp) (
		LONG cp,
		int iDir,
		LONG &cpDelta,
		HBLOB &hblob
	) PURE;

	STDMETHOD(CpFromBlob) (
		HBLOB hblob,
		LONG &cp
	) PURE;

	STDMETHOD(BlobsFromPoint) (
		POINT pt,
		IEnumTextBlobs **ppEnumBlobs
	) PURE;

	STDMETHOD(EnumBlobs) (
		IEnumTextBlobs **ppEnumBlobs
	) PURE;

	STDMETHOD(GetBlobPosRect) (
		HBLOB hblob,
		BOOL fOverlapping,
		RECT &rcPosRect
	) PURE;

	STDMETHOD(OnBlobPropertiesChange) (
		HBLOB hblob
	) PURE;

	STDMETHOD(RemoveBlob) (
		HBLOB hblob
	) PURE;

	// ------- Tracker Methods ----------------------

	STDMETHOD(InsertTracker) (
		LONG cpMin,
		LONG cpMost,
		ITextTracker *pTracker,
		HTRACKER &htracker
	) PURE;

	STDMETHOD(GetTrackerRange) (
		HTRACKER htracker,
		LONG &cpMin,
		LONG &cpMost
	) PURE;

	// SetTrackerRange can be used to delete a tracker by specifying -1, -1 for cps.
	STDMETHOD(SetTrackerRange) (
		HTRACKER htracker,
		LONG cpMin,
		LONG cpMost
	) PURE;

	STDMETHOD(TrackersFromCp) (
		LONG cp,
		IEnumTextTrackers **ppEnumTrackers
	) PURE;

	STDMETHOD(EnumTrackers) (
		IEnumTextTrackers **ppEnumTrackers
	) PURE;

	STDMETHOD(TrackerFromHTracker) (
		HTRACKER htracker,
		ITextTracker **ppTracker
	) PURE;

	STDMETHOD(RemoveTracker) (
		HTRACKER htracker
	) PURE;

	// ------- Boundary Methods ----------------------

	STDMETHOD(InsertBoundary) (
		SHORT nVertices,
		LPPOINT lpPoints,
		HBOUNDARY &hboundary
	) PURE;

	STDMETHOD(DeleteBoundary) (
		HBOUNDARY hboundary
	) PURE;

	// ------- Text Geometry Methods ----------------------

	enum TEXTRECT_FLAGS
	{
		TEXTRECT_ORIGIN_CLIENT			= 0x00000000,
		TEXTRECT_ORIGIN_CONTENT			= 0x00000001,
		TEXTRECT_ORIGIN_LAYOUT			= 0x00000002,
		TEXTRECT_EXCLUDE_MARGINS		= 0x00000004,
		TEXTRECT_EXCLUDE_SELECTIONBAR	= 0x00000008,
		TEXTRECT_EXCLUDE_BORDER			= 0x00000010,
		TEXTRECT_EXCLUDE_PADDING		= 0x00000020,
		TEXTRECT_EXCLUDE_INDENT			= 0x00000040,
		TEXTRECT_EXCLUDE_BULLET			= 0x00000080,
		TEXTRECT_EXCLUDE_TRAILING_GAP	= 0x00000100,
		TEXTRECT_EXCLUDE_LINE_SPACING	= 0x00000200,
		TEXTRECT_INCLUDE_OVERLAP		= 0x00000400,
		TEXTRECT_INCLUDE_TEXTHANG		= 0x00000800,
		TEXTRECT_EXCLUDE_TRAILING_SPACES= 0x00001000,
		TEXTRECT_COLLAPSE_TO_BASELINE	= 0x80000000,

		TEXTRECT_TEXT_ONLY =	TEXTRECT_EXCLUDE_MARGINS |
								TEXTRECT_EXCLUDE_SELECTIONBAR |
								TEXTRECT_EXCLUDE_BORDER |
								TEXTRECT_EXCLUDE_PADDING |
								TEXTRECT_EXCLUDE_INDENT |
								TEXTRECT_EXCLUDE_BULLET |
								TEXTRECT_EXCLUDE_TRAILING_GAP,

		TEXTRECT_VIEW_WIDTH =	TEXTRECT_EXCLUDE_MARGINS |
								TEXTRECT_EXCLUDE_SELECTIONBAR,

		TEXTRECT_VIEW_NATURAL =
								TEXTRECT_EXCLUDE_TRAILING_GAP |
								TEXTRECT_EXCLUDE_TRAILING_SPACES,

		TEXTRECT_VIEW_RENDERING =
								TEXTRECT_VIEW_WIDTH |
								TEXTRECT_INCLUDE_OVERLAP
	};

	enum TEXTHITTEST_FLAGS
	{
		TEXTHITTEST_BLOBS_IGNORE = 0,
		TEXTHITTEST_BLOBS_NOOVERLAP = 1,
		TEXTHITTEST_BLOBS_OVERLAP = 2,
		TEXTHITTEST_ON_GAP_LEFT = 4,
		TEXTHITTEST_ON_GAP_RIGHT = 8,
		TEXTHITTEST_ON_GAP_UP = 16,
		TEXTHITTEST_ON_GAP_DOWN = 32,
		TEXTHITTEST_ON_GAP_BACKWARD = 0,
		TEXTHITTEST_ON_GAP_FORWARD = 64
	};


	// TEXTPOSFROMPOINTDETAILS is an auxiliary data type used for pvDetails
	// argument of ITextMarkContainer::GetTextPosFromPoint.
	struct TEXTPOSFROMPOINTDETAILS
	{
		UINT	cbSize;
		DWORD	dwFlags;
		LONG	lAux1;
	};

	STDMETHOD(GetRectFromTextPos) (
		LONG cpTextPos,
		DWORD dwTextPosFlags,
		DWORD dwTextRectFlags,
		RECT &rect
	) PURE;

	enum HTDETAILS_FLAGS
	{
		HTDETAILS_LINE = 1
	};

	STDMETHOD(GetTextPosFromPoint) (
		POINT pt,
		DWORD dwHittestFlags,
		DWORD dwTextPosFlags,
		LONG &cpTextPos,
		TEXTPOSFROMPOINTDETAILS *ptpfpd
	) PURE;

	typedef HRESULT (WINAPI *BlobFromDataObjectProc)(IDataObject*, CHARRANGE *, LPVOID);

	STDMETHOD(SetBlobFromDataObjectFactory) (
		BlobFromDataObjectProc pBlobFromDataObjectFunction
	) PURE;

	STDMETHOD(BlobsFromRect) (
		RECT rc,
		IEnumTextBlobs **ppEnumBlobs
	) PURE;
};


// This interface enables clients to do custom rendering. Return FALSE for
// GetCharWidthW and RichEdit will call the OS to fetch character widths.
interface ICustomTextOut
{
	virtual BOOL WINAPI ExtTextOutW(HDC, int, int, UINT, CONST RECT *, LPCWSTR, UINT, CONST INT *) = 0;
	virtual BOOL WINAPI GetCharWidthW(HDC, UINT, UINT, LPINT) = 0;
	virtual BOOL WINAPI NotifyCreateFont(HDC) = 0;
	virtual void WINAPI NotifyDestroyFont(HFONT) = 0;
	virtual void WINAPI GetKerningPairs(HDC, DWORD, LPKERNINGPAIR) = 0;
};
STDAPI SetCustomTextOutHandlerEx(ICustomTextOut **ppcto, DWORD dwFlags);

#endif // __cplusplus

// Window messages
#define EM_LASTREMSG			(WM_USER + 255)

// Formatting enhancements (mostly realted to East Asian typography)
#ifndef EM_SETPARAFORMATOPTIONS
#define EM_SETPARAFORMATOPTIONS	(WM_USER + 128)
#define EM_GETPARAFORMATOPTIONS	(WM_USER + 129)
#define EM_SETCHAREFFECTS		(WM_USER + 130)
#define EM_GETCHAREFFECTS		(WM_USER + 131)
#define EM_SETEMPHASISFONTPROC	(WM_USER + 132)
#define EM_GETEMPHASISFONTPROC	(WM_USER + 133)
#endif

// The host can implement an extended interface ITextHostEx containing some
// special operations. Its lifetime is host's lifetime (exceeding RichEdit's
// lifetime), so RichEdit can rely on a pointer provided by the client via
// EM_SETITEXTHOSTEX message (the interface pointer goes in lparam).
#define EM_SETITEXTHOSTEX		(WM_USER + 134)

// Adjust text procedure : Can be used by the client to provide substitute text
#define EM_GETADJUSTTEXTPROC	(WM_USER + 233)
#define EM_SETADJUSTTEXTPROC	(WM_USER + 234)
#define EM_CALLADJUSTTEXTPROC	(WM_USER + 255)

// Adjust text callback
typedef int (WINAPI *AdjustTextProc)(LANGID langid, const WCHAR *pszBefore, WCHAR *pszAfter, LONG cchAfter, LONG *pcchReplaced);

#define ATP_NOCHANGE			0
#define ATP_CHANGE				1
#define ATP_NODELIMITER			2

#ifndef EM_GETAUTOCORRECTPROC
// Remove when clients have been changed for new name
#define EM_GETAUTOCORRECTPROC	(WM_USER + 233)
#define EM_SETAUTOCORRECTPROC	(WM_USER + 234)
typedef BOOL (WINAPI *AutoCorrectProc)(LANGID langid, const WCHAR *pszBefore, WCHAR *pszAfter, LONG cchAfter, LONG *pcchReplaced);
#endif

// Special inline object insertion
#ifndef EM_INSERTOBJECT
#define EM_INSERTOBJECT 		(WM_USER + 245)
#endif

// Special inline object insertion

enum OBJECTTYPE {			// EM_INSERTOBJECT wparam
	tomSimpleText = 0,
	tomRuby,
	tomHorzVert,
	tomWarichu,

	tomEq = 9,
	tomMath = 10,
	tomAbs = tomMath,
	tomArray,
	tomBox,
	tomBraces,
	tomBrackets,
	tomDisplace,
	tomFence,
	tomFraction,
	tomIntegral,
	tomList,
	tomOverscore,
	tomOverstrike,
	tomParens,
	tomProduct,
	tomRadical,
	tomSubscript,
	tomSummation,
	tomSuperscript,
	tomUnderscore,
	tomObjectMax = tomUnderscore
};

#define RUBYBELOW		0x80			// OR with rubycharjust for lparam

// Message to complete/cancel IME/CTF composition.
#define	EM_TERMINATECOMPOSITION		(WM_USER + 246)
#define STC_COMPLETECOMPOSITION		0
#define STC_CANCELCOMPOSITION		1

// Extra RE Combobox setting.  Only apply to RE Combobox controls.
#define EM_SETCOMBOBOXSTYLE			(WM_USER + 247)
#define SCB_NOAUTOCOMPLETEONSIZE	1
#define SCB_SEARCHEXACTSTRING		2

#define EM_GETCOMBOBOXSTYLE			(WM_USER + 248)

#define EM_CREATCTFRANGE			(WM_USER + 249)

#define EM_GETRANGEFORMAT			(WM_USER + 250)
#define EM_SETRANGEFORMAT			(WM_USER + 251)

typedef struct _getrangeformat{
	CHARRANGE	charRange;
	DWORD		flags;
} GETRANGEFORMAT;

// HTML Stream in support
// Depends on Structure Storage IStream interface
#define EM_ISTREAMIN				(WM_USER + 252)
#define ISF_HTML	0x0001

// Get the parent of RE Listbox
#define EM_GETLBPARENT				(WM_USER + 253)

#define	EM_HANDLELINKNOTIFICATION	(WM_USER + 254)

// Extra RE Combobox/Listbox messages
#define CB_GETNATURALDROPPEDWIDTH	(WM_USER + 255)
#define LB_GETNATURALWIDTH			(WM_USER + 256)

#define EM_REFLECTCHANGE			(WM_USER + 257)

// Message pass to tracker during pre-cp change notification.
#define EM_QUERYCPCHANGE			(WM_USER + 258)

//	Bits used inside TMPDISPLAYATTRIB bMask
#define TDA_CRTEXTCOLOR			1
#define	TDA_CRBACKCOLOR			2
#define	TDA_CRUNDERLINECOLOR	4
#define	TDA_UNDERLINETYPE		8

typedef struct __tmpdisplayattrib{
    COLORREF	crTextColor;
    COLORREF	crBackColor;
    COLORREF	crUnderlineColor;
    BYTE		bUnderlineType;
    BYTE		bMask;		// indicate which fields are valid
    BYTE		fMixed;		// TRUE - indicate if there is a mixed tmp. display
							//	attribs in the cpRange.
} TMPDISPLAYATTRIB;


// GETRANGEFORMAT flags
#define	GRF_CHARACTER			1	// lparam contains (CHARFORMAT *)  To be deleted soon
#define GRF_CHARFORMAT			1	// lparam contains (CHARFORMAT *)
#define	GRF_TMPDISPLAYATTRIB	2	// lparam contains (TMPDISPLAYATTRIB *)
#define GRF_PARAFORMAT			3	// lparam contains (PARAFORMAT *)



// Formatting enhancements (mostly realted to East Asian typography)

// Masks and values for EM_SETPARAFORMATOPTIONS, EM_GETPARAFORMATOPTIONS
#define PFOM_JUSTIFYNEWSPAPER		0x00010000
#define PFOM_JUSTIFYNEWSPAPEREA		0x00020000
#define PFOM_JUSTIFYLASTLINE		0x00040000
#define PFOM_JUSTIFYSOFTENDLINE		0x00080000
#define PFOM_NOHANGINGPUNCT			0x00100000
#define PFOM_PUNCTSTARTLINE			0x00200000
#define PFOM_WIDCTLPAR				0x00400000

// PARAFORMAT OPTIONS "ALL" masks
#define PFOM_ALLEA	(PFOM_JUSTIFYNEWSPAPER | PFOM_JUSTIFYNEWSPAPEREA | \
					 PFOM_JUSTIFYLASTLINE  | PFOM_JUSTIFYSOFTENDLINE | \
					 PFOM_NOHANGINGPUNCT   | PFOM_PUNCTSTARTLINE)
#define PFOM_ALL	(PFOM_ALLEA | PFOM_WIDCTLPAR)

#define PFO_JUSTIFYNEWSPAPER		PFOM_JUSTIFYNEWSPAPER
#define PFO_JUSTIFYNEWSPAPEREA		PFOM_JUSTIFYNEWSPAPEREA
#define PFO_JUSTIFYLASTLINE			PFOM_JUSTIFYLASTLINE
#define PFO_JUSTIFYSOFTENDLINE		PFOM_JUSTIFYSOFTENDLINE
#define PFO_NOHANGINGPUNCT			PFOM_NOHANGINGPUNCT
#define PFO_PUNCTSTARTLINE			PFOM_PUNCTSTARTLINE
#define PFO_WIDCTLPAR				PFOM_WIDCTLPAR

// Auxiliary type for CHAREFFECTS
typedef struct _opentypefeature
{
	DWORD	dwFeatureTag;
	DWORD	dwParam;
} OPENTYPEFEATURE;

// CHAREFFECTS - data type for EM_SETCHAREFFECTS, EM_GETCHAREFFECTS
typedef struct _chareffects
{
	DWORD	dwMask;
	DWORD	dwEffects;
	OPENTYPEFEATURE*	prgOpenTypeFeatures;	
	BYTE	cOpenTypeFeatures;
	BYTE	bUnderlinePosition;	// see CEP_ constants
	BYTE	bEmphasisPosition;	// see CEP_ constants
	BYTE	bEmphasisStyle;		// see EMPHS_ constants
	BYTE	bCompressionMode;	// see CMPM_ constants
	BYTE	bReserved1;
	BYTE	bReserved2;
	BYTE	bReserved3;
} CHAREFFECTS;

// Effects mask bits
#define CEM_MODWIDTHPAIRS			0x00000001
#define CEM_MODWIDTHSPACE			0x00000002
#define CEM_AUTOSPACEALPHA			0x00000004
#define CEM_AUTOSPACENUMERIC		0x00000008
#define CEM_AUTOSPACEPARENTHESIS	0x00000010
#define CEM_EMBEDDEDFONT			0x00000020
#define CEM_DOUBLESTRIKE			0x00000040

// Other mask bits
#define CEM_UNDERLINEPOSITION		0x80000000
#define CEM_EMPHASISPOSITION		0x40000000
#define CEM_EMPHASISSTYLE			0x20000000
#define CEM_COMPRESSIONMODE			0x10000000

// Composite masks
#define CEM_AUTOSPACE		(CEM_AUTOSPACEALPHA |\
							CEM_AUTOSPACENUMERIC |\
							CEM_AUTOSPACEPARENTHESIS)
#define CEM_EFFECTSEA		(CEM_AUTOSPACE |\
							CEM_MODWIDTHSPACE |\
							CEM_MODWIDTHPAIRS)
#define CEM_EFFECTS			(CEM_EFFECTSEA |\
							CEM_EMBEDDEDFONT |\
							CEM_DOUBLESTRIKE)
#define CEM_ALL				(CEM_EFFECTS |\
							CEM_UNDERLINEPOSITION |\
							CEM_EMPHASISPOSITION |\
							CEM_EMPHASISSTYLE |\
							CEM_COMPRESSIONMODE)

#define CEP_AUTO					0
#define CEP_BELOW					1
#define CEP_ABOVE					2

#define CMPM_NONE					0
#define CMPM_PUNCTUATION			1
#define CMPM_PUNCTANDKANA			2

// Effect bits
#define CEF_MODWIDTHPAIRS			CEM_MODWIDTHPAIRS
#define CEF_MODWIDTHSPACE			CEM_MODWIDTHSPACE
#define CEF_AUTOSPACEALPHA			CEM_AUTOSPACEALPHA
#define CEF_AUTOSPACENUMERIC		CEM_AUTOSPACENUMERIC
#define CEF_AUTOSPACEPARENTHESIS	CEM_AUTOSPACEPARENTHESIS
#define CEF_EMBEDDEDFONT			CEM_EMBEDDEDFONT
#define CEF_DOUBLESTRIKE			CEM_DOUBLESTRIKE

// Predefined emphasis styles
#define EMPHS_NONE				0
#define EMPHS_DOT				1
#define EMPHS_COMMA				2
#define EMPHS_LASTPREDEFINED	2
#define EMPHS_MAX				255

// Predefined indexes in the vector of emphasis symbols
#define EMPHI_DOT				0
#define EMPHI_COMMA				1
#define EMPHI_RING				2
#define EMPHI_LASTPREDEFINED	2
#define EMPHI_MAX				255

// Auxiliary type for emphasis font callback
typedef struct _emphasissymbol
{
	WCHAR	wchHorizontal;
	WCHAR	wchVertical;
} EMPHASISSYMBOL;

// Emphasis font callback - fills out the emphasis font name (no more than
// LF_FACESIZE wide chars including the terminating zero) and an array of
// characters representing emphasis marks in the given font. This array is
// indexed by the emphasis type, so it contains no more than 255 symbols.
// The result is number of symbols put into prgEmphasisSymbols.
typedef LONG (CALLBACK* EMPHASISFONTPROC)(WCHAR* szEmphasisFont, EMPHASISSYMBOL* rgEmphasisSymbols);

// RichEdit binary format
// (other stream formats SF_TEXT, SF_RTF, SF_RTFNOOBJS, SF_TEXTIZED are defined
// in richedit.h)

#define SF_BINARY		0x0008

// Subpixel positioning support for custom text out

#define CTO_USESUBPIXELS	1

// Advanved options for EM_SETTYPOGRAPHYOPTIONS 

#define TO_ADVANCEDPAGELAYOUT	0x0010
#define TO_EAFORMATTING			0x0020
#define TO_INLINETEXTOBJECTS	0x0040

// Table insertion
#ifndef EM_INSERTTABLE
#define EM_INSERTTABLE			(WM_USER + 232)

// Data type defining table rows for EM_INSERTTABLE
typedef struct _tableRowParms
{							// EM_INSERTTABLE wparam is a (TABLEROWPARMS *)
	BYTE	cbRow;			// Count of bytes in this structure
	BYTE	cbCell; 		// Count of bytes in TABLECELLPARMS
	BYTE	cCell;			// Count of cells
	BYTE	cRow;			// Count of rows
	LONG	dxCellMargin;	// Cell left/right margin (\trgaph)
	LONG	dxIndent;		// Row left (right if fRTL indent (similar to \trleft)
	LONG	dyHeight;		// Row height (\trrh)
	DWORD	nAlignment:3;	// Row alignment (like PARAFORMAT::bAlignment, \trql, trqr, \trqc)
	DWORD	fRTL:1; 		// Display cells in RTL order (\rtlrow)
	DWORD	fKeep:1;		// Keep row together (\trkeep}
	DWORD	fKeepFollow:1;	// Keep row on same page as following row (\trkeepfollow)
	DWORD	fWrap:1;		// Wrap text to right/left (depending on bAlignment)
							// (see \tdfrmtxtLeftN, \tdfrmtxtRightN)
	DWORD	fIdentCells:1;	// lparam points at single struct valid for all cells
} TABLEROWPARMS;

// Data type defining table cells for EM_INSERTTABLE
typedef struct _tableCellParms
{							// EM_INSERTTABLE lparam is a (TABLECELLPARMS *)
	LONG	dxWidth;		// Cell width (\cellx)
	WORD	nVertAlign:2;	// Vertical alignment (0/1/2 = top/center/bottom
							//	\clvertalt (def), \clvertalc, \clvertalb)
	WORD	fMergeTop:1;	// Top cell for vertical merge (\clvmgf)
	WORD	fMergePrev:1;	// Merge with cell above (\clvmrg)
	WORD	fVertical:1;	// Display text top to bottom, right to left (\cltxtbrlv)
	WORD	wShading;		// Shading in .01%		(\clshdng) e.g., 10000 flips fore/back

	SHORT	dxBrdrLeft; 	// Left border width	(\clbrdrl\brdrwN) (in twips)
	SHORT	dyBrdrTop;		// Top border width 	(\clbrdrt\brdrwN)
	SHORT	dxBrdrRight;	// Right border width	(\clbrdrr\brdrwN)
	SHORT	dyBrdrBottom;	// Bottom border width	(\clbrdrb\brdrwN)
	COLORREF crBrdrLeft;	// Left border color	(\clbrdrl\brdrcf)
	COLORREF crBrdrTop; 	// Top border color 	(\clbrdrt\brdrcf)
	COLORREF crBrdrRight;	// Right border color	(\clbrdrr\brdrcf)
	COLORREF crBrdrBottom;	// Bottom border color	(\clbrdrb\brdrcf)
	COLORREF crBackPat; 	// Background color 	(\clcbpat)
	COLORREF crForePat; 	// Foreground color 	(\clcfpat)
} TABLECELLPARMS;

#endif // EM_INSERTTABLE

enum
{
	CARET_NONE		= 0,	// Normal Western caret (blinking vertical bar)
	CARET_CUSTOM	= 1,	// Use custom bitmap
	CARET_RTL		= 2,
	CARET_THAI		= 4,
	CARET_INDIC		= 8,
	CARET_LTR		= 16,
	CARET_ITALIC	= 32,
	CARET_NULL		= 64,	// Nondegenerate selection: use empty bitmap
	CARET_ROTATE90	= 128	// Rotate 90 degrees clockwise
};

// REC-specific data or view aspects in addition to standard DVASPECT bits -
// to use in dwDrawAspect argument of ITextServices::TxDraw.
enum
{
	DVASPECT_RE_NOPLAINCONTENT	= 0x01000000,
	DVASPECT_RE_NOSPECIALBLOBS	= 0x02000000,
};

typedef struct _CTFContext
{
	IUnknown	*pIC;
	DWORD		editCookie;
} CTFCONTEXT;

// Additional SES_ values
#define SES_AUTOMARGINS			0x40000000	// turn on/off auto-margin feature.
#define SES_MOUSECHECK			0x80000000	// clients want IME mouse operation

// Special return value for TxNotify EN_LINK/WM_CHAR case.  Client needs to return
// this value for both WM_LBUTTONDOWN and WM_LBUTTONUP notification before Richedit
// will perform default action.
#define EN_LINK_DO_DEFAULT	0x00444546

// Option for EM_SETLANGOPTIONS and EM_GETLANGOPTIONS 
#define IMF_DISABLEAUTOBIDIAUTOKEYBOARD	 0x0100 // Use language neutral auto-keyboard switching logic

// New search flags for FindText
#define FR_MATCHWIDTH		0x10000000	// Single-byte and double-byte versions of the
										// same character are considered to be not equal.
#define FR_IGNOREKANATYPE	0x08000000	// Corresponding Hiragana and Katakana characters,
										// when compared, are considered to be equal.

// New option for EM_GETIMECOMPTEXT to get the composition string
#define ICT_COMPSTR		2

// Notification structure for EN_ENDCOMPOSITION
typedef struct _endcomposition
{
	NMHDR	nmhdr;
	DWORD	dwCode;
} ENDCOMPOSITIONNOTIFY;

// Constants for ENDCOMPOSITIONNOTIFY dwCode
#define ECN_ENDCOMPOSITION	0x0001
#define ECN_NEWTEXT			0x0002

// Additional flags for the GETTEXTEX data structure 
#define GT_INCLUDENUMBERING  16	// Include list numbering / bullets

// BIDIOPTIONS new masks
#define	BOM_LBCONTEXTREADING	0x0020	// Listbox Context reading order

// BIDIOPTIONS new effects
#define	BOE_LBCONTEXTREADING	0x0020	// Listbox Context reading order

// E_KILLBLOBONLOAD is a special error code for ITextBlob::Save. The blob can
// return this code as a signal that it should not be re-created from the stream
// during the load operation; instead, it should be replaced by a space.
#define	E_KILLBLOBONLOAD		E_FAIL

#endif // !_MSFTEDIT_ 
