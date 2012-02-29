/****************************************************************************
	Msooch.h

	Owner: ChrisMcB
 	Copyright (c) 1999 Microsoft Corporation

	Declarations for functions related to Services on the Web
****************************************************************************/

#ifndef MSOOCH_H
#define MSOOCH_H 1
#include <mshtmhst.h>

/*****************************************************************************
	MsoHwndCreateHTMLWindow

	Creates and shows a window that hosts trident.
	bstrUrl is the url to navigate
	pidhuih is a dic host ui handler that gets attached to the window, can be NULL
	sw should be either SW_SHOW or SW_SHOWMAXIMIZED

*****************************************************************************/
MSOAPI_(HWND) MsoHwndCreateHTMLWindow(int x, int y, int dx, int dy, HWND hwndParent,
							HMENU hmenu, HINSTANCE hinst, LPVOID pvParam,
							BSTR bstrUrl, IDocHostUIHandler *pidhuih, int sw);



/*****************************************************************************
// Enteries below are for the och control
*****************************************************************************/

// ------------------- Window messages for OC Host --------------------

// IUnknown::QueryInterface the hosted OC
typedef struct _QIMSG {
    const IID * qiid;
    void ** ppvObject;
} QIMSG, FAR * LPQIMSG;

// ................. Query Interface Message ..........
#define OCM_QUERYINTERFACE      (WM_USER+0)

#ifdef __cplusplus
inline HRESULT OCHost_QueryInterface(HWND hwndOCH, REFIID riid, LPVOID* ppv) \
{ QIMSG qimsg = {&riid, ppv}; \
  return (HRESULT)SNDMSG((hwndOCH), OCM_QUERYINTERFACE, (WPARAM)sizeof(qimsg), (LPARAM)&qimsg); \
}
#else
#define OCHost_QueryInterface(hwndOCH, riid, ppv) \
{ QIMSG qimsg = {&riid, ppv}; \
  SNDMSG((hwndOCH), OCM_QUERYINTERFACE, (WPARAM)sizeof(qimsg), (LPARAM)&qimsg); \
}
#endif

// ................ Initialize and activate the OC ...............
#define OCM_INITIALIZE      (WM_USER+1)
#define OCM_INITOC          OCM_INITIALIZE
#define OCHost_InitOC(hwndOCH, lpOCS) \
  (HRESULT)SNDMSG((hwndOCH), OCM_INITOC, 0, (LPARAM)lpOCS)


// ............... give ochost a parent IUnknown .......
#define OCM_SETOWNER            (WM_USER+2)
#define OCHost_SetOwner(hwndOC, punk) \
  (HRESULT)SNDMSG((hwndOC), OCM_SETOWNER, 0, (LPARAM)(IUnknown*)(punk))

// ............... DoVerb the OC .......
// n.b. iVerb is technically a long, WPARAM might truncate it
#define OCM_DOVERB              (WM_USER+3)
#define OCHost_DoVerb(hwndOC, iVerb, lpMsg) \
  (HRESULT)SNDMSG((hwndOC), OCM_DOVERB, (WPARAM)iVerb, (LPARAM)lpMsg)

#define OCM_SETDOCHOST              (WM_USER+4)
#define OCHost_SetDocHost(hwndOC, lpMsg) \
  (HRESULT)SNDMSG((hwndOC), OCM_SETDOCHOST, 0, (LPARAM)lpMsg)

#define OCM_SETIDISP                (WM_USER+5)
#define OCHost_SetIDisp(hwndOC, lpMsg) \
  (HRESULT)SNDMSG((hwndOC), OCM_SETIDISP, 0, (LPARAM)lpMsg)
// ------------------ Window Notify messages from OC Host --------------

#define OCN_FIRST               0x1300
#define OCN_COCREATEINSTANCE    (OCN_FIRST + 1)

typedef struct _OCNCOCREATEMSG {
    NMHDR nmhdr;
    CLSID clsidOC;
    IUnknown ** ppunk;
} OCNCOCREATEMSG, FAR * LPOCNCOCREATEMSG;

// NOTE: return values are defined as the following
// If the handler of OCN_COCREATEINSTANCE Notify message returns OCNCOCREATE_ALREADYCREATED,
// on return the (*ppvObj) is assumed to have the value of the OC's IUnkown pointer
#define OCNCOCREATE_CONTINUE       0
#define OCNCOCREATE_HANDLED       -1


#define OCN_PERSISTINIT         (OCN_FIRST + 2)
// NOTE: return values are defined as the following
// If the handler of OCN_PERSISTINIT Notify message returns OCNPERSIST_ABORT,
// the OCHOST will abort IPersist's initialization.
#define OCNPERSISTINIT_CONTINUE    0
#define OCNPERSISTINIT_HANDLED    -1

// The return value on the following notify messages are ignored.
#define OCN_ACTIVATE            (OCN_FIRST + 3)
#define OCN_DEACTIVATE          (OCN_FIRST + 4)
#define OCN_EXIT                (OCN_FIRST + 5)
#define OCN_ONPOSRECTCHANGE     (OCN_FIRST + 6)

typedef struct _OCNONPOSRECTCHANGEMSG {
    NMHDR nmhdr;
    LPCRECT prcPosRect;
} OCNONPOSRECTCHANGEMSG, FAR * LPOCNONPOSRECTCHANGEMSG;

#define OCN_ONUIACTIVATE        (OCN_FIRST + 7)
typedef struct _OCNONUIACTIVATEMSG {
    NMHDR nmhdr;
    IUnknown *punk;
} OCNONUIACTIVATEMSG, FAR * LPOCNONUIACTIVATEMSG;

#define OCNONUIACTIVATE_HANDLED       -1

#define OCN_ONSETSTATUSTEXT     (OCN_FIRST + 8)
typedef struct _OCNONSETSTATUSTEXT {
    NMHDR nmhdr;
    LPCOLESTR pwszStatusText;
} OCNONSETSTATUSTEXTMSG, FAR * LPOCNONSETSTATUSTEXTMSG;

/*****************************************************************************
	Constants
*****************************************************************************/

#define MSOOCHOST_CLASSA   "MsoOCHost" 	//class name for registering OLE control host 
#define OLECMDERR_E_FIRST            (OLE_E_LAST+1)
#define OLECMDERR_E_NOTSUPPORTED 	(OLECMDERR_E_FIRST)
#define OLECMDERR_E_DISABLED         (OLECMDERR_E_FIRST+1)
#define OLECMDERR_E_NOHELP           (OLECMDERR_E_FIRST+2)
#define OLECMDERR_E_CANCELED         (OLECMDERR_E_FIRST+3)
#define OLECMDERR_E_UNKNOWNGROUP     (OLECMDERR_E_FIRST+4)

#ifndef DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE
#define DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE (0x4000)
#endif

#endif
			