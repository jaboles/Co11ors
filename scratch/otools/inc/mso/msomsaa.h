/*------------------------------------------------------------------------*
 * msomsaa.h                                                              *
 *  -- Mso generic Accessible framework.                                  *
 *  -- No C++ constructs are allowed in this file!                        *
 *                                                                        *
 * Please do not modify or check in this file without contacting PaulCole.*
 *------------------------------------------------------------------------*/


#ifndef __MSOMSAA_H__
#define __MSOMSAA_H__

enum
{
	msoDispPfn_get_accParent = 0,
	msoDispPfn_get_accChildCount,
	msoDispPfn_get_accChild,
	msoDispPfn_get_accName,
	msoDispPfn_get_accValue,
	msoDispPfn_get_accRole,
	msoDispPfn_get_accState,
	msoDispPfn_get_accKeyboardShortcut,
	msoDispPfn_get_accFocus,
	msoDispPfn_get_accHelp,
	msoDispPfn_get_accHelpTopic,
	msoDispPfn_get_accSelection,
	msoDispPfn_get_accDefaultAction,
	msoDispPfn_get_accDescription,
	msoDispPfn_accSelect,
	msoDispPfn_accLocation,
	msoDispPfn_accNavigate,
	msoDispPfn_accHitTest,
	msoDispPfn_accDoDefaultAction,

	// Add new call backs above here.
	msoDispPfnMax,
};

typedef void * PFNMSODISP;

// When you define these in your application, please use STDMETHODIMP_(int) to define them
typedef HRESULT (__stdcall *PFN_get_accParent)(void *, IDispatch **);
typedef HRESULT (__stdcall *PFN_get_accChildCount)(void *, long *);
typedef HRESULT (__stdcall *PFN_get_accChild)(void *, VARIANT, IDispatch **);
typedef HRESULT (__stdcall *PFN_get_accName)(void *, VARIANT, BSTR*);
typedef HRESULT (__stdcall *PFN_get_accRole)(void *, VARIANT, VARIANT*);
typedef HRESULT (__stdcall *PFN_get_accValue)(void *, VARIANT, BSTR*);
typedef HRESULT (__stdcall *PFN_get_accState)(void *, VARIANT, VARIANT*);
typedef HRESULT (__stdcall *PFN_get_accKeyboardShortcut)(void *, VARIANT, BSTR*);
typedef HRESULT (__stdcall *PFN_get_accFocus)(void *, VARIANT*);
typedef HRESULT (__stdcall *PFN_get_accHelp)(void *, VARIANT, BSTR *);
typedef HRESULT (__stdcall *PFN_get_accHelpTopic)(void *, BSTR *, VARIANT, long *);
typedef HRESULT (__stdcall *PFN_get_accSelection)(void *, VARIANT * pvarSelectedChildren);
typedef HRESULT (__stdcall *PFN_get_accDefaultAction)(void *, VARIANT, BSTR *);
typedef HRESULT (__stdcall *PFN_get_accDescription)(void *, VARIANT, BSTR *);
typedef HRESULT (__stdcall *PFN_accSelect)(void *, long, VARIANT);
typedef HRESULT (__stdcall *PFN_accLocation)(void *, long*, long*, long*, long*, VARIANT);
typedef HRESULT (__stdcall *PFN_accNavigate)(void *, long, VARIANT, VARIANT*);
typedef HRESULT (__stdcall *PFN_accHitTest)(void *, long, long, VARIANT*);
typedef HRESULT (__stdcall *PFN_accDoDefaultAction)(void *, VARIANT);

MSOAPI_(void *) MsoDispCreate(HWND hwnd, HMSOINST hmsoinst, 
								void * pUserData, IAccessible *pAccessible);

/* MsoDispSetPfn
	The pfn parameter must be able to be typecast as one of the above PFN_...
	types.  Also the msoDispPfn parameter must be one of the above 
	msoDispPfn_... enumerations, and must correspond to the pfn parameter.  */
MSOAPI_(PFNMSODISP) MsoDispSetPfn(void * pDispMso, PFNMSODISP pfn, int msoDispPfn);



enum
{
	msoMSAADoubleClick,
	msoMSAAClose,
	msoMSAAOpen,
	msoMSAAAlt,
	msoMSAAAltDownArrow,
};

MSOAPI_(BOOL) MsoFGetMSAAWtz(int mm, _Out_z_cap_(cch) WCHAR * wtz, int cch);

#endif // __MSOMSAA_H__
