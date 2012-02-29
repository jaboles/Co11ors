/*----------------------------------------------------------------------------
	%%File: FECOM.H
	%%Unit: FECOM
	%%Contact: seijia

	Header file of MSIME COM for applications.  Defines IFECommon.

----------------------------------------------------------------------------*/


#ifndef __FECOM_H__
#define __FECOM_H__

#undef  INTERFACE
#define INTERFACE   IFEClassFactory

typedef struct _IMEDLG
{
	int			cbIMEDLG;				//size of this structure
	HWND		hwnd;					//parent window handle
	LPWSTR		lpwstrWord;				//optional string
	int			nTabId;					//specifies a tab in dialog
} IMEDLG;

////////////////////////////////
// The IFEClassFactory Interface
////////////////////////////////

DECLARE_INTERFACE_(IFEClassFactory, IClassFactory)
{
	// IUnknown members
    STDMETHOD(QueryInterface)	(THIS_ REFIID refiid, VOID **ppv) PURE;
    STDMETHOD_(ULONG,AddRef)	(THIS) PURE;
    STDMETHOD_(ULONG,Release)	(THIS) PURE;

	// IFEClassFactory members
    STDMETHOD(CreateInstance)	(THIS_ LPUNKNOWN, REFIID, void **) PURE;
    STDMETHOD(LockServer)		(THIS_ BOOL) PURE;
};


#undef  INTERFACE
#define INTERFACE   IFECommon

//////////////////////////
// The IFECommon Interface
//////////////////////////

//no more entries in the dictionary
#define IFEC_S_ALREADY_DEFAULT			MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x7400)

DECLARE_INTERFACE_(IFECommon, IUnknown)
{
	// IUnknown members
    STDMETHOD(QueryInterface)	(THIS_ REFIID refiid, VOID **ppv) PURE;
    STDMETHOD_(ULONG,AddRef)	(THIS) PURE;
    STDMETHOD_(ULONG,Release)	(THIS) PURE;

	// IFECommon members
    STDMETHOD(IsDefaultIME)	(THIS_
							_Out_z_cap_(cszName) CHAR *szName,				//(out) name of MS-IME
							INT cszName					//(in) size of szName
							) PURE;
    STDMETHOD(SetDefaultIME)	(THIS) PURE;
	STDMETHOD(InvokeWordRegDialog) (THIS_
							IMEDLG *pimedlg				//(in) parameters
							) PURE;
	STDMETHOD(InvokeDictToolDialog) (THIS_
							IMEDLG *pimedlg				//(in) parameters
							) PURE;
};

#ifdef __cplusplus
extern "C" {
#endif

// The following API replaces CoCreateInstance(), when CLSID is not used.
HRESULT WINAPI CreateIFECommonInstance(VOID **ppvObj);
typedef HRESULT (WINAPI *fpCreateIFECommonInstanceType)(VOID **ppvObj);

#ifdef __cplusplus
} /* end of 'extern "C" {' */
#endif

#endif //__FECOM_H__
