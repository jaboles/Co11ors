
#ifndef _SAVE93_H_
#define _SAVE93_H_

#ifdef __cplusplus
interface IVbaProvideStorage;
#else
typedef interface IVbaProvideStorage IVbaProvideStorage;
#endif

// Magic cookie passed to DllVbeInit and DllVbeGetHashOfCode
#define MSO_COOKIE 0x0FF1CE

// 6d5140db-7436-11ce-8034-00aa006009fa
DEFINE_GUID(IID_IVbaSaveAsOld, 0x6d5140db, 0x7436, 0x11ce, 0x89, 0x34, 0x00, 0xaa, 0x00, 0x60, 0x09, 0xfa);

// Passed in reserved field of init structure to get Esc key to break.
#define	VBEINITINFO_EscFlag     0x8000

// Passed in reserved field of init structure to get Option Compare Database.
#define	VBEINITINFO_OptCmpDatabase     0x4000

// Passed in VBEINITINFO::dwFlags to Disable the exception handler
// surrounding the IInvoke Calls.
// NOTE: Meant for Debug only
#define VBEINITINFO_DisableExceptionHandler        0x1000


// host index values
#define VBA_FMT_EXCEL93   1
#define VBA_FMT_PROJECT93 2
#define VBA_FMT_PWRPNT93  3

#undef  INTERFACE
#define INTERFACE  IVbaSaveAsOld
DECLARE_INTERFACE_(IVbaSaveAsOld, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaSaveAsOld methods ***
	STDMETHOD(SaveAsOld)(THIS_ IStorage *pstg, USHORT usHostIndex) PURE;
};



// {E03686C0-03DF-11d1-8F9C-00A0C9110057}
DEFINE_GUID(IID_IVbaPersistOldAccess,
0xe03686c0, 0x3df, 0x11d1, 0x8f, 0x9c, 0x0, 0xa0, 0xc9, 0x11, 0x0, 0x57);


#undef  INTERFACE
#define INTERFACE  IVbaPersistOldAccess
DECLARE_INTERFACE_(IVbaPersistOldAccess, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaPersistOldAccess methods ***
	STDMETHOD(LoadOldAccess)(THIS_   IStorage* lpStg, IVbaProvideStorage* lpPrStg,
					 REFGUID refGuid, IStorage* lpProjStg)  PURE;

	STDMETHOD(SaveOldAccess)(THIS_   IVbaProvideStorage* lpPrStg, ITypeInfo* pITypeInfo,
			        IStorage* lpIStorage) PURE;

};


// {98BFF7F0-15C6-11d1-8FAE-00A0C9110057}
DEFINE_GUID(IID_IVbaProjItemEx,
0x98bff7f0, 0x15c6, 0x11d1, 0x8f, 0xae, 0x0, 0xa0, 0xc9, 0x11, 0x0, 0x57);
	
#undef  INTERFACE
#define INTERFACE  IVbaProjItemEx
DECLARE_INTERFACE_(IVbaProjItemEx, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	STDMETHOD(Revert)(THIS) PURE;
	STDMETHOD(GetUniqueIdentifier)(THIS_ LPBSTR pbstrIdent) PURE;
	STDMETHOD(SetDirtyEx)(THIS_ BOOL fDirty) PURE;
};


// {799CE0DF-22BB-11d2-9939-00A0C9702442}
DEFINE_GUID(IID_IVbaProjectSiteEx, 0x799ce0df, 0x22bb, 0x11d2, 0x99, 0x39, 0x0, 0xa0, 0xc9, 0x70, 0x24, 0x42);

// Type of project-level dialog for CanOpenDialog
typedef enum {
        VBA_DLG_REFERENCES = 1,
        VBA_DLG_PROJPROPS  = 2
} VBADIALOG;

#undef  INTERFACE
#define INTERFACE  IVbaProjectSiteEx
DECLARE_INTERFACE_(IVbaProjectSiteEx, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaProjectSiteEx methods ***
	STDMETHOD(GetMsoDigSigData)(THIS_ LPVOID *ppvDigSigStore, LPVOID *ppvDoc) PURE;
        STDMETHOD(CanOpenCodeWindow)(THIS_ HWND hwndOwner, LPVOID pvVbaProjItem, BOOL *pfCanOpen) PURE;
        STDMETHOD(CanOpenDialog)(THIS_ HWND hwndOwner, VBADIALOG VbaDlg, BOOL *pfCanOpen) PURE;
        STDMETHOD(CanChangeName)(THIS_ LPVOID pvVbaProjItem, BOOL *pfCanChange) PURE; 
        STDMETHOD(CanDeleteProjItem)(THIS_ LPVOID pvVbaProjItem, BOOL *pfCanDelete) PURE; 
};

// {978684B1-FF8C-11cf-8D08-00A0C90F2732}
DEFINE_GUID(IID_IVbaProjectEx, 0x978684b1, 0xff8c, 0x11cf, 0x8d, 0x8, 0x0, 0xa0, 0xc9, 0xf, 0x27, 0x32);
// {5A21752C-1D54-11d4-B6AD-00C04FB17665}
DEFINE_GUID(IID_IVbaProjectEx2, 0x5a21752c, 0x1d54, 0x11d4, 0xb6, 0xad, 0x0, 0xc0, 0x4f, 0xb1, 0x76, 0x65);

// Compute the hash for either the source or compiled code
typedef enum {
        VBA_HASH_SOURCE = 1,
        VBA_HASH_ALL    = 2
} VBAHASHCODE;

// Passed to IVbaProjectEx::HashProject
typedef struct {
        VBAHASHCODE vbahashcode;
        BSTR bstrHash;
        BOOL fLoadedFromText;
} VBAHASH;

// EntryPoint to compute hash of a VBA storage
STDAPI DllVbeGetHashOfCode(DWORD dwCookie, IStorage *pStg, VBAHASH *pVbaHash);

#undef  INTERFACE
#define INTERFACE  IVbaProjectEx

// calendar values
#define VBA_CAL_GREG 0
#define VBA_CAL_HIJRI 1

DECLARE_INTERFACE_(IVbaProjectEx, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaProjectEx methods ***
        STDMETHOD(SetHelpFile)(THIS_ LPCOLESTR lpstrHelpFile) PURE;
        STDMETHOD(GetHelpFile)(THIS_ BSTR *pbstrHelpFile) PURE;

        STDMETHOD(GetLcidOriginal)(THIS_ LCID *plcid) PURE;
        STDMETHOD(GetHashOfCode)(THIS_ VBAHASH *pVbaHash) PURE;
        STDMETHOD(GetTypeLib)(THIS_ ITypeLib **pptlib) PURE;
	STDMETHOD(SetDatabaseLcid)(THIS_ ULONG lcid) PURE;
	STDMETHOD(SetCalendar)(THIS_ ULONG ulCalendar) PURE;
	STDMETHOD(GetLibIdOfRefLib)(THIS_ WORD index, BSTR* pbstrLibId, BOOL* pfDefault) PURE;
	STDMETHOD(MakeMDE)(LPSTORAGE pstgNew) PURE;
	STDMETHOD_(BOOL, IsMDE)(THIS) PURE;

        STDMETHOD(SetConstValues)(THIS_ LPCOLESTR lpstrConstValues) PURE;
        STDMETHOD(GetConstValues)(THIS_ BSTR *pbstrConstValues) PURE;
};

DECLARE_INTERFACE_(IVbaProjectEx2, IUnknown) // access only. See Vegas 6-63075
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaProjectEx2 methods ***
        STDMETHOD(RemoveReference)(_In_z_ THIS_ LPOLESTR lpstrRefFilePath) PURE;
};
// {978684B2-FF8C-11cf-8D08-00A0C90F2732}
DEFINE_GUID(IID_IVbaProcsEx, 0x978684b2, 0xff8c, 0x11cf, 0x8d, 0x8, 0x0, 0xa0, 0xc9, 0xf, 0x27, 0x32);

#undef  INTERFACE
#define INTERFACE  IVbaProcsEx
DECLARE_INTERFACE_(IVbaProcsEx, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaProcsEx methods ***
        STDMETHOD(SetHelpContext)(THIS_ LPCOLESTR lpstrProcName, DWORD dwHelpContextId) PURE;
        STDMETHOD(GetHelpContext)(THIS_ LPCOLESTR lpstrProcName, DWORD *pdwHelpContextId) PURE;
};

// {D5026F30-3894-11d2-994A-00A0C9702442}
DEFINE_GUID(IID_IVbaSiteEx, 0xd5026f30, 0x3894, 0x11d2, 0x99, 0x4a, 0x0, 0xa0, 0xc9, 0x70, 0x24, 0x42);

#undef  INTERFACE
#define INTERFACE  IVbaSiteEx
DECLARE_INTERFACE_(IVbaSiteEx, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaSiteEx methods ***
        STDMETHOD(ModulePassword)(THIS_ BSTR bstrModName, BSTR bstrModPassword) PURE;
        STDMETHOD(IsStandardModule)(THIS_ LPCOLESTR lpstrModName, BOOL *pfIsStdMod) PURE;
};

// {3B82B7DF-3A30-11d1-98C3-00A0C9702442}
DEFINE_GUID(IID_IVbaEx, 0x3b82b7df, 0x3a30, 0x11d1, 0x98, 0xc3, 0x0, 0xa0, 0xc9, 0x70, 0x24, 0x42);

#undef  INTERFACE
#define INTERFACE  IVbaEx

// BrkOnErrOption values
typedef enum _VBA_BRKONERR {
    VBA_BRKONERR_All,
    VBA_BRKONERR_ClassModule,
    VBA_BRKONERR_Unhandled
} VBA_BRKONERR;


DECLARE_INTERFACE_(IVbaEx, IUnknown)
{
	BEGIN_INTERFACE
	// *** IUnknown methods ****
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;

	// *** IVbaEx methods ***
        STDMETHOD(GetExprSrv)(THIS_ LPVOID FAR *ppvExprSrv,
                                    USHORT usVerMajExprSrv,
                                    USHORT usVerMinExprSrv) PURE;

        STDMETHOD(SetCommandLine)(_In_z_ THIS_ LPOLESTR szCmdLine) PURE;
	STDMETHOD(GetInfoOfLib)(_In_z_ THIS_ LPOLESTR szLibId, GUID* pguid, WORD* pwMaj, WORD *pwMin,
				LCID *plcid) PURE;

        STDMETHOD(SetBrkOnErrOption)(THIS_ VBA_BRKONERR vbaBrkOnErr) PURE;
        STDMETHOD(GetBrkOnErrOption)(THIS_ VBA_BRKONERR *pvbaBrkOnErr) PURE;

        STDMETHOD(EnterRunMode)(THIS_ BOOL fEnter) PURE;

	STDMETHOD(ViewImmediateWindow)(THIS)PURE;
};

#endif // _SAVE93_H_
