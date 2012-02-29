#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __std_h__
#define __std_h__
#ifndef _MAC 
#include "oleidl.h" 
#endif

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

typedef interface ISimpleTabularDataEvents ISimpleTabularDataEvents;


typedef interface ISimpleTabularData ISimpleTabularData;


typedef interface IForms96Binder IForms96Binder;


typedef interface IForms96BinderDispenser IForms96BinderDispenser;



#ifndef _MAC 
#include "oaidl.h"
#include "olectl.h"
#endif
//#include "objext.h"

/****************************************
 * Generated header for interface: __MIDL__intf_0000
 * at Mon Sep 25 15:03:06 1995
 * using MIDL 2.00.92
 ****************************************/
/* [auto_handle][local] */ 


#define STD_IndexLabel      (0)
#define STD_IndexAll        (~0ul)
#define STD_IndexUnknown    (~0ul)

			/* size is 2 */
typedef 
enum _STDRW
    {	STDRW_DEFAULT	= 1,
	STDRW_READONLY	= 0,
	STDRW_READWRITE	= 1,
	STDRW_MIXED	= 2
    }	STDRW;

			/* size is 2 */
typedef 
enum _STDFIND
    {	STDFIND_DEFAULT	= 1,
	STDFIND_EXACT	= 1,
	STDFIND_PREFIX	= 2,
	STDFIND_NEAREST	= 4,
	STDFIND_LAST	= 8,
	STDFIND_UP	= 16,
	STDFIND_CASESENSITIVE	= 32
    }	STDFIND;





/****************************************
 * Generated header for interface: ISimpleTabularDataEvents
 * at Mon Sep 25 15:03:06 1995
 * using MIDL 2.00.92
 ****************************************/
/* [uuid][unique][object][local] */ 



EXTERN_C const IID IID_ISimpleTabularDataEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface ISimpleTabularDataEvents : public IUnknown
    {
    public:
        virtual HRESULT __stdcall CellChanged( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn) = 0;
        
        virtual HRESULT __stdcall DeletedRows( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows) = 0;
        
        virtual HRESULT __stdcall InsertedRows( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows) = 0;
        
        virtual HRESULT __stdcall DeletedColumns( 
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns) = 0;
        
        virtual HRESULT __stdcall InsertedColumns( 
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ISimpleTabularDataEventsVtbl
    {
        
        HRESULT ( __stdcall *QueryInterface )( 
            ISimpleTabularDataEvents * This,
            /* [in] */ REFIID riid,
            /* [out] */ void **ppvObject);
        
        ULONG ( __stdcall *AddRef )( 
            ISimpleTabularDataEvents * This);
        
        ULONG ( __stdcall *Release )( 
            ISimpleTabularDataEvents * This);
        
        HRESULT ( __stdcall *CellChanged )( 
            ISimpleTabularDataEvents * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn);
        
        HRESULT ( __stdcall *DeletedRows )( 
            ISimpleTabularDataEvents * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows);
        
        HRESULT ( __stdcall *InsertedRows )( 
            ISimpleTabularDataEvents * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows);
        
        HRESULT ( __stdcall *DeletedColumns )( 
            ISimpleTabularDataEvents * This,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns);
        
        HRESULT ( __stdcall *InsertedColumns )( 
            ISimpleTabularDataEvents * This,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns);
        
    } ISimpleTabularDataEventsVtbl;

    interface ISimpleTabularDataEvents
    {
        CONST_VTBL struct ISimpleTabularDataEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISimpleTabularDataEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ISimpleTabularDataEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ISimpleTabularDataEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ISimpleTabularDataEvents_CellChanged(This,iRow,iColumn)	\
    (This)->lpVtbl -> CellChanged(This,iRow,iColumn)

#define ISimpleTabularDataEvents_DeletedRows(This,iRow,cRows)	\
    (This)->lpVtbl -> DeletedRows(This,iRow,cRows)

#define ISimpleTabularDataEvents_InsertedRows(This,iRow,cRows)	\
    (This)->lpVtbl -> InsertedRows(This,iRow,cRows)

#define ISimpleTabularDataEvents_DeletedColumns(This,iColumn,cColumns)	\
    (This)->lpVtbl -> DeletedColumns(This,iColumn,cColumns)

#define ISimpleTabularDataEvents_InsertedColumns(This,iColumn,cColumns)	\
    (This)->lpVtbl -> InsertedColumns(This,iColumn,cColumns)

#endif /* COBJMACROS */


#endif 	/* C style interface */

















/****************************************
 * Generated header for interface: ISimpleTabularData
 * at Mon Sep 25 15:03:06 1995
 * using MIDL 2.00.92
 ****************************************/
/* [uuid][unique][object][local] */ 



EXTERN_C const IID IID_ISimpleTabularData;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface ISimpleTabularData : public IUnknown
    {
    public:
        virtual HRESULT __stdcall GetDimensions( 
            /* [out] */ ULONG *pcRows,
            /* [out] */ ULONG *pcColumns) = 0;
        
        virtual HRESULT __stdcall GetRWStatus( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [out] */ STDRW *prwStatus) = 0;
        
        virtual HRESULT __stdcall GetVariant( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [out] */ VARIANTARG *pVar) = 0;
        
        virtual HRESULT __stdcall SetVariant( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [in] */ VARIANTARG *pVar) = 0;
        
        virtual HRESULT __stdcall GetString( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cchBuf,
            /* [out] */ OLECHAR *pchBuf,
            /* [out] */ ULONG *pcchActual) = 0;
        
        virtual HRESULT __stdcall SetString( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cchBuf,
            /* [in] */ OLECHAR *pchBuf) = 0;
        
        virtual HRESULT __stdcall DeleteRows( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows,
            /* [out] */ ULONG *pcRowsDeleted) = 0;
        
        virtual HRESULT __stdcall InsertRows( 
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows,
            /* [out] */ ULONG *pcRowsInserted) = 0;
        
        virtual HRESULT __stdcall DeleteColumns( 
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns,
            /* [out] */ ULONG *pcColumnsDeleted) = 0;
        
        virtual HRESULT __stdcall InsertColumns( 
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns,
            /* [out] */ ULONG *pcColumnsInserted) = 0;
        
        virtual HRESULT __stdcall FindPrefixString( 
            /* [in] */ ULONG iRowStart,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cchBuf,
            /* [in] */ OLECHAR *pchBuf,
            /* [in] */ DWORD findFlags,
            /* [out] */ STDFIND *foundFlag,
            /* [out] */ ULONG *piRowFound) = 0;
        
        virtual HRESULT __stdcall SetAdviseEvent( 
            /* [in] */ ISimpleTabularDataEvents *pEvent) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ISimpleTabularDataVtbl
    {
        
        HRESULT ( __stdcall *QueryInterface )( 
            ISimpleTabularData * This,
            /* [in] */ REFIID riid,
            /* [out] */ void **ppvObject);
        
        ULONG ( __stdcall *AddRef )( 
            ISimpleTabularData * This);
        
        ULONG ( __stdcall *Release )( 
            ISimpleTabularData * This);
        
        HRESULT ( __stdcall *GetDimensions )( 
            ISimpleTabularData * This,
            /* [out] */ ULONG *pcRows,
            /* [out] */ ULONG *pcColumns);
        
        HRESULT ( __stdcall *GetRWStatus )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [out] */ STDRW *prwStatus);
        
        HRESULT ( __stdcall *GetVariant )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [out] */ VARIANTARG *pVar);
        
        HRESULT ( __stdcall *SetVariant )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [in] */ VARIANTARG *pVar);
        
        HRESULT ( __stdcall *GetString )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cchBuf,
            /* [out] */ OLECHAR *pchBuf,
            /* [out] */ ULONG *pcchActual);
        
        HRESULT ( __stdcall *SetString )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cchBuf,
            /* [in] */ OLECHAR *pchBuf);
        
        HRESULT ( __stdcall *DeleteRows )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows,
            /* [out] */ ULONG *pcRowsDeleted);
        
        HRESULT ( __stdcall *InsertRows )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRow,
            /* [in] */ ULONG cRows,
            /* [out] */ ULONG *pcRowsInserted);
        
        HRESULT ( __stdcall *DeleteColumns )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns,
            /* [out] */ ULONG *pcColumnsDeleted);
        
        HRESULT ( __stdcall *InsertColumns )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cColumns,
            /* [out] */ ULONG *pcColumnsInserted);
        
        HRESULT ( __stdcall *FindPrefixString )( 
            ISimpleTabularData * This,
            /* [in] */ ULONG iRowStart,
            /* [in] */ ULONG iColumn,
            /* [in] */ ULONG cchBuf,
            /* [in] */ OLECHAR *pchBuf,
            /* [in] */ DWORD findFlags,
            /* [out] */ STDFIND *foundFlag,
            /* [out] */ ULONG *piRowFound);
        
        HRESULT ( __stdcall *SetAdviseEvent )( 
            ISimpleTabularData * This,
            /* [in] */ ISimpleTabularDataEvents *pEvent);
        
    } ISimpleTabularDataVtbl;

    interface ISimpleTabularData
    {
        CONST_VTBL struct ISimpleTabularDataVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISimpleTabularData_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ISimpleTabularData_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ISimpleTabularData_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ISimpleTabularData_GetDimensions(This,pcRows,pcColumns)	\
    (This)->lpVtbl -> GetDimensions(This,pcRows,pcColumns)

#define ISimpleTabularData_GetRWStatus(This,iRow,iColumn,prwStatus)	\
    (This)->lpVtbl -> GetRWStatus(This,iRow,iColumn,prwStatus)

#define ISimpleTabularData_GetVariant(This,iRow,iColumn,pVar)	\
    (This)->lpVtbl -> GetVariant(This,iRow,iColumn,pVar)

#define ISimpleTabularData_SetVariant(This,iRow,iColumn,pVar)	\
    (This)->lpVtbl -> SetVariant(This,iRow,iColumn,pVar)

#define ISimpleTabularData_GetString(This,iRow,iColumn,cchBuf,pchBuf,pcchActual)	\
    (This)->lpVtbl -> GetString(This,iRow,iColumn,cchBuf,pchBuf,pcchActual)

#define ISimpleTabularData_SetString(This,iRow,iColumn,cchBuf,pchBuf)	\
    (This)->lpVtbl -> SetString(This,iRow,iColumn,cchBuf,pchBuf)

#define ISimpleTabularData_DeleteRows(This,iRow,cRows,pcRowsDeleted)	\
    (This)->lpVtbl -> DeleteRows(This,iRow,cRows,pcRowsDeleted)

#define ISimpleTabularData_InsertRows(This,iRow,cRows,pcRowsInserted)	\
    (This)->lpVtbl -> InsertRows(This,iRow,cRows,pcRowsInserted)

#define ISimpleTabularData_DeleteColumns(This,iColumn,cColumns,pcColumnsDeleted)	\
    (This)->lpVtbl -> DeleteColumns(This,iColumn,cColumns,pcColumnsDeleted)

#define ISimpleTabularData_InsertColumns(This,iColumn,cColumns,pcColumnsInserted)	\
    (This)->lpVtbl -> InsertColumns(This,iColumn,cColumns,pcColumnsInserted)

#define ISimpleTabularData_FindPrefixString(This,iRowStart,iColumn,cchBuf,pchBuf,findFlags,foundFlag,piRowFound)	\
    (This)->lpVtbl -> FindPrefixString(This,iRowStart,iColumn,cchBuf,pchBuf,findFlags,foundFlag,piRowFound)

#define ISimpleTabularData_SetAdviseEvent(This,pEvent)	\
    (This)->lpVtbl -> SetAdviseEvent(This,pEvent)

#endif /* COBJMACROS */


#endif 	/* C style interface */































/****************************************
 * Generated header for interface: IForms96Binder
 * at Mon Sep 25 15:03:06 1995
 * using MIDL 2.00.92
 ****************************************/
/* [uuid][unique][object][local] */ 



EXTERN_C const IID IID_IForms96Binder;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface IForms96Binder : public IUnknown
    {
    public:
        virtual HRESULT __stdcall GetName( 
            /* [out] */ OLECHAR **ppName) = 0;
        
        virtual HRESULT __stdcall BindToObject( 
            /* [in] */ REFIID riidResult,
            /* [out] */ void **ppvResult) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IForms96BinderVtbl
    {
#ifdef MAC
		BEGIN_INTERFACE
#endif
        HRESULT ( __stdcall *QueryInterface )( 
            IForms96Binder * This,
            /* [in] */ REFIID riid,
            /* [out] */ void **ppvObject);
        
        ULONG ( __stdcall *AddRef )( 
            IForms96Binder * This);
        
        ULONG ( __stdcall *Release )( 
            IForms96Binder * This);
        
        HRESULT ( __stdcall *GetName )( 
            IForms96Binder * This,
            /* [out] */ OLECHAR **ppName);
        
        HRESULT ( __stdcall *BindToObject )( 
            IForms96Binder * This,
            /* [in] */ REFIID riidResult,
            /* [out] */ void **ppvResult);
        
    } IForms96BinderVtbl;

    interface IForms96Binder
    {
        CONST_VTBL struct IForms96BinderVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IForms96Binder_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IForms96Binder_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IForms96Binder_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IForms96Binder_GetName(This,ppName)	\
    (This)->lpVtbl -> GetName(This,ppName)

#define IForms96Binder_BindToObject(This,riidResult,ppvResult)	\
    (This)->lpVtbl -> BindToObject(This,riidResult,ppvResult)

#endif /* COBJMACROS */


#endif 	/* C style interface */











/****************************************
 * Generated header for interface: IForms96BinderDispenser
 * at Mon Sep 25 15:03:06 1995
 * using MIDL 2.00.92
 ****************************************/
/* [uuid][unique][object][local] */ 



EXTERN_C const IID IID_IForms96BinderDispenser;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface IForms96BinderDispenser : public IUnknown
    {
    public:
        virtual HRESULT __stdcall ParseName( 
            /* [in] */ OLECHAR *pszName,
            /* [out] */ IForms96Binder **ppBinder) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IForms96BinderDispenserVtbl
    {
#ifdef MAC
		BEGIN_INTERFACE
#endif
  
        HRESULT ( __stdcall *QueryInterface )( 
            IForms96BinderDispenser * This,
            /* [in] */ REFIID riid,
            /* [out] */ void **ppvObject);
        
        ULONG ( __stdcall *AddRef )( 
            IForms96BinderDispenser * This);
        
        ULONG ( __stdcall *Release )( 
            IForms96BinderDispenser * This);
        
        HRESULT ( __stdcall *ParseName )( 
            IForms96BinderDispenser * This,
            /* [in] */ OLECHAR *pszName,
            /* [out] */ IForms96Binder **ppBinder);
        
    } IForms96BinderDispenserVtbl;

    interface IForms96BinderDispenser
    {
        CONST_VTBL struct IForms96BinderDispenserVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IForms96BinderDispenser_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IForms96BinderDispenser_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IForms96BinderDispenser_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IForms96BinderDispenser_ParseName(This,pszName,ppBinder)	\
    (This)->lpVtbl -> ParseName(This,pszName,ppBinder)

#endif /* COBJMACROS */


#endif 	/* C style interface */








/****************************************
 * Generated header for interface: __MIDL__intf_0093
 * at Mon Sep 25 15:03:06 1995
 * using MIDL 2.00.92
 ****************************************/
/* [auto_handle][local] */ 


// External functions:
STDAPI FormsCreateRowset(ISimpleTabularData *pSTD, IUnknown **ppUnk);
STDAPI FormsRowsetIID(IID *pIIDRowset);

// External testing functions:
STDAPI FormsSTDCreate(ISimpleTabularData **ppSTD);
STDAPI FormsSTDCreatePopulate(
    ULONG cColumns, ULONG cchBuf, OLECHAR *pchBuf, BOOL fLabels,
    ISimpleTabularData **ppSTD);



/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
