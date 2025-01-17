#ifndef _MSADERR_H_
#define _MSADERR_H_
#ifndef FACILITY_WINDOWS
//+---------------------------------------------------------------------------
//
//  Microsoft OLE DB
//  Copyright (C) Microsoft Corporation, 1994 - 1997.
//
//----------------------------------------------------------------------------


//
// ***************************************************************************
// The comment after this block is a message compiler relic and should be
// ignored.  See the OLE 2 documentation for the correct layout:
//
// |S|       Context       |  Fac  |             Code              |
// |3|3 2 2 2 2 2 2 2 2 2 2|1 1 1 1|1 1 1 1 1 1                    |
// |1|0 9 8 7 6 5 4 3 2 1 0|9 8 7 6|5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0|
//
// S       - severity code; 0= success and 1= error
// Context - reserved for future use, may or may not be 0
// Fac     - facility code
// Code    - facility's status code
// ***************************************************************************
//

//
//  Values are 32 bit values layed out as follows:
//
//   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
//   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
//  +---+-+-+-----------------------+-------------------------------+
//  |Sev|C|R|     Facility          |               Code            |
//  +---+-+-+-----------------------+-------------------------------+
//
//  where
//
//      Sev - is the severity code
//
//          00 - Success
//          01 - Informational
//          10 - Warning
//          11 - Error
//
//      C - is the Customer code flag
//
//      R - is a reserved bit
//
//      Facility - is the facility code
//
//      Code - is the facility's status code
//
//
// Define the facility codes
//
#define FACILITY_WINDOWS                 0x8
#define FACILITY_ITF                     0x4


//
// Define the severity codes
//
#define STATUS_SEVERITY_SUCCESS          0x0
#define STATUS_SEVERITY_COERROR          0x2


//
// MessageId: DB_E_BOGUS
//
// MessageText:
//
//  Dummy error - need this error so that mc puts the above defines
//  inside the FACILITY_WINDOWS guard, instead of leaving it empty
//
#define DB_E_BOGUS                       ((HRESULT)0x80040EFFL)

#endif // FACILITY_WINDOWS

//
// Codes 0x0e00-0x0eff are reserved for the OLE DB group of
// interfaces.
//
// Free codes are:
//
//		Error:
//			-none-
//
//		Success:
//			0x0ed7
//


//
// OLEDBVER
//	OLE DB version number (0x0100); to use version 2.0 features,
//	#define OLEDBVER 0x0200 before including this file.
//

// If OLEDBVER is not defined, assume version 1.0
#ifndef OLEDBVER
#define OLEDBVER 0x0100
#endif

//
// MessageId: DB_E_BADACCESSORHANDLE
//
// MessageText:
//
//  Invalid accessor
//
#define DB_E_BADACCESSORHANDLE           ((HRESULT)0x80040E00L)

//
// MessageId: DB_E_ROWLIMITEXCEEDED
//
// MessageText:
//
//  Creating another row would have exceeded the total number of active
//  rows supported by the rowset
//
#define DB_E_ROWLIMITEXCEEDED            ((HRESULT)0x80040E01L)

//
// MessageId: DB_E_READONLYACCESSOR
//
// MessageText:
//
//  Unable to write with a read-only accessor
//
#define DB_E_READONLYACCESSOR            ((HRESULT)0x80040E02L)

//
// MessageId: DB_E_SCHEMAVIOLATION
//
// MessageText:
//
//  Given values violate the database schema
//
#define DB_E_SCHEMAVIOLATION             ((HRESULT)0x80040E03L)

//
// MessageId: DB_E_BADROWHANDLE
//
// MessageText:
//
//  Invalid row handle
//
#define DB_E_BADROWHANDLE                ((HRESULT)0x80040E04L)

//
// MessageId: DB_E_OBJECTOPEN
//
// MessageText:
//
//  An object was open
//
#define DB_E_OBJECTOPEN                  ((HRESULT)0x80040E05L)

//@@@+ V1.5
#if( OLEDBVER >= 0x0150 )
//
// MessageId: DB_E_BADCHAPTER
//
// MessageText:
//
//  Invalid chapter
//
#define DB_E_BADCHAPTER                  ((HRESULT)0x80040E06L)

#endif // OLEDBVER >= 0x0150
//@@@- V1.5

//
// MessageId: DB_E_CANTCONVERTVALUE
//
// MessageText:
//
//  A literal value in the command could not be converted to the
//  correct type due to a reason other than data overflow
//
#define DB_E_CANTCONVERTVALUE            ((HRESULT)0x80040E07L)

//
// MessageId: DB_E_BADBINDINFO
//
// MessageText:
//
//  Invalid binding info
//
#define DB_E_BADBINDINFO                 ((HRESULT)0x80040E08L)

//
// MessageId: DB_SEC_E_PERMISSIONDENIED
//
// MessageText:
//
//  Permission denied
//
#define DB_SEC_E_PERMISSIONDENIED        ((HRESULT)0x80040E09L)

//
// MessageId: DB_E_NOTAREFERENCECOLUMN
//
// MessageText:
//
//  Specified column does not contain bookmarks or chapters
//
#define DB_E_NOTAREFERENCECOLUMN         ((HRESULT)0x80040E0AL)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_LIMITREJECTED
//
// MessageText:
//
//  Some cost limits were rejected
//
#define DB_E_LIMITREJECTED               ((HRESULT)0x80040E0BL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_NOCOMMAND
//
// MessageText:
//
//  No command has been set for the command object
//
#define DB_E_NOCOMMAND                   ((HRESULT)0x80040E0CL)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_COSTLIMIT
//
// MessageText:
//
//  Unable to find a query plan within the given cost limit
//
#define DB_E_COSTLIMIT                   ((HRESULT)0x80040E0DL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_BADBOOKMARK
//
// MessageText:
//
//  Invalid bookmark
//
#define DB_E_BADBOOKMARK                 ((HRESULT)0x80040E0EL)

//
// MessageId: DB_E_BADLOCKMODE
//
// MessageText:
//
//  Invalid lock mode
//
#define DB_E_BADLOCKMODE                 ((HRESULT)0x80040E0FL)

//
// MessageId: DB_E_PARAMNOTOPTIONAL
//
// MessageText:
//
//  No value given for one or more required parameters
//
#define DB_E_PARAMNOTOPTIONAL            ((HRESULT)0x80040E10L)

//
// MessageId: DB_E_BADCOLUMNID
//
// MessageText:
//
//  Invalid column ID
//
#define DB_E_BADCOLUMNID                 ((HRESULT)0x80040E11L)

//
// MessageId: DB_E_BADRATIO
//
// MessageText:
//
//  Invalid ratio
//
#define DB_E_BADRATIO                    ((HRESULT)0x80040E12L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_BADVALUES
//
// MessageText:
//
//  Invalid value
//
#define DB_E_BADVALUES                   ((HRESULT)0x80040E13L)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_ERRORSINCOMMAND
//
// MessageText:
//
//  The command contained one or more errors
//
#define DB_E_ERRORSINCOMMAND             ((HRESULT)0x80040E14L)

//
// MessageId: DB_E_CANTCANCEL
//
// MessageText:
//
//  The executing command cannot be canceled
//
#define DB_E_CANTCANCEL                  ((HRESULT)0x80040E15L)

//
// MessageId: DB_E_DIALECTNOTSUPPORTED
//
// MessageText:
//
//  The provider does not support the specified dialect
//
#define DB_E_DIALECTNOTSUPPORTED         ((HRESULT)0x80040E16L)

//
// MessageId: DB_E_DUPLICATEDATASOURCE
//
// MessageText:
//
//  A data source with the specified name already exists
//
#define DB_E_DUPLICATEDATASOURCE         ((HRESULT)0x80040E17L)

//
// MessageId: DB_E_CANNOTRESTART
//
// MessageText:
//
//  The rowset was built over a live data feed and cannot be restarted
//
#define DB_E_CANNOTRESTART               ((HRESULT)0x80040E18L)

//
// MessageId: DB_E_NOTFOUND
//
// MessageText:
//
//  No key matching the described characteristics could be found within
//  the current range
//
#define DB_E_NOTFOUND                    ((HRESULT)0x80040E19L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_CANNOTFREE
//
// MessageText:
//
//  Ownership of this tree has been given to the provider
//
#define DB_E_CANNOTFREE                  ((HRESULT)0x80040E1AL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_NEWLYINSERTED
//
// MessageText:
//
//  The provider is unable to determine identity for newly inserted
//  rows
//
#define DB_E_NEWLYINSERTED               ((HRESULT)0x80040E1BL)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_GOALREJECTED
//
// MessageText:
//
//  No nonzero weights specified for any goals supported, so goal was
//  rejected; current goal was not changed
//
#define DB_E_GOALREJECTED                ((HRESULT)0x80040E1CL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_UNSUPPORTEDCONVERSION
//
// MessageText:
//
//  Requested conversion is not supported
//
#define DB_E_UNSUPPORTEDCONVERSION       ((HRESULT)0x80040E1DL)

//
// MessageId: DB_E_BADSTARTPOSITION
//
// MessageText:
//
//  lRowsOffset would position you past either end of the rowset,
//  regardless of the cRows value specified; cRowsObtained is 0
//
#define DB_E_BADSTARTPOSITION            ((HRESULT)0x80040E1EL)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_NOQUERY
//
// MessageText:
//
//  Information was requested for a query, and the query was not set
//
#define DB_E_NOQUERY                     ((HRESULT)0x80040E1FL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_NOTREENTRANT
//
// MessageText:
//
//  Provider called a method from IRowsetNotify in the consumer and	the
//  method has not yet returned
//
#define DB_E_NOTREENTRANT                ((HRESULT)0x80040E20L)

//
// MessageId: DB_E_ERRORSOCCURRED
//
// MessageText:
//
//  Errors occurred
//
#define DB_E_ERRORSOCCURRED              ((HRESULT)0x80040E21L)

//
// MessageId: DB_E_NOAGGREGATION
//
// MessageText:
//
//  A non-NULL controlling IUnknown was specified and the object being
//  created does not support aggregation
//
#define DB_E_NOAGGREGATION               ((HRESULT)0x80040E22L)

//
// MessageId: DB_E_DELETEDROW
//
// MessageText:
//
//  A given HROW referred to a hard- or soft-deleted row
//
#define DB_E_DELETEDROW                  ((HRESULT)0x80040E23L)

//
// MessageId: DB_E_CANTFETCHBACKWARDS
//
// MessageText:
//
//  The rowset does not support fetching backwards
//
#define DB_E_CANTFETCHBACKWARDS          ((HRESULT)0x80040E24L)

//
// MessageId: DB_E_ROWSNOTRELEASED
//
// MessageText:
//
//  All HROWs must be released before new ones can be obtained
//
#define DB_E_ROWSNOTRELEASED             ((HRESULT)0x80040E25L)

//
// MessageId: DB_E_BADSTORAGEFLAG
//
// MessageText:
//
//  One of the specified storage flags was not supported
//
#define DB_E_BADSTORAGEFLAG              ((HRESULT)0x80040E26L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_BADCOMPAREOP
//
// MessageText:
//
//  The comparison operator was invalid
//
#define DB_E_BADCOMPAREOP                ((HRESULT)0x80040E27L)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_BADSTATUSVALUE
//
// MessageText:
//
//  The specified status flag was neither DBCOLUMNSTATUS_OK nor
//  DBCOLUMNSTATUS_ISNULL
//
#define DB_E_BADSTATUSVALUE              ((HRESULT)0x80040E28L)

//
// MessageId: DB_E_CANTSCROLLBACKWARDS
//
// MessageText:
//
//  The rowset cannot scroll backwards
//
#define DB_E_CANTSCROLLBACKWARDS         ((HRESULT)0x80040E29L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_BADREGIONHANDLE
//
// MessageText:
//
//  Invalid region handle
//
#define DB_E_BADREGIONHANDLE             ((HRESULT)0x80040E2AL)

//
// MessageId: DB_E_NONCONTIGUOUSRANGE
//
// MessageText:
//
//  The specified set of rows was not contiguous to or overlapping the
//  rows in the specified watch region
//
#define DB_E_NONCONTIGUOUSRANGE          ((HRESULT)0x80040E2BL)

//
// MessageId: DB_E_INVALIDTRANSITION
//
// MessageText:
//
//  A transition from ALL* to MOVE* or EXTEND* was specified
//
#define DB_E_INVALIDTRANSITION           ((HRESULT)0x80040E2CL)

//
// MessageId: DB_E_NOTASUBREGION
//
// MessageText:
//
//  The specified region is not a proper subregion of the region
//  identified by the given watch region handle
//
#define DB_E_NOTASUBREGION               ((HRESULT)0x80040E2DL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_MULTIPLESTATEMENTS
//
// MessageText:
//
//  The provider does not support multi-statement commands
//
#define DB_E_MULTIPLESTATEMENTS          ((HRESULT)0x80040E2EL)

//
// MessageId: DB_E_INTEGRITYVIOLATION
//
// MessageText:
//
//  A specified value violated the integrity constraints for a column or
//  table
//
#define DB_E_INTEGRITYVIOLATION          ((HRESULT)0x80040E2FL)

//
// MessageId: DB_E_BADTYPENAME
//
// MessageText:
//
//  The given type name was unrecognized
//
#define DB_E_BADTYPENAME                 ((HRESULT)0x80040E30L)

//
// MessageId: DB_E_ABORTLIMITREACHED
//
// MessageText:
//
//  Execution aborted because a resource limit has been reached; no
//  results have been returned
//
#define DB_E_ABORTLIMITREACHED           ((HRESULT)0x80040E31L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_ROWSETINCOMMAND
//
// MessageText:
//
//  Cannot clone a command object whose command tree contains a rowset
//  or rowsets
//
#define DB_E_ROWSETINCOMMAND             ((HRESULT)0x80040E32L)

//
// MessageId: DB_E_CANTTRANSLATE
//
// MessageText:
//
//  Cannot represent the current tree as text
//
#define DB_E_CANTTRANSLATE               ((HRESULT)0x80040E33L)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_DUPLICATEINDEXID
//
// MessageText:
//
//  The specified index already exists
//
#define DB_E_DUPLICATEINDEXID            ((HRESULT)0x80040E34L)

//
// MessageId: DB_E_NOINDEX
//
// MessageText:
//
//  The specified index does not exist
//
#define DB_E_NOINDEX                     ((HRESULT)0x80040E35L)

//
// MessageId: DB_E_INDEXINUSE
//
// MessageText:
//
//  The specified index was in use
//
#define DB_E_INDEXINUSE                  ((HRESULT)0x80040E36L)

//
// MessageId: DB_E_NOTABLE
//
// MessageText:
//
//  The specified table does not exist
//
#define DB_E_NOTABLE                     ((HRESULT)0x80040E37L)

//
// MessageId: DB_E_CONCURRENCYVIOLATION
//
// MessageText:
//
//  The rowset was using optimistic concurrency and the value of a
//  column has been changed since it was last read
//
#define DB_E_CONCURRENCYVIOLATION        ((HRESULT)0x80040E38L)

//
// MessageId: DB_E_BADCOPY
//
// MessageText:
//
//  Errors were detected during the copy
//
#define DB_E_BADCOPY                     ((HRESULT)0x80040E39L)

//
// MessageId: DB_E_BADPRECISION
//
// MessageText:
//
//  A specified precision was invalid
//
#define DB_E_BADPRECISION                ((HRESULT)0x80040E3AL)

//
// MessageId: DB_E_BADSCALE
//
// MessageText:
//
//  A specified scale was invalid
//
#define DB_E_BADSCALE                    ((HRESULT)0x80040E3BL)

//
// MessageId: DB_E_BADID
//
// MessageText:
//
//  Invalid table ID
//
#define DB_E_BADID                       ((HRESULT)0x80040E3CL)

//
// MessageId: DB_E_BADTYPE
//
// MessageText:
//
//  A specified type was invalid
//
#define DB_E_BADTYPE                     ((HRESULT)0x80040E3DL)

//
// MessageId: DB_E_DUPLICATECOLUMNID
//
// MessageText:
//
//  A column ID was occurred more than once in the specification
//
#define DB_E_DUPLICATECOLUMNID           ((HRESULT)0x80040E3EL)

//
// MessageId: DB_E_DUPLICATETABLEID
//
// MessageText:
//
//  The specified table already exists
//
#define DB_E_DUPLICATETABLEID            ((HRESULT)0x80040E3FL)

//
// MessageId: DB_E_TABLEINUSE
//
// MessageText:
//
//  The specified table was in use
//
#define DB_E_TABLEINUSE                  ((HRESULT)0x80040E40L)

//
// MessageId: DB_E_NOLOCALE
//
// MessageText:
//
//  The specified locale ID was not supported
//
#define DB_E_NOLOCALE                    ((HRESULT)0x80040E41L)

//
// MessageId: DB_E_BADRECORDNUM
//
// MessageText:
//
//  The specified record number is invalid
//
#define DB_E_BADRECORDNUM                ((HRESULT)0x80040E42L)

//
// MessageId: DB_E_BOOKMARKSKIPPED
//
// MessageText:
//
//  Although the bookmark was validly formed, no row could be found to
//  match it
//
#define DB_E_BOOKMARKSKIPPED             ((HRESULT)0x80040E43L)

//
// MessageId: DB_E_BADPROPERTYVALUE
//
// MessageText:
//
//  The value of a property was invalid
//
#define DB_E_BADPROPERTYVALUE            ((HRESULT)0x80040E44L)

//
// MessageId: DB_E_INVALID
//
// MessageText:
//
//  The rowset was not chaptered
//
#define DB_E_INVALID                     ((HRESULT)0x80040E45L)

//
// MessageId: DB_E_BADACCESSORFLAGS
//
// MessageText:
//
//  Invalid accessor
//
#define DB_E_BADACCESSORFLAGS            ((HRESULT)0x80040E46L)

//
// MessageId: DB_E_BADSTORAGEFLAGS
//
// MessageText:
//
//  Invalid storage flags
//
#define DB_E_BADSTORAGEFLAGS             ((HRESULT)0x80040E47L)

//
// MessageId: DB_E_BYREFACCESSORNOTSUPPORTED
//
// MessageText:
//
//  By-ref accessors are not supported by this provider
//
#define DB_E_BYREFACCESSORNOTSUPPORTED   ((HRESULT)0x80040E48L)

//
// MessageId: DB_E_NULLACCESSORNOTSUPPORTED
//
// MessageText:
//
//  Null accessors are not supported by this provider
//
#define DB_E_NULLACCESSORNOTSUPPORTED    ((HRESULT)0x80040E49L)

//
// MessageId: DB_E_NOTPREPARED
//
// MessageText:
//
//  The command was not prepared
//
#define DB_E_NOTPREPARED                 ((HRESULT)0x80040E4AL)

//
// MessageId: DB_E_BADACCESSORTYPE
//
// MessageText:
//
//  The specified accessor was not a parameter accessor
//
#define DB_E_BADACCESSORTYPE             ((HRESULT)0x80040E4BL)

//
// MessageId: DB_E_WRITEONLYACCESSOR
//
// MessageText:
//
//  The given accessor was write-only
//
#define DB_E_WRITEONLYACCESSOR           ((HRESULT)0x80040E4CL)

//
// MessageId: DB_SEC_E_AUTH_FAILED
//
// MessageText:
//
//  Authentication failed
//
#define DB_SEC_E_AUTH_FAILED             ((HRESULT)0x80040E4DL)

//
// MessageId: DB_E_CANCELED
//
// MessageText:
//
//  The change was canceled during notification; no columns are changed
//
#define DB_E_CANCELED                    ((HRESULT)0x80040E4EL)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_CHAPTERNOTRELEASED
//
// MessageText:
//
//  The rowset was single-chaptered and the chapter was not released
//
#define DB_E_CHAPTERNOTRELEASED          ((HRESULT)0x80040E4FL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_E_BADSOURCEHANDLE
//
// MessageText:
//
//  Invalid source handle
//
#define DB_E_BADSOURCEHANDLE             ((HRESULT)0x80040E50L)

//
// MessageId: DB_E_PARAMUNAVAILABLE
//
// MessageText:
//
//  The provider cannot derive parameter info and SetParameterInfo has
//  not been called
//
#define DB_E_PARAMUNAVAILABLE            ((HRESULT)0x80040E51L)

//
// MessageId: DB_E_ALREADYINITIALIZED
//
// MessageText:
//
//  The data source object is already initialized
//
#define DB_E_ALREADYINITIALIZED          ((HRESULT)0x80040E52L)

//
// MessageId: DB_E_NOTSUPPORTED
//
// MessageText:
//
//  The provider does not support this method
//
#define DB_E_NOTSUPPORTED                ((HRESULT)0x80040E53L)

//
// MessageId: DB_E_MAXPENDCHANGESEXCEEDED
//
// MessageText:
//
//  The number of rows with pending changes has exceeded the set limit
//
#define DB_E_MAXPENDCHANGESEXCEEDED      ((HRESULT)0x80040E54L)

//
// MessageId: DB_E_BADORDINAL
//
// MessageText:
//
//  The specified column did not exist
//
#define DB_E_BADORDINAL                  ((HRESULT)0x80040E55L)

//
// MessageId: DB_E_PENDINGCHANGES
//
// MessageText:
//
//  There are pending changes on a row with a reference count of zero
//
#define DB_E_PENDINGCHANGES              ((HRESULT)0x80040E56L)

//
// MessageId: DB_E_DATAOVERFLOW
//
// MessageText:
//
//  A literal value in the command overflowed the range of the type of
//  the associated column
//
#define DB_E_DATAOVERFLOW                ((HRESULT)0x80040E57L)

//
// MessageId: DB_E_BADHRESULT
//
// MessageText:
//
//  The supplied HRESULT was invalid
//
#define DB_E_BADHRESULT                  ((HRESULT)0x80040E58L)

//
// MessageId: DB_E_BADLOOKUPID
//
// MessageText:
//
//  The supplied LookupID was invalid
//
#define DB_E_BADLOOKUPID                 ((HRESULT)0x80040E59L)

//
// MessageId: DB_E_BADDYNAMICERRORID
//
// MessageText:
//
//  The supplied DynamicErrorID was invalid
//
#define DB_E_BADDYNAMICERRORID           ((HRESULT)0x80040E5AL)

//
// MessageId: DB_E_PENDINGINSERT
//
// MessageText:
//
//  Unable to get visible data for a newly-inserted row that has not
//  yet been updated
//
#define DB_E_PENDINGINSERT               ((HRESULT)0x80040E5BL)

//
// MessageId: DB_E_BADCONVERTFLAG
//
// MessageText:
//
//  Invalid conversion flag
//
#define DB_E_BADCONVERTFLAG              ((HRESULT)0x80040E5CL)

//
// MessageId: DB_E_BADPARAMETERNAME
//
// MessageText:
//
//  The given parameter name was unrecognized
//
#define DB_E_BADPARAMETERNAME            ((HRESULT)0x80040E5DL)

//
// MessageId: DB_E_MULTIPLESTORAGE
//
// MessageText:
//
//  Multiple storage objects can not be open simultaneously
//
#define DB_E_MULTIPLESTORAGE             ((HRESULT)0x80040E5EL)

//
// MessageId: DB_E_CANTFILTER
//
// MessageText:
//
//  The requested filter could not be opened
//
#define DB_E_CANTFILTER                  ((HRESULT)0x80040E5FL)

//
// MessageId: DB_E_CANTORDER
//
// MessageText:
//
//  The requested order could not be opened
//
#define DB_E_CANTORDER                   ((HRESULT)0x80040E60L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_E_VOLUMENOTFOUND
//
// MessageText:
//
//  The specified volume was not found on the server
//
#define DB_E_VOLUMENOTFOUND              ((HRESULT)0x80040F00L)

//
// MessageId: DB_E_RESOURCENAMEINVALID
//
// MessageText:
//
//  The specified resource name is not valid
//
#define DB_E_RESOURCENAMEINVALID         ((HRESULT)0x80040F01L)

//
// MessageId: DB_E_RESOURCENOTFOUND
//
// MessageText:
//
//  The specified resource was not found on the server
//
#define DB_E_RESOURCENOTFOUND            ((HRESULT)0x80040F02L)

//
// MessageId: DB_E_RESOURCELOCKED
//
// MessageText:
//
//  Could not open resource because another process has it locked
//
#define DB_E_RESOURCELOCKED              ((HRESULT)0x80040F03L)

//
// MessageId: DB_E_RESOURCEEXISTS
//
// MessageText:
//
//  Could not complete operation because specified resource already exists
//
#define DB_E_RESOURCEEXISTS              ((HRESULT)0x80040F04L)

//
// MessageId: DB_E_CANTMOVE
//
// MessageText:
//
//  Can not move resource out of volume unless MOVERESOURCE_COPY_ALLOWED is specified
//
#define DB_E_CANTMOVE                    ((HRESULT)0x80040F05L)

//
// MessageId: DB_S_REDIRECTED
//
// MessageText:
//
//  The bind succeeded, but the URL of the resource has changed.
//
#define DB_S_REDIRECTED                  ((HRESULT)0x00040F06L)

//
// MessageId: DB_E_DENYNOTSUPPORTED
//
// MessageText:
//
//  The server or protocol provider does not support locks.
//
#define DB_E_DENYNOTSUPPORTED            ((HRESULT)0x80040F07L)

//
// MessageId: DB_E_DENYTYPENOTSUPPORTED
//
// MessageText:
//
//  The requested deny type is not supported by the server.
//
#define DB_E_DENYTYPENOTSUPPORTED        ((HRESULT)0x80040F08L)

//
// MessageId: DB_E_RESERVENOTSUPPORTED
//
// MessageText:
//
//  The server or protocol does not support reservations.
//
#define DB_E_RESERVENOTSUPPORTED         ((HRESULT)0x80040F09L)

//
// MessageId: DB_E_RESERVETYPENOTSUPPORTED
//
// MessageText:
//
//  The requested reservation type is not supported by the server.
//
#define DB_E_RESERVETYPENOTSUPPORTED     ((HRESULT)0x80040F0AL)

//
// MessageId: DB_E_UNINITIALIZED
//
// MessageText:
//
//  The bind failed because the binding object has not been initialized.
//
#define DB_E_UNINITIALIZED               ((HRESULT)0x80040F0BL)

//
// MessageId: DB_E_NOTCOLLECTION
//
// MessageText:
//
//  The client requested a object which is only valid for collections, but the resource is not a collection.
//
#define DB_E_NOTCOLLECTION               ((HRESULT)0x80040F0CL)

//
// MessageId: DB_E_READONLY
//
// MessageText:
//
//  The client requested write access to a read-only resource.
//
#define DB_E_READONLY                    ((HRESULT)0x80040F0DL)

//
// MessageId: DB_E_ASYNCNOTSUPPORTED
//
// MessageText:
//
//  The provider does not support asynchronous binding.
//
#define DB_E_ASYNCNOTSUPPORTED           ((HRESULT)0x80040F0EL)

//
// MessageId: DB_E_CANTCONNECT
//
// MessageText:
//
//  The provider could not connect to the server for the resource.
//
#define DB_E_CANTCONNECT                 ((HRESULT)0x80040F0FL)

//
// MessageId: DB_E_TIMEOUT
//
// MessageText:
//
//  The attempt to bind to the resource timed out.
//
#define DB_E_TIMEOUT                     ((HRESULT)0x80040F10L)

//
// MessageId: DB_E_CANTCHANGETYPE
//
// MessageText:
//
//  The type specified in the variant does not match the existing type for the property.
//
#define DB_E_CANTCHANGETYPE              ((HRESULT)0x80040F11L)

//
// MessageId: DB_E_WAITFORINITNOTSUPPORTED
//
// MessageText:
//
//  The provider does not support the BINDFLAG_WAITFORINIT option.
//
#define DB_E_WAITFORINITNOTSUPPORTED     ((HRESULT)0x80040F12L)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_S_ROWLIMITEXCEEDED
//
// MessageText:
//
//  Fetching requested number of rows would have exceeded total number
//  of active rows supported by the rowset
//
#define DB_S_ROWLIMITEXCEEDED            ((HRESULT)0x00040EC0L)

//
// MessageId: DB_S_COLUMNTYPEMISMATCH
//
// MessageText:
//
//  One or more column types are incompatible; conversion errors will
//  occur during copying
//
#define DB_S_COLUMNTYPEMISMATCH          ((HRESULT)0x00040EC1L)

//
// MessageId: DB_S_TYPEINFOOVERRIDDEN
//
// MessageText:
//
//  Parameter type information has been overridden by caller
//
#define DB_S_TYPEINFOOVERRIDDEN          ((HRESULT)0x00040EC2L)

//
// MessageId: DB_S_BOOKMARKSKIPPED
//
// MessageText:
//
//  Skipped bookmark for deleted or non-member row
//
#define DB_S_BOOKMARKSKIPPED             ((HRESULT)0x00040EC3L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_S_ERRORSINTREE
//
// MessageText:
//
//  Errors found in validating tree
//
#define DB_S_ERRORSINTREE                ((HRESULT)0x00040EC4L)

//
// MessageId: DB_S_NONEXTROWSET
//
// MessageText:
//
//  There are no more rowsets
//
#define DB_S_NONEXTROWSET                ((HRESULT)0x00040EC5L)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_S_ENDOFROWSET
//
// MessageText:
//
//  Reached start or end of rowset or chapter
//
#define DB_S_ENDOFROWSET                 ((HRESULT)0x00040EC6L)

//
// MessageId: DB_S_COMMANDREEXECUTED
//
// MessageText:
//
//  The provider re-executed the command
//
#define DB_S_COMMANDREEXECUTED           ((HRESULT)0x00040EC7L)

//
// MessageId: DB_S_BUFFERFULL
//
// MessageText:
//
//  Variable data buffer full
//
#define DB_S_BUFFERFULL                  ((HRESULT)0x00040EC8L)

//
// MessageId: DB_S_NORESULT
//
// MessageText:
//
//  There are no more results
//
#define DB_S_NORESULT                    ((HRESULT)0x00040EC9L)

//
// MessageId: DB_S_CANTRELEASE
//
// MessageText:
//
//  Server cannot release or downgrade a lock until the end of the
//  transaction
//
#define DB_S_CANTRELEASE                 ((HRESULT)0x00040ECAL)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_S_GOALCHANGED
//
// MessageText:
//
//  Specified weight was not supported or exceeded the supported limit
//  and was set to 0 or the supported limit
//
#define DB_S_GOALCHANGED                 ((HRESULT)0x00040ECBL)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//@@@+ V1.5
#if( OLEDBVER >= 0x0150 )
//
// MessageId: DB_S_UNWANTEDOPERATION
//
// MessageText:
//
//  Consumer is uninterested in receiving further notification calls for
//  this reason
//
#define DB_S_UNWANTEDOPERATION           ((HRESULT)0x00040ECCL)

#endif // OLEDBVER >= 0x0150
//@@@- V1.5

//
// MessageId: DB_S_DIALECTIGNORED
//
// MessageText:
//
//  Input dialect was ignored and text was returned in different
//  dialect
//
#define DB_S_DIALECTIGNORED              ((HRESULT)0x00040ECDL)

//
// MessageId: DB_S_UNWANTEDPHASE
//
// MessageText:
//
//  Consumer is uninterested in receiving further notification calls for
//  this phase
//
#define DB_S_UNWANTEDPHASE               ((HRESULT)0x00040ECEL)

//
// MessageId: DB_S_UNWANTEDREASON
//
// MessageText:
//
//  Consumer is uninterested in receiving further notification calls for
//  this reason
//
#define DB_S_UNWANTEDREASON              ((HRESULT)0x00040ECFL)

//@@@+ V1.5
#if( OLEDBVER >= 0x0150 )
//
// MessageId: DB_S_ASYNCHRONOUS
//
// MessageText:
//
//  The operation is being processed asynchronously
//
#define DB_S_ASYNCHRONOUS                ((HRESULT)0x00040ED0L)

#endif // OLEDBVER >= 0x0150
//@@@- V1.5

//
// MessageId: DB_S_COLUMNSCHANGED
//
// MessageText:
//
//  In order to reposition to the start of the rowset, the provider had
//  to reexecute the query; either the order of the columns changed or
//  columns were added to or removed from the rowset
//
#define DB_S_COLUMNSCHANGED              ((HRESULT)0x00040ED1L)

//
// MessageId: DB_S_ERRORSRETURNED
//
// MessageText:
//
//  The method had some errors; errors have been returned in the error
//  array
//
#define DB_S_ERRORSRETURNED              ((HRESULT)0x00040ED2L)

//
// MessageId: DB_S_BADROWHANDLE
//
// MessageText:
//
//  Invalid row handle
//
#define DB_S_BADROWHANDLE                ((HRESULT)0x00040ED3L)

//
// MessageId: DB_S_DELETEDROW
//
// MessageText:
//
//  A given HROW referred to a hard-deleted row
//
#define DB_S_DELETEDROW                  ((HRESULT)0x00040ED4L)

//@@@+ V2.0
#if( OLEDBVER >= 0x0200 )
//
// MessageId: DB_S_TOOMANYCHANGES
//
// MessageText:
//
//  The provider was unable to keep track of all the changes; the client
//  must refetch the data associated with the watch region using another
//  method
//
#define DB_S_TOOMANYCHANGES              ((HRESULT)0x00040ED5L)

#endif // OLEDBVER >= 0x0200
//@@@- V2.0

//
// MessageId: DB_S_STOPLIMITREACHED
//
// MessageText:
//
//  Execution stopped because a resource limit has been reached; results
//  obtained so far have been returned but execution cannot be resumed
//
#define DB_S_STOPLIMITREACHED            ((HRESULT)0x00040ED6L)

//
// MessageId: DB_S_LOCKUPGRADED
//
// MessageText:
//
//  A lock was upgraded from the value specified
//
#define DB_S_LOCKUPGRADED                ((HRESULT)0x00040ED8L)

//
// MessageId: DB_S_PROPERTIESCHANGED
//
// MessageText:
//
//  One or more properties were changed as allowed by provider
//
#define DB_S_PROPERTIESCHANGED           ((HRESULT)0x00040ED9L)

//
// MessageId: DB_S_ERRORSOCCURRED
//
// MessageText:
//
//  Errors occurred
//
#define DB_S_ERRORSOCCURRED              ((HRESULT)0x00040EDAL)

//
// MessageId: DB_S_PARAMUNAVAILABLE
//
// MessageText:
//
//  A specified parameter was invalid
//
#define DB_S_PARAMUNAVAILABLE            ((HRESULT)0x00040EDBL)

//
// MessageId: DB_S_MULTIPLECHANGES
//
// MessageText:
//
//  Updating this row caused more than one row to be updated in the
//  data source
//
#define DB_S_MULTIPLECHANGES             ((HRESULT)0x00040EDCL)

#endif // _OLEDBERR_H_
