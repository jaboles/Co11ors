/****************************************************************************
*                                                                           *
* Copyright (c) 1999 - 2001, Microsoft Corporation.  All rights reserved.   *
*                                                                           *
* File       - WsmDef.h                                                     *
*                                                                           *
* Purpose    - Contains Windows State Management public state engine        *
*              definitions                                                  *
*                                                                           *
* History    -     Date        Alias         Comment                        *
*              09-Jul-2001     State         Created                        *
*                                                                           *
****************************************************************************/

#ifndef _WSMDEF_H_
#define _WSMDEF_H_

#include "WsmId.h"

// State mangager timestamp definitions
#ifndef _WSMTIMESTAMP_
#define _WSMTIMESTAMP_
typedef enum wsmTimestamp   // WSM timestamp, seconds since 1/1/1980
{
    wsmtsBase = 0,          // base date = 1/1/1980
    wsmtsNull = 0xFFFFFFFF, // value reserved to indicate no timestamp
    wsmtsLocal= 0xFFFFFFFE, // data already or not to be committed to store
};
#endif // _WSMTIMESTAMP_

// Base datatypes - low-order 4 bits of type integer

typedef enum wsmdtEnum
{
    wsmdtUntyped       =  0,  // value not present, null, deleted, invalidated, trunctated
    wsmdtBoolean       =  1,  // boolean passed through API, stored internally as 0 or 1
    wsmdtInteger       =  2,  // integer 2, 4, 8 bytes, signed or unsigned
    wsmdtReal          =  3,  // 8- or 4-byte floating point value
    wsmdtDate          =  4,  // datetime, as Win32 datetime or wsmTimestamp
    wsmdtBinary        =  5,  // uninterpreted byte string
    wsmdtString        =  6,  // UTF-8 string stored without null
    wsmdtComposite     =  7,  // heterogeneous composite data, header stored for each item
    wsmdtBitFlag       =  8,  // indexed bit flag within integer 1-8 byte
    wsmdtBooleanList   =  9,  // list of named boolean values
    wsmdtIntegerList   = 10,  // list of named integer values
    wsmdtRealList      = 11,  // list of named floating point values
    wsmdtDateList      = 12,  // list of named data values
    wsmdtBinaryList    = 13,  // list of named binary byte strings
    wsmdtStringList    = 14,  // list of named UTF-8 strings
    wsmdtCompositeList = 15,  // list of named composite data values
};

// Datatype-specific options, storage format options or legacy conversion options

#define wsmdoNoFalseBool     0x00010000  // wsmdtBoolean only, legacy, store only True values, remove value if False
#define wsmdoBoolAsyesno     0x00020000  // wsmdtBoolean only, legacy, store in registry as "yes" or "no"
#define wsmdoBoolAsDigit     0x00040000  // wsmdtBoolean only, legacy, store in registry or INI as "0" or "1"
#define wsmdoBoolAsKey       0x00080000  // wsmdtBoolean only, legacy, Element is key name, absent or present

#define wsmdoShortInt        0x00000010  // wsmdtInt only, 2-byte integer
#define wsmdoDoubleInt       0x00000020  // wsmdtInt only, 8-byte integer
#define wsmdoUnsignedInt     0x00000040  // wsmdtInt only, treat as unsigned, used only where this matters
#define wsmdoBigEndian       0x00000080  // wsmdtInt only, non-Intel byte ordering in legacy store
#define wsmdoIntAsString     0x00010000  // wsmdtInt only, legacy, store integer in registry or INI as string
#define wsmdoIntAsBinary     0x00020000  // wsmdtInt only, legacy, store integer in registry as binary

#define wsmdoSingleReal      0x00000010  // wsmdtReal only, 4-byte float instead of 8-byte
#define wsmdoCurrency        0x00000020  // wsmdtReal only, interpret as currency

#define wsmdoTimestamp       0x00000010  // wsmdtDate only, WSM timestamp, seconds since 1/1/1908
#define wsmdoAutoDate        0x00010000  // wsmdtDate only, legacy, store as automation data (double)
                                       
#define wsmdoReservedString  0x00000010  // wsmdtString only, reserved to indicate string expression
#define wsmdoExpandString    0x00000020  // wsmdtString only, contains environment variables to expand
#define wsmdoMultiString     0x00000040  // wsmdtString only, contains embedded nulls, double null at end
#define wsmdoPathString      0x00000080  // wsmdtString only, need to provide relocatable directory root
#define wsmdoStringReference 0x00000100  // wsmdtString only, string stored as separate file
#define wsmdoNoEmptyString   0x00010000  // wsmdtString only, legacy, store only non-empty strings, remove value if empty
#define wsmdoUnicodeString   0x00020000  // wsmdtString only, legacy, store string in registry as Unicode REG_BINARY
#define wsmdoUTF8String      0x00040000  // wsmdtString only, legacy, store string in registry as UTF8 on Win9X
#define wsmdoBytePrefix      0x00001000  // wsmdtString only, not in the metastore, used at runtime (INTERNAL)
#define wsmdoShortPrefix     0x00002000  // wsmdtString only, not in the metastore, used at runtime (INTERNAL)
#define wsmdoLongPrefix      0x00004000  // wsmdtString only, not in the metastore, used at runtime (INTERNAL)

//#define wsmdoReservedBinary  0x0010
#define wsmdoGuidAsString    0x00000010  // wsmdtBinary only, binary GUID as string in VARIANT
#define wsmdoSecureBinary    0x00000020  // wsmdtBinary only, encrypted binary data
#define wsmdoBinaryReference 0x00000100  // wsmdtBinary only, data stored as separate file
#define wsmdoBytePrefix      0x00001000  // wsmdtBinary only, not in the metastore, used at runtime (INTERNAL)
#define wsmdoShortPrefix     0x00002000  // wsmdtBinary only, not in the metastore, used at runtime (INTERNAL)
#define wsmdoLongPrefix      0x00004000  // wsmdtBinary only, not in the metastore, used at runtime (INTERNAL)

//#define wsmdoReservedComposite 0x0010
#define wsmdoCompositeReference 0x00000100  // wsmdtComposite only, string stored as separate file
#define wsmdoCompositeBlob      0x00010000  // wsmdtComposite only, legacy, by default data packed as binary blob
#define wsmdoBlobAsInt          0x00020000  // wsmdtComposite only, legacy, data packed as integer blob

#define wsmdoBitFlagIndex    0x00000FF0  // wsmdtBitFlag only, bit to access

#define wsmdoNullValue       0x00000010  // wsmdtUntyped only, value is null in the database sense, VT_NULL
#define wsmdoInvalidated     0x00000020  // wsmdtUntyped only, data invalid, forces refresh on synch, NOTHING
#define wsmdoDeletedItem     0x00000040  // wsmdtUntyped only, removed from local and legacy stores, VT_EMPTY
#define wsmdoTruncatedItem   0x00000080  // wsmdtUntyped only, removed from local store, but not legacy store

// data transfer option bits used with datatype in store and transfer document

#define wsmxoNotifyOwner     0x00001000  // notify registered callback on data change
#define wsmxoNoHistory       0x00002000  // non-critical setting for which change history is not important
#define wsmxoNoRoam          0x00004000  // local data only, doesn't roam but may be backed up
#define wsmxoNoMerge         0x00008000  // never copied to store, backed up for informational use
#define wsmxoMask            0x0000F000  // data transafer options mask

#ifdef __cplusplus

// Datatype access inline functions
const int wsmdtListFlag = 8;  // type bit that indicates List or bitflag array
inline bool WsmIsListType(wsmdtEnum wsmdt)  { return wsmdt >  wsmdtBitFlag; }
inline bool WsmIsUntyped(UINT iDataType)    { return (iDataType & 15) == wsmdtUntyped; }
inline bool WsmIsNullValue(UINT iDataType)  { return (iDataType & 0xFF) == wsmdtUntyped + wsmdoNullValue; }
inline bool WsmIsDeleted(UINT iDataType)    { return (iDataType & 0xFF) == wsmdtUntyped + wsmdoDeletedItem; }
inline bool WsmIsRemoved(UINT iDataType)    { return (iDataType & 15) == wsmdtUntyped && (iDataType & wsmdoNullValue) == 0; }
inline bool WsmIsListType(UINT iDataType)   { return (iDataType & 15) >  wsmdtBitFlag; }
inline bool WsmIsComposite(UINT iDataType)  { return (iDataType & 15) == wsmdtComposite; }
inline bool WsmIsAggregate(UINT iDataType)  { return (iDataType & 15) == wsmdtComposite || (iDataType & 15) > wsmdtBitFlag; }
inline wsmdtEnum WsmTypeFromTypeBits(UINT iDataType) { return (wsmdtEnum)(iDataType & 15); }
inline UINT      WsmTypeWithOptions(UINT iDataType) { return (iDataType & 0x000F03FF); }
inline UINT      WsmTypeWithBaseOptions(UINT iDataType) { return (iDataType & 0x03FF); } // no legacy bits
inline UINT      WsmOptionsFromTypeBits(UINT iDataType) { return (iDataType & 0x000F03F0); }
inline UINT      WsmBaseOptionsFromTypeBits(UINT iDataType) { return (iDataType & 0x000003F0); }
inline UINT      WsmRemoveTypeOptions(UINT iDataType) { return (iDataType & 0xFFF0FC00); }
inline bool WsmIsReference(UINT idt)  { return ((WsmTypeFromTypeBits(idt) == wsmdtString) && (idt & wsmdoStringReference)) || ((WsmTypeFromTypeBits(idt) == wsmdtBinary) && (idt & wsmdoBinaryReference)) || ((WsmTypeFromTypeBits(idt) == wsmdtComposite) && (idt & wsmdoCompositeReference)); }
inline UINT WsmBitFlagIndexFromTypeBits(UINT iDataType) { return (iDataType & wsmdoBitFlagIndex) >> 4; }

#endif

// Options affecting conversion of state data to VARIANT for IWsmNamespace::GetVariant and PutVariant only
#define wsmvoBinaryAsBin64   0x0100   // Get/PutVariant only: use base64 encoding to transfer binary, instead of UI1 array
#define wsmvoBinaryAsHex     0x0200   // Get/PutVariant only: use hex encoding to transfer binary, instead of UI1 array
#define wsmvoTimestampAsInt  0x0400   // Get/PutVariant only: return wsmTimestamp as VT_UI4, instead of VT_DATE  
#define wsmvoRawDataBytes    0x1000   // Get/PutVariant only: raw data, strings as UTF-8, as VT_UI1 array
#define wsmvoScriptTypes     0x2000   // GetVariant only: convert VT_I8,VT_UI8,VT_UI4,VT_UI2 to other automation types
#define wsmvoListAsArray     0x4000   // GetVariant only: pass entire List item as two-dimentional array of name,value (not currently implemented)
#define wsmvoForceTypeMask   0x00FF   // PutVariant only: low byte use to specify an explicit WSM type (performance improvement)


//
// State contexts, associated with manifest Context, GUIDs, and state 
// file locations
//

typedef enum wsmscEnum
{
    wsmscMetadata = 0,  // compiled manifest metadata, no setting data accessed
    wsmscUser     = 1,  // Context='user',    per-user state
    wsmscMachine  = 2,  // Context='machine', machine state
    wsmscShared   = 3,  // Context='allUsers' shared user state
    wsmscSystem   = 4,  // Context='system'   reserved for system components
    wsmscArchive  = 5,  // state archive files
    wsmscLogFiles = 6,  // diagnostic log files
};

// URI property types for external resources
typedef enum wsmUriPropEnum {
        wsmUriPropVar = 0, // namespace property/variable
        wsmUriPropLoc = 1, // resolved shell folder or setup detection directory
        wsmUriPropEnv = 2, // environment variable
        wsmUriPropPol = 3, // policy client parameter access
        wsmUriPropReg = 4, // registry item
        wsmUriPropIni = 5, // INI file item
        wsmUriPropSys = 6, // system property
        wsmUriPropCom = 7, // COM object
    };

//
// Options for processing context guid & store folder
//

#define wsmcpQuery        0  // Query already initialized context guid/folder
#define wsmcpRegister     1  // Initialize context guid/folder if non-existent
#define wsmcpUnregister   2  // Remove context guid/folder


//
// Commit options used in Namespace, Handler & Store
//

#define wsmcoCommit      0x0000    // (default) commit pending changes
#define wsmcoRollback    0x0001    // rollback changes, else commit
#define wsmcoReload      0x0002    // synchronize external changes
#define wsmcoForce       0x0004    // force a transaction record, even if no 
                                       // pending changes
#define wsmcoFullCommit  0x0010    // Commit full namespace i.e. all the
                                       // isolated child store/dirty packed, blob data
#define wsmcoReserved    0x0020    // Reserved flag, used internally

#define wsmcoAllowed     (wsmcoRollback + wsmcoReload + wsmcoForce + wsmcoFullCommit + wsmcoReserved)

#endif // _WSMDEF_H_
