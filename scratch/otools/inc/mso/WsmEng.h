/****************************************************************************
*                                                                           *
* Copyright (c) 1999 - 2001, Microsoft Corporation.  All rights reserved.   *
*                                                                           *
* File       - WsmEng.h                                                     *
*                                                                           *
* Purpose    - Contains Windows State Management public state engine        *
*              definitions                                                  *
*                                                                           *
* History    -     Date        Alias         Comment                        *
*              09-Jul-2001     State         Created                        *
*                                                                           *
****************************************************************************/

#ifndef _WSMENG_H_
#define _WSMENG_H_
#ifndef RC_INVOKED

#include "WsmDef.h"
#include "WsmMeta.h"

//____________________________________________________________________________
//
// Enum and bit flag definitions used in different Api's
//____________________________________________________________________________

//
// Namespace open mode
//

typedef unsigned int wsmonEnum; // Definition of open mode for NS
#define wsmonOpenNamespace  wsmonEnum

#define wsmonReadOnly      0x00000001 // data is read only, PutValue fails
#define wsmonLocalOnly     0x00000002 // data can change locally, commit fails
#define wsmonReadWrite     0x00000004 // data can be read, written, committed
#define wsmonFullIsolated  0x00000100 // Ignore isolated view of nested NS
#define wsmonManaged       0x00000200 // data is accessed through the managed store, not the lagecy store
#define wsmonNoReconcile   0x00000400 // no timestamp reconciliation happens with PutValue

//
// Manifest attributes for settings or members of composites, used in 
// GetInfo method
//

typedef enum wsmiiEnum
{
    //
    // metadata attributes for setting
    //

    wsmiiDatatype         = 1,  // datatype code and attributes
    wsmiiNullable         = 2,  // setting accepts an explicit Null value
    wsmiiContext          = 3,  // storage context, one of enum wsmscEnum
    wsmiiUpdate           = 4,  // write permissions, one of enum wsmupEnum
    wsmiiDescription      = 5,  // descriptions from manifest
    wsmiiLegacyName       = 6,  // manifest-defined legacy name for setting
    wsmiiQualifier        = 7,  // qualifying property for setting
    wsmiiValidation       = 8,  // name of validation expression, if present
    wsmiiForceValidation  = 9,  // validation cannot be suppresed
    wsmiiReconcilableUnit = 10, // True if item is stored as an atomic unit
    wsmiiRoam             = 11, // setting is roamed between machines
    wsmiiVolatile         = 12, // setting not merged from backup or roaming  
    wsmiiNotify           = 13, // setting marked for notification if changed
    wsmiiDefault          = 14, // default value for setting
    wsmiiCustomType       = 15, // name of the composite 
    wsmiiUriName          = 16, // URI form of the setting name

    //
    // store attributes for setting stored in managed state store
    //

    wsmiiLastTransactId  = 0x020, // transaction ID of specified item, UINT
    wsmiiTimestamp       = 0x021, // timestamp of current value, UINT
    wsmiiParentInstance  = 0x022, // returned parent instance name as stored 
                                  // in managed store

    //
    // store attributes used by store handler only, not available via 
    // IWsmNamespace API except by properties
    //

    wsmiiSyncTimeStamp    = 0x100, // timestamp of last transaction of current 
                                   //     type (no item name)
    wsmiiSyncTransactId   = 0x101, // transaction ID of last transaction of 
                                  //     current type (no item name)
    wsmiiLastSyncData     = 0x102, // data for the specified item on or before 
                                   //    last transaction of current type (no 
                                   //    item name for transaction data)
    wsmiiLastSyncView     = 0x103, // Cache the last sync view of current
                                   //    transact mode (Legacy & Remote sync)
    wsmiiIsInLastSyncView = 0x104  // determine whether setting is within last 
                                   //    sync view (Legacy & Remote sync)
} wsmiiEnum;

//
// Update permissions for settings
//

typedef enum wsmupEnum
{
    wsmupConst   = 0,   // default value used, store not accessed
    wsmupAdmin   = 1,   // writable only with administrators privilege
    wsmupProgram = 2,   // writable only by owning program or with 
                        //     administrators privilege
    wsmupUser    = 3,   // writable by anyone with write permission to store
} wsmupEnum;

typedef enum wsmNotifyEnum
{
    wsmNotifyDefault      = 0,
    wsmNotifyNoChangeEnum = 1,  // don't create change enumeration object
} wsmNotifyEnum;

//
// Execution detection flags - only one type bit per criterion
//
#define wsmedAssembly       0x00000001 // assembly name 
#define wsmedComponentId    0x00000002 // MSI component ID
#define wsmedTaskName       0x00000004 // EXE name
#define wsmedRunningObject  0x00000008 // moniker for ROT
#define wsmedRequired       0x40000000 // required criterion

//
// Setup detection flags - only one type bit per criterion
//
#define wsmsdAssembly               0x00000001 // assembly name
#define wsmsdComponentId            0x00000002 // MSI component ID
#define wsmsdFileExists             0x00000010 // file path with virtual folder root
#define wsmsdFolderExists           0x00000020 // folder path with virtual folder root
#define wsmsdRegistryFile           0x00000040 // registry entry pointing to file
#define wsmsdRegistryFolder         0x00000080 // registry entry pointing to folder
// Version detection flags
#define wsmsdComments               0x00010000 // version info: comments
#define wsmsdCompanyName            0x00020000 // version info: company name
#define wsmsdFileDescription        0x00040000 // version info: file description
#define wsmsdFileVersion            0x00080000 // version info: file version
#define wsmsdInternalName           0x00100000 // version info: internal name
#define wsmsdLegalCopyright         0x00200000 // version info: legal copyright
#define wsmsdLegalTrademarks        0x00400000 // version info: legal trademarks
#define wsmsdOriginalFilename       0x00800000 // version info: original file name
#define wsmsdProductName            0x01000000 // version info: product name
#define wsmsdProductVersion         0x02000000 // version info: product version
#define wsmsdPrivateBuild           0x04000000 // version info: private build
#define wsmsdSpecialBuild           0x08000000 // version info: special build
#define wsmsdVersionInfoMask        0x0fff0000 // version info group mask

#define wsmsdRequired               0x40000000 // required criterion

//
// Maintaining detection flags - only one type bit per criterion
//
#define wsmmdLegacySync     0x00000001 // Legacy sync
#define wsmmdRemoteSync     0x00000002 // Roaming
#define wsmmdCompaction     0x00000004 // Store is compacting
#define wsmmdWait           0x40000000 // Wait for the maintaining to be finished

//
// Data processing options passed only to settings manager through the 
// external API, used by Get methods only (low order bits used for info 
// request, wsmiiEnum options
//

// don't resolve relocatable string paths
#define wsmpoNoPathResolve  0x00010000

// don't expand environment variables within strings
#define wsmpoNoEnvironment  0x00020000

// don't return default values if data not present in store
#define wsmpoNoDefault      0x00040000

// return default value only, S_FALSE if no default
#define wsmpoDefaultOnly    0x00080000

// don't validate, except if ForceValidation attribute set
#define wsmpoNoValidation   0x00100000

// don't store data changes where new value same as current
#define wsmpoOnlyIfChanged  0x00200000

// determine is setting is writable to the store
#define wsmpoTestWritable   0x00400000

// perform validation but don't write value
#define wsmpoValidateOnly   0x00800000

// mask process option bits affecting pre/post-processing by group
#define wsmpoProcessMask    0x00FF0000


//____________________________________________________________________________
//
// Standard state engine environment property names, to be used in expressions
//____________________________________________________________________________

#define WSMEP_UserLangId   L"UserLangId"
#define WSMEP_SystemLangId L"SystemLangId"
#define WSMEP_ScreenWidth  L"ScreenWidth"
#define WSMEP_ScreenHeight L"ScreenHeight"
#define WSMEP_ColorBits    L"ColorBits"
#define WSMEP_Unicode      L"Unicode"  
#define WSMEP_LogonUser    L"LogonUser"
#define WSMEP_ComputerName L"ComputerName"
#define WSMEP_Platform     L"Platform"  
#define WSMEP_OSVersion    L"OSVersion"
#define WSMEP_ProfileName  L"ProfileName" 

//
// Dynammic namespace properties obtained from the managed state store
//

// timestamp of last transaction of current type
#define WSMNP_TransactTimestamp "TransactTimestamp"

// transaction ID of last transaction of current type
#define WSMNP_TransactId        "TransactId"

// latest transaction ID of for synch enumerator
#define WSMNP_LastTransactId    "LastTransactId"

// Commit data for the last synch transaction
#define WSMNP_LastTransactData  "LastTransactData"



//
// Length value to indicate null-terminated string
//

#ifdef __cplusplus
// used for PutString and utility functions
const UINT wsmNullTerminated = UINT(-1);
#else
#define wsmNullTerminated ((UINT)(-1));
#endif



#ifdef __cplusplus

#ifndef DLLVER_PLATFORM_NT   // defined in shlwapi.h
typedef struct _DllVersionInfo
{
        DWORD cbSize;
        DWORD dwMajorVersion;                   // Major version
        DWORD dwMinorVersion;                   // Minor version
        DWORD dwBuildNumber;                    // Build number
        DWORD dwPlatformID;                     // DLLVER_PLATFORM_*
} DLLVERSIONINFO;
#define DLLVER_PLATFORM_WINDOWS         0x00000001      // Windows 95
#define DLLVER_PLATFORM_NT              0x00000002      // Windows NT
#endif


//____________________________________________________________________________
//
// Application-level interfaces to state engine
//____________________________________________________________________________

//
// Forward reference to class that manipulates a reference string
//

class IWsmNamespace;

class IWsmChangeEnum : public IUnknown
{
public:

    //
    // GetNamespace -  Returns the namespace for which the changes can 
    //                 be enumerated by this interface. This will be a 
    //                 new instance of the namespace constructed for 
    //                 notification purposes.
    //  ppNamespace -  returne namespace pointer, must release after use
    //

    virtual HRESULT __stdcall GetNamespace (
        OUT IWsmNamespace **ppNamespace) = 0;

    //
    // NextChange   -  This method is used to enumerate for the settings that 
    //                 have changed in this namespace, it can also return the 
    //                 child IWsmChangeEnum Interface for Sub-namespaces.
    // ppwszSetting -  return changed setting name, don't release
    // ppChgEnum    -  return IWsmChangeEnum interface if setting ppwszSetting
    //                 is a sub namespace, must release after use

    virtual HRESULT __stdcall NextChange (
        OUT WCHAR**           ppwszSetting,
        OUT IWsmChangeEnum**  ppChgEnum) = 0;

    //
    // Reset        -  This method is used to reset the change enumerator
    //

    virtual HRESULT __stdcall Reset() = 0;
};

//
// Notification callback, passed to OpenNamespace, specific to that namespace
//

class IWsmNotify : public IUnknown
{
public:

    //
    // NamespaceChanged -  Callback api implemented by client, used to notify
    //                     client for changes in namespace
    // pChgEnum         -  namespace change enumerator, must release after use
    //

    virtual HRESULT __stdcall NamespaceChanged (
        IN IWsmChangeEnum *pChgEnum) = 0;
};


//
// State namespace application interface
//

class IWsmNamespace : public IUnknown
{
public:

    // 
    // GetVariant  -  Returns the value of the item in a variant
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   wsmvoOptions -  variant conversion options
    //   pvData       -  output variant, must do VariantClear after use
    //

    virtual HRESULT __stdcall GetVariant(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  UINT         wsmvoOptions,
        OUT VARIANT*     pvData) = 0;

    // 
    // PutVariant  -  Puts the value of the item from a variant
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   wsmvoOptions -  variant conversion options
    //   pvData       -  input variant value
    //

    virtual HRESULT __stdcall PutVariant(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  UINT         wsmvoOptions,
        IN  VARIANT*     pvData) = 0;

    // 
    // GetNamespace   -  Returns the sub namespace pointer
    //   szItem       -  name of sub namespace 
    //   uiReserved   -  reserved, must be 0
    //   ppNamespace  -  return sub namespace pointer, must release after use
    //

    virtual HRESULT __stdcall GetNamespace(
        IN  const WCHAR*    szItem,
        IN  UINT            uiReserved,
        OUT IWsmNamespace** ppNamespace) = 0;

    // 
    // GetBool  -  Returns the value of boolean item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pfData       -  output boolean data
    //

    virtual HRESULT __stdcall GetBool(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        OUT bool*        pfData) = 0;

    // 
    // PutBool  -  Puts the boolean value of the item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   fData        -  input boolean value
    //

    virtual HRESULT __stdcall PutBool(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  bool         fData) = 0;

    // 
    // GetInteger  -  Returns the value of integer item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pfData       -  output integer data
    //

    virtual HRESULT __stdcall GetInteger(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        OUT int*         piData) = 0;

    // 
    // PutInteger  -  Puts the integer value of the item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   fData        -  input integer value
    //

    virtual HRESULT __stdcall PutInteger(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  int          iData) = 0;

    // 
    // GetLargeInteger  -  Returns the value of large integer item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pliData      -  output large integer data
    //

    virtual HRESULT __stdcall GetLargeInteger(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        OUT LONGLONG*    pliData) = 0;

    // 
    // PutLargeInteger  -  Puts the large integer value of the item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   fData        -  input large integer value
    //

    virtual HRESULT __stdcall PutLargeInteger(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  LONGLONG     liData) = 0;

    // 
    // GetDouble  -  Returns the double value of item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pliData      -  output double value
    //

    virtual HRESULT __stdcall GetDouble(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        OUT double*      pdblData) = 0;

    // 
    // PutDouble  -  Puts the double value of the item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   fData        -  input double value
    //

    virtual HRESULT __stdcall PutDouble(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  double       dblData) = 0;

    // 
    // GetString  -  Returns the string value of item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pchBuf       -  output buffer to hold string, owned by caller
    //   cchBuf       -  output buffer size
    //   pcchData     -  size of output string
    //

    virtual HRESULT __stdcall GetString(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        OUT WCHAR*       pchBuf,
        IN  UINT         cchBuf,
        OUT UINT*        pcchData) = 0;

    // 
    // PutString  -  Puts the string value of item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   szData       -  input string buffer
    //   cchData      -  input string buffer size
    //

    virtual HRESULT __stdcall PutString(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  const WCHAR* szData,
        IN  UINT         cchData) = 0;

    // 
    // GetBinary  -  Returns the binary value of item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pbBuf        -  output buffer to hold binary data, owned by caller
    //   cbBuf        -  output buffer size
    //   pcbData      -  size of output binary data
    //

    virtual HRESULT __stdcall GetBinary(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        OUT BYTE*        pbBuf,
        IN  UINT         cbBuf,
        OUT UINT*        pcbData) = 0;

    // 
    // PutBinary  -  Puts the binary value of item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //   pbData       -  input binary data
    //   cbData       -  input binary data size
    //

    virtual HRESULT __stdcall PutBinary(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions,
        IN  const BYTE*  pbData,
        IN  UINT         cbData) = 0;

    // 
    // PutNull  -  Puts a null value for item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //

    virtual HRESULT __stdcall PutNull(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions) = 0;

    // 
    // RemoveItem  -  Deletes the item
    //   szItem       -  item name 
    //   iDataOptions -  item options, wsmpo* + wsmii*
    //

    virtual HRESULT __stdcall RemoveItem(
        IN  const WCHAR* szItem,
        IN  UINT         iDataOptions) = 0;

    // 
    // Commit  -  Commit the dirty data in root namespace or isolated 
    //            sub namespace, no op for any other sub namespace
    //   wsmcoOptions   -  commit options 
    //   szTaskInfo     -  commit data written with transaction
    //

    virtual HRESULT __stdcall Commit(
        IN  UINT         wsmcoOptions,
        IN  const WCHAR* szTaskInfo) = 0;

    //
    // NextItem   -  Retunrs next item in the namespace
    //   pchBuf   -  output buffer to hold the item name, owned by caller
    //   cchBuf   -  size of output buffer
    //   pcchItem -  size of item name
    //

    virtual HRESULT __stdcall NextItem(
        OUT WCHAR* pchBuf,
        IN  UINT   cchBuf,
        OUT UINT*  pcchItem) = 0;

    //
    // Reset   - Reset the namespace item name enumerator
    //

    virtual HRESULT __stdcall Reset()=0;

    //
    // RegisterNotification  -  Registers notification for setting in 
    //                          current namespace
    //   pwszSetting    -  particular setting name or filter, NULL to receive 
    //                     notification for any setting change under current 
    //                     namespace
    //   pNotify        -  notification callback interface
    //   wsmNotifyFlags -  notification filter flags
    //

    virtual HRESULT __stdcall RegisterNotification (
        IN  const WCHAR*   pwszSetting,
        IN  IWsmNotify*    pNotify,
        IN  UINT           wsmNotifyFlags) = 0;

    //
    // UnregisterNotification  -  Unregisters notification for setting in 
    //                            current namespace
    //   pwszSetting    -  particular setting name or filter 
    //   pNotify        -  notification callback interface
    //

    virtual HRESULT __stdcall UnRegisterNotification (
        IN  const WCHAR*  pwszSetting,
        IN  IWsmNotify*   pNotify) = 0;

};


//
// Namespace Enumerator interface
//

class IWsmNamespaceEnum : public IUnknown
{
public:

    //
    // NextNamespace - return the next namespace 
    //     pszuriNamespace - Out parameter to receive namespace URI
    //     piNamespaceVers - Returned namespace version :
    //                                  (Major << 8) + Minor version
    //     piNamespaceLang - Returned namespace LangId, 0 if language neutral

    virtual HRESULT __stdcall NextNamespace(
        OUT WCHAR** pszuriNamespace,
        OUT USHORT* piNamespaceVers,
        OUT LANGID* piNamespaceLang) = 0;

    //
    // Reset - Reset the namespace enumerator
    //

    virtual HRESULT __stdcall Reset() = 0;
};


//
//  Manifest compiler flags
//

#define WSMCOMPILER_NOINDEX             0x0001
#define WSMCOMPILER_DELETEMETA          0x0002


//
//  Manifest compiler exported error context
//

class IWsmCompilerError : public IUnknown
{
public:

    //
    //  GetErrorCode - get the HRESULT reported by the compiler
    //

    virtual HRESULT __stdcall GetErrorCode(
        OUT HRESULT*  pHr) = 0;


    // 
    //  GetErrorLine - Get the line, and position within that line,
    //      where the error occured
    //

    virtual HRESULT __stdcall GetErrorLine(
        OUT DWORD* pdwLine, 
        OUT DWORD* pdwPosition) = 0;


    //
    //  GetErrorNodeName - Returns the name of the XML node in which
    //      the error occured
    //

    virtual HRESULT __stdcall GetErrorNodeName(
        OUT BSTR* pbstrNodeName) = 0;

    // 
    //  GetErrorNodeValue - Returns the value of the node in which the
    //      error occured
    //

    virtual HRESULT __stdcall GetErrorNodeValue(
        OUT BSTR* pbstrNodeValue) = 0;

    //
    //  GetErrorMessage - formats an error message based on the error
    //      record and returns that string.
    //

    virtual HRESULT __stdcall GetErrorMessage(
        OUT BSTR* pbstrMessage) = 0;
};


//
//  Manifest compiler interface
//

class IWsmManifestCompiler : public IUnknown
{
public:

    //
    //  CompileStream - Takes an IStream to an XML manifest and
    //      compiles it to meta data
    //

    virtual HRESULT __stdcall CompileStream(
        IN IStream*  pInput,
        IN UINT      uiCompileOptions) = 0;


    //
    //  CompileFile - Takes the filename of an XML manifest and
    //      compiles it to meta data
    //

    virtual HRESULT __stdcall CompileFile(
        IN const WCHAR*  pInput,
        IN UINT          uiCompileOptions) = 0;


    //
    //  ResumeCompilation - resumes compilation after CompileStream()
    //      reports an error.
    //

    virtual HRESULT __stdcall ResumeCompilation() = 0;


    //
    //  GetErrorContext - returns a WsmCompilerError object after
    //      CompileStream or ResumeCompilation fails
    //

    virtual HRESULT __stdcall GetErrorContext(
        OUT IWsmCompilerError**  ppContext) = 0;
};


//
// Global state engine access, created by class factory or by exported 
// API function
//

class IWsmEngine : public IUnknown 
{
public:

    //
    // GetVersion - returns version info for state engine
    //     pverInfo - non-null location for returned version info
    //

    virtual HRESULT __stdcall GetVersion(
        OUT DLLVERSIONINFO* pverInfo) = 0;

    //
    // GetLanguage - returns default language for state engine
    //     pLangId - non-null location for return value of language ID
    //

    virtual HRESULT __stdcall GetLanguage(
        OUT LANGID* pLangId) = 0;

    //
    // SetLanguage - sets default language for state engine
    //     LangId - Language ID to set
    //

    virtual HRESULT __stdcall SetLanguage(
        IN  LANGID LangId) = 0;

    //
    // CreateNamespaceEnumerator - return list of namespaces present in context
    //     wsmscContext    - Context for enumeration
    //     uiReserved      - Reserved field, must be 0
    //     ppNamespaceEnum - Returned enumerator, must release when done
    //

    virtual HRESULT __stdcall CreateNamespaceEnumerator(
        IN  wsmscEnum           wsmscContext,
        IN  UINT                uiReserved,
        OUT IWsmNamespaceEnum** ppNamespaceEnum) = 0;

    //
    // OpenNamespcae - Open the namespace in specified mode
    //     szuriNamespace   - namespace URI for settings store
    //     iVersion         - manifest version, major*256 + minor, zero to 
    //                        use latest version
    //     iLangId          - manifest language, best match used if iVersion 
    //                        is zero
    //     wsmonMode        - read/write mode
    //     pNotify          - non-null location for returned pointer to opened 
    //                        namespace
    //     ppiRootNamespace - non-null location for returned pointer to opened
    //                        namespace
    //

    virtual HRESULT __stdcall OpenNamespace(
        IN  const WCHAR* szuriNamespace,
        IN  USHORT iVersion,
        IN  LANGID iLangId,
        IN  wsmonEnum wsmonMode,
        OUT IWsmNamespace** ppiRootNamespace) = 0;

    //
    // SynchronizeNamespace - Synchronize the namespace settings between 
    //                        lagacy and managed store
    //     szuriNamespace - namespace URI for settings store
    //     iVersion       - manifest version, major*256 + minor, zero to use 
    //                      latest version
    //     iLangId        - manifest language, best match used if iVersion is 
    //                      zero
    //

    virtual HRESULT __stdcall SynchronizeNamespace(
        IN  const WCHAR* szuriNamespace,
        IN  USHORT iVersion,
        IN  LANGID iLangId) = 0;


    //
    // GetManifestCompiler - Retrieves an instance of the manifest compiler
    //     ppCompiler     - [Output] Manifest compiler instance
    //

    virtual HRESULT __stdcall GetManifestCompiler(
        OUT IWsmManifestCompiler** ppCompiler) = 0;


    //
    // IsRunning - Detect the execution of the owner of a specific namespace
    //     szuriNamespace - namespace URI for execution detection
    //     iVersion       - manifest version, major*256 + minor, zero to use 
    //                      latest version
    //     iLangId        - manifest language, best match used if iVersion is 
    //                      zero
    //     wsmedFlags     - flags for different type of detection
    //

    virtual HRESULT __stdcall IsRunning(
        IN  const WCHAR* szuriNamespace,
        IN  USHORT iVersion,
        IN  LANGID iLangId,
        IN  UINT wsmedFlags) = 0;

    //
    // IsInstalled - Detect if the owner of a specific namespace is installed
    //     szuriNamespace - namespace URI for execution detection
    //     iVersion       - manifest version, major*256 + minor, zero to use 
    //                      latest version
    //     iLangId        - manifest language, best match used if iVersion is 
    //                      zero
    //     wsmsdFlags     - flags for different type of detection
    //

    virtual HRESULT __stdcall IsInstalled(
        IN  const WCHAR* szuriNamespace,
        IN  USHORT iVersion,
        IN  LANGID iLangId,
        IN  UINT wsmsdFlags) = 0;

    //
    // IsMaintaining - Detect if the namespace is doing some maintaining/synchronizing work
    //                 such as legacy sync or store compaction
    //     szuriNamespace - namespace URI for maintaining detection
    //     iVersion       - manifest version, major*256 + minor, zero to use 
    //                      latest version
    //     iLangId        - manifest language, best match used if iVersion is 
    //                      zero
    //     wsmmdFlags     - flags for different type of detection
    //

    virtual HRESULT __stdcall IsMaintaining(
        IN  const WCHAR* szuriNamespace,
        IN  USHORT iVersion,
        IN  LANGID iLangId,
        IN  UINT wsmmdFlags) = 0;

};


//____________________________________________________________________________
//
// Global functions to create base objects, also available from class factory
//____________________________________________________________________________

extern "C"
{
#ifndef _WSMIMP
#define _WSMIMP __declspec(dllimport)
#endif
_WSMIMP HRESULT __stdcall WsmEngineCreate(OUT IWsmEngine** ppWsmEngine);
} // extern "C"

//
// Interface IDs exported for state engine
//

extern "C" const  GUID      IID_IWsmEngine;
extern "C" const  GUID      IID_IWsmNamespace;
extern "C" const  GUID      IID_IWsmNamespaceEnum;
extern "C" const  GUID      IID_IWsmNotify;
extern "C" const  GUID      IID_IWsmChangeEnum;

#endif // __cplusplus


//____________________________________________________________________________
//
// Handle-based wrapper functions for use with C programs
// Handles are validated at each used, and must be explicitly closed
// Can be used with C++, but the performance is slower than with the classes
// Error returns are identical to those for the C++ wrapper classes
//____________________________________________________________________________

#ifdef __cplusplus
struct WsmEngineWrapper;
struct WsmNamespaceWrapper;
struct WsmNamespaceEnumWrapper;
extern "C" {
#else // !__cplusplus - cannot forward ref class for C compilation
typedef struct WsmEngineWrapper {int m_;} WsmEngineWrapper;
typedef struct WsmNamespaceWrapper {int m_;} WsmNamespaceWrapper;
typedef struct WsmNamespaceEnumWrapper { int m_;} WsmNamespaceEnumWrapper;
#endif

typedef WsmEngineWrapper* WsmEngineHandle;
typedef WsmNamespaceWrapper* WsmNamespaceHandle;
typedef WsmNamespaceEnumWrapper* WsmNamespaceEnumHandle;

//
// Callback function used for wrapper classes/functions
//

typedef HRESULT (__stdcall *wsmfChangeNotify)(IN const WCHAR* szSetting);


//
// Wrapper functions for use with C programs
//

// returns a engine handle that used in subsequent calls
HRESULT WsmEngineOpen(
    OUT WsmEngineHandle* phwsmEngine);

// returns default language for state engine
HRESULT WsmEngineGetLanguage(
    IN  WsmEngineHandle hwsmEngine,
    OUT LANGID*         pLangId);

// sets default language for state engine
HRESULT WsmEngineSetLanguage(
    IN  WsmEngineHandle hwsmEngine,
    IN  LANGID          LangId);

// close the engine handle
HRESULT WsmEngineClose(
    IN  WsmEngineHandle hwsmEngine);

// Namespace enumerator for a given context
HRESULT WsmEngineNamespaceEnum(
    IN  WsmEngineHandle         hwsmEngine,
    IN  wsmscEnum               wsmscContext,
    IN  UINT                    uiReserved,
    OUT WsmNamespaceEnumHandle* phwsmNamespaceEnum);

// Open the namespace handle in specified mode
HRESULT WsmNamespaceOpen(
    IN  WsmEngineHandle     hwsmEngine,
    IN  const WCHAR*        szuriNamespace,
    IN  USHORT              iVersion,
    IN  LANGID              iLangId,
    IN  wsmonEnum           wsmonMode,
    IN  const void*         pvNotifyContext,
    IN  wsmfChangeNotify    fChangeNotify,
    OUT WsmNamespaceHandle* phwsmNamespace);

// Synchronize the namespace settings between lagacy and managed store
HRESULT WsmNamespaceSynchronize(
    IN  WsmEngineHandle hwsmEngine,
    IN  const WCHAR*    szuriNamespace,
    IN  USHORT          iVersion,
    IN  LANGID          iLangId);

// Detect the execution of the owner of a specific namespace
HRESULT WsmNamespaceIsRunning(
    IN  WsmEngineHandle hwsmEngine,
    IN  const WCHAR*    szuriNamespace,
    IN  USHORT          iVersion,
    IN  LANGID          iLangId,
    IN  UINT            wsmedFlags);

// Detect if the owner of a specific namespace is installed
HRESULT WsmNamespaceIsInstalled(
    IN  WsmEngineHandle hwsmEngine,
    IN  const WCHAR*    szuriNamespace,
    IN  USHORT          iVersion,
    IN  LANGID          iLangId,
    IN  UINT            wsmsdFlags);

// Detect if the specific namespace is doing some maintaining work
HRESULT WsmNamespaceIsMaintaining(
    IN  WsmEngineHandle hwsmEngine,
    IN  const WCHAR*    szuriNamespace,
    IN  USHORT          iVersion,
    IN  LANGID          iLangId,
    IN  UINT            wsmmdFlags);

// Reset the namespace enumerator handle
HRESULT WsmNamespaceEnumReset(
    IN  WsmNamespaceEnumHandle hwsmNamespaceEnum);

// Close the namespace enumerator handle
HRESULT WsmNamespaceEnumClose(
    IN  WsmNamespaceEnumHandle hwsmNamespaceEnum);

// close the namespace handle
HRESULT WsmNamespaceClose(
    IN  WsmNamespaceHandle hwsmNamespace);

HRESULT WsmNamespaceGetInteger(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    OUT int*               piData);

HRESULT WsmNamespacePutInteger(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  int                iData);

HRESULT WsmNamespaceGetLargeInteger(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    OUT LONGLONG*          pliData);

HRESULT WsmNamespacePutLargeInteger(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  LONGLONG           liData);

HRESULT WsmNamespaceGetBool(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    OUT BOOL*              pfData);

HRESULT WsmNamespacePutBool(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  BOOL               fData);

HRESULT WsmNamespaceGetString(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    OUT WCHAR*             pchBuf,
    IN  UINT               cchBuf,
    OUT UINT*              pcchData);

HRESULT WsmNamespacePutString(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  const WCHAR*       szData,
    IN  UINT               cchData);

HRESULT WsmNamespacePutBinary(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  const BYTE*        pbData,
    IN  UINT               cbData);

HRESULT WsmNamespaceGetDouble(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    OUT double*            pdblData);

HRESULT WsmNamespacePutDouble(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  double             dblData);

HRESULT WsmNamespaceGetNamespace(
    IN  WsmNamespaceHandle  hwsmNamespace,
    IN  const WCHAR*        szItem,
    OUT WsmNamespaceHandle* phwsmNamespace);

HRESULT WsmNamespaceGetVariant(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  UINT               wsmvoOptions,
    OUT VARIANT*           pvData);

HRESULT WsmNamespacePutVariant(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions,
    IN  UINT               wsmvoOptions,
    IN  VARIANT*           pvData);

HRESULT WsmNamespacePutNull(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions);

HRESULT WsmNamespaceRemoveItem(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  const WCHAR*       szItem,
    IN  UINT               iDataOptions);

HRESULT WsmNamespaceReset(
    IN  WsmNamespaceHandle hwsmNamespace);

HRESULT WsmNamespaceCommit(
    IN  WsmNamespaceHandle hwsmNamespace,
    IN  UINT               wsmcoOptions,
    IN  const WCHAR*       szTaskInfo);

#ifdef __cplusplus
} // extern "C"
#endif // __cplusplus

#endif // !RC_INVOKED

#endif // _WSMENG_H_
