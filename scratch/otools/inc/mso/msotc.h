/****************************************************************************
	Msotc.h

	Owner: EricSchr
 	Copyright (c) 1997 Microsoft Corporation

	Declarations for functions related to TCO
****************************************************************************/

#pragma once

#ifndef MSOTC_H
#define MSOTC_H 1

enum
{
	msoNoPathPreference	= 0,
	msoPathPreferLetter	= 1,
	msoPathPreferUNC	= 2
};


/*---------------------------------------------------------------------------
	MsoAppendToPath

	Append a string, ensuring that there's the proper slash in between.
------------------------------------------------------------------ JJames -*/
MSOAPI_(WCHAR *) MsoAppendToPath(const WCHAR *wzSub, _Inout_z_cap_(cchMax) WCHAR *wzPath, int cchMax);

/*-----------------------------------------------------------------------------
	MsoWzAfterPath

	Return a pointer after the last backslash, or the start if there is none.
------------------------------------------------------------------ JJames ---*/
MSOAPI_(WCHAR *) MsoWzAfterPath(const WCHAR *wzPathName);

/*-----------------------------------------------------------------------------
   MsoSzAfterPathCch

   Return a pointer after the last backslash, or the start if there is none.
   Since we already know the cch, we can traverse the string backwards which
   will most likely be faster.
--------------------------------------------------------------------dgray----*/
MSOAPIX_(const char *) MsoSzAfterPathCch(const char *szPathName, DWORD cch);

/*-----------------------------------------------------------------------------
	MsoWzBeforeExt

	Return a pointer just before extension.
	Return NULL if there is no extension.
------------------------------------------------------------------ IgorZ ---*/
MSOAPI_(WCHAR *) MsoWzBeforeExt(const WCHAR *wzPathName);

/*-----------------------------------------------------------------------------
	MsoFFileExist

	Returns fTrue iff the file exists and is not a directory.
------------------------------------------------------------------ JJames ---*/
MSOAPI_(BOOL) MsoFFileExist(const WCHAR *wzFile);


/*---------------------------------------------------------------------------
	MsoFDirExist

	Return fTrue if and only if wzDir exists and is a directory.
---------------------------------------------------------------- EricSchr -*/
MSOAPI_(BOOL) MsoFDirExist(const WCHAR *wzDir);


/****************************************************************************
	Office layer over Darwin (GimmeFile and friends).

	* DO NOT EVER CALL DARWIN DIRECTLY. ALWAYS GO THROUGH OFFICE. *
	All MsoGimme* APIs are provided for overall performance, extra resiliency,
	and for your convenience. Gimme(TM) is a trademark of KirkG.

	The structures below add up to a MSOTCFCF structure which you have
	to give to IMsoUser::FHookDarwinTables.

	This structure holds tables which describe to MsoFGimmeFile and
	associated APIs the structure of files, components, and features, as
	authored in Setup. From the app's point of view, instead of accessing
	files by name, you access them by fid (file id). Some files are
	known by Office and the id is built in (see msotcdar.h). Others used
	only by your app need to be hooked up to Office so that the Office
	code can apply the same resiliency rules as with its own files.

	The tcmsi.exe tool (in otools) takes as input files which describe the
	Setup database layout, and output two files:
	- a header, which lists fid's local to your app (and cid's for
		components, and ftid's for features).
	- a C file, which contains the MSOTCFCF structure for your app.

	Note that Office itself uses this tool to generate msotcdar.h (included
	below) and tcdar.inc (included in the bowels of Office code).

	Expect the internals of the MSOTCFCF structure to move around a lot
	as we figure out boot, string loading, etc.
****************************************************************************/


// this flag controls whether we build a single table with enums (1)
// or externs to individual structs per row (0)
#define MSOGIMME_INDEXIDS 1

typedef int msofidT;
typedef int msocidT;
typedef int msoftidT;

// must be in ssync with otcdarmake's output
#define msoidstcoGeneric 0

typedef enum
	{
	msotcidmin = 0,
	msofid = 0,
	msocid,
	msoftid,
	msoqcid,
	msoqfid,
	msotcidmax = msoqfid
	} MSOTCID;

typedef	enum { // default language
	msolangNone,
	msolangInstall,
	msolangUI,
	msolangHelp,
	msolangInstallFlavor,
	msolangPreviousUI,
	msolangPreviousInstallFlavor
	} msolangT;

typedef struct _msotcfileinfo {
	CHAR *szFilename;		// filename
	msocidT cid;			// component ID
} MSOTCFILEINFO;

typedef struct _msotccomponentinfo {
	GUID msoguid;
	CHAR *szKeyFile;		// name of keyfile or static qualifier
	msoftidT ftid;			// feature ID, -1 if belongs to multiple features
	int idsInstall;			// string id, msoidstcoGeneric if none
	int idsRepair;			// TODO(JBelt): delete (otcdarmake, ssync with VB)
	unsigned langDefault : 2;
	unsigned fLcidQualified : 1;
	unsigned fFilenameQualified : 1;
	unsigned fOtherQualified : 1;
	unsigned fStaticQualified : 1;
} MSOTCCOMPONENTINFO;

typedef struct _msotcfeatureinfo {
	CHAR *szFtid;			// GUID
	msocidT qcid;			// publish component for cross-product features, -1 if none
	int idsInstall;			// string id, msoidstcoGeneric if none
	int idsRepair;			// TODO(JBelt): delete (otcdarmake, ssync with VB)
} MSOTCFEATUREINFO;

typedef struct _msotcclassinfo {
	CLSID clsid;	// GUID constant
	DWORD dwClsCtx;	// Class Context
	msocidT cid;	// Feature ID;
} MSOTCCLASSINFO;

typedef struct _msotcbackdoorinfo {
	msocidT cid;
	void *pReserved;
	CHAR *szRelativePath;
	CHAR *szQualifier;
	CHAR *szAppData;
} MSOTCBACKDOORINFO;

#define MSOTCFILEINFO_NIL { "invalidFid", msocidNil }
#define MSOTCCOMPONENTINFO_NIL { {0,0,0,0}, "invalidCid", msoftidNil, -1 }
#define MSOTCFEATUREINFO_NIL { "invalidFtid", msoftidNil }
#define MSOTCCLASSINFO_NIL { {0,0,0,0}, 0, msocidNil }
#define MSOTCBACKDOORINFO_NIL { msocidNil, 0, "", "", "" }

typedef struct _msotcfcf {
	DWORD dwVersion;				// version
	int iTableIndex;
	const MSOTCFILEINFO *rgfi;		// file table
	const MSOTCCOMPONENTINFO *rgci;	// component table
	const MSOTCFEATUREINFO *rgfti;	// feature table
	const MSOTCBACKDOORINFO *rgbdi;	// backdoor data
	int cfi, cci, cfti, cbdi;		// counts
} MSOTCFCF;

/* Edit language info */
typedef struct _msoeli
	{
	UCHAR	fExplicit;
	LCID	lcid;
	}MSOELI;

// include file, component, and feature id's from Darwin
#if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
#include "msotcdar.h"
#else
#define msocidNil (-1)
#endif
// Maximum length of a GUID string in the standard format,
// "{000Cxxxx-0000-0000-C000-000000000046}"
#define MAX_GUID 39


/*---------------------------------------------------------------------------
	dwGimmeFlags

	These flags specify options for the MsoFGimme*Ex functions below.
	Each functions always checks the install state of the object.
------------------------------------------------------------------ JJames -*/
#define msotcogfDemandInstall			0x0001	// install if not already
#define msotcogfSearchForFile			0x0002	// check file system if darwin fails
#define msotcogfTryOtherLanguages		0x0004	// try backup lcid's if requested one fails
#define msotcogfVerifyFileExists		0x0008	// verify that the requested file exists
#define msotcogfFixIfNecessary			0x0010	// call darwin fix functions if necessary
#define msotcogfForceFix				0x0020	// call darwin fix functions
#define msotcogfTrueIfAdvertised		0x0040	// return TRUE if the object is advertised
#define msotcogfForceFixMachineRegistry    0x0080	// force repair of user registry data
#define msotcogfNoInstallUI				0x0100	// no install confirmation UI (assume Yes)
#define msotcogfNoRepairUI				0x0200	// no repair confirmation UI (assume Yes)
#define msotcogfNoRetryUI				0x0400	// no retry on busy UI (assume Cancel)
#define msotcogfNoDisabledUI			0x0800	// no disabled feature UI (assume Ok)
#define msotcogfValidate				0x1000  // call MsiUseFeature instead of QueryFeature
#define msotcogfUninstall				0x2000  // change feature to advertised
#define msotcogfForceFixUserRegistry	0x4000  // force repair of machine registry data
#define msotcogfSearchFirst             0x8000  // check file system before querying Darwin (boot perf.)
#define msotcogfNoSourceDialog		   	0x10000  // suppress Darwin's source dialog
#define msotcogfNoCustomUI			   	0x20000  // Disable Office demandinstallUI dialog
#define msotcogfNoAutoApprove		   	0x40000  // Don't say "Yes" automatically
#define msotcogfNoAutoReject		   	0x80000  // Don't say "No" automatically
#define msotcogfForceUI				   0x100000  // Ignore the app callback
#define msotcogfJustThisProduct		   0x200000  // only query this product (qualfied only)
// reserved for Gimme internal     		0xF0000000

// preserve ForceFixRegistry option
#define msotcogfForceFixRegistry (msotcogfForceFixUserRegistry | msotcogfForceFixMachineRegistry)

#define msotcogfResiliency (msotcogfSearchForFile | msotcogfTryOtherLanguages)

// test whether the associated feature is enabled for install
#define msotcogfEnabled (msotcogfTrueIfAdvertised | msotcogfResiliency)

// test whether the file or component is already installed on the machine
#define msotcogfInstalled (msotcogfResiliency)
#define msotcogfInstalledNoResiliency (0)

// request the file or component, installing if necessary
#define msotcogfProvide (msotcogfDemandInstall | msotcogfResiliency)

// do whatever it takes to get the file or component
#define msotcogfRequired (msotcogfDemandInstall | msotcogfResiliency | msotcogfVerifyFileExists | msotcogfFixIfNecessary)
#define msotcogfRequiredNoResiliency (msotcogfDemandInstall | msotcogfVerifyFileExists | msotcogfFixIfNecessary)

// don't display any UI
// TODO(JBelt): merge all UI flags into one?
#define msotcogfQuiet (msotcogfNoInstallUI | msotcogfNoRepairUI | msotcogfNoRetryUI | msotcogfNoDisabledUI)

// don't automatically approve or deny an install or repair
#define msotcogfNoAutoReponse (msotcogfNoAutoApprove | msotcogfNoAutoReject)

/*--------------------------------------------------------------------------
	Gimme table indexes for each gimme client, moved from tcutil.h
-------------------------------------------------------------------jixinwu*/
// we're making references that assume order here.  If we're going to do that,
// we might as well use an enumeration for readability.
enum { TABLEID_CORE = 0, TABLEID_APP, TABLEID_VBE, TABLEID_MAPI, TABLEID_EPAPER, TABLEID_STSLIST };

/*---------------------------------------------------------------------------
	MsoFGimmeFeatureEx

	Perform an operation on a Darwin feature, as specified in
	%otools%\inc\misc\tcinuse.txt.  Will demand install, search, fix, check
	advertisement, etc. according to options in dwGimmeFlags.

	Returns TRUE on success, which usually means the feature is installed and
	enabled.  However, some flags affect this return.

	Use the Wz version only when absolutely necessary.
------------------------------------------------------------------ JJames -*/
MSOAPI_(BOOL) MsoFGimmeFeatureEx(msoftidT ftid, DWORD dwGimmeFlags);
#if 0 //$[VSMSO]
MSOAPI_(BOOL) MsoFGimmeFeatureExWz(const WCHAR *wzFeature, DWORD dwGimmeFlags);
#endif //$[VSMSO]

#define MsoFGimmeFeature(ftid) MsoFGimmeFeatureEx(ftid, msotcogfProvide)
#if 0 //$[VSMSO]
#define _MsoFGimmeFeature(wzFeature) MsoFGimmeFeatureExWz(wzFeature, msotcogfProvide)
#define MsoFEnabledFeature(ftid) MsoFGimmeFeatureEx(ftid, msotcogfEnabled)
#define _MsoFEnabledFeature(wzFeature) MsoFGimmeFeatureExWz(wzFeature, msotcogfEnabled)
#define MsoFInstalledFeature(ftid) MsoFGimmeFeatureEx(ftid, msotcogfInstalled)
#define _MsoFInstalledFeature(wzFeature) MsoFGimmeFeatureExWz(wzFeature, msotcogfInstalled)
#define MsoFFixFeature(ftid) MsoFGimmeFeatureEx(ftid, msotcogfForceFix)

/*---------------------------------------------------------------------------
	MsoFGimmeComponentEx

	Returns a full pathname to the component keyfile specified by the component
	id with possible language and string qualification according to the cid
	specification found in %otools%\inc\misc\tcinuse.txt.  Will demand install,
	search, fix, check advertisement, etc. according to options in dwGimmeFlags.

	Since components don't always have keyfiles (or they can be misauthored),
	you should generally use MsoFGimmeFileEx if you are relying on the pathname return.

	wzPath must be NULL or MAX_PATH characters in size.
	Returns TRUE on success, which usually means the component is installed and
	enabled.  However, some flags affect this return.
------------------------------------------------------------------ JJames -*/
MSOAPI_(BOOL) MsoFGimmeComponentEx(msocidT cid, LCID lcid, const WCHAR *wzQualifier, WCHAR *wzPath, int cchPath, DWORD dwGimmeFlags);

#define MsoFGimmeComponent(cid, wzPath, cchPath) MsoFGimmeComponentEx(cid, 0, NULL, wzPath, cchPath, msotcogfProvide)
#define MsoFGimmeComponentQualified(cid, wzQualifier, wzPath, cchPath) MsoFGimmeComponentEx(cid, 0, wzQualifier, wzPath, cchPath, msotcogfProvide)
#define MsoFGimmeLocalizedComponent(cid, lcid, wzPath, cchPath) MsoFGimmeComponentEx(cid, lcid, NULL, wzPath, cchPath, msotcogfProvide)
#define MsoFFixLocalizedComponent(cid, lcid, wzPath, cchPath) MsoFGimmeComponentEx(cid, lcid, NULL, wzPath, cchPath, msotcogfForceFix)
#endif //$[VSMSO]

/*---------------------------------------------------------------------------
	MsoFGimmeFileEx

	Returns a full pathname to the file specified by the file id with possible
	language and string qualification according to the fid specification found
	in %otools%\inc\misc\tcinuse.txt.  Will demand install,	search, fix, check
	advertisement, etc. according to options in dwGimmeFlags.

	wzPath must be NULL or MAX_PATH characters in size.
	Returns TRUE on success, which usually means the file is installed and
	enabled.  However, some flags affect this return.
------------------------------------------------------------------ JJames -*/
MSOAPI_(BOOL) MsoFGimmeFileEx(msofidT fid, LCID lcid, const WCHAR *wzQualifier, _Out_z_cap_(cchMax) WCHAR *wzPath, int cchMax, DWORD dwGimmeFlags);

#define MsoFGimmeFile(fid, wzPath, cchPath) MsoFGimmeFileEx(fid, 0, NULL, wzPath, cchPath, msotcogfProvide)
#if 0 //$[VSMSO]
#define MsoFGimmeFileQualified(fid, wzQualifier, wzPath, cchPath) MsoFGimmeFileEx(fid, 0, wzQualifier, wzPath, cchPath, msotcogfProvide)
#define MsoFGimmeLocalizedFile(fid, lcid, wzPath, cchPath) MsoFGimmeFileEx(fid, lcid, NULL, wzPath, cchPath, msotcogfProvide)
#define MsoFGimmeAdvertisedFile(qcid, wzFilename, wzPath, cchPath, fDemandInstall) MsoFGimmeComponentEx(qcid, 0, wzFilename, wzPath, cchPath, \
	(fDemandInstall) ? msotcogfProvide : msotcogfInstalled)
#define MsoFFixFile(fid, wzPath, cchPath) MsoFGimmeFileEx(fid, 0, NULL, wzPath, cchPath, msotcogfForceFix)
#define MsoFEnabledFile(fid) MsoFGimmeFileEx(fid, 0, NULL, NULL, 0, msotcogfEnabled)
#define MsoFInstalledFile(fid) MsoFGimmeFileEx(fid, 0, NULL, NULL, 0, msotcogfInstalled)
#endif //$[VSMSO]

/*-----------------------------------------------------------------------------
	MsoFGimmeFileVersion

	In addition to grabbing file (modeled after MsoFGimmeFileFull), this checks
	version of the file that we get and calls a more agressive repair if the
	version is less that the one expected.

	*** Inherited from MsoFGimmeFileFull ***
	Returns a full pathname to the file specified by the file id with possible
	language and string qualification according to the fid specification found
	in %otools%\inc\misc\tcinuse.txt.  Will demand install, search, fix, check
	advertisement, etc. according to options in dwGimmeFlags.

	wzPath must be NULL or MAX_PATH characters in size.
	Returns the language used in *plcid.
	*** Inherited from MsoFGimmeFileFull ***

	Returns TRUE on success, which usually means the file is installed, enabled,
	and at least the version requested.  However, some flags affect this return.

---------------------------------------------------------------- RFlaming ---*/
MSOAPI_(BOOL) MsoFGimmeFileVersion(msofidT fid, LCID lcid, const WCHAR *wzQualifier,
	_Out_z_cap_(cchMax) WCHAR *wzPath, int cchMax, DWORD dwGimmeFlags, DWORD dwTargetVersionMS, DWORD dwTargetVersionLS);


/*----------------------------------------------------------------------------
	MsoFGimmeComponentPathEx

	Thin wrapper around MsoFGimmeComponent, which strips out the keyfile if
	there is one.
-------------------------------------------------------------------- JBelt --*/
MSOAPI_(BOOL) MsoFGimmeComponentPathEx(msocidT cid, LCID lcid, const WCHAR *wzQualifier, WCHAR *wzPath, int cchPath, DWORD dwGimmeFlags);
#define MsoFGimmeComponentPath(cid, wzPath, cchPath) MsoFGimmeComponentPathEx(cid, 0, NULL, wzPath, cchPath, msotcogfProvide)


#if 0 //$[VSMSO]
/*---------------------------------------------------------------------------
	MsoFEnumComponentQualifiers

	Enumerate the qualifiers advertised under this qcid.
	Begin with iIndex = 0 and increase until the function returns FALSE.
	String args except wzQualifier can be NULL.
	Can pass qcid and wzQualifier to MsoFGimmeComponentQualified.
	Returns FALSE when there are no more to be had.
---------------------------------------------------------------- JJames -*/
MSOAPI_(BOOL) MsoFEnumComponentQualifiers(msocidT qcid, DWORD iIndex, WCHAR *wzQualifier, DWORD cchQualifier,
	WCHAR *wzAppData, DWORD cchAppData);


/*---------------------------------------------------------------------------
	MsoFEnumGraphicFilters

	Specialized graphic filter enumeration routine.

	Input:
	- qcid: qualified component to enumerate
	- piIndex: fill with 0 prior to initial call

	Output (all strings must be 256 chars, including terminator)
	- piIndex: incremented internally, do not modify
	- wzClass: class name
	- wzName: friendly display name
	- wzDarwinPath: Gimme token representing the path. Give this path to
		MsoFGimmeComponentQcidQualifierEx to get the real path, and possibly
		install on demand. This string is guaranteed to start with '{'.
	- wzExtensions: extensions, separated by spaces. No lowercase / uppercase
		assumptions can be made.
	- pgfo: bit field representing graphic filter options. Use msogfoxxx flags
		below. May be NULL.
-------------------------------------------------------------------- JBelt --*/
#define msogfoShowOptionsDialog  0x0001
#define msogfoShowProgressDialog 0x0002
MSOAPI_(BOOL) MsoFEnumGraphicFilters(msocidT qcid, int *piIndex,
	WCHAR *wzClass, int cchClass, WCHAR *wzName, int cchName, WCHAR *wzDarwinPath, int cchDarwin, WCHAR *wzExtensions, int cchExtensions,
	DWORD *pgfo);

/*---------------------------------------------------------------------------
	MsoFEnumNativeGraphics

	Specialized native graphic type enumeration routine.

	Input:
	- fImport: bool indicating native import or export graphic types
	- piIndex: fill with 0 prior to initial call

	Output:
	- piIndex: incremented internally, do not modify
	- wzName: friendly display name (cchMax must include NULL)
	- wzExtensions: extensions, separated by spaces. No lowercase / uppercase
		assumptions can be made. (cchMax must include NULL)
-------------------------------------------------------------------- MatWood --*/
MSOAPI_(BOOL) MsoFEnumNativeGraphics(BOOL fImport, int *piIndex,
	WCHAR *wzName, int cchName, WCHAR *wzExtensions, int cchExtensions);
#endif //$[VSMSO]

/*---------------------------------------------------------------------------
	MsoEnumComponentQualifiersEx

	Enumerate the qualifiers advertised under this qcid.
	lcid identifies a language for doubly qualified components and should
	be the same throughout a sequence of calls.	wzAppData can be null.
	Begin with *pdwIterator = 0, the function will increment on success,
	possibly by more than +1.

	Pass the lcid and wzQualifier to MsoFGimmeComponentQualifiedEx to
	retrieve the component.
---------------------------------------------------------------- JJames -*/
MSOAPI_(UINT) MsoEnumComponentQualifiersEx(
	msocidT qcid,         // gimme id
	LCID lcid,            // language id for double-qualified components, 0 if not
	WCHAR *wzQualifier,   // buffer for to receive qualifier
	DWORD *pcchQualifier, // pointer to size of buffer, receives resulting size
	WCHAR *wzAppData,     // buffer to receive application data (can be NULL)
	DWORD *pcchAppData,   // pointer to size of buffer, receives resulting size
	DWORD *pdwIterator);  // internally incremented iterator

#if 0 //$[VSMSO]
#if 0
MSOAPI_(UINT) MsoEnumComponentQualifiersExEx(
	msocidT qcid,         // gimme id
	LCID lcid,            // language id for double-qualified components, 0 if not
	WCHAR *wzQualifier,   // buffer for to receive qualifier
	DWORD *pcchQualifier, // pointer to size of buffer, receives resulting size
	WCHAR *wzAppData,     // buffer to receive application data (can be NULL)
	DWORD *pcchAppData,   // pointer to size of buffer, receives resulting size
	DWORD *pdwIterator,   // internally incremented iterator
	DWORD fAnsiCPConversion);
#endif

/*---------------------------------------------------------------------------
	MsoFGimmeComponentQualifiedData

	Get the wzAppData field for this qualified component.
	Returns TRUE and sets wzAppData if the the qualifier was found.
------------------------------------------------------------------ JJames -*/
MSOAPI_(BOOL) MsoFGimmeComponentQualifiedData(msocidT qcid, const WCHAR *wzQualifier, WCHAR *wzAppData, int cchData);

/*---------------------------------------------------------------------------
	MsoFGimmeAdvertisedName

	Returns the filename qualifier for a given pathname, verifying that the
	darwin entry for that qualifier is installed at that path.
	Reverse of MsoFGimmeAdvertisedFile.
	Returns TRUE if the path proved to be a darwin aware file.
------------------------------------------------------------------ JJames -*/
MSOAPIX_(BOOL) MsoFGimmeAdvertisedName(msocidT qcid, const WCHAR *wzPath, WCHAR *wzQualifier, int cchQualifier);

/*---------------------------------------------------------------------------
	MsoFGimmeProductCode

	Copies the 39 character product code into wzPath if true is returned
------------------------------------------------------------------ AndrewH -*/
MSOAPI_(BOOL) MsoFGimmeProductCode(WCHAR *wzPath, int cchPath);


/*-----------------------------------------------------------------------------
	MsoFGimmeOleServer

	Demand load a server based on an OLE object.
------------------------------------------------------------------ JJames ---*/
MSOAPI_(BOOL) MsoFGimmeOleServer(IOleObject *pOleObject, DWORD dwGimmeFlags);
#endif //$[VSMSO]


/*----------------------------------------------------------------------------
	MsoFidToFilename

	Returns the filename corresponding to qcid. The filename can be NULL.
------------------------------------------------------------------ JJames --*/
MSOAPI_(VOID) MsoFidToFilename(msofidT fid, _Out_z_cap_(cchMax) WCHAR *wzFilename, int cchMax);
#if 0 //$[VSMSO]
MSOAPI_(VOID) MsoCidToFilename(msocidT cid, WCHAR *wzFilename, int cch);

#define MsoQfidToFilename MsoFidToFilename
#define MsoQcidToFilename MsoCidToFilename
#endif //$[VSMSO]


/*----------------------------------------------------------------------------
	MsoFidToGuid

	Returns the GUID corresponding to the component the file belongs to.
	wzGuid must should be at least MAX_GUID characters long (39)
-------------------------------------------------------------------- JBelt --*/
MSOAPIX_(void) MsoFidToGuid(msofidT fid, _Out_z_cap_(cchGuid) WCHAR *wzGuid, int cchGuid);

#if 0 //$[VSMSO]
/*----------------------------------------------------------------------------
	MsoFindFid

	Returns the fid corresponding to a known filename.
	Returns msofidNil if not found.
	These should be used ONLY when filenames are given from outside sources.
------------------------------------------------------------------ JJames --*/
MSOAPIX_(msofidT) MsoFindFid(const WCHAR *wzFile);

#define MsoFindQfid MsoFindFid

/*----------------------------------------------------------------------------
	MsoFindFidInList

	Returns one of the fid's in the msofidNil terminated list according
	to equivalent filenames.  Returns msofidNil if not found.
	This should be used ONLY when filenames are given from outside sources.
------------------------------------------------------------------ JJames --*/
MSOAPI_(msofidT) MsoFindFidInList(const WCHAR *wzFile, const msofidT *pfid);


/*----------------------------------------------------------------------------
	MsoFFirstRun

	Performs Office first run if necessary. Call this only if your app has
	already detected that *it* needs to do a first run. This saves one boot
	registry lookup. If this returns FALSE, you must refuse to boot.
------------------------------------------------------------------- JBelt --*/
MSOAPI_(BOOL) MsoFFirstRun(HMSOINST hinst);


/*----------------------------------------------------------------------------
	MsoFReinstallProduct

	Reinstall the product.
-------------------------------------------------------------------- KirkG --*/
MSOAPIX_(BOOL) MsoFReinstallProduct(void);

/*---------------------------------------------------------------------------
	MsoFGetUserInfo

	Returns name, company, and serial number (CD key). Each string must be at
	as long msocch[Username|UserInitials|Company|Serial]Max. Pass in NULL if
	not interested in a particular string.
------------------------------------------------------------------- JBelt -*/
MSOAPI_(BOOL) MsoFGetUserInfo(WCHAR *wzName, int cchName, WCHAR *wzInitials, int cchInitials,
	WCHAR *wzCompany, int cchCompany, WCHAR *wzSerial, int cchSerial);
#endif //$[VSMSO]


// string size limits, not including null terminator
#define msocchUsernameMax		52
#define msocchUserInitialsMax	9
#define msocchCompanyMax		52
#define msocchSerialMax			23	// RPCNO-LOC-SERIALX-SEQNC

// for compatibility
#define cbCDUserNameMax 		msocchUsernameMax
#define cbCDOrgNameMax  		msocchCompanyMax
#define cbFormattedPID  		msocchSerialMax


#if 0 //$[VSMSO]
/*---------------------------------------------------------------------------
	MsoLGetProductInfo

	An Office wrapper around MsiGetProductInfoW().
---------------------------------------------------------------- EricSchr -*/
MSOAPIX_(LONG) MsoLGetProductInfo(const WCHAR *wzProperty,
	WCHAR *wzValueBuf, DWORD *pcchValueBuf);
#endif //$[VSMSO]


/*-----------------------------------------------------------------------------
	MsoLoadLocalizedLibraryFull

	LoadLibraryEx's the file fid and language plcid (if non NULL).

	If dwFlags is zero, does the equivalent of a simple LoadLibrary.

	If wzFullPath is non NULL, returns the path of the module loaded (max
	length MAX_PATH + null char).
------------------------------------------------------------------- JBelt ---*/
MSOAPI_(HMODULE) MsoLoadLocalizedLibraryFull(msofidT fid, LCID *plcid,
	const DWORD dwFlags, _Out_opt_z_cap_(cchMax) WCHAR *wzFullPath, int cchMax);


/*-----------------------------------------------------------------------------
	MsoLoadLocalizedLibraryEx

	Thin wrapper around MsoLoadLocalizedLibraryFull.
------------------------------------------------------------------- JBelt ---*/
MSOAPI_(HMODULE) MsoLoadLocalizedLibraryEx(msofidT fid, LCID lcid, DWORD dwFlags);

#define MsoLoadLibrary(fid) MsoLoadLocalizedLibraryEx(fid, 0, 0)
#define MsoLoadLibraryEx(fid, dwFlags) MsoLoadLocalizedLibraryEx(fid, 0, dwFlags)
#if 0 //$[VSMSO]
#define MsoLoadLocalizedLibrary(fid, lcid) MsoLoadLocalizedLibraryEx(fid, lcid, 0)
#endif //$[VSMSO]


/*----------------------------------------------------------------------------
	MsoFGimmeComponentQcidQualifierEx

	Decodes the qcid and qualifier encoded in wzQC, and calls
	MsoFGimmeComponentQualifier on the results. wzPath must be at least
	MAX_PATH+1 characters long.
-------------------------------------------------------------------- JBelt --*/
MSOAPI_(BOOL) MsoFGimmeComponentQcidQualifierEx(const WCHAR *wzQC, WCHAR *wzPath, int cchPath,
	DWORD dwGimmeFlags);

#define MsoFGimmeComponentQcidQualifier(wzQC, wzPath) MsoFGimmeComponentQcidQualifierEx(wzQC, wzPath, msotcogfProvide)


/*----------------------------------------------------------------------------
	MsoFGimmeComponentQcidQualifierDp

	Wrapper around MsoFGimmeComponentQcidQualifierEx, which also takes into
	account the dp variable passed in (see msotcodpxxx below). Used in loops,
	to carry the user choice to the first Gimme prompt through all subsequent
	prompts.
-------------------------------------------------------------------- JBelt --*/
MSOAPI_(BOOL) MsoFGimmeComponentQcidQualifierDp(const WCHAR *wzQC, WCHAR *wzPath, int cchPath,
	int *pdp);

#define msotcodpNotAsked    0
#define msotcodpInstall     1
#define msotcodpDontInstall 2


#if 0 //$[VSMSO]
/*-----------------------------------------------------------------------------
	MsoForgetLastGimme

	The Gimme API automatically approves or rejects subsequent calls
	within 10 seconds and before the event monitor picks up a user action.
	Successful installs turn on automatic approval, failures and user-cancels
	turn on automatic rejection.  This function resets that memory and should
	be used in cases where events don't get to the monitor.
------------------------------------------------------------------ JJames ---*/
MSOAPI_(void) MsoForgetLastGimme(void);


/*----------------------------------------------------------------------------
	MsoFFormatMessage

	Formats an appropriate error message for the last Gimme error.
	Returns FALSE if unknown.
	NOT REALLY USED.
------------------------------------------------------------------ JJames --*/
MSOAPI_(BOOL) MsoFFormatMessage(DWORD dwError, WCHAR *wzMessage, int cchMax);


/*------------------------------------------------------------------------
	MsoLaunchFid

	Given a FID for a file that represents a setup-installed EXE,
	ShellExec it and return the hinstance.  Pases arguments specified
	in character string, if any (may be NULL)
---------------------------------------------------------------- MikeKell -*/
MSOAPI_(HINSTANCE) MsoLaunchFid(msofidT fid, const WCHAR *wzArguments, int sw);


/*------------------------------------------------------------------------
	MsoFLaunchMsInfo

	Launches MSInfo.  Does some extra checking like bringing it forward
	if if is already there.  wzArguments should contain the name of the
	application invoking MSInfo.
---------------------------------------------------------------- MikeKell -*/
MSOAPI_(BOOL) MsoFLaunchMsInfo(const WCHAR *wzArguments);


/*----------------------------------------------------------------------------
	MsoFEnsureUserData

	If the registry referenced by rid than 1, ensures the user data in ftid
	is on the machine, then writes 1 in the registry. Call this after initing
	ORAPI, but before reading anything, to make sure Setup-time user data has
	been written for this particular user.
-------------------------------------------------------------------- JBelt --*/
MSOAPI_(BOOL) MsoFEnsureUserData(int rid, msoftidT ftid);


/*----------------------------------------------------------------------------
	MsoFEnsureFeature

	If the registry referenced by rid does not exist, then reinstall ftid.
	Call this after initing	ORAPI and Gimme. 
---------------------------------------------------------------- (EricLam) --*/
MSOAPI_(BOOL) MsoFEnsureFeature(int rid, msoftidT ftid);


/*----------------------------------------------------------------------------
	MsoFEnsureTypelib

	1) Detect Mso Typelib key existence:
	[HKEY_CLASSES_ROOT\TypeLib\{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}\2.2\0\win32]
    2) call Darwin to repair the feature that contains the typelib (ProductFiles)
	Note: Call this after initing ORAPI and Gimme.
---------------------------------------------------------------- (EricLam) --*/
MSOAPI_(BOOL) MsoFEnsureMsoTypelib();


/*----------------------------------------------------------------------------
	MsoFInGimme

	Returns TRUE if we're currently stuck in a potentially long Darwin call.
	Be patient if you get called in a message filter and this returns TRUE.
	Fix Office 9 24103.
-------------------------------------------------------------------- JBelt --*/
MSOAPI_(BOOL) MsoFInGimme();


/*----------------------------------------------------------------------------
	MsoDwGetGimmeTableVersion

	Returns the version number of the passed-in Gimme table.
-------------------------------------------------------------------- JBelt --*/
MSOAPIX_(DWORD) MsoDwGetGimmeTableVersion(const MSOTCFCF *pfcf);


/*------------------------------------------------------------------------
	MsoFIsOLEAwareDarwin

	Returns TRUE if OS recognizes Darwin descriptors in the OLE registry.
---------------------------------------------------------------- WesYang -*/
MSOAPIX_(BOOL) MsoFIsOLEAwareDarwin ();
#endif //$[VSMSO]


/*------------------------------------------------------------------------
	Obsolete Gimme wrapper to the standard OLE calls
---------------------------------------------------------------- JJames -*/

#define MsoFGimmeCoCreateInstance(rclsid,pUnkOuter,dwClsContext,riid,ppv) \
		         CoCreateInstance(rclsid,pUnkOuter,dwClsContext,riid,ppv)
#define MsoFGimmeCoGetClassObject(rclsid,dwClsContext,pServerInfo,riid,ppv) \
		         CoGetClassObject(rclsid,dwClsContext,pServerInfo,riid,ppv)
#define MsoHrCoCreateInstance(rclsid,pUnkOuter,dwClsContext,riid,ppv) \
		     CoCreateInstance(rclsid,pUnkOuter,dwClsContext,riid,ppv)
#define MsoHrCoGetClassObject(rclsid,dwClsContext,pServerInfo,riid,ppv) \
		     CoGetClassObject(rclsid,dwClsContext,pServerInfo,riid,ppv)
#define MsoHrOleCreate(rclsid,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject) \
		     OleCreate(rclsid,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject)
#define MsoHrOleCreateLink(pmkLinkSrc,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObj) \
		     OleCreateLink(pmkLinkSrc,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObj)
#define MsoHrOleCreateFromFile(rclsid,lpszFileName,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject) \
		     OleCreateFromFile(rclsid,lpszFileName,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject)
#define MsoHrOleCreateLinkToFile(lpszFileName,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject) \
		     OleCreateLinkToFile(lpszFileName,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject)
#define MsoHrOleCreateFromData(pSrcDataObj,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject) \
		     OleCreateFromData(pSrcDataObj,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject)
#define MsoHrOleCreateStaticFromData(pSrcDataObj,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject) \
		     OleCreateStaticFromData(pSrcDataObj,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject)
#define MsoHrOleCreateLinkFromData(pSrcDataObj,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject) \
		     OleCreateLinkFromData(pSrcDataObj,rrid,renderopt,pFormatEtc,pClientSite,pStg,ppvObject)
#define MsoHrOleLoad(pStg,riid,pClientSite,ppvObj) \
		     OleLoad(pStg,riid,pClientSite,ppvObj)
#define MsoHrOleLoadFromStream(pStm,iidInterface,ppvObj) \
		     OleLoadFromStream(pStm,iidInterface,ppvObj)


/*---------------------------------------------------------------------------
	IMsoGimmeUser callback for Gimme clients.
------------------------------------------------------------------- JBelt -*/

#undef  INTERFACE
#define INTERFACE  IMsoGimmeUser

#define	cchMaxMsoGimmeUserFGetString	300	// maximum length of returned string
											// from IMsoGimmeUser::FGetString
#define GIMMEUSER_INSTALL	0x01	// ok to install
#define GIMMEUSER_GIMMEUI	0x02	// ok to display Gimme UI (install prompt, etc)
#define GIMMEUSER_DARWINUI	0x04	// ok to display Darwin UI
#define GIMMEUSER_DEFAULT   (GIMMEUSER_INSTALL | GIMMEUSER_GIMMEUI | GIMMEUSER_DARWINUI)

DECLARE_INTERFACE(IMsoGimmeUser)
{

	/* Debugging interface for this interface */
	MSODEBUGMETHOD

	/* If Darwin is on the machine, but your app doesn't have a product code,
		Office will look for a MSI to install on the fly, and call you.
		If you return FALSE, Office immediately gives up looking for a MSI.
		If you return TRUE, Office will look for the MSI.
		- If you set *pfPattern to TRUE, wtzPattern is assumed to be a file
			pattern (with wildcards). If you set *pfPattern to FALSE, wtzPattern
			is assumed to be a fully qualified path to the file, and is
			used as is.
		- Office prefills wtzPattern with a suggested search pattern, and
			pfPattern with TRUE. */
	MSOMETHOD_(BOOL, FSearchMSI) (THIS_ WCHAR *wtzPattern, BOOL *pfPattern) PURE;

	/* Gimme is ready to demand install something. Return a combination of
		GIMMEUSER_xxx flags declared above, or GIMMEUSER_DEFAULT. */
	MSOMETHOD_(DWORD, DwInstallBehavior) (THIS) PURE;

	/* The Gimme layer needs the string corresponding to the string identifier
		corresponding to an id you have specified in your tcinuse.txt. Fill
		wtz with the string and return TRUE; return FALSE if unable to fill
		the string, and Gimme will default to a generic string.  Note the
		buffer is a wtz, i.e. the first entry is the length of the string,
		not including the terminating null.  The size of the buffer is
		cchMaxMsoGimmeUserFGetString.  */
	MSOMETHOD_(BOOL, FGetString) (THIS_ int ids, WCHAR *wtz, int cch) PURE;

	/* Return the directory in which most of your files live (your main, or bin,
		directory), and TRUE. The Gimme layer uses this directory for search for
		files for dev override (oprep machines), and resiliency (if Darwin is
		dead, there's a MSI mismatch, etc). Office prefills wtzDir with the
		directory where the EXE which launched the process lives. wtzDir is
		MAX_PATH+1 long. Returning FALSE turns off override / resiliency.
		Hint: use MsoGetModuleFilenameW to get your main DLL / EXE path. */
	MSOMETHOD_(BOOL, FGetRootDirectory) (THIS_ WCHAR *wtzDir) PURE;
};

#if 0 //$[VSMSO]
/*----------------------------------------------------------------------------
	MsoDwGimmeUserInstallBehavior

	This code generates the behavior flags needed for DarwinOK
------------------------------------------------------------------- ARSHADA --*/
MSOAPI_(DWORD) MsoDwGimmeUserInstallBehavior(WORD FeatureInstall, BOOL fDisplayAlerts);

// OBSOLETE, DO NOT CALL. Use MsoDwGimmeUserInstallBehavior instead
// TODO(JBelt): delete, ssync with VB
MSOAPIX_(BOOL) MsoFGimmeUserInstallBehavior(WORD FeatureInstall, BOOL fDisplayAlerts, WORD *pwFlag);


/*---------------------------------------------------------------------------
	MsoInitGimme

	Hook up your own Gimme tables, if you have any. If you don't, you don't
	need to call this API.

	cidCore is a Gimme component ID to your core component.  It must be given.
	pGimmeTables is build by otcdarmake.exe as "vfcf" and can be NULL.
	If your cid is not in the Office tables, then you must build your own
	using otcdarmake.exe from a version of %otools%\inc\misc\tcinuse.txt.
	See http://officedev/tco/gimmehelp.htm for details.

	See IMsoGimmeUser above for description of pigu. You must implement this
	interface, or some Darwin features will be disabled.
------------------------------------------------------------------- JBelt -*/
MSOAPI_(void) MsoInitGimmeEx(msocidT cidCore, MSOTCFCF *pGimmeTables,
		IMsoGimmeUser *pigu, msocidT cidFull, DWORD dwUnused);
MSOAPI_(void) MsoInitGimme(msocidT cidCore, MSOTCFCF *pGimmeTables,
	IMsoGimmeUser *pigu);


//
// This enum list which is used for the MsoGimmePublishComponentString() API
// below is to be mapped with the %OTOOLS%\inc\misc\msistr.pp list of
// string identifiers.  This list is on a "fill in as you go" basis if
// other clients wish to access certain strings of the MSI string table.
enum {
	msiIndexDesignTemplates = 4, // msiidsDesignTemplates
};

MSOAPI_(void) MsoGimmePublishComponentString(DWORD dwIndex, LPWSTR pwzBuffer, DWORD *pcchBuffer);
#endif //$[VSMSO]

/****************************************************************************
 World-Wide Exe
****************************************************************************/

/*----------------------------------------------------------------------------
	MsoInitPluggableUI

	Init and cache the language settings used by pluggable UI
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(BOOL) MsoInitPluggableUI(void);

/*-----------------------------------------------------------------------------
	MsoSetupFontLink

	Setup the global Fontlink switch : vfDoFontLink
--------------------------------------------------------------------- ZIYIW -*/
MSOAPIX_(void) MsoSetupFontLink(LCID lcidUINew);

/*----------------------------------------------------------------------------
	MsoAnsiCodePageLimited

	Declare that your application is code page limited.
	Must be called before MsoInitPluggableUI.
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(VOID) MsoAnsiCodePageLimited(BOOL fLimited);

/*----------------------------------------------------------------------------
	MsoFAnsiCodePageSupportsLCID

	Test whether the code page supports the lcid for an ANSI application.
	Typically, cp = GetACP().
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(BOOL) MsoFAnsiCodePageSupportsLCID(UINT cp, LCID lcid);


/*----------------------------------------------------------------------------
	MsoFValidLocale

	Test whether this lcid is valid on this machine.
------------------------------------------------------------------- JJames --*/
MSOAPI_(BOOL) MsoFValidLocale(LCID lcid);

#define ILS_NOTCHANGED				0
#define ILS_CHANGED_NOT_PROCESSED	1
#define ILS_CHANGED_PROCESSED		2
#define ILS_CHANGED_PROCESSING		3

#define APPID_WORD					0
#define APPID_XL					1
#define APPID_PPT					2
#define APPID_ACCESS				3
#define APPID_OUTLOOK				4
#define APPID_FRONTPAGE				5
#define APPID_PUBLISHER				6
#define APPID_PROJECT				7
#define APPID_DESIGNER				8
#define APPID_XDOCS					9
#define APPID_ONENOTE				10

#if 0 //$[VSMSO]
/*----------------------------------------------------------------------------
	MsoAppSetChangeInstallLanguageState
	MsoAppGetChangeInstallLanguageState

	Get/Set Application based install language change state.
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(int) MsoAppSetChangeInstallLanguageState(int idApp, int ils);
MSOAPI_(int) MsoAppGetChangeInstallLanguageState(int idApp);
#endif //$[VSMSO]

/*----------------------------------------------------------------------------
	MsoGetInstallLcid

	return the cached office install lcid
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(LCID) MsoGetInstallLcid(void);

#if 0 //$[VSMSO]
/*----------------------------------------------------------------------------
	MsoGetInstallLcid

	return the cached office install lcid
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(LCID) MsoGetInstallLcid2000Compatible(void);

/*-----------------------------------------------------------------------------
	MsoGetInstallFlavor

	return the cached office install flavor lcid
--------------------------------------------------------------------- ZIYIW -*/
MSOAPI_(LCID) MsoGetInstallFlavor(void);
#endif //$[VSMSO]

/*-----------------------------------------------------------------------------
	MsoFLangChanged

	return a flag whether or not the UI lang has been changed since last time
--------------------------------------------------------------------- ZIYIW -*/
MSOAPI_(int) MsoFLangChanged(LCID *plcid);

/*-----------------------------------------------------------------------------
	MsoGetPreviousUILcid

	return the cached previous ui lcid
--------------------------------------------------------------------- ZIYIW -*/
MSOAPIX_(LCID) MsoGetPreviousUILcid(void);

/*----------------------------------------------------------------------------
	MsoGetUILcid

	return the cached office UI lcid
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(LCID) MsoGetUILcid(void);

/*----------------------------------------------------------------------------
	MsoGetHelpLcid

	return the cached office Help lcid
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(LCID) MsoGetHelpLcid(void);

#if 0 //$[VSMSO]
/*-----------------------------------------------------------------------------
	MsoGetExeModeLcid

	return the cached ExeMode lcid
-------------------------------------------------------------------- IrfanGo -*/
MSOAPI_(LCID) MsoGetExeModeLcid(void);

/*----------------------------------------------------------------------------
	MsoGetSKULcid

	return the cached office installed SKU lcid
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(LCID) MsoGetSKULcid(void);
#endif //$[VSMSO]

/*-----------------------------------------------------------------------------
	MsoGetWebLocale -  return the cached weblocale lcid
-------------------------------------------------------------------- ZIYIW -*/
MSOAPI_(LCID) MsoGetWebLocale(void);

/*----------------------------------------------------------------------------
	MsoEnumEditLcid

	enumerate throught the cached office edit lcids
------------------------------------------------------------------- ZIYIW --*/
MSOAPI_(BOOL) MsoEnumEditLcid(LCID*, int);

/*----------------------------------------------------------------------------
	MsoLcidGetLanguages

	read various language settings from the registry throught ORAPI
----------------------------------------------------------- ZIYIW/irfango --*/
#if 0 //$[VSMSO]
BOOL MsoLcidGetLanguages(LCID*, LCID*, LCID*, MSOELI* rgeli, int* pceliMax, MSOELI* rgeliOff, int* pceliOffMax, LCID*);
#else //$[VSMSO]
BOOL MsoLcidGetLanguages(LCID*, LCID*, LCID*, MSOELI* rgeli, int* pceliMax, MSOELI* rgeliOff, int* pceliOffMax);
#endif //$[VSMSO]

/*-----------------------------------------------------------------------------
	MsoDialogFontNameLid

	get the Localized/EUC dialog font name based on UI lid passed in.
--------------------------------------------------------------------- ZIYIW -*/
MSOAPI_(void) MsoDialogFontNameLid(_Out_z_cap_(cchMax) WCHAR *wzName, int cchMax, LCID lid);


/*-----------------------------------------------------------------------------
	MsoSetPureReg/MsoGetPureReg

	Operations on the Pure language resource registry
--------------------------------------------------------------------- ZIYIW -*/
#define REG_PURE_UNKNOWN		0
#define REG_PURE_OFF			1
#define REG_PURE_COMPLETED		2
#define REG_PURE_PROHIBITED		3
#define REG_PURE_ON             REG_PURE_PROHIBITED

MSOAPIX_(int) MsoSetPureReg(int iState);
MSOAPI_(int) MsoGetPureReg(void);

/****************************************************************************
 Migration
****************************************************************************/

enum
{
	msoInvalidApp = -1, // used for initialization
	msoFirstApp  = 0,
	msoWord      = 0,
	msoExcel     = 1,
	msoAccess    = 2,
	msoPPT       = 3,
	msoOffice    = 4,
	msoGraph     = 5,
	msoOutlook   = 6,
	msoFrontPage = 7,
	msoPublisher = 8,
	msoProject   = 9,
	msoVisio     = 10,
	msoDesigner  = 11,
	msoOSA		 = 12,
	msoJot       = 13,
	msoXDocs     = 14,
	msoLastApp   = msoXDocs
};

#define msoNoCmwPopulate 0x80000000

// WARNING: There is code in word that assumes the ordering of the versions:
// versions should be sorted newest to oldest (code in word\src\htmlin.c)
enum
{
	msoOfficeCurrentVersion = 0,
	msoOffice11Version		= 0,
	msoOffice10Version      = 1,
	msoOffice9Version       = 2,
	msoOffice97Version      = 3
};

enum
{
	msoNoMigration		= 0,
	msoOfficeVersion6	= 6,
	msoOfficeVersion7	= 7,
	msoOfficeVersion8	= 8,
	msoOfficeVersion9	= 9,
	msoOfficeVersionXP	= 10
};

MSOAPI_(BOOL) MsoMigrate(int iApp, DWORD *pdwMigrationVersion);


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                        O R A P I   A P I s
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

#if !defined(STATIC_LIB_DEF) || defined(ZENSTAT_LIB_DEF)
#include "msoreg.h"
#endif

/*-----------------------------------------------------------------------------
	ORAPI Cache data
--------------------------------------------------------------------dgray----*/
typedef struct KEYNODE_S {
	HKEY	hKey;                   // handle to the key
	int		keyID;                  // enum ID number for this key
	union
		{
		int		Options;            // ORAPICacheOptionFlags
		struct
			{
			ULONG fPersist  : 1;
			ULONG fRWAccess : 1;
			ULONG fIsPolicy : 1;
			ULONG fIsApp    : 1;
			ULONG fIsValid  : 1;
			// if the ref count is changed, make sure
			// ORAPI_MAX_REF_COUNT in tcorapi.cpp matches
			ULONG nRefCount  : 3;
			// TODO DGray : Pad this to 32 bits?
			// int pad        : 24;
			};
		};
#ifdef DEBUG
	CHAR	szKeyName[MAX_PATH];  // Name of the key
	int		nTimesUsed;           // Number of times this key has been hit
#endif // DEBUG
	struct KEYNODE_S* pNext;      // The next keyID in the cache
	struct KEYNODE_S* pPrev;      // The next keyID in the cache
} KEYNODE;


/*

 PERSIST    - When set, Do not remove this key from the cache
 RW_ACCESS  - When set, key opened with READ/WRITE access, otherwise just read
 IS_POLICY  - Set, this key exists in the policy tree, otherwise user tree
 IS_APP_KEY - Set, this key is from the APP, otherwise from MSO
 KEY_VALID  - The hkey attached to this node is valid

*/
enum ORAPICacheOptionFlags
{
	PERSIST			= 0x01,
	RW_ACCESS		= 0x02,
	IS_POLICY		= 0x04,
	IS_APP_KEY		= 0x08,
	KEY_VALID		= 0x10,

	MASK_PERSIST	= 0xFFFFFFFE,
	MASK_RW			= 0xFFFFFFFD,
	MASK_POLICY		= 0xFFFFFFFB,
	MASK_APP		= 0xFFFFFFF7,
};



/*-----------------------------------------------------------------------------
	MsoFOrapiPrimeKeyCache

	Prime a root of the registry tree in the cache
	Used only in the init
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFOrapiPrimeKeyCache(int keyID, int Options, HKEY hKey,
                                     BOOL fHoldRef, KEYNODE** ppkn, PCSTR sz);
//#ifdef DEBUG
//MSOAPI_(BOOL) MsoFOrapiPrimeKeyCache(int keyID, int Options, HKEY hKey,
//                                     BOOL fHoldRef, KEYNODE** ppkn,
//                                     PCSTR sz);
//#else  // ! DEBUG
//#define MsoFOrapiPrimeKeyCache(keyID,Options,hKey,fHoldRef,ppkn,sz) \
//        MsoFOrapiPrimeKeyCache(keyID,Options,hKey,fHoldRef,ppkn)
//MSOAPI_(BOOL) MsoFOrapiPrimeKeyCache(int keyID, int Options, HKEY hKey,
//                                     BOOL fHoldRef, KEYNODE** ppkn);
//#endif // ! DEBUG


/*-----------------------------------------------------------------------------
	ORAPI App init routines
--------------------------------------------------------------------dgray----*/

/*-----------------------------------------------------------------------------
   MsoFRegHookAppTables

   Hooks the application data tables into ORAPI.  Needs to be called before
   any ORAPI calls are made by an application.
   Returns TRUE if policy is in effect, FALSE if no policy.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFRegHookAppTables(const void* pAppReg, const void* pAppOrkey,
                                   int cNumRegs, int cNumKeys);



/*-----------------------------------------------------------------------------
	These ORAPI functions are used to get a handle to a HKEY.  This should
	only be used when you are using orapi to open the key, and then you are
	using the HKEY to do something ORAPI doesn't handle, such as enumeration.

	Using MsoRegOpenKeyEx2 is the preferred function to open.

	ALWAYS close the key through ORAPI with MsoRegCloseKeyHkey().
--------------------------------------------------------------------dgray----*/
MSOAPI_(LONG) MsoRegOpenKey    (int msorid, PHKEY phkResult);
MSOAPI_(LONG) MsoRegOpenKeyEx  (int msorid, REGSAM samDesired, PHKEY phkResult);
MSOAPI_(LONG) MsoRegOpenKeyEx2 (int msorid, REGSAM samDesired, PHKEY phkResult, BOOL fPolicy);
MSOAPI_(LONG) MsoRegCreateKeyEx(int msorid, PHKEY phkResult,
                                LPDWORD lpdwDisposition);
MSOAPI_(LONG) MsoRegCreateKey  (int msorid, PHKEY phkResult);




/*-----------------------------------------------------------------------------
   MsoRegDeleteValue

   This function deletes a value from the registry.
--------------------------------------------------------------------dgray----*/
MSOAPI_(LONG) MsoRegDeleteValue(int msorid);



/*-----------------------------------------------------------------------------
   MsoRegDeleteKey

   This function deletes a registry key from the users registry tree.
   It also clears the value in the cache if there is one.

--------------------------------------------------------------------dgray----*/
MSOAPI_(LONG) MsoRegDeleteKey(int msorid);



/*-----------------------------------------------------------------------------
	MsoCbRegGetBufferSize*

	Use these functions to mimic the win32 RegQueryValueEx() call to get
	buffer size.  RegQueryValueEx(phkey, pValueName, NULL, NULL, NULL, &Size);
	Use the function suited to the type accessed.

  MsoCbRegGetBufferSizeCore
  MsoCbRegGetBufferSizeDefaultCore
    Don't use these.  Other funcions wrap or #define to them.

  MsoCbRegGetBufferSizeSz
    Use this function to query for the size you need to allocate in order
    to query for a REG_SZ in ansi space.  Returns the size in bytes.

  MsoCbRegGetBufferSizeDefaultSz
    Use this function to query for the size you need to allocate in order
    to query for the default value of a REG_SZ in ansi space.
    Returns the size in bytes.

  MsoCbRegGetBufferSizeWz
    Obsolete, use MsoCchRegGetBufferSizeWz

  MsoCbRegGetBufferSizeDefaultWz
    Obsolete, use MsoCchRegGetBufferSizeDefaultWz

  MsoCchRegGetBufferSizeWz
    Use this function to query for the size you need to allocate in order
    to query for a REG_SZ in unicode space.  Returns the size in charater count.

  MsoCchRegGetBufferSizeDefaultWz
    Use this function to query for the size you need to allocate in order
    to query for the default value of a REG_SZ in unicode space.
    Returns the size in charater count.

  MsoCbRegGetBufferSizeBinary
    Use this function to query for the size you need to allocate in order
    to query for a REG_BINARY value data.  Returns size in bytes.

  MsoCbRegGetBufferSizeMultiSz
    Use this function to query for the size you need to allocate in order
    to query for a REG_MULTI_SZ value data.  Returns size in bytes.

--------------------------------------------------------------------dgray----*/
MSOAPI_(DWORD) MsoCbRegGetBufferSizeCore(int msorid);
MSOAPIX_(DWORD) MsoCbRegGetBufferSizeDefaultCore(int msorid);
MSOAPI_(DWORD) MsoCbRegGetBufferSizeSz(int msorid);
MSOAPIX_(DWORD) MsoCbRegGetBufferSizeDefaultSz(int msorid);
MSOAPI_(DWORD) MsoCbRegGetBufferSizeWz(int msorid);
MSOAPIX_(DWORD) MsoCbRegGetBufferSizeDefaultWz(int msorid);
MSOAPI_(DWORD) MsoCchRegGetBufferSizeWz(int msorid);
MSOAPIX_(DWORD) MsoCchRegGetBufferSizeDefaultWz(int msorid);
MSOAPI_(DWORD) MsoCbRegGetBufferSizeMultiWz(int msorid);

#ifdef DEBUG
MSOAPI_(DWORD) MsoCbRegGetBufferSizeBinary(int msorid);
#else  // ! DEBUG
#define MsoCbRegGetBufferSizeBinary(msorid) MsoCbRegGetBufferSizeCore(msorid)
#endif // ! DEBUG



/*-----------------------------------------------------------------------------
   MsoRegForceWriteDefaultValue

   DO NOT USE THIS FUNCTION!  It is not safe for general consumption.  This
   should only be used by JJames.  See the comments in tcorapi.cpp
--------------------------------------------------------------------dgray----*/
MSOAPIX_(LONG) MsoRegForceWriteDefaultValue(int msorid);



/*-----------------------------------------------------------------------------
	MsoFRegGetDw

	Gets the REG_DWORD value for this msorid and puts it in *pdwData.
	MsoFRegGetDw
		Returns TRUE if succeeded, FALSE if failed.  Assert if default value
		is not NO_DEFAULT
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegGetDwCore(int msorid, DWORD* pdwData);
#ifdef DEBUG
MSOAPI_(BOOL)  MsoFRegGetDw(int msorid, DWORD* pdwData);
#else  // ! DEBUG
#define MsoFRegGetDw(msorid, pdwData) MsoFRegGetDwCore(msorid, pdwData)
#endif // ! DEBUG



/*-----------------------------------------------------------------------------
	MsoDwRegGetDw

	Returns the REG_DWORD value for this msorid.
	Does not check for failure, which is fine if default values are defined.
--------------------------------------------------------------------dgray----*/
MSOAPI_(DWORD) MsoDwRegGetDw(int msorid);



/*-----------------------------------------------------------------------------
	These should be used in only select cases!  These functions should be
	used only in cases where we know that a DWORD may be written out as a
	REG_BINARY type.  In general, we will accept a REG_BINARY if it is the
	correct size, but we will assert if the type does not match.  These
	functions turn that assert off and back on.  Use like so:
		MsoRegDwTypeMatchAssertOff();
		dw = MsoDwRegGetDw(msoridFoo);
		MsoRegDwTypeMatchAssertOn();

	ALWAYS CALL THIS IN PAIRS SO THE ASSERT GETS RE-ENABLED!
--------------------------------------------------------------------dgray----*/
#ifdef DEBUG
MSOAPI_(void) MsoRegDwTypeMatchAssertOff();
MSOAPI_(void) MsoRegDwTypeMatchAssertOn();
#else  // ! DEBUG
#define MsoRegDwTypeMatchAssertOff()
#define MsoRegDwTypeMatchAssertOn()
#endif // ! DEBUG



/*-----------------------------------------------------------------------------
	These should be used in only select cases!  These functions should be
	used when we're calling MsoFRegGetDw, but the rid may have Orapi
	default-value-data.  In general, MsoFRegGetDw will assert if the rid has
	def-value-data, since the code could call MsoDwRegGetDw instead (guaranteed
	not to fail).
	These functions turn that assert off and back on.  Use like so:
		MsoRegDefValAssertOff();
		if (MsoFDwRegGetDw(msoridFoo)) blah;
		MsoRegDefValAssertOn();

	ALWAYS CALL THIS IN PAIRS SO THE ASSERT GETS RE-ENABLED!
--------------------------------------------------------------------dgray----*/
#ifdef DEBUG
MSOAPI_(void) MsoRegDefValAssertOff();
MSOAPI_(void) MsoRegDefValAssertOn();
#else  // ! DEBUG
#define MsoRegDefValAssertOff()
#define MsoRegDefValAssertOn()
#endif // ! DEBUG



/*-----------------------------------------------------------------------------
	MsoFRegSetDw

	Sets the REG_DWORD value for this msorid.
	Returns TRUE if succeeded, FALSE if failed.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegSetDw(int msorid, DWORD dwData);



/*-----------------------------------------------------------------------------
	MsoFRegGetBinary

	Gets the REG_BINARY value for this msorid. pcbData should be set to the size
	of the buffer passed in.  (pcbData needed can be queried with GetBufferSize
	functions)

	Returns:	TRUE if succeeded, FALSE if failed.
	Sides:		*pbData is filled with the binary data retrieved.
				*pcbData is filled with the size of the returned binary.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegGetBinary(int msorid, LPBYTE pbData, DWORD* pcbData);



/*-----------------------------------------------------------------------------
	MsoFRegSetBinary

	Sets the REG_BINARY value for this msorid. Cb should be set to the size
	of the buffer passed in to be written.

	Returns:	TRUE if succeeded, FALSE if failed.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegSetBinary(int msorid, const BYTE *pbData, DWORD cbData);


/*-----------------------------------------------------------------------------
	MsoFRegGetMultiWz / MsoFRegSetMultiWz

	Gets/sets the REG_MULTI_SZ value for this msorid.  cch should be set to the
	count of characters (including NULLs) allowed in the buffer passed in.
	Note: There are no RegGetMultiSz versions of these functions.

	Returns:    TRUE if succeeded, FALSE if failed.
	   (get)    pcch will be filled in with the count of characters written
	            to the buffer.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFRegGetMultiWz(int msorid, _Out_z_cap_(*pcch) WCHAR* pwch, DWORD* pcch);
MSOAPI_(BOOL) MsoFRegSetMultiWz(int msorid, const WCHAR *pwch, DWORD cch);


/*-----------------------------------------------------------------------------
	Mso*RegReadSz

	Gets the REG_SZ value for this msorid.

	Input Parameters:
						Cb should be set to the size of the buffer passed
							in in chars
						sz should be the character buffer to be filled.

	Sides:		*sz is filled with the ansi string data retrieved.

	MsoFRegReadSz
		Returns TRUE if success, FALSE if failed.  Asserts is default value
		is not NO_DEFAULT
	MsoRegReadSz
		No return value.  Used if there is a default value in the database.
--------------------------------------------------------------------dgray----*/

MSOAPI_(VOID)  MsoRegReadSz (int msorid, _Out_z_cap_(cch) char* sz, DWORD cch);
MSOAPI_(BOOL)  MsoFRegReadSz(int msorid, _Out_z_cap_(cch) char* sz, DWORD cch);

#define MsoFRegGetSz(msorid, sz, Cb) MsoFRegReadSz(msorid, sz, Cb)
#define MsoRegGetSz(msorid, sz, Cb)  MsoRegReadSz(msorid, sz, Cb)


/*-----------------------------------------------------------------------------
	Mso*RegReadWz

	Gets the REG_SZ value for this msorid.

	Input Parameters:
						cwch should be set to the size of the buffer passed
							in in characters
						wz should be the character buffer to be filled.

	Sides:		*wz is filled with the wide string data retrieved.

	MsoFRegReadWz
		Returns TRUE if success, FALSE if failed.  Asserts is default value
		is not NO_DEFAULT
	MsoRegReadWz
		No return value.  Used if there is a default value in the database.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegReadWz(int msorid, _Out_z_cap_(cwch) WCHAR* wz, DWORD cwch);
MSOAPI_(VOID)  MsoRegReadWz(int msorid, _Out_z_cap_(cwch) WCHAR* wz, DWORD cwch);

__inline BOOL  MsoFRegReadWzNonEmpty(int msorid, _Out_z_cap_(cwch) WCHAR* wz, DWORD cwch)
{
	return MsoFRegReadWz(msorid, wz, cwch) && wz[0] != '\0';	//useful for old win.ini msorid's
}

#define MsoFRegGetWz( msorid, wz, Cb ) MsoFRegReadWz( msorid, wz, (Cb)/sizeof(WCHAR) )
#define MsoRegGetWz( msorid, wz, Cb ) MsoRegReadWz( msorid, wz, (Cb)/sizeof(WCHAR) )


/*-----------------------------------------------------------------------------
	MsoFRegSetSz

	Sets the REG_SZ registry value for this msorid using sz as the input.

	Input Parameters:
						*sz is the buffer containing the ansi string to write.

	Returns:	TRUE if succeeded, FALSE if failed.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegSetSz(int msorid, PCSTR sz);



/*-----------------------------------------------------------------------------
	MsoFRegSetWz

	Sets the REG_SZ registry value for this msorid using a wz as the input.

	Input Parameters:
						*wz is the buffer containing the wide string to write.

	Returns:	TRUE if succeeded, FALSE if failed.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL)  MsoFRegSetWz(int msorid, PCWSTR wz);



/*-----------------------------------------------------------------------------
	MsoFRegKeyExists

	Returns true if the key for this msorid exists in the registry.
	This could be in either the policy tree or the user reg tree.
	Returns false if there is no value in the registry.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFRegKeyExists(int msorid);



/*-----------------------------------------------------------------------------
	Mso*RegGetDefault*

	Retrieves the DEFAULT value for a particular rid, in the same method as the
	other retrieval functions above.

	False if failed (no default value)
--------------------------------------------------------------------dgray----*/
MSOAPIX_(BOOL) MsoFRegGetDefaultSz(int msorid, _Out_z_cap_(cch) PSTR sz, DWORD cch);
MSOAPI_(BOOL) MsoFRegReadDefaultWz(int msorid, _Out_z_cap_(cwch) PWSTR wz, DWORD cwch);
MSOAPI_(DWORD) MsoDwRegGetDefaultDw(int msorid);

// DON'T USE THIS API --  Use MsoFRegReadDefaultWz instead
MSOAPI_(BOOL) MsoFRegGetDefaultWz(int msorid,  _Out_bytecap_(Cb) PWSTR wz, DWORD Cb);



/*-----------------------------------------------------------------------------
	MsoFRegValueExists

	Returns true if value data for this msorid exists in the registry.
	This could be in either the policy tree or the user reg tree.
	Returns false if there is no value in the registry.

--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFRegValueExists(int msorid);



/*-----------------------------------------------------------------------------
	MsoFRegDefaultValueExists

	Returns > 0 if default value exists, returns 0 if no default exists.
--------------------------------------------------------------------dgray----*/
MSOAPIDBG_(BOOL) MsoFRegValueExistsDefault(int msorid);



#if 0 //$[VSMSO]
/*-----------------------------------------------------------------------------
	MsoFRegPolicyValueExists

	Returns TRUE if there is a value in the Policies tree to return.
	FALSE otherwise.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFRegPolicyValueExists(int msorid);
#endif //$[VSMSO]



/*-----------------------------------------------------------------------------
	OrapiSetVal

	Sets a Generic ORAPI value

	Don't use this function if you just check for ERROR_SUCCESS.  It is
	wrapped for each type.

	Returns:	Win32 error code
	Sides:		None.
--------------------------------------------------------------------dgray----*/
MSOAPIX_(LONG) OrapiSetVal(int msorid, const BYTE* pbData, DWORD Cb);



/*-----------------------------------------------------------------------------
	OrapiQueryVal

	Queries for a Generic ORAPI value, not a string
	Order of how it gets the value
			1) Query Software/Policy Tree
			2) Query Software Tree
			3) Use Default Value
			4) Fill with empty value (0; 0x00, 0x00;)

	Returns:	Win32 error code (ERROR_SUCCESS except in case 4, but never
					use to check for ERROR_MORE_DATA, or
					ERROR_INSUFFICIENT_BUFFER.
					(Use the GetBufferSize functions)
	Sides:		Fills wzData with a valid wz string
				Fills *pcbData with the size in bytes returned.

REVIEW DGray:
		Is there a way I can make this function even more compact by
		using a pointer to a function to make the query calls,
		and dynamically setting it to either the W or A version, based
		on vfUnicodeAPI?  The only thing holding my up is the type checking for
		the wzValueName or rgMsoReg[msorid].szValue
--------------------------------------------------------------------dgray----*/
MSOAPIX_(LONG) OrapiQueryVal(int msorid, LPBYTE pbData, LPDWORD pcbData);



/*-----------------------------------------------------------------------------
	OrapiGetRid

	Given a value name, what is the rid associated with it.
--------------------------------------------------------------------dgray----*/
MSOAPIX_(BOOL) OrapiGetRidForValueEx(const char* pszValueName, const char* pszKeyName,
                                    BOOL fUseApp, DWORD *pdwMsoRid,
                                    DWORD *pdwRegType);
MSOAPI_(BOOL) OrapiGetRidForValueExW(const WCHAR *pwzValueName, const WCHAR *pwzKeyName, BOOL fUseApp, DWORD *pdwMsoRid, DWORD *pdwRegType);
#define OrapiGetRidForValue(pszValueName, pszKeyName, fUseApp, pdwMsoRid) \
	OrapiGetRidForValueEx(pszValueName, pszKeyName, fUseApp, pdwMsoRid, NULL)


/*-----------------------------------------------------------------------------
	FOfficePolicyKeyExists

	Determines if registry policy is being applied by reading the hkey.  Also
	inserts the key into the cache.
--------------------------------------------------------------------dgray----*/
BOOL FOfficePolicyKeyExists(HKEY hHive, int keyID);


/*-----------------------------------------------------------------------------
	When dealing with shared Excel and Graph code, use this to distinguish
	between places where a single reg call is used to access graph in some
	cases and excel in others.  Tag on a rid that is msorid*XL to see an
	example.
--------------------------------------------------------------------dgray----*/
#ifdef EXCEL_BUILD
#ifdef GRAF
#define GetRid(x) x##GR
#else
#define GetRid(x) x##XL
#endif
#endif


/*-----------------------------------------------------------------------------
	MsoSzRegGetKeyName
-----------------------------------------------------------------------------*/
MSOAPI_(LPCSTR) MsoSzRegGetKeyName(int msorid, int *pcchName);

/*-----------------------------------------------------------------------------
	MsoFRegGetKeyPath
-----------------------------------------------------------------------------*/
MSOAPI_(BOOL) MsoFRegGetKeyPath(int msorid, int msoridRoot, _Out_z_cap_(cchBuf) CHAR* szBuf, int cchBuf);

#ifdef DEBUG
/*-----------------------------------------------------------------------------
	MsoSzRegGetDebugRoot
-----------------------------------------------------------------------------*/
MSOAPI_(LPCSTR) MsoSzRegGetDebugRoot();
#endif //DEBUG

/*-----------------------------------------------------------------------------
	MsoFRegCheckKeyPath

	Fills a buffer with the Ansi key path to a rid.  Used to make sure
	certain paths don't get accidentally changed.  Note: This path does NOT
	include the root, that is, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc.
	So "Software\\Policies" would match "HKEY_CURRENT_USER\\Software\\Policies"
	and also "HKEY_LOCAL_MACHINE\\Software\\Policies", but that is not where
	these key mishaps that we are trying to catch actually happen, so this is
	fine.  If msoridRoot is specified, we build the string up only until that
	key.

	Returns:
		TRUE  if the rid's key matches the expected string
		FALSE if the rid's key does not match the expected string.
--------------------------------------------------------------------dgray----*/
#ifdef DEBUG
MSOAPI_(BOOL) MsoFRegCheckKeyPath(int msorid, int msoridRoot, const CHAR* sz);
#endif // DEBUG



/*----------------------------------------------------------------------------
	Experimental code to detect unused rids.
-------------------------------------------------------------------- KirKG--*/
// REVIEW: KirkG (DGray)  This is gone now, so we should rip it out, yes?
#ifdef ORAPI_RIDCHECK
#define MsoRegFGetSz(rid,b,c,d)	(ORAPI_##rid(), MsoFRegGetSzCore(rid,b,c,d))
MSOAPI_(LONG) MsoFRegGetSzCore(int msorid, LPDWORD lpType, LPBYTE lpData, LPDWORD lpcbData);
#endif


#if 0 //$[VSMSO]
/*-----------------------------------------------------------------------------
   MsoRegGetRid

   Use this function to retrieve RIDs stored in MSO.  Sometimes it is
   necessary for an application, OLE control, etc. to have a rid without
   being an ORAPI client. Store and retrieve those rids here.  Add the
   rid to ORAPI per http://officedev/tco/orapi.  Add an identifier to the
   enum list in otools\inc\mso\msotc.h and then call this function
   with that enum identifier.
--------------------------------------------------------------------DGRAY----*/
MSOAPI_(int) MsoRegGetRid(int id);
enum
	{
	ridwrapStsListMttf = 1,
	ridPersonaMenuAllowOutlookProperties,
	ridPersonaMenuAllowContact,
	ridPersonaMenuAllowCreateRule,
	ridPersonaMenuNonBuddyStatus,
	ridNameCtrlMttf,
	};
#endif //$[VSMSO]



/*-----------------------------------------------------------------------------
   MsoFlushOrapiCache

   Flush all the keys out of the ORAPI cache.  If you do this, you'll need to
   re-prime the orapi cache or all orapi H E double hockey sticks will break
   out WRT registry.  You prime the cache per whereever you are calling
   MsoFRegHookAppTables among other things.

   Be careful about using this function. Use like so:

      MsoUseOrapiStrictCriticalSections();
      MsoFRegHookAppTables(...);
      ...
      MsoLockOrapiCache();
      MsoFlushOrapiCache();
         // prime all my new keys
      MsoUnlockOrapiCach();
         // continue about my business
--------------------------------------------------------------------dgray----*/
MSOAPI_(void) MsoFlushOrapiCache();
MSOAPI_(void) MsoLockOrapiCache();
MSOAPI_(void) MsoUnlockOrapiCache();


/*-----------------------------------------------------------------------------
   MsoUseOrapiStrictCriticalSections

   Use more-strict critical section protection of ORAPI cached keys.  Used
   when cache flushing is necessary by the app.
   Normally, the cache is only protected when adding/removing items, and is
   protected using the OFFICE_CRITICAL_SECTION object.
   When cache flushing is necessary, we must also protect references to stored
   key values.  Strictly speaking this would be necessary anyway, but since
   it is a large enough temporal cache, in reality hkeys don't drop off while
   we are referencing them.  Protection of references to the cache is offered
   by the ORAPI_CRITICAL_SECTION object.
--------------------------------------------------------------------dgray----*/
MSOAPI_(void) MsoUseOrapiStrictCriticalSections();



/****************************************************************************
	Binary Policy Routines (for Word and Excel)
****************************************************************************/

/*---------------------------------------------------------------------------
	MsoPolicyApplyBinary

	Applies policy settings to bitmapped members of a structure.  Stores
	a history list of changes so that policy changes are rolled back and
	don't infect the app's user preference settings in the registry.
-------------------------------------------------------------------JoelDow-*/
MSOAPI_(void) MsoPolicyApplyBinary(void* pvStruct, int msoridKey,
	WORD wcbStruct, HANDLE* phRestore);


/*---------------------------------------------------------------------------
	MsoApplyAppBinarySettings

	Similar in concept to MsoPolicyApplyBinary, this API will apply, on a ONE
	time basis, bits to certain app structures. After it finishes, it deletes
	the key containing the bits. ( msoridKey ) Used for CMW settings.
-------------------------------------------------------------------MattP-*/
MSOAPI_(void) MsoApplyAppBinarySettings(void* pvStruct, int msoridKey,
	WORD wcbStruct);

/*---------------------------------------------------------------------------
	MsoPolicyRestoreBinary

	"Undoes" previous changes to bitmapped members of a structure to
	prevent policy settings from infecting a user's configuration
	options in the registry.
-------------------------------------------------------------------JoelDow-*/
MSOAPI_(void) MsoPolicyRestoreBinary(void* pvStruct, WORD wcbStruct, HANDLE hRestore);

#ifdef DEBUG
/*---------------------------------------------------------------------------
	MsoPolicyDumpBinary

	Outputs a list of active policy overrides for the context provided
	in hRestore.  Intended for use in a debug-only status dump.
-------------------------------------------------------------------JoelDow-*/

MSOAPI_(void) MsoPolicyDumpBinary(HANDLE fhOut, HANDLE hRestore);
#endif


/****************************************************************************
	Terminal Server (Hydra) Support/Detection
****************************************************************************/

/*------------------------------------------------------------

	MsoFIsTerminalServer

	Old API to detect Hydra.  I'm keeping this function here
	to maintain binary compatibility with all MSO clients.

	IT IS OBSOLETE.  PLEASE DO NOT USE THIS FUNCTION.

	This function will assert and ask everyone to use the
	new behavior modification API's. However, to maintain
	some semblance of backward compatibility, we will return
	assume that what the caller wants to know is whether or
	not this code is running on an AppServer, since that's
	really the only interesting way WTS 4 was used and WTS 5
	hadn't shipped when Office 9 did.

	It's just an incomplete answer, since now there are more
	"flavors" of Terminal Server, and we may want to modify
	our behavior on the Console vs. when we're running with
	graphics over the wire, etc.

----------------------------------------------- (FrankRon) -*/
MSOAPI_(BOOL) MsoFIsTerminalServer(void);



/*------------------------------------------------------------

	MsoFTSAppServer

	Use this routine to fork behavior based on whether or not
	we're running on a regular TS App Server machine (console
	or not).  TS-Lite/Remote-Admin and regular non-TS work-
	stations return FALSE here.
----------------------------------------------- (FrankRon) -*/
MSOAPI_(BOOL) MsoFTSAppServer(void);

/*------------------------------------------------------------

	MsoFEnableComplexGraphics

	Use this routine to fork behavior on animation/sound-
	intensive features (e.g., splashes) to minimize unnecessary
	graphics "candy."  Basically this is important for Hydra
	systems were we're transmitting lots of graphics bits
	over the wire, but it's also useful if the Shell ever
	implements a "Simplify Graphics" key in their Display
	Properties that may actually be exposed on all systems
----------------------------------------------- (FrankRon) -*/
MSOAPI_(BOOL) MsoFEnableComplexGraphics(void);

/*------------------------------------------------------------

	MsoFAllowIOD

	Use this routine to determine if Darwin is going to let
	us do an Install-On-Demand. Especially useful for TS.

	Basically, we allow IOD for non-TS systems, TS-Lite/RA,
	or when the user is an admin and the Policy to allow
	remote installs on TS is set.

	Note that usually Darwin will allow an admin to do an
	install at the console even if the policy is not set,
	but we don't allow this.  Dynamic installation of Office
	is generally not a good idea, and we will respect admin
	policy, but we aren't going to allow admin console install
	without the policy setting.
----------------------------------------------- (FrankRon) -*/
MSOAPI_(BOOL) MsoFAllowIOD(void);

/*------------------------------------------------------------

	MsoFRemoteUI

	Use this routine if a feature needs to know if it's UI
	is running remotely (as is the case on TS non-Console
	sessions).
----------------------------------------------- (FrankRon) -*/
MSOAPI_(BOOL) MsoFRemoteUI(void);

/*------------------------------------------------------------

	MsoFDisconnectedRemoteUI

	Use this routine if a feature needs to know if its UI
	is running remotely, AND if its session is disconnected.
	BitBlt fails on a disconnected remote session in Win.Net
	server (but not on Win2k server).
-------------------------------------------------- (TVial) -*/
MSOAPI_(BOOL) MsoFDisconnectedRemoteUI(void);

/*-----------------------------------------------------------------------------
	MsoFSupportThisEditLID

	return whether the editing of specified lang is supported
-------------------------------------------------------------------- IrfanGo -*/
MSOAPI_(BOOL) MsoFSupportThisEditLID(UINT lid);

/*----------------------------------------------------------------------------
MSOAPI_(WCHAR*) MsoWzGetAppNameFromPath(WCHAR * wzAppPath)

Given application path wzAppPath as stored in the registry
returns its name without a path
------------------------------------------------------------------ AnzhelN -*/
MSOAPI_(WCHAR*) MsoWzGetAppNameFromPath(_Inout_z_ WCHAR * wzAppPath);

/*-----------------------------------------------------------------------------
MSOAPI_(int) MsoFShowEulaOnBoot()

Returns TRUE if Eula must be shown on first boot according to sku reduction data
--------------------------------------------------------------------AnzhelN -*/
MSOAPI_(BOOL) MsoFShowEulaOnBoot(void);


/*----------------------------------------------------------------------------
MSOAPI_(void) MsoLicensingRFMAlert(void);

Displays an Alert message if command fails due to RFM restrictions
To be used only where we can't disable the command itself
------------------------------------------------------------------ AnzhelN -*/
MSOAPI_(void) MsoLicensingRFMAlert(void);


/*---------------------------------------------------------------------------
MsoFProFeatures

Returns yes if Pro features should be turned on based on licensing/sku data
------------------------------------------------------------------ AnzhelN -*/
MSOAPI_(BOOL) MsoFProFeatures(VOID);
#endif // MSOTC_H

