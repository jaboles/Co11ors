#pragma once

/*
	Include the constant data structures and types for ORAPI.
*/


#if !defined(_ORAPI_H_)
#define _ORAPI_H_ 1

#include "msokey.inc"

//REVIEW  DGray 4?
#pragma pack(push,4)

//Roaming Types
enum {
	RoamingFalse,		// 0
	RoamingInt,			// 1
	RoamingBool,		// 2
	RoamingString,		// 3
	RoamingBinary, 		// 4
	RoamingList,			// 5
	RoamingBoolAsDigit,	// 6
	RoamingIntAsBinary,	// 7
	RoamingIntAsString, //  8
	RoamingBlob,		// 9
	RoamingTypeMax,	// 10	
};

typedef struct _RoamingAppIds
{
	char * szAppId;
}	RoamingAppIds;

typedef struct _orkey
	{
#ifdef DEBUG
	BOOL fReadOnly;
	int keyID;
	PCSTR szKeyID;
#endif
	PCSTR szKey;
	WORD cchKey;
	WORD iOrkeyParent;		// TODO dgray high bit for ReadOnly?
	} ORKEY;

typedef struct _msoreg
	{
#ifdef DEBUG
	int msorid;
	PCSTR szMsorid;
#endif
	PCSTR szRoamingRid; //szMsorid without the msorid prefix for better perf, defined only for roaming rids
	PCSTR szRoamingParent; //defined only for rids that are member of composites; very few settings have this set
	PCSTR szValue;	// value name
	union
		{
		PSTR	szDefault;
		DWORD	dwDefault;
		BYTE*	pbDefault;
		};
	WORD idKey       : 11; // currently we have 792 keys, but reserve 2048; Review: as of 07/11/2002 we have 1130 keys
	WORD nType       : 4;
	WORD fIgnoreSafe : 1;  // Get value from registry, even while in safe mode
	WORD iRoamingType : 4; //different State Apis are called depending on the type
	DWORD dwRoamingApps; //bits corresponding to apps using this rid are turned on.
	
	} MSOREG;

#include "msoreg.inc"

#pragma pack(pop)

#endif    // _ORAPI_H_

#if 0 //$[VSMSO]
extern BOOL vfPolicyExists;
#endif //$[VSMSO]

