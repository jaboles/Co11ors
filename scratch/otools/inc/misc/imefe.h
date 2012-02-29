/*----------------------------------------------------------------------------
	%%File: IMEFE.H
	%%Unit: IMEFE
	%%Contact: seijia

	Date		Alias		Description
	====		=====		===========
	10/22/97	seijia		Initial definition

	version independent FarEast common definitions for IME COM 
----------------------------------------------------------------------------*/

#ifndef __IMEFE_H__
#define __IMEFE_H__

#include "fecom.h"
#include "felang.h"
#include "fedict.h"
#include "imeapp.h"


///////////////////////////
// HKEY_CLASSES_ROOT values 
///////////////////////////

#define szImeJapan			"MSIME.Japan"
#define szImeKorea			"MSIME.Korea"
#define szImeChina			"MSIME.China"
#define szImeTaiwan			"MSIME.Taiwan"


////////////////
// CLSID and IID
////////////////

// Interface ID for IFECommon
// {019F7151-E6DB-11d0-83C3-00C04FDDB82E}
DEFINE_GUID(IID_IFECommon, 
0x19f7151, 0xe6db, 0x11d0, 0x83, 0xc3, 0x0, 0xc0, 0x4f, 0xdd, 0xb8, 0x2e);

// Interface ID for IFELanguage
// {019F7152-E6DB-11d0-83C3-00C04FDDB82E}
DEFINE_GUID(IID_IFELanguage, 
0x19f7152, 0xe6db, 0x11d0, 0x83, 0xc3, 0x0, 0xc0, 0x4f, 0xdd, 0xb8, 0x2e);

// Interface ID for IFELanguage2
// {21164102-C24A-11d1-851A-00C04FCC6B14}
DEFINE_GUID(IID_IFELanguage2,
0x21164102, 0xc24a, 0x11d1, 0x85, 0x1a, 0x0, 0xc0, 0x4f, 0xcc, 0x6b, 0x14);

// Interface ID for IFEDictionary
// {019F7153-E6DB-11d0-83C3-00C04FDDB82E}
DEFINE_GUID(IID_IFEDictionary, 
0x19f7153, 0xe6db, 0x11d0, 0x83, 0xc3, 0x0, 0xc0, 0x4f, 0xdd, 0xb8, 0x2e);


#endif //__IMEFE_H__
