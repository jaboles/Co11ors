/*----------------------------------------------------------------------------
	%%File: IMEJPN.H
	%%Unit: IMEJPN
	%%Contact: seijia

	Japanese specific definitions

	Date		Alias		Description
	====		=====		===========
	06/17/97	seijia		Initial definition
----------------------------------------------------------------------------*/

#ifndef __IMEJPN_H__
#define __IMEJPN_H__

#include "imefe.h"


///////////////////////////
// HKEY_CLASSES_ROOT values 
///////////////////////////

#define szImeJapan98		"MSIME.Japan.6"


////////////////
// CLSID and IID
////////////////

// Class ID for FE COM interfaces
// {019F7150-E6DB-11d0-83C3-00C04FDDB82E}
DEFINE_GUID(CLSID_MSIME_JAPANESE, 
0x19f7150, 0xe6db, 0x11d0, 0x83, 0xc3, 0x0, 0xc0, 0x4f, 0xdd, 0xb8, 0x2e);


#endif //__IMEJPN_H__
