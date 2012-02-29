#pragma once
/****************************************************************************
	MSOWC.H

	Owner: cathysax
	Copyright (c) 1997 Microsoft Corporation

	Exported declarations for all Office Web Client stuff.
****************************************************************************/

#ifndef MSOWC_H
#define MSOWC_H

#include <shlwapi.h>

/*-----------------------------------------------------------------------------
	MsoHrCLSIDFromString
-------------------------------------------------------------------- HAILIU -*/
MSOAPI_(HRESULT) MsoHrCLSIDFromString(LPCOLESTR wz, CLSID *pclsid);

#endif // MSOWC_H
