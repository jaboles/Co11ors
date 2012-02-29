/****************************************************************************
	Sv.h

	Owner: JBelt
 	Copyright (c) 1999 Microsoft Corporation

	Declarations for functions related to Services on the Web
****************************************************************************/

#ifndef SV_H
#define SV_H 1

#ifndef STATIC_LIB_DEF
#include "msosvtl.h"
#endif
#ifndef OFFICE_BUILD
#define OFFICE_BUILD
#endif
#include "msostd.h"

MSOAPI_(BOOL) MsoFCreateWDWizard(HMODULE hModule, interface IMsoWDWizard **ppiwdw);

enum
{
svnavNA = 0,
svnavBack,
svnavComplete,
svnavCancelled,
svnavError,
};

#endif // SV_H

