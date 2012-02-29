#pragma once

/*************************************************************************
	mso.h

	Owner: petero
	Copyright (c) 1994-1997 Microsoft Corporation

	The main Office header file.
*************************************************************************/

#ifndef MSO_H
#define MSO_H

// Basic platform defines, etc.
#include <msostd.h>

#ifndef RC_INVOKED
	/* REVIEW: what should these pack values be? */
	#pragma pack(push,8)
#endif

// Office GUIDs
#include <msoguids.h>

// Office-global initialization and interfaces
#include <msouser.h>

#ifdef STATIC_LIB_DEF
#include <msolib.h>
#endif

// Office Debugging
#include <msodebug.h>

// Office memory and global variables
#include <msoalloc.h>

// Office strings
#include <msostr.h>

// Office dynamic strings
#include <msodstr.h>

// Office resources
#include <msores.h>

// Office error messages
#include <msoerr.h>

// Office Toolbars
#include <msotb.h>

// Office UI
#include <msoui.h>

// Office palette management
#include <msohpal.h>

// Office Drawing
#ifndef RC_INVOKED
#include <msodr.h>
#endif

// Office Assistant
#include <msofc.h>

// Office Component Integration
#include <msoci.h>

#include "mso95api.h"

// Office Internationalization stuff
#include <msointl.h>

// Office Configuration Management
#include <msocm.h>

// Web client stuff
#include <msowc.h>

// Office Url stuff
#include <msourl.h>

// Office TCO
#include <msotc.h>

// Trident in a dialog
#include <msooch.h>

#ifndef ZENSTAT_LIB_DEF
// Office Complex Scripts APIs
#include <msousp.h>
#endif //ZENSTAT_LIB_DEF

// Office Unicode wrapper
#include <msowrap.h>

// Office Unicode stuff
#include <msounicd.h>

// Office Programmability
#include <msopm.h>

// Office Data Integration (including Mail Merge)
//#include <msodata.h>

// Office Asserts
#include <msoassert.h>

// Office On-Obj UI
#include <msoooui.h>

// Office Logging
#include <msouls.h>

// Security related APIs
#include <msosecure.h>

#ifndef RC_INVOKED
	#pragma pack(pop)
#endif

#endif // MSO_H
