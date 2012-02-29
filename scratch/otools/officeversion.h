//
// OfficeVersion.h
//
// Central location for updating version numbers and legal copyright strings.
//

#ifndef _OFFICEVERSION_H_
#define _OFFICEVERSION_H_


// --- Version Numbers ---

#define rmj		11
#define rmm		0
#define rup		5308
#define szVerName	"Internal Version"
#define szVerUser	"Internal User"

//  iBuildIteration starts at 0.  In the build lab, we will bump this
//  by 1 after each CMS release run commences.  Subsequent builds will
//  then be distinguishable within a given daily build cycle (Self Test
//  files will be .0 and Self Host files (that have changed) will be .1)
#define iBuildIteration 0

// --- Macros ---

#define QQ(x)	QQ2(x)
#define QQ2(x)	#x

#if defined(MANAGED_CODE_PP)
#define _COMBINE2_(x,y) x + y
#define _COMBINE3_(x,y,z) x + y + z
#else
#define _COMBINE2_(x,y) x y
#define _COMBINE3_(x,y,z) x y z
#endif

// --- Beta Toggle ---

//  NOTE:  Ready for Office 11 RTM.  Remove in Office 12 after integration.
//  #define VER_BETA
//  NOTE:  Very important to remove the ifdef BETA before ship.
//  NOTE:  VER_REMOVE_BETA is to allow one offs of sub projects, not the
//  official way to build Office as non-Beta.
#ifdef VER_REMOVE_BETA
#undef VER_BETA
#undef VER_ALPHA
#endif

// --- Default Resource Strings ---

#define VER_FILE_OS                VOS_NT_WINDOWS32

#define VER_COMPANY_NAME           "Microsoft Corporation"
#ifndef VER_PRODUCT_NAME
//#define VER_PRODUCT_NAME           "Microsoft Office 11 Beta"
#define VER_PRODUCT_NAME           "Microsoft Visual Studio"
#endif
#ifndef VER_FILE_DESCRIPTION
//#define VER_FILE_DESCRIPTION       "Microsoft Office 11 Beta component"
#define VER_FILE_DESCRIPTION       "Microsoft Visual Studio component"
#endif

// --- Manually Updated Constants ---

//  FUTURE: move more such values (e.g. BETA stuff) here from officeversion.rc
//  and elsewhere

// VER_COPYRIGHT_END is used to determine how much of the copyright text is to
// be printed by DrawSplashText in splash.cpp.  Update to keep in sync with
// actul copyright strings to be matched. (DHsu)  
// FUTURE: Have distinct copyright strings for use on Terminal Server, rather
// than generating them from original string based on this #define
#define VER_COPYRIGHT_END "2003"

// First three digits of the copyright end year.  Used to find the end of
// copyright text in a way that won't break if intl releases of Office ship in
// a different year from the core product. (JixinWu)
#define VER_COPYRIGHT_END_SHORT "200"

// VER_LEGACY_YEAR is used as first number in year.major.build.iteration format
// versions.  Update to reflect current year only until RTM.  I.e. this year
// should not be updated to current year for QFE builds occuring in years after
// RTM. (SMueller)
#define VER_LEGACY_YEAR 2003

// --- Copyright Resource Strings ---

#ifndef VER_LEGAL_COPYRIGHT_YEARS
#ifdef VER_COPYRIGHT_START
#define	VER_LEGAL_COPYRIGHT_YEARS  _COMBINE3_(VER_COPYRIGHT_START, "-", VER_COPYRIGHT_END)
#else
#define	VER_LEGAL_COPYRIGHT_YEARS  _COMBINE2_("1983-", VER_COPYRIGHT_END)
#endif
#endif

#ifndef RC_INVOKED
#if defined(VER_LEGAL_COPYRIGHT)
#error VER_LEGAL_COPYRIGHT unexpectedly defined
#endif
#endif // RC_INVOKED

#if defined(MANAGED_CODE_PP)
#define LEGAL_COPYRIGHT            _COMBINE3_("Copyright \u00A9 ", VER_LEGAL_COPYRIGHT_YEARS, " Microsoft Corporation.  All rights reserved.")
#else
#define LEGAL_COPYRIGHT            _COMBINE3_("Copyright \251 ", VER_LEGAL_COPYRIGHT_YEARS, " Microsoft Corporation.  All rights reserved.")
#endif

#if defined(VER_BETA) || defined(VER_ALPHA)
#define VER_LEGAL_COPYRIGHT        _COMBINE2_("Unpublished work.  ", LEGAL_COPYRIGHT)
#else
#define VER_LEGAL_COPYRIGHT        LEGAL_COPYRIGHT
#endif

// --- Legal Trademark Strings ---

#ifndef RC_INVOKED
#if defined(VER_LEGAL_TRADEMARK1)
#error VER_LEGAL_TRADEMARK1 unexpectedly defined
#endif
#if defined(VER_LEGAL_TRADEMARK2)
#error VER_LEGAL_TRADEMARK2 unexpectedly defined
#endif
#if defined(VER_LEGAL_TRADEMARK10)
#error VER_LEGAL_TRADEMARK10 unexpectedly defined
#endif
#if defined(VER_LEGAL_TRADEMARK11)
#error VER_LEGAL_TRADEMARK11 unexpectedly defined
#endif
#endif // RC_INVOKED

#if defined(MANAGED_CODE_PP)
#define VER_LEGAL_TRADEMARK1       "Microsoft\u00AE is a registered trademark of Microsoft Corporation."
#define VER_LEGAL_TRADEMARK2       "Windows\u00AE is a registered trademark of Microsoft Corporation."
#define VER_LEGAL_TRADEMARK        _COMBINE3_(VER_LEGAL_TRADEMARK1, " ", VER_LEGAL_TRADEMARK2)
#else
#define VER_LEGAL_TRADEMARK1       "Microsoft\256 is a registered trademark of Microsoft Corporation."
#define VER_LEGAL_TRADEMARK2       "Windows\256 is a registered trademark of Microsoft Corporation."
#endif

#endif // _OFFICEVERSION_H_
