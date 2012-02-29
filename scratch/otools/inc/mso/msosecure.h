/******************************************************************************
    MSOSECURE.H

    See http://officehack/hack.htm for Office security information
    Security questions - mailto:odevhack - Office Security Team

    Owner: DGray
    Copyright (c) 2002 Microsoft Corporation

******************************************************************************/

#pragma once
#ifndef MSOSECURE_H
#define MSOSECURE_H
#include <winnt.h>     // PACL, ACCESS_MASK



// SECURITY

// Exclude the sorts of rights that you would not want to grant to EVERYONE.
// These values are meant to be &-ed with other values.
// Example: (KEY_ALL_ACCESS & SECURE_ACCESS)
// See Writing Secure Code, 1st Edition, p. 110:

#define SECURE_ACCESS ~(WRITE_DAC | WRITE_OWNER | DELETE | GENERIC_ALL | ACCESS_SYSTEM_SECURITY)

//
// NULL DACLs
//
// For details on NULL DACLs, see Writing Secure Code, 1st Ed, p.108
// A NULL DACL is when NULL is used for the PACL of a security descriptor.
// It commonly looks like this:
//  1. SetSecurityDescriptorDacl(&sd, TRUE, NULL, FALSE);
//                    There is a DACL--/      \---- DACL is NULL
//  or
//  2. SECURITY_DESCRIPTOR sd = {
//      SECURITY_DESCRIPTOR_REVISION, 0x0, SE_DACL_PRESENT, 0x0,0x0,0x0,0x0};
//                             DACL present-----/    NULL DACL-----------/
//
// When the security descriptor itself is NULL, then the default security
// descriptor is inherited.  This typically looks like this, using psa as
// the security attributes input later:
//     SECURITY_ATTRIBUTES* psa;
//  1. SECURITY_ATTRIBUTES sa = {sizeof(SECURITY_ATTRIBUTES), NULL, FALSE};
//            lpSecurityDescriptor is NULL, use default--------/
//     psa = &sa;
//  or
//  2. psa = NULL;
//
// There is a lot of legacy code that just uses the default because it just
// works.  This may cause unintended access denied error messages.  To 
// explicitly state your intentions to use the default security descriptor, use
// these #defines to specify that you really do want the default.
#define DEFAULT_SECURITY_DESCRIPTOR NULL
#define DEFAULT_SECURITY_ATTRIBUTES NULL


MSOAPI_(PACL) MsoCreateAcl(const ACCESS_MASK rggrfDeny[], int cDeny, const ACCESS_MASK rggrfAllow[], int cAllow, PSID rgpsid[]);

MSOAPIX_(PACL) MsoCreateEveryoneAcl(DWORD grfPermissions);
MSOAPIX_(PACL) MsoCreateUserAcl(DWORD grfPermissions);


/*-----------------------------------------------------------------------------
   MsoFCreateImpersonatedAcl

   Create an ACL with the specified permissions for the current thread 
   impersonated user.

   http://msdn.microsoft.com/library/en-us/security/Security/windows_2000_windows_nt_access_mask_format.asp

   It is recommended that whatever permissions you grant
   be ANDed with SECURE_ACCESS.  (See %OTOOLS%\inc\mso\msostd.h)

   Returns FALSE on failure.
   Returns TRUE and fills *ppacl on success.

   On success, the Acl should be free'd with GlobalFree.
--------------------------------------------------------------------dgray----*/
MSOAPIX_(BOOL) MsoFCreateImpersonatedAcl(PACL* ppacl, DWORD grfPermissions);


//
// Standard office SECURITY_ATTRIBUTES structures for standard user groups
//

/*-----------------------------------------------------------------------------
   MsoPsaEveryoneDacl

   Returns the security attributes structure for the "Everyone" user
   minus dangerous SECURE_ACCESS rights.

   If fAllAccess flag is set then we allow ALL_ACCESS.
   
   !!! DO NOT SET fAllAccess TO TRUE TO BYPASS SECURE_ACCESS !!!  This flag has
   been added to allow for certain backwards compatibility scenarios with older
   versions of Office (ex. Ole menu merging).
--------------------------------------------------------------------dgray----*/
MSOAPI_(LPSECURITY_ATTRIBUTES) MsoPsaEveryoneDacl(BOOL fAllAccess);

/*-----------------------------------------------------------------------------
   MsoPsaUserDacl

   Returns the security attributes structure for the process owner user
   minus dangerous SECURE_ACCESS rights.
--------------------------------------------------------------------dgray----*/
MSOAPI_(LPSECURITY_ATTRIBUTES) MsoPsaUserDacl();

/*-----------------------------------------------------------------------------
   MsoPsaInteractiveUserDacl

   Returns the security attributes structure for interactive desktop users
   minus dangerous SECURE_ACCESS rights.
--------------------------------------------------------------------dgray----*/
MSOAPI_(LPSECURITY_ATTRIBUTES) MsoPsaInteractiveUserDacl();

/*-----------------------------------------------------------------------------
   MsoPsaAdministratorsDacl

   Returns the security attributes structure for the Administrators group
   minus dangerous SECURE_ACCESS rights.
--------------------------------------------------------------------dgray----*/
MSOAPI_(LPSECURITY_ATTRIBUTES) MsoPsaAdministratorsDacl();

/*-----------------------------------------------------------------------------
   MsoFGetCurrentImpersonatedUserSa

   Allocate a security attributes object and fill it with the information for
   the current impersonated user.

   Success :  Return TRUE  and fills lppSecurity with security object
   Failure :  Return FALSE and fills lppSecurity with NULL.

   You must call MsoFreeImpersonatedUserSa on lppSecurity when you are
   finished with it, like so;

   if (MsoFGetCurrentImpersonatedUserSa(&lpSecurity))
      {
      ...
      MsoFreeImpersonatedUserSa(&lpSecurity);
      }
--------------------------------------------------------------------dgray----*/
MSOAPIX_(BOOL) MsoFGetCurrentImpersonatedUserSa(LPSECURITY_ATTRIBUTES* lppSecurity);



/*-----------------------------------------------------------------------------
   MsoFreeImpersonatedUserSa

   Free an Sa that was allocated by MsoFGetCurrentImpersonatedUserSa
--------------------------------------------------------------------dgray----*/
MSOAPIX_(void) MsoFreeImpersonatedUserSa(LPSECURITY_ATTRIBUTES* lppSecurity);


/*-----------------------------------------------------------------------------
   MsoGetUserSid

   Get the Sid for the process user.

   Warning: The returned PSID should NOT be freed.

   DON'T USE THIS API, use MsoFGetProcessUserSid instead.
--------------------------------------------------------------------jeffmit--*/
MSOAPI_(PSID) MsoGetUserSid(void);



/*-----------------------------------------------------------------------------
   MsoFGetProcessUserSid

   Get the Sid for the process user.

   Warning: The returned PSID should NOT be freed.

   This can fail if the user sid is not yet cached and you are currently
   impersonating another user.
--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFGetProcessUserSid(PSID* ppSid);


/*-----------------------------------------------------------------------------
   MsoFGetImpersonatedUserSid

   Get the SID of the current impersonated user.

   This Sid should be freed with GlobalFree when you are finished with it.
--------------------------------------------------------------------dgray----*/
MSOAPIX_(BOOL) MsoFGetImpersonatedUserSid(PSID* ppsid);


/*-----------------------------------------------------------------------------

   Flags for named objects, used with APIs below

--------------------------------------------------------------------dgray----*/

// 
// Consider the following scenarios:  NO = named object
//
// Scenario A: User1 logs into machine1 and uses NO.  User1 starts a second
//            process that uses NO in the same session. (easy case)
// Scenario B: User1 logs into machine1 and uses NO.  User1 connects a
//            second session using Terminal services to machine1 and uses NO.
//            Should the same NO be used across terminal server sessions?
//              yes - use MNO_GLOBAL
//              no  - use MNO_LOCAL
// Scenario C: User1 logs into machine1 and uses NO.  User1 uses runas.exe to
//            spawn an application as User2.  User2 uses NO.
//            Should the NO be the same for both users?
//              yes - use MNO_NAME_EVERYONE, use MNO_PERMISSION_EVERYONE or
//                      MNO_PERMISSION_INTERACTIVE_USER.
//              no  - use MNO_NAME_USER, use MNO_PERMISSION_USER or other.
// Scenario D: User1 logs into machine1 console.  User2 logs into machine1
//            terminal server.  What resources should they share?
//              consider EVERYONE flags.
// Scenario E: User1 logs into machine1 console.  Service running as
//            system on machine.  What resources should they share?
//              consider EVERYONE flags.
//
// Scenario F: Consider legacy applications.  If you are naming your objects
//              using the same convention in the past, these APIs will create
//              the named object with SECURE_ACCESS.  That means that
//              attempting to open a named object by creating it will probably
//              fail.
//              Example - CreateMutex.  If the mutex exists, CreateMutex
//              attempts to open it with MUTEX_ALL_ACCESS.  Office creates
//              a shared mutex MSO97BStripMutex.  Since VisualStudio uses
//              the mso.dll, the name of the newly secured object needs to
//              change so that VisualStudio doesn't fail 0x5 on MutexCreate.

enum 
	{
	// A named object can be global or local in scope.  If it is global, then
	// it is shared across all terminal server sessions.  If it is local, then
	// it applies only to the current terminal server session.
	// Global named objects are prepended with "Global\"
	// Local named objects are prepended with "Local\"
	MNO_GLOBAL               = (1 <<  0),
	MNO_LOCAL                = (1 <<  1),

	// If the NAME_USER flag is set, then a user identifier will be appended to
	// the name.  This means that 2 named objects running under 2 different
	// user contexts will have different names.
	// e.g. Suppose Word runas DGray, Excel runas ChakMLi.  Specifying the 
	// MNO_NAME_USER flag indicates that a user identifier should be injected 
	// into a mutex name so that there are 2 different mutexes.
	// If MNO_NAME_EVERYONE is specified, then DGray and ChakMLi will attempt
	// to use the same mutex, with the same name.
	MNO_NAME_EVERYONE          = (1 <<  3), // "Global\MyMutex"
	MNO_NAME_USER              = (1 <<  4), // "Global\MyMutex_S-1-4-1234-5678"
	MNO_NAME_IMPERSONATED_USER = (1 <<  5), // "Global\MyMutex_S-1-4-2222-5678"
	// tack on the office version to the end of the name. MajorMinorBld
	// 11.0.4311 = 1104311
	MNO_NAME_OVERSION        = (1 <<  6),
	// It is sometimes desireable to have an object without a name
	// Typically the object is created and the handle is passed
	// A malicious user can't squat on an unnamed object
	MNO_NAME_UNNAMED         = (1 <<  7),
	// Add the %ApplicationName% to the end of the mutex name,
	//   instances of 'Microsoft Access' will share the named object
	MNO_NAME_APPLICATION     = (1 <<  8),	
	
	// NYI   (FUTURE: implement and add to valid flags list)
	// By default, the named object should have a machine specific identifier
	// appended to the name.  The identifier should not be predictable by
	// other machines.
	//MNO_NAME_NOCANARY      = (1 <<  9),

	// Uses the name exactly as is...  However, the name expected should be
	// of the following forms (case sensitive on prefixes):
	//		Local\BaseName
	//		Global\BaseName
	// Anything else will assert in debug...
	MNO_NAME_DONTMODIFY      = (1 << 10),
	MNO_NAME_FLAGS       = (MNO_NAME_USER | MNO_NAME_EVERYONE |
	                        MNO_NAME_OVERSION | MNO_NAME_UNNAMED |
	                        MNO_NAME_APPLICATION | MNO_NAME_DONTMODIFY |
	                        MNO_NAME_IMPERSONATED_USER),

	// Named Object Permissions
	// Share this mutex with all users
	// (Everyone Full control minus SECURE_ACCESS)
	MNO_PERMISSION_EVERYONE          = (1 << 12),
	// Object for use with with current user only
	// (User Full control minus SECURE_ACCESS)
	MNO_PERMISSION_USER              = (1 << 13),
	// Object for use with with current desktop interactive users
	// (Interactive Users Full control minus SECURE_ACCESS)
	MNO_PERMISSION_INTERACTIVE_USER  = (1 << 14),
	// Object for use with with builtin Administrators group
	// (Administrators Full control minus SECURE_ACCESS)
	MNO_PERMISSION_ADMINISTRATORS    = (1 << 15),
	// Set permission to Create with ALL_ACCESS and Open with SECURE_ACCESS.
	// This currently should only be used MNO_PERMISSION_EVERYONE.
	// !!!DO NOT ADD THIS FLAG TO BYPASS SECURE_ACCESSS!!!  This flag is used
	// to allow for backwards compatibility with older versions
	// of OFFICE: O11#231885
	MNO_PERMISSION_TRANSITIONAL      = (1 << 16),
	// Use the permissions specified in lpSecurity
	MNO_PERMISSION_SPECIFIED         = (1 << 17),
	// Permission the object as the currently impersonated user on the
	// current thread.  If the thread is not impersonating, use the process
	// user permissions.
	MNO_PERMISSION_IMPERSONATED      = (1 << 18),

	MNO_PERMISSION_FLAGS = (
		  MNO_PERMISSION_EVERYONE
		| MNO_PERMISSION_USER
		| MNO_PERMISSION_INTERACTIVE_USER
		| MNO_PERMISSION_ADMINISTRATORS
		| MNO_PERMISSION_TRANSITIONAL
		| MNO_PERMISSION_SPECIFIED
		| MNO_PERMISSION_IMPERSONATED
		),

	// Handle options
	// If set, a process created by CreateProcess can inherit this handle
	// otherwise the handle can not be inherited
	MNO_INHERITHANDLE                = (1 << 23),

	MNO_VALID_FLAGS      = ( 0
		| MNO_GLOBAL
		| MNO_LOCAL
		| MNO_NAME_FLAGS
		| MNO_PERMISSION_FLAGS
		| MNO_INHERITHANDLE
		 ),
	MNO_INVALID_FLAGS    = (~(MNO_VALID_FLAGS)),
	// first 24 bits for generic named object settings
	MNO_GENERIC_FLAGS    = 0x00FFFFFF,
	// last 8 bits for specific object settings
	MNO_SPECIFIC_FLAGS   = 0xFF000000,

	};



/*-----------------------------------------------------------------------------
   MsoAssertMNOFlags

   Check that the flags aren't conflicting
--------------------------------------------------------------------dgray----*/
#ifdef DEBUG
MSOAPIX_(void) MsoAssertMNOFlags(DWORD grfMNO);
#else // !DEBUG
#define MsoAssertMNOFlags(grfMNO)
#endif

/*-----------------------------------------------------------------------------
   MsoFMNOResolveSecurity

   Return the appropriate security attributes pointer based on grfMNO
   lpSecurity is returned if MNO_PERMISSION_SPECIFIED is in grfMNO.

   Must be used in conjuction with MsoMNOReleaseSecurity, like so:

      if (MsoFMNOResolveSecurity(grfMNO, &lpSecurity))
         {
         ...
         MsoMNOReleaseSecurity(grfMNO, &lpSecurity);
         }

--------------------------------------------------------------------dgray----*/
MSOAPI_(BOOL) MsoFMNOResolveSecurity(DWORD grfMNO,
	LPSECURITY_ATTRIBUTES* lppSecurity);



/*-----------------------------------------------------------------------------
   MsoMNOReleaseSecurity

   Closes and Frees security objects returned from MsoFMNOResolveSecurity
--------------------------------------------------------------------dgray----*/
MSOAPI_(void) MsoMNOReleaseSecurity(DWORD grfMNO,
	LPSECURITY_ATTRIBUTES* lppSecurity);



/*-----------------------------------------------------------------------------
   MsoSzCreateNamedObjectNameSz

   Fill the szOut buffer with an appropriate name for a named object, using 
   flags for named objects above.

   Returns a pointer that should be used as the name of the object, which is
   usually to the start of the passed in buffer, but will return NULL if the
   object should be unnamed.

   String is "<Global|Local>\<szBaseName><oversion>[_<SID>]" or NULL if unnamed

   e.g.:
     szBaseName = SqmMutex, grfMNO = MNO_GLOBAL | MNO_NAME_USER | MNO_PERMISSION_USER
        output will look like "Global\SqmMutex_S-1-2-12345-67890"
     szBaseName = HeapMutex, grfMNO = MNO_LOCAL | MNO_NAME_EVERYONE | MNO_PERMISSION_EVERYONE
        output will look like "Local\HeapMutex"
     szBaseName = HeapMutex, grfMNO = MNO_LOCAL | MNO_NAME_EVERYONE | MNO_NAME_OVERSION | MNO_PERMISSION_EVERYONE
        output will look like "Local\HeapMutex1104311"
     szBaseName = NULL, grfMNO = MNO_NAME_UNNAMED | MNO_PERMISSION_EVERYONE
        output will look like "" and return NULL
--------------------------------------------------------------------dgray----*/
MSOAPI_(char*) MsoSzCreateNamedObjectNameSz(_Out_z_cap_(cchOut) char* szOut, DWORD cchOut, 
	const char* szBaseName, DWORD grfMNO);

// semi-arbitrary max for cch of named objects
enum { CCH_NAMED_OBJECT_MAX = 1024 };


#endif // MSOSECURE_H
