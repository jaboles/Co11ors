#pragma once

/****************************************************************************

	msocm.h

	Owner: MattWe
	Copyright (c) 1995 Microsoft Corporation

   Configuration Management Features for MSO96.dll

   Current features:
		MsoRegistryAdd

   Licensing API:	contact MattWe.  This feature
	has currently been cut from 96 release.

*****************************************************************************/


#ifndef __MSOCM_H__
#define __MSOCM_H__

#ifdef __cplusplus
extern "C" {
#endif

/****************************************************************************
	MsoFRegistryAdd
		Parameters:
			pwzFileName:
				UNICODE full path name of registry file
					If NULL => no action, automatic success
			pwzRootKey:
				When Non-NULL, only keys from pwzFileName beginning with this
				will be registered.  Particularly useful for choosing between
				root keys (HKEY_LOCAL_USER vs HKEY_USERS)
					If NULL => all keys from reg file used.
					If ""	=> all keys from reg file used.
		Returns:
			TRUE for success,
			FALSE for any catastrophic failure.  Additional information is returned via
				GetLastError:


				msoerrNYI				an unhandled error, that will soon
									have a more specific error.

				msoerrInvalidArg                        file not found

				msoerrTooFewBytes                       Error parsing reg file
										Premature end of file?

		Assumes:
			pwzFileName points to a readable file name, according to 
			established spec.  See \\office\config\Acme96\SelfRegistration.doc
			For the destination directory substitution, the directory will be
			determined from the directory of the .reg file specified.
			
			These are the types of data supported by the registry:
				(These are the 3 most common, in order, currerntly used data
				types.  These will be the first implemented)
				REG_SZ  
					A null-terminated string. This value will be a Unicode 
					or ANSI string depending on whether you use the Unicode or      
					ANSI functions.
				REG_DWORD       
					A 32-bit number.        


				This is just about implemented.  The tester isn't on it yet.

				REG_BINARY      
					Binary data in any form.

----------------------->	Not Yet Implemented:
	
				(These are considerably more esoteric, and will be implemented
				if needed.  Mail MattWe)
				REG_DWORD_BIG_ENDIAN    
					A 32-bit number in big-endian format. In big-endian format, 
					the most significant byte of a word is the low-order byte.
				REG_DWORD_LITTLE_ENDIAN 
					A 32-bit number in little-endian format (same as REG_DWORD). 
					In little-endian format, the most significant byte of a 
					word is the high-order byte. This is the most common format 
					for computers running Windows NT.
				REG_EXPAND_SZ   
					A null-terminated string that contains unexpanded references
					to environment variables (for example, %PATH%). This value 
					will be a Unicode or ANSI string, depending on whether you 
					use the Unicode or ANSI functions. Note that the registry 
					always stores strings internally as Unicode strings. 
				REG_LINK        
					A Unicode symbolic link.
				REG_MULTI_SZ    
					An array of null-terminated strings, terminated by two null 
					characters.
				REG_NONE        
					No defined value type.
				REG_RESOURCE_LIST       
					A device-driver resource list.
			
	
*****************************************************************************/
MSOAPI_(BOOL) MsoFRegistryAdd(LPCWSTR pwzFileName, LPCWSTR pwzRootKey);

/****************************************************************************
	MsoFRegistryDelete
		Parameters:
			pwzFileName	fully qualified path to a SRG file

			pwzRootKey	if non-null, all keys that begin with this
					substring will be processed.
		Description:
			Removes keys from the registry, and all the specified 
			key's subkeys.  Keys specified in the SRG should *not*
			be any of the root keys, as *all* of the subkeys will
			be deleted.  
*****************************************************************************/
MSOAPI_(BOOL) MsoFRegistryDelete(LPCWSTR pwzFileName, LPCWSTR pwzRootKey);

#ifdef __cplusplus
}
#endif
#endif
