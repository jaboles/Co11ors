#pragma once

 // msosrver.h - Public header for translating server errors
//  
//	Owner: ThomasOl
//	Copyright (c) 2000 Microsoft Corporation
//
//
//
//

#ifndef __MSOSRVER_H__
#define __MSOSRVER_H__

// ----------------------------------------------------------------------------
// Types and constants.

typedef enum
	{
	msoSrvErrNoErr = -1,
	msoSrvErrUnknown = 0,
	msoSrvErrBadUrl,
	msoSrvErrRenameDupeErr,
	msoSrvErrInvalidDirectory,			//"Invalid directory"
	msoSrvErrSaveFileExistsAsDir,
	msoSrvErrSourceFileConflict,		//" The file in your FrontPage web conflicts with the file in the source control repository.  Rename your file to another file name, then use the Recalculate Hyperlinks command to see the repository's version, then checkout the file and merge the conflicting files manually."
	msoSrvErrFileAlreadyExists,			//<filename> already exists.  Please enter a different file name.
	msoSrvErrFileCannotBeAccessed,		//<filename> cannot be accessed.  The file may not exist or may be in use or the web server is temporarily busy.
	msoSrvErrFileCannotBeCreated,		//<filename> cannot be created.  The file may already exist or the web server is temporarily busy.
	msoSrvErrFileCannotBeRemoved,		//<filename> cannot be removed.  The file may not exist or may be in use or the web server is temporarily busy.
	msoSrvErrFileCannotBeRenamed,		//<filename> cannot be renamed.  The file may not exist or may be in use or the web server is temporarily busy.
	msoSrvErrFileDoesNotExist,			//<filename> does not exist.  Please check your typing.
	msoSrvErrFileAutoCheckOut,			//<filename> has been checked out automatically.
	msoSrvErrFileUnderSourceControl,	//<filename> is already under source control.
	msoSrvErrFileCheckedOut,			//<filename> is checked out by someone else.
	msoSrvErrFileInUse,					//<filename> is currently in use.  Please try again later.
	msoSrvErrFileNotCheckedOut,			//<filename> is not checked out.
	msoSrvErrFileNotUnderSourceControl,	//<filename> is not under source control.
	msoSrvErrFolderCannotBeCreated,		//<folder> cannot be created.  The folder may already exist or the web server is temporarily busy.
	msoSrvErrFolderCannotBeRemoved,		//<folder> cannot be removed.  The folder may not exist or may be in use or the web server is temporarily busy.
	msoSrvErrFolderDoesNotExist,		//<folder> does not exist.  Please check your typing.
	msoSrvErrCreateFolder,				//<folder> does not exist.  Please create the folder and try again.
	msoSrvErrNotAFolder,				//<folder> is not a folder.  Please check your typing.
	msoSrvErrFPWebNotConfigured,		//A FrontPage web component is incorrectly configured.
	msoSrvErrThemeAlreadyExists,		//A theme already exists on the server.
	msoSrvErrCannotCreateWeb,			//A web cannot be created at this web address.
	msoSrvErrCannotAccessWeb,			//An error occurred accessing your FrontPage web files.  Please contact the server administrator.
	msoSrvErrAnonUploadNotAvailable,	//Anonymous upload is not available on this web server.
	msoSrvErrCannotOverwriteItself,		//Cannot copy <filename> to itself.
	msoSrvErrCannotCopyFiles,			//Cannot copy files.
	msoSrvErrCannotCopyFolder,			//Cannot copy folder <folder1> to folder <folder2>.
	msoSrvErrCannotLinkFiles,			//Cannot link files.
	msoSrvErrCannotMarkFolderExecutable,//Cannot mark a folder executable on this server.
	msoSrvErrCannotRenameLinks,			//Could not rename links in <filename>; the file is currently in use.
	msoSrvErrFilesCannotBeSavedFolder,	//Files cannot be saved to this folder.
	msoSrvErrCannotRecalcLinks,			//Links within this file could not be recalculated.
	msoSrvErrExclusiveCheckoutRequired,	//Only exclusive checkout  is supported with FrontPage-based locking.
	msoSrvErrSubwebRequiresRefresh,		//"Operation interrupted because a subweb was created, renamed, or deleted during the operation.  You may need to refresh your view.  Please try again."
	msoSrvErrIEOffline,					//"Page cannot be accessed because Microsoft Internet Explorer is offline.  Run Internet Explorer, uncheck the File menu's ""Work Offline"" item, and retry."
	msoSrvErrFilesAutoCheckout,			//Some files have been checked out automatically. The FrontPage server extensions on this web do not support getting previous versions.
	msoSrvErrServerNotAvailable,		//The web server is not available at this time.  Please try again later.
	msoSrvErrFolderLimitExceeded,		//"The size of files uploaded has exceeded, or may exceed, the limit set for the folder. "
	msoSrvErrWebAddressTooLong,			//The web address is too long.
	msoSrvErrWebConfigChanged,			//The web configuration has changed.  Please refresh your view (F5) and try again.
	msoSrvErrWebServerBusy,				//The web server is busy.  Please try again later.
	msoSrvErrWebServerProblem,			//There is a problem with the web server.  Please try again later or contact the server administrator.
	msoSrvErrNotEnoughDiskSpace,		//There is not enough disk space.
	msoSrvErrCollabFeaturesNotAvailable,//This feature requires the Microsoft Office Collaboration Extensions to be installed on the web server.
	msoSrvErrFolderCannotBeModified,	//"This folder containing supporting files.  It cannot be moved, copied, renamed or deleted."
	msoSrvErrBrowseOnlyWeb,				//This is a browse-only web.  Please enter another web address.
	msoSrvErrFolderContainsSpaces,		//This server doesn't support folder names that contain spaces.
	msoSrvErrCannotChangePermissions,	//Unable to change web permissions.
	msoSrvErrCannotConnectToServer,		//Unable to connect to the web server.
	msoSrvErrCannotConvertWebToFolder,	//Unable to convert the web into a folder because it does not use the same permissions as its parent web.
	msoSrvErrCannotCopyFolderWithSubweb,//Unable to copy folder.  Folders with subwebs cannot be copied.
	msoSrvErrCannotDeleteFolderWithSubweb,//Unable to delete folder.  Folders with subwebs cannot be deleted.
	msoSrvErrCannotDeleteOrConvertWeb,	//Unable to delete web or convert it into a folder.
	msoSrvErrCannotDeleteWeb,			//Unable to delete web.  Webs with subwebs cannot be renamed.
	msoSrvErrCannotRenameFolderWithSubweb,//Unable to rename folder.  Folders with subwebs cannot be renamed.
	msoSrvErrCannotRenameWeb,			//Unable to rename web.
	msoSrvErrCannotRenameWebWithParentUrl,//Unable to rename web.  The parent URL cannot be changed when renaming a web.
	msoSrvErrCannotRenamewebWithSubweb,	//Unable to rename web.  Webs with subwebs cannot be renamed.
	msoSrvErrSourceSafeMustBeAbsolute,	//"Visual Source Safe projects must be absolute (they must start with ""$/"") and have no other embedded ""*"", ""$"", ""\\\\"", or ""?"" characters."
	msoSrvErrCannotUploadEmptyFile,		//You attempted to upload an empty file.  Please check your typing.
	msoSrvErrCannotConvertVirtDir,		//You cannot convert a virtual directory into a subweb remotely.
	msoSrvErrCannotAnonUploadToRootWeb,	//You cannot use anonymous upload to the root of a web.  Please use a subfolder for anonymous uploads.
	msoSrvErrIncorrectPermissions,		//You do not have the correct permissions.  Please contact the server administrator.
	msoSrvErrClientTooOld,				//The version of the FrontPage Server Extensions running on the server is more recent than the version of FrontPage you are using.  You need a more recent version of FrontPage.
	msoSrvErrVersioningNotSupported,		//The FrontPage server extensions on this web do not support getting previous versions.

	msoSrvErrOverQuota,             // You are over your quota.
	msoSrvErrBlockedFileType,       // The file type is blocked.

	msoSrvErrVirusInfectedUpload,
	msoSrvErrVirusCleanedUpload,
	msoSrvErrVirusInfectedDownload,
	msoSrvErrVirusCleanedDownload,
	msoSrvErrVirusDeletedDownload,
	msoSrvErrVirusBlockedDownload,

	msoSrvErrLockTimeMismatch, // The lock is not valid on the server anymore.
	msoSrvErrBadChars, // The url contained invalid characters (determined by the server)
	msoSrvErrFileTooLarge, // The file being uploaded is too large


	} MSOSRVERRTYPE;

MSOAPI_(MSOSRVERRTYPE) MsoServerErrorTypeGet(BSTR * pBstr);

MSOAPI_(void) MsoServerErrorAlert(MSOSRVERRTYPE errType, const WCHAR *wzArg1, const WCHAR *wzArg2);

#if 0 //$[VSMSO]
MSOAPI_(BOOL) MsoExtractVirusInfo(MSOSRVERRTYPE err, const BSTR bstr, WCHAR *rgch, int cch);
#endif //$[VSMSO]

#endif //__MSOSRVER_H__
