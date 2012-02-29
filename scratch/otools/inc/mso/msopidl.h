/****************************************************************************
	Msopidl.h

	Owner: JBelt
 	Copyright (c) 1999 Microsoft Corporation

	Shared structures between Office and NSE
****************************************************************************/
#ifndef MSOPIDL_H
#define MSOPIDL_H


#define HTTPID_LEAD 		0x504c
#define HTTPID_TRAIL		0x5742

	// The pidl version.  Don't change this!
#define HTTPID_VER		0x0

#define HTTPID_PLACE		0x01	/* A place */
#define HTTPID_URL			0x02	/* A real item, either file or folder */
#define HTTPID_PLACEWIZ		0x04	/* The places wizard item */
#define HTTPID_NOEXIST		0x08	/* An item that looks valid but doesn't exist on the server */
#define HTTPID_WEBDRIVE		0x10	/* WebDrive */
#define HTTPID_WDPLACE		(HTTPID_WEBDRIVE | HTTPID_PLACE)
#define HTTPID_WDWIZ			(HTTPID_WEBDRIVE | HTTPID_PLACEWIZ)
#define HTTPID_ABSURL		(HTTPID_ABSOLUTE | HTTPID_URL)
#define HTTPID_ABSOLUTE		0x20	/* The url is an absolute URL */
#define HTTPID_TAHOE		0x40

#define FILE_ATTR_COMP_MASK	FILE_ATTRIBUTE_DIRECTORY

#ifdef OFFLINE
#define FILE_ON_OFFLINELIST	0x80000000 /* is file on Ragent's offline list? */
#define OFFLINE_STATE_FRESH	0x20000000 /* the information given by FILE_ON_OFFLINELIST is current */
#endif

#define FOLDER_IS_SUBWEB		0x40000000 /* the folder is an fp subweb */
#define FILE_THICKET_SHOW		0x08000000 /* the file should show the show files context menu */
#define FILE_THICKET_HIDE		0x04000000 /* the file should show the hide files context menu */
#define FOLDER_IS_DOCLIB        0x02000000 /* folder is a doc lib. show special icons */

	// Our Http ItemID struct.  This must look like a SHITEMID!
	// Sadly, this struct (in the v0 format) is duplicated in office 2000, dmufilob.cpp.  Do not change
	// this in an incompatible way - the v0 ILHttpFromPidl is what office 2000 has built-in.
#pragma pack(push,1)	 	
typedef struct tagHTTPITEMID
{
		// Header block
   USHORT            cb;
   USHORT            wId;
	BYTE					bVer;
	BYTE					bType;
   USHORT            wPadId;     // Padding, stick an id in it for now.

		// File info, only meaningful if bType & HTTPID_URL
	FILETIME				ftLastMod;
	unsigned __int64	ul64Size;
	DWORD					dwAttr;

	WCHAR					wtzName[2];
	WCHAR					wtzLocation[2];
	WCHAR					wtzMetaTags[2];

} HTTPITEMID, *LPHTTPITEMID;
#pragma pack(pop)

	// One of our pidls must be at least this size, it will likely be bigger.
#define CBHTTPITEMID_MIN	sizeof(HTTPITEMID)


#endif // MSOPIDL_H
