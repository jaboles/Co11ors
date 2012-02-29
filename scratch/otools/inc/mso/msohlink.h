/****************************************************************************
	msohlink.h

	Owner: AndrewQ
	Copyright (c) 1999 Microsoft Corporation

	MSO's hyperlink object.
****************************************************************************/

#pragma once

#if !defined(MSOHLINK_H)
#define MSOHLINK_H

#if !defined(MSOURL_H)
#include <msourl.h>
#endif

#if 0 //$[VSMSO]
#ifndef CONST_METHOD
// to declare const methods in C++ */
#if defined(__cplusplus) && !defined(CINTERFACE)
	#define CONST_METHOD const
#else
	#define CONST_METHOD
#endif
#endif

/****************************************************************************
	IMsoHyperlink

	Direct inheritance from IMsoUrl, rather than requiring a QI(), to
	strengthen up the conceptual relationship.
*****************************************************************************/

#undef  INTERFACE
#define INTERFACE IMsoHyperlink
DECLARE_INTERFACE_(IMsoHyperlink, IMsoUrl)
{
	// ----- IUnknown methods

	MSOMETHOD (QueryInterface) (THIS_ REFIID riid, VOID **ppvObj) PURE;
	MSOMETHOD_(ULONG, AddRef) (THIS) PURE;
	MSOMETHOD_(ULONG, Release) (THIS) PURE;

	// ----- IMsoUrl methods - see <msourl.h> for details

	// Is the URL valid?
	MSOMETHOD_(BOOL, FValid) (THIS) CONST_METHOD PURE;

	// URL set methods.
	MSOMETHOD (HrSetFromUser) (THIS_ LPCWSTR wzUrl, DWORD cp, const IMsoUrl *piurlBase, DWORD grfurl) PURE;
	MSOMETHOD (HrSetFromUserRgwch) (THIS_ LPCWSTR rgwchUrl, int cchUrl, DWORD cp, const IMsoUrl *piurlBase, DWORD grfurl) PURE;
	MSOMETHOD (HrSetFromCanonicalUrl) (THIS_ LPCWSTR wzUrl, DWORD cp, const IMsoUrl *piurlBase) PURE;

	// Lock/unlock, use when using methods returning LPCWSTR's.
	MSOMETHOD_(VOID, Lock) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(VOID, Unlock) (THIS) CONST_METHOD PURE;

	// URL get methods.
	MSOMETHOD (HrGetDisplayForm) (THIS_ LPWSTR wzBuf, int *pcchBuf, DWORD grfurldf) CONST_METHOD PURE;
	MSOMETHOD (HrGetCanonicalForm) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetCustomForm) (THIS_ DWORD grfurlc, LPWSTR wzBuf, int *pcchBuf, DWORD grfurlcf) CONST_METHOD PURE;
	MSOMETHOD_(DWORD, CpGetCodePage) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, WzCanonicalForm) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(int, CchCanonicalForm) (THIS) CONST_METHOD PURE;

	// URL component accessor methods.
	MSOMETHOD_(MSOURLSCHEME, UrlsGetScheme) (THIS) CONST_METHOD PURE;
	MSOMETHOD (HrGetScheme) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetAuthority) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetUserName) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetPassword) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetServer) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetPort) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetPath) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetDir) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetFileLeaf) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetFileName) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetFileExt) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetQuery) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD (HrGetFragment) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchScheme) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchAuthority) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchUserName) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchPassword) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchServer) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchPort) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchPath) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchDir) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchFileLeaf) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchFileName) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchFileExt) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchQuery) (THIS_ int *pcch) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, RgwchFragment) (THIS_ int *pcch) CONST_METHOD PURE;

	// Helper functions for file management.
	MSOMETHOD_(BOOL, FIsHttp) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FIsFtp) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FIsLocal) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FIsUNC) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(BOOL, FIsODMA) (THIS) CONST_METHOD PURE;

	// Local path accessor methods (only work for file: URLs).
	MSOMETHOD (HrGetLocalPath) (THIS_ LPWSTR wzBuf, int *pcchBuf) CONST_METHOD PURE;
	MSOMETHOD_(LPCWSTR, WzLocalPath) (THIS) CONST_METHOD PURE;
	MSOMETHOD_(int, CchLocalPath) (THIS) CONST_METHOD PURE;

	// Equality check.
	MSOMETHOD_(BOOL, FIsEqual) (THIS_ const IMsoUrl *piurl) CONST_METHOD PURE;

	// Component equality method.
	MSOMETHOD_(BOOL, FComponentsAreEqual) (THIS_ DWORD grfurlc, const IMsoUrl *piurl) CONST_METHOD PURE;

	// Does this url subsume piurl (is it's folder location a direct hierarchical ancestor of piurl)?
	MSOMETHOD_(BOOL, FSubsumes) (THIS_ const IMsoUrl *piurl) CONST_METHOD PURE;

	// Get the URL relativity state.
	MSOMETHOD_(MSOURLRELATIVITY, UrlrGetRelativity) (THIS) CONST_METHOD PURE;

	// Update the URL relativity state.
	MSOMETHOD (HrSetRelativity) (THIS_ MSOURLRELATIVITY urlr) PURE;

	// Resolve this URL to an absolute URL.
	MSOMETHOD (HrResolve) (THIS_ IMsoUrl **ppiurlAbsolute) CONST_METHOD PURE;

	// Rebase this URL with another (link fixup).
	MSOMETHOD (HrRebase) (THIS_ const IMsoUrl *piurlBase) PURE;

	// Download this URL for local data access.
	MSOMETHOD (HrDownload) (THIS_ LPWSTR wzBuf, int *pcchBuf, DWORD grfurld) CONST_METHOD PURE;

	// Upload the file specified in wzFile to the location specified by this url.
	MSOMETHOD (HrUpload) (THIS_ LPCWSTR wzFile, DWORD grfurlu) CONST_METHOD PURE;

	// Delete the file specified by this url.
	MSOMETHOD (HrDelete) (THIS) CONST_METHOD PURE;

	// Check if this url exists.
	MSOMETHOD_(BOOL, FExists) (THIS_ DWORD grfurle) CONST_METHOD PURE;

	// Return the base url of this url.
	MSOMETHOD (HrGetBase) (THIS_ const IMsoUrl **ppiurlBase) CONST_METHOD PURE;

 	// FIsMhtml - does the url point to an mhtml file, and in the form "mhtml:X[!Y]" ?
	MSOMETHOD_(BOOL, FIsMhtml) (THIS) CONST_METHOD PURE;

	// Returns the reference to the bodypart this mhtml url is referring to
	MSOMETHOD_(LPCWSTR, WzMhtmlBodypart) (THIS) CONST_METHOD PURE;

	// Returns the cch of the reference to the bodypart this mhtml url is referring to
	MSOMETHOD_(int, CchMhtmlBodypart) (THIS) CONST_METHOD PURE;

	// ----- IMsoHyperlink methods

	// FDebugMessage method
	MSODEBUGMETHOD

	// Direct access to internal IHlink object. Only use if absolutely needed!
	// (Our eventual goal is for no MSO client to know about IHlink.)
	MSOMETHOD(HrSetHlink) (THIS_ IHlink *pihl) PURE;
	MSOMETHOD(HrGetHlink) (THIS_ IHlink **ppihl) CONST_METHOD PURE;

	// Additional component get methods.
	MSOMETHOD(HrGetFriendlyText) (THIS_ LPWSTR wzText, int *pcchText) CONST_METHOD PURE;
	MSOMETHOD(HrGetTargetFrame) (THIS_ LPWSTR wzFrame, int *pcchFrame) CONST_METHOD PURE;
	MSOMETHOD(HrGetScreenTip) (THIS_ LPWSTR wzTip, int *pcchTip) CONST_METHOD PURE;
	MSOMETHOD(HrGetHlinkSite) (THIS_ IHlinkSite **ppihlsite, DWORD *pdwSiteData) CONST_METHOD PURE;

	// Additional component set methods.
	MSOMETHOD(HrSetFriendlyText) (THIS_ LPCWSTR wzText) PURE;
	MSOMETHOD(HrSetTargetFrame) (THIS_ LPCWSTR wzFrame) PURE;
	MSOMETHOD(HrSetScreenTip) (THIS_ LPCWSTR wzTip) PURE;
	MSOMETHOD(HrSetHlinkSite) (THIS_ IHlinkSite *pihlsite, DWORD dwSiteData) PURE;

	// Write a hyperlink to an IStream.
	MSOMETHOD(HrSaveToStream) (THIS_ IStream *pistm, BOOL fClearDirty) CONST_METHOD PURE;

	// Another IHlink passthrough.
	MSOMETHOD(HrNavigate) (THIS_ DWORD grfHLNF, LPBC pibc, IBindStatusCallback *pibsc, IHlinkBrowseContext *pihlbc) PURE;

	MSOMETHOD_(BPSC, BpscBulletProof) (THIS_ MSOBPCB *pmsobpcb) PURE;

	MSOMETHOD_(ULONG, Free) (THIS) PURE;
};

/****************************************************************************
	MsoHrCreateHyperlink, etc.

	Various ways to create an IMsoHyperlink.

	Note that MsoHrCreateHyperlinkFromData will also read in the screentip
	following the hyperlink in the stream.
	
	This representation is consistent with Office 2000's representation
	of a hyperlink (IHlink, then screen tip).
*****************************************************************************/

MSOAPIX_(HRESULT) MsoHrCreateHyperlink(
	IMsoHyperlink **ppimhl);

MSOAPIX_(HRESULT) MsoHrCreateHyperlinkFromUser(
	IMsoHyperlink **ppimhl,
	LPCWSTR wzUrl,
	DWORD cp,
	IMsoUrl *piurlBase,
	DWORD grfurl,
	IHlinkSite *pihlsite,
	DWORD dwSiteData);

MSOAPIX_(HRESULT) MsoHrCreateHyperlinkFromCanonicalUrl(
	IMsoHyperlink **ppimhl,
	LPCWSTR wzUrl,
	DWORD cp,
	IMsoUrl *piurlBase,
	IHlinkSite *pihlsite,
	DWORD dwSiteData);

MSOAPIX_(HRESULT) MsoHrCreateHyperlinkFromData(
	IMsoHyperlink **ppimhl,
	IDataObject *piDataObj,
	IHlinkSite *pihlsite,
	DWORD dwSiteData);

MSOAPIX_(HRESULT) MsoHrCreateHyperlinkFromHlink(
	IMsoHyperlink **ppimhl,
	IHlink *pihl,
	IHlinkSite *pihlsite,
	DWORD dwSiteData);

MSOAPIX_(HRESULT) MsoHrCreateHyperlinkFromMoniker(
	IMsoHyperlink **ppimhl,
	IMoniker *pimkTrgt,
	LPCWSTR pwzLocation,
	IHlinkSite *pihlsite,
	DWORD dwSiteData);

/****************************************************************************
	MsoHrHyperlinkClone

	Copy an IMsoHyperlink.
*****************************************************************************/

MSOAPIX_(HRESULT) MsoHrHyperlinkClone(
	IMsoHyperlink *pimhl,
	IHlinkSite *pihlsite, 
	DWORD dwSiteData, 
	IMsoHyperlink **ppimhlClone);

/****************************************************************************
	IMsoHyperlink C interface
*****************************************************************************/

#ifndef __cplusplus

// ----- IUnknown methods
#define IMsoHyperlink_QueryInterface(This, riid, ppvObj) \
	(This)->lpVtbl->QueryInterface(This, riid, ppvObj)

#define IMsoHyperlink_AddRef(This) \
	(This)->lpVtbl->AddRef(This)

#define IMsoHyperlink_Release(This) \
	(This)->lpVtbl->Release(This)

// ----- IMsoUrl methods
#define IMsoHyperlink_FValid(This) \
	(This)->lpVtbl->FValid(This)

#define IMsoHyperlink_HrSetFromUser(This, wzUrl, cp, piurlBase, grfurl) \
	(This)->lpVtbl->HrSetFromUser(This, wzUrl, cp, piurlBase, grfurl)

#define IMsoHyperlink_HrSetFromUserRgwch(This, rgwchUrl, cchUrl, cp, piurlBase, grfurl) \
	(This)->lpVtbl->HrSetFromUserRgwch(This, rgwchUrl, cchUrl, cp, piurlBase, grfurl)

#define IMsoHyperlink_HrSetFromCanonicalUrl(This, wzUrl, cp, piurlBase) \
	(This)->lpVtbl->HrSetFromCanonicalUrl(This, wzUrl, cp, piurlBase)

#define IMsoHyperlink_Lock(This) \
	(This)->lpVtbl->Lock(This)

#define IMsoHyperlink_Unlock(This) \
	(This)->lpVtbl->Unlock(This)

#define IMsoHyperlink_HrGetDisplayForm(This, wzBuf, pcchBuf, grfurldf) \
	(This)->lpVtbl->HrGetDisplayForm(This, wzBuf, pcchBuf, grfurldf)

#define IMsoHyperlink_HrGetCanonicalForm(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetCanonicalForm(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetCustomForm(This, grfurlc, wzBuf, pcchBuf, grfurlcf) \
	(This)->lpVtbl->HrGetCustomForm(This, grfurlc, wzBuf, pcchBuf, grfurlcf)

#define IMsoHyperlink_CpGetCodePage(This) \
	(This)->lpVtbl->CpGetCodePage(This)

#define IMsoHyperlink_WzCanonicalForm(This) \
	(This)->lpVtbl->WzCanonicalForm(This)

#define IMsoHyperlink_CchCanonicalForm(This) \
	(This)->lpVtbl->CchCanonicalForm(This)

#define IMsoHyperlink_UrlsGetScheme(This) \
	(This)->lpVtbl->UrlsGetScheme(This)

#define IMsoHyperlink_HrGetScheme(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetScheme(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetAuthority(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetAuthority(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetUserName(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetUserName(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetPassword(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetPassword(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetServer(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetServer(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetPort(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetPort(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetPath(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetPath(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetDir(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetDir(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetFileLeaf(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetFileLeaf(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetFileName(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetFileName(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetFileExt(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetFileExt(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetQuery(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetQuery(This, wzBuf, pcchBuf)

#define IMsoHyperlink_HrGetFragment(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetFragment(This, wzBuf, pcchBuf)

#define IMsoHyperlink_RgwchScheme(This, pcch) \
	(This)->lpVtbl->RgwchScheme(This, pcch)

#define IMsoHyperlink_RgwchAuthority(This, pcch) \
	(This)->lpVtbl->RgwchAuthority(This, pcch)

#define IMsoHyperlink_RgwchUserName(This, pcch) \
	(This)->lpVtbl->RgwchUserName(This, pcch)

#define IMsoHyperlink_RgwchPassword(This, pcch) \
	(This)->lpVtbl->RgwchPassword(This, pcch)

#define IMsoHyperlink_RgwchServer(This, pcch) \
	(This)->lpVtbl->RgwchServer(This, pcch)

#define IMsoHyperlink_RgwchPort(This, pcch) \
	(This)->lpVtbl->RgwchPort(This, pcch)

#define IMsoHyperlink_RgwchPath(This, pcch) \
	(This)->lpVtbl->RgwchPath(This, pcch)

#define IMsoHyperlink_RgwchDir(This, pcch) \
	(This)->lpVtbl->RgwchDir(This, pcch)

#define IMsoHyperlink_RgwchFileLeaf(This, pcch) \
	(This)->lpVtbl->RgwchFileLeaf(This, pcch)

#define IMsoHyperlink_RgwchFileName(This, pcch) \
	(This)->lpVtbl->RgwchFileName(This, pcch)

#define IMsoHyperlink_RgwchFileExt(This, pcch) \
	(This)->lpVtbl->RgwchFileExt(This, pcch)

#define IMsoHyperlink_RgwchQuery(This, pcch) \
	(This)->lpVtbl->RgwchQuery(This, pcch)

#define IMsoHyperlink_RgwchFragment(This, pcch) \
	(This)->lpVtbl->RgwchFragment(This, pcch)

#define IMsoHyperlink_FIsHttp(This) \
	(This)->lpVtbl->FIsHttp(This)

#define IMsoHyperlink_FIsFtp(This) \
	(This)->lpVtbl->FIsFtp(This)

#define IMsoHyperlink_FIsLocal(This) \
	(This)->lpVtbl->FIsLocal(This)

#define IMsoHyperlink_FIsUNC(This) \
	(This)->lpVtbl->FIsUNC(This)

#define IMsoHyperlink_FIsODMA(This) \
	(This)->lpVtbl->FIsODMA(This)

#define IMsoHyperlink_HrGetLocalPath(This, wzBuf, pcchBuf) \
	(This)->lpVtbl->HrGetLocalPath(This, wzBuf, pcchBuf)

#define IMsoHyperlink_WzLocalPath(This) \
	(This)->lpVtbl->WzLocalPath(This)

#define IMsoHyperlink_CchLocalPath(This) \
	(This)->lpVtbl->CchLocalPath(This)

#define IMsoHyperlink_FIsEqual(This, piurl) \
	(This)->lpVtbl->FIsEqual(This, piurl)

#define IMsoHyperlink_FComponentsAreEqual(This, grfurlc, piurl) \
	(This)->lpVtbl->FComponentsAreEqual(This, grfurlc, piurl)

#define IMsoHyperlink_FSubsumes(This, piurl) \
	(This)->lpVtbl->FSubsumes(This, piurl)

#define IMsoHyperlink_UrlrGetRelativity(This) \
	(This)->lpVtbl->UrlrGetRelativity(This)

#define IMsoHyperlink_HrResolve(This, ppiurlAbsolute) \
	(This)->lpVtbl->HrResolve(This, ppiurlAbsolute)

#define IMsoHyperlink_HrRebase(This, piurlBase) \
	(This)->lpVtbl->HrRebase(This, piurlBase)

#define IMsoHyperlink_HrDownload(This, wzBuf, pcchBuf, grfurld) \
	(This)->lpVtbl->HrDownload(This, wzBuf, pcchBuf, grfurld)

#define IMsoHyperlink_HrUpload(This, wzFile, grfurlu) \
	(This)->lpVtbl->HrUpload(This, wzFile, grfurlu)

#define IMsoHyperlink_HrDelete(This) \
	(This)->lpVtbl->HrDelete(This)

#define IMsoHyperlink_FFileExists(This, grfurle) \
	(This)->lpVtbl->FFileExists(This, grfurle)

#define IMsoHyperlink_HrGetBase(This, ppiurlBase) \
	(This)->lpVtbl->HrGetBase(This, ppiurlBase)

#define IMsoHyperlink_FIsMhtml(This) \
	(This)->lpVtbl->FIsMhtml(This, ppiurlBase)

#define IMsoHyperlink_WzMhtmlBodypart(This) \
	(This)->lpVtbl->WzMhtmlBodypart(This)

#define IMsoHyperlink_CchMhtmlBodypart(This) \
	(This)->lpVtbl->CchMhtmlBodypart(This)

// ----- IMsoHyperlink methods
#define IMsoHyperlink_HrSetHlink(This, pihl) \
	(This)->lpVtbl->HrSetHlink(This, pihl)

#define IMsoHyperlink_HrGetHlink(This, ppihl) \
	(This)->lpVtbl->HrGetHlink(This, ppihl)

#define IMsoHyperlink_HrGetFriendlyText(This, wzText, pcchText) \
	(This)->lpVtbl->HrGetFriendlyText(This, wzText, pcchText)

#define IMsoHyperlink_HrGetTargetFrame(This, wzFrame, pcchFrame) \
	(This)->lpVtbl->HrGetTargetFrame(This, wzFrame, pcchFrame)

#define IMsoHyperlink_HrGetScreenTip(This, wzTip, pcchTip) \
	(This)->lpVtbl->HrGetScreenTip(This, wzTip, pcchTip)

#define IMsoHyperlink_HrGetHlinkSite(This, ppihlsite, pdwSiteData) \
	(This)->lpVtbl->HrGetHlinkSite(This, ppihlsite, pdwSiteData)

#define IMsoHyperlink_HrSetFriendlyText(This, wzText) \
	(This)->lpVtbl->HrSetFriendlyText(This, wzText)

#define IMsoHyperlink_HrSetTargetFrame(This, wzFrame) \
	(This)->lpVtbl->HrSetTargetFrame(This, wzFrame)

#define IMsoHyperlink_HrSetScreenTip(This, wzTip) \
	(This)->lpVtbl->HrSetScreenTip(This, wzTip)

#define IMsoHyperlink_HrSetHlinkSite(This, pihlsite, dwSiteData) \
	(This)->lpVtbl->HrSetHlinkSite(This, pihlsite, dwSiteData)

#define IMsoHyperlink_HrSaveToStream(This, pistm, fClearDirty) \
	(This)->lpVtbl->HrSaveToStream(This, pistm, fClearDirty)

#define IMsoHyperlink_HrNavigate(This, grfHLNF, pibc, pibsc, pihlbc) \
	(This)->lpVtbl->HrNavigate(This, grfHLNF, pibc, pibsc, pihlbc)

#endif // __cplusplus

#endif //$[VSMSO]
#endif // MSOHLINK_H
