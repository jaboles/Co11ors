#pragma once

/*************************************************************************
 	msoerr.h

 	Owner: rickp
 	Copyright (c) 1994 Microsoft Corporation

	Office error return values.
*************************************************************************/

#ifndef MSOERR_H
#define MSOERR_H

#if !defined(MSOSTD_H)
	#include <msostd.h>
#endif

/*************************************************************************
	Error return code manipulation macros
*************************************************************************/

/* Severity codes */
#define msoerrSevMask 0xC0000000
#define msoerrSevFail 0xC0000000
#define msoerrSevWarn 0x80000000
#define msoerrSevInfo 0x40000000
#define msoerrSevSuccess 0x00000000

/* The Windows customer error bit */
#define msoerrCustomer 0x20000000

/* Office facility code */
#define msoerrFacMask 0x1FFF0000
#define msoerrFacOffice (FACILITY_ITF << 16)

#define msoerrStatusMask 0x0000FFFF

/* Convenience macros */
#define MsoErrFail(err) (msoerrSevFail|msoerrCustomer|msoerrFacOffice|(err))
#define MsoErrWarn(err) (msoerrSevWarn|msoerrCustomer|msoerrFacOffice|(err))
#define MsoErrInfo(err) (msoerrSevInfo|msoerrCustomer|msoerrFacOffice|(err))
#define MsoErrSuccess(err) (msoerrSevSuccess|msoerrCustomer|msoerrFacOffice|(err))

/* Given a mac error (negative number) fabricate a status code - this only
	takes the low fifteen bits of the error code however the code is a short
	and is always negative, so this macro should also preserve the whole
	code.  As a consequence the upper half of the Office space for mac error
	codes. */
#define MsoErrMac(err)  MsoErrFail(0x8000 | (0x7FFF & (err)))

/* Get the last error and convert to an HRESULT. */
#define MsoHrGetLastError() (ResultFromScode(GetLastError()))


/*************************************************************************
	The actual error values
*************************************************************************/

#define msoerrSuccess 0 // no error occured
#define msoerrNoIntl 1 // International DLL not found
#define msoerrNoMem 14 // out of memory
#define msoerrOleInit 15 // OleInitialize failed
#define msoerrInvalidArg 87 // one or more arguments are invalid


/* Return the following when an Office API fails (e.g. returns NULL) but
	GetLastError() returns a non-FAILED code - this allows recovery when
	implementing an OLE interface which must return an HRESULT, clearly not
	doing SetLastError() is a serious internal Office error in this case. */
#define msoerrNoErrorSet       0x100

/* Set the following for internal errors where recovery is possible (i.e.
	on returning from an "impossible" case after an Assert). */
#define msoerrInternal         0x101

#define msoerrNoShapeSel       0x201
#define msoerrNotTwoShapeSel   0x202
#define msoerrNoGroupSel       0x203
#define msoerrNoObjectSel      0x204
#define msoerrNoPictureSel     0x205
#define msoerrNoPolygonSel     0x206
#define msoerrNoShapeClipboard 0x207
#define msoerrNoPrevUngroup    0x208
#define msoerrNeedExactlyOneShapeSelected 0x209
#define msoerrInvalidSelMode      0x210
#define msoerrInvalidAddPolygonPt 0x211
#define msoerrNestedBeginFreeform 0x212
#define msoerrNotYetImplemented   0x213

#define msoerrLocale		0x214

#define msoerrNoTwoPICanvasSel    0x215

#define msoerrLockAgainstGroup 0x220
#define msoerrLockAgainstUngroup 0x221

#define msoerrPathNotStarted 0x301  // Paths must start with a MoveTo
#define msoerrPathBadBezier  0x302  // Bad point count in a bezier path
#define msoerrPathBadLineTo  0x303  // Bad point count in a line path
#define msoerrPathBad        0x304  // GEL does not handle path element
#define msoerrPathBadType    0x305  // PolyDraw path control byte bad
#define msoerrPathClientEsc  0x306  // Client escape handling failed
#define msoerrPathEscape     0x307  // Escape handling failed
#define msoerrPathBadEscape  0x308  // Unexpected or unhandleable escape
#define msoerrPathBadPath    0x309  // Incorrect point count for an MsoPath
#define msoerrPathExtension  0x30A  // Escape extension not terminated
#define msoerrPathQuadratic  0x30C  // Unterminated quadratic Bezier
#define msoerrPathNoPath     0x30D  // No path is being written
#define msoerrPathQuadFail   0x30E  // Could not handle a quadratic Bezier
#define msoerrPathInternal   0x30F  // Some internal error (bug)

#define msoerrImageAbsent    0x321  // No image property to display
#define msoerrNoBlipLoaded   0x322  // Blip created but not inited
#define msoerrBlipOverwrite  0x323  // Attempt to overwrite a blip
#define msoerrBlipConvert    0x324  // Invalid Blip convertion attempted
#define msoerrBlipUnimpl     0x325  // Unsupported blip operation (internal)
#define msoerrBlipOverrun    0x326  // Overran end of data - blip corrupted?
#define msoerrBlipIOError    0x327  // Rewind/sync error (internal)
#define msoerrBlipNoData     0x328  // No data in blip (internal)
#define msoerrBlipInternal   0x329  // Programming error (bug)
#define msoerrBlipTooBig     0x32A  // File of blip data is too big to load!
#define msoerrBlipTooSmall   0x32B  // File of blip data is too small
#define msoerrBlipData       0x32C  // Error in blip data (file corrupt!)
#define msoerrBlipCallback   0x32D  // Client callback failed
#define msoerrBlipDIBOverflow 0x32E // FDrawToDib cannot fit data into target
#define msoerrBlipHlink      0x32F  // Internal hlink resolution error
#define msoerrBlipFBH        0x330  // File block header invalid

#define msoerrBlipPNGFormat  0x331  // PNG file format error
#define msoerrBlipGIFFormat  0x332  // GIF file format error
#define msoerrBlipBMPFormat  0x333  // BMP file format error
#define msoerrBlipJPEGFormat 0x334  // JPEG file format error
#define msoerrBlipWMFFormat  0x335  // WMF file format error
#define msoerrBlipEMFFormat  0x336  // EMF file format error
#define msoerrBlipPICTFormat 0x337  // PICT file format error
#define msoerrBlipTIFFFormat 0x338  // TIFF file format error
#define msoerrBlipDIBFormat  0x339  // DIB data format error
#define msoerrBlipMFPFormat  0x33A  // METAFILEPICT bad
#define msoerrBlipDIBNosup   0x33B  // DIB data format not supported (JPEG)

#define msoerrFPUnderflow    0x341  // Loss of precision in FP operation
#define msoerrFPOverflow     0x342  // FP overflow detected
#define msoerrFPDivideBy0    0x343  // FP division by zero detected
#define msoerrFPInvalid      0x344  // Invalid FP operation performed
#define msoerrFPUnknown      0x345  // Unknown FP problem - should not occur
#define msoerrIntOverflow    0x346  // Integer overflow

#define msoerrQuickDraw      0x350  // Some problem in QuickDraw (Mac only)

#define msoerrGtrPUnderflow  0x361  // Loss of precision in perspective
#define msoerrGtrPOverflow   0x362  // Overflow or 0 weight in perspective
#define msoerrGtrShadow      0x363  // Bad shadow type
#define msoerrGtrPerspective 0x364  // Bad perspective type

#define msoerrRMFNoError        0x370  // A success code!
#define msoerrRMFError          0x371  // Unspecified error
#define msoerrRMFBadParam       0x372  // Function parameter is out of range
#define msoerrRMFNotImplemented 0x373  // Routine not written
#define msoerrRMFMemoryError    0x374  // Problem allocating memory
#define msoerrRMFOutOfBuffer    0x375  // Supplied buffer is too small
#define msoerrRMFNoSuchData     0x376  // Requested data isn't in file
#define msoerrRMFUnexpectedEnd  0x377  // No EOL at end of buffer
#define msoerrRMFDeltaFound     0x378  // Unsupported RLE8 delta command found

#define msoerrUnknownObject     0x381  // Unexpected GDI object encountered
#define msoerrInvalidObject     0x382  // Inappropriate GDI object encountered
#define msoerrRenderingFailed   0x383  // GEL effect failed to render

#define msoerrDcCacheInvalid 0x391  // Improperly nested MSODC caching

#define msoerrPaletteTooBig   0x394 // Cannot get colors from true color pal
#define msoerrPaletteTooSmall 0x395 // Palette cannot hold system colors

#define msoerrCMSNotFound     0x3A0 // Unrecognized CMS value
#define msoerrCMSIdNotFound   0x3A1 // Unrecognized CMS Id
#define msoerrCMSNameNotFound 0x3A2 // Unrecognized CMS color name
#define msoerrCMSInvalid      0x3A3 // Badly formed CMS in color string
#define msoerrCMSIdInvalid    0x3A4 // Badly formed CMS ID in color string
#define msoerrCMSNameInvalid  0x3A5 // Badly formed CMS color name in string
#define msoerrCMSInternal     0x3A6 // Internal error (unexpected)

#define msoerrFontNotInTable 0x401  // RTF parsing generated bad font index
#define msoerrNoDefaultFont  0x402  // RTF parsing not given default font
#define msoerrNoFonts        0x403  // Could not find ANY fonts
#define msoerrFontLost       0x404  // A valid font family has gone away
#define msoerrNoRTFOrFont    0x405  // Spletter call with no RTF or font
#define msoerrGlyphOverflow  0x406  // Glyph data not representable
#define msoerrGlyphBMErr     0x407  // Internal error building bitmap glyph
#define msoerrGlyphTTBad     0x408  // Invalid TrueType glyph data in font
#define msoerrGlyphPath      0x409  // Badly formed glyph path (internal)
#define msoerrGlyphNotFound  0x40A  // No such glyph in font
#define msoerrGlyphDataBad   0x40B  // Data describing glyph invalid

#define msoerrZlibBase       0x410  // Base of error codes for Zlib
#define msoerrZlibIO         0x411  // IO error detected below Z compression
#define msoerrZlibStream     0x412  // Zlib driver error (internal)
#define msoerrZlibData       0x413  // Zlib compressed data corrupt
#define msoerrZlibMemory     0x414  // Out of memory in Zlib
#define msoerrZlibBuffer     0x415  // Buffer full in Zlib (internal?)
#define msoerrZlibInvalid    0x416  // Invalid error code (internal)
#define msoerrZlibOverflow   0x417  // Memory output buffer overflow
#define msoerrZlibMismatch   0x418  // Compressed data length unexpected

#define msoerrStringTooLong   0x440  // Uncode->ANSI would not fit in buffer
#define msoerrFileNameTooLong 0x441  // File name won't fit in buffer

#define msoerrNoCompMgrReg   0x501  // No IMsoComponentManager registered
#define msoerrTrackCompSet   0x502	// A track comp has already been set
#define msoerrNotTrackComp   0x503  // NonTrack comp tried to end tracking
#define msoerrACompIsXActive 0x504  // Activation attempted while a comp is
                                    // exclusively active

#define msoerrWrongVersion	  0x601  // Data is from a different version
#define msoerrTooFewBytes    0x602	// Premature end of file?
#define msoerrBadPassword    0x603  // Password given is incorrect

#define msoerrNotAutoshape   0x701  // Invalid auto shape number

/* WARNING: The JPEG support code uses all 256 of the following error codes,
	do not define any other error of the form 0x8..! */
#define msoerrIJG            0x800  // First IJG error code.
#define msoerrIJGInternal    0x8FE  // Some internal library problem
#define msoerrIJGNoImage     0x8FF  // No image (just tables) in file

#define msoerrUserAbort      0xFFE  // Set by client on user abort
#define msoerrNYI            0xFFF  // Not Yet Implemented

//mei's stuff
//TODO: I'll set all failure to this value, KC will actually define
//error values
#define msoerrFunctionRetureFailure   0x900
//end mei's stuff
#define msoerrCanNotCreateDialog   0x901  // Error creating Dialog
#define msoerrImportResult         0x903  // Can not get result from import filter
#define msoerrGetMetaFileHandle    0x904  // Can not get valid Metafile handle
#define msoerrGetEnhMetaFileHandle 0x905 	// Can not get valid Enh Metafile handle
#define msoerrPictDlgInitFailure   0x907  // Failure initializing InsertPicture Dlg
#define msoerrPathHasNoExtension   0x908  // The path has no extension
#define msoerrFilterNotFound       0x909  // Graphics filter not found
#define msoerrSettingupFilterInfo  0x910  // Can not read registry to create filter info
#define msoerrWrongFilter          0x911  // Wrong picture filter
#define msoerrImportPict			  0x912  // Error in using filter to import picture.
#define msoerrExportPict			  0x913  // Error in using filter to export drawing.
#define msoerrNoCag					  0x914	// Error launching CAG

// Picture Disassembly errors.
#define msoerrTooManyObj			  0xA00  // Too many objects in a picture.
#define msoerrDisassembly          0xA01  // Failure in disassembling picture.
#define msoerrInitDisState         0xA02  // Fail to init DisassembleState.
#define msoerrSetPenFill           0xA03  // Fail to set line and fill style.

#if 0 //$[VSMSO]
// Web errors
#define msoerrWebNoStartPage       0xB00  // No Web Start Page registered
#define msoerrWebNoSearchPage      0xB01  // No Web Search Page registered
#define msoerrWebNoHistoryFolder   0xB02  // No Web History Folder registered
#define msoerrWebUnexpected        0xB04  // Generic Unexpected Web error
#define msoerrWebAddToFavFail      0xB05  // AddToFavorites generic error
#define msoerrWebSetStartPageFail  0xB06  // SetStartPage generic error
#define msoerrWebSetSearchPageFail 0xB07  // SetSearchPage generic error
#endif //$[VSMSO]

// Beta Errors
#define msoerrBetaExpired 				0xC02 // beta build expired
#define	msoerrOEMExpirationWarning 		0xC03
#define msoerrOEMExpirationWarningOneDay 0xc04
#define	msoerrOEMExpired 				0xc05

// Licensing Errors
#define msoerrAlreadyRegistered			0xc80
#define msoerrQuitRestart				0xc81
#define msoerrAppDisabled				0xc82

#define MSO_E_APP_DISABLED (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrAppDisabled))

// Darwin/Gimme Errors
#define msoerrMsiNoService							0xD00
#define msoerrMsiNoService_ConfirmInstall		0xD01
#define msoerrMsiNoDatabase						0xD02
#define msoerrMsiNoDatabase_ConfirmInstall	0XD03
#define msoerrMsiDisabledFeature				0xD04  // Fail because feature is disabled
#define msoerrMsiConfirmInstall					0xD06
#define msoerrMsiConfirmRepair					0xD08
#define msoerrMsiNoHydraFixes					0xD0F
#define msoerrMsiInstallerBusy					0xD10
#define msoerrMsiInstallerError					0xD11  // get error info from system
#define msoerrMsiUnknownError					0xD12
#define msoerrMsiHydraDisabledFeature			0xD13

// XML Vector Graphic parse errors - 0xE.. is reserved for Vector Graphic
// error codes.
#define msoerrXMLTextUnexpected    0xE01  // Text unexpected in element
#define msoerrXMLUnexpected        0xE02  // Sub-element unexpected
#define msoerrXMLInvalidEnum       0xE03  // Enumeration value unknown
#define msoerrXMLInvalidInteger    0xE04  // Text did not parse as integer
#define msoerrXMLIntegerOverflow   0xE05  // Number too big
#define msoerrXMLInvalidfraction   0xE06  // Text did not parse as fixed
#define msoerrXMLfractionOverflow  0xE07  // Fixed number too big
#define msoerrXMLInvalidlength     0xE08  // Text did not parse as length
#define msoerrXMLlengthOverflow    0xE09  // Length too big (in EMU)
#define msoerrXMLInvalidmeasure    0xE0A  // Text did not parse as measure
#define msoerrXMLmeasureOverflow   0xE0B  // Measure too big (in EMU or frac)
#define msoerrXMLInvalidangle      0xE0C  // Text did not parse as angle
#define msoerrXMLangleOverflow     0xE0D  // Angle too big (in fixed degrees)
#define msoerrXMLInvalidboolean    0xE0E  // VG boolean value invalid
#define msoerrXMLstyle             0xE0F  // Unexpected stuff in CSS style

// More specific cases
#define msoerrXMLNoColor             0xE11  // <C> element with no color
#define msoerrXMLBadColor            0xE12  // Unrecognized color
#define msoerrXMLBadColorParameter   0xE13  // Color op parameter too large
#define msoerrXMLBadColorCombination 0xE14  // Color op/color combination bad
#define msoerrXMLRepeatedColor       0xE15  // Color element repeated
#define msoerrXMLBadColorOrder       0xE16  // Color order in shade wrong
#define msoerrXMLNoDefault           0xE18  // Defaulted value where not allowed
#define msoerrXMLRuleTypeInvalid     0xE19  // Unrecognized rule type
#define msoerrXMLTooFewPoints        0xE21  // Missing points in path (v)
#define msoerrXMLTooManyPoints       0xE22  // Extra points in path (v)
#define msoerrXMLNoOpCode            0xE23  // Points with no opcode (v)
#define msoerrXMLBadOpCode           0xE24  // Bad long opcode (v)
#define msoerrXMLBad2COpCode         0xE25  // Bad two character opcode (v)
#define msoerrXMLBadPathCharacter    0xE26  // Unknown character in path (v)
#define msoerrXMLNoPValue            0xE27  // No value in <p> element
#define msoerrXMLNoFInP              0xE28  // No <f> element in <p> element
#define msoerrXMLFValueTooBig        0xE29  // <f> value too big
#define msoerrXMLFMisplaced          0xE2A  // @ when no number required
#define msoerrXMLIncompleteRect      0xE30  // some but not all of left, top, width, height set
#define msoerrXMLTooManyAdj          0xE31  // more than 8 adj values inside a <var>
#define msoerrXMLTooManyGuide        0xE32  // more than 128 guides inside a <formula>
#define msoerrXMLCantFixupPath       0xE33  // Path fixup for O97 failed
#define msoerrXMLTooManyAdjHandle    0xE34  // more than 4 adjust handles in <v:handles>

#define msoerrXMLGuideConstTooBig    0xE40  // constant in <guide> won't fit in 16 bits
#define msoerrXMLGuideMissingParen   0xE41  // ( or ) not found where expected in <guide>
#define msoerrXMLGuideIndexTooBig    0xE42  // guide(x) is a forward reference (guide or user)
#define msoerrXMLGuideBadChar        0xE43  // Unknown character in formula (guide)
#define msoerrXMLGuideBadFormula     0xE44  // formula in <guide> is unknown
#define msoerrXMLGuideBadArgument    0xE45  // argument in <guide> is unknown
#define msoerrXMLGuideString         0xE46  // unexpected string in <guide>
#define msoerrXMLGuideNumber         0xE47  // unexpected number in <guide>
#define msoerrXMLAdjustIndexTooBig   0xE48  // adjust(x) x not in [0,7] (guide or user)
#define msoerrXMLGuideMissingArg     0xE49  // formula does not have enough args (guide)
#define msoerrXMLGuideBogusParen     0xE4A  // ( or ) missing from <guide>
#define msoerrXMLUserNotAllowed      0xE4B  // adjust(n) etc. in pin or range (user)
#define msoerrXMLUserBadArg          0xE4C  // invalid handle.position   (user)

#define msoerrXMLUnknownOOM          0xE51  // OOM storing unknown XML

#define msoerrXMLfont                0xE61  // Bad XML font attribute
#define msoerrXMLvtext               0xE62  // Bad XML vtext attribute

#define msoerrXMLIncompatiblePattern 0xE70  // Blip tag cannot be set to prevent O97 crash
#define msoerrXMLBadBitmapTag        0xE71  // something inconsistent in <bitmap>


// Data Integration errors
#define msoerrDataDsCantCreateTemp   0xF11 /* Cannot create a temporary file */
#define msoerrDataDsZeroColumns      0xF12 /* Zero columns in the Data Store */
#define msoerrDataDsCantInitDialog   0xF13 /* Cannot initialize the Dialog */
#define msoerrDataDsTooManyFields    0xF14 /* Too many fields in the DS */
#define msoerrDataDsCantReadData     0xF15 /* Can't read data from the DS */
#define msoerrDataDsCantWriteData    0xF16 /* Can't write data into the DS */
#define msoerrDataDsCantInsertRow    0xF17 /* Can't insert row into the DS */
#define msoerrDataDsCantDeleteRow    0xF18 /* Can't delete row from the DS */
#define msoerrDataDsNotValid         0xF19 /* Not a valid Data Store */
#define msoerrDataDsGeneric          0xF1A /* Generic DS error */
#define msoerrDataDsCantSave         0xF1B /* Cannot save the data source file */
#define msoerrDataDsCustomize        0xF1C /* Failed to customize the data source */
#define msoerrDataSecurityFailure    0xF1D /* DB_SEC_E_AUTH_FAILED failure - password failure */
#define msoerrDataCantReadRow        0xF1E /* Cannot read row data - */
#define msoerrDataCantInitWdbImp     0xF1F /* Cannot initialize the wdbimp dll */
#define msoerrDataWorksError         0xF20 /* An error in reading the works file */
#define msoerrDataErrorDisplayed     0xF21 /* An error message has been displayed */
#define msoerrDataMapiInit           0xF22 /* We could not load the mapi dll correctly */
#define msoerrDataFilterSortErr      0xF23 /* We could not apply the desired Filter/Sort.*/
#define msoerrDataDsZeroTables       0xF24 /* Zero tables in the Data Store */
#define msoerrDataNotOleDB           0xF25 /* The user is trying to open a non-OleDB data source */
#define msoerrDataDontOpenOal        0xF26 /* The user is trying to open an OAL that they won't be able to open */
#define msoerrDataContinueSearch     0xF27 /* Continue the Search */
#define msoerrDataOpenSpecial        0xF28 /*  */
#define msoerrDataConnOLDoc          0xF29 /* */
#define msoerrDataOpenUseQuery       0xF2A /* */
#define msoerrDataConnFileDamaged    0xF2B /* */

#define msoerrNotAllInk              0xF2C /* Not all shapes in a selection are ink shapes */

#define S_ERROR_DISPLAYED (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataErrorDisplayed))
#define E_ERROR_DISPLAYED (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataErrorDisplayed))
#define E_DS_ZEROCOLUMNS (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataDsZeroColumns))
#define E_MAPIINIT (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataMapiInit))
#define E_DATAFILTERSORT (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataFilterSortErr))
#define E_DS_ZEROTABLES (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataDsZeroTables))
#define E_NOT_OLEDB     (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataNotOleDB))
#define E_DONT_OPENOAL (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataDontOpenOal))
#define S_DS_CONTINUE_SEARCH (MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, msoerrDataContinueSearch))
#define E_INVALID_DATA HRESULT_FROM_WIN32(ERROR_INVALID_DATA)

#define S_OPEN_SPECIAL MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, msoerrDataOpenSpecial)
#define E_OLDOC MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataConnOLDoc)
#define S_USE_QUERY     MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, msoerrDataOpenUseQuery)
#define E_QUERY_FILE_DAMAGED MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDataConnFileDamaged)



// Crypt Features (digital signatures and strong encryptions)
#define msoerrCspNotFound            0x1001
#define msoerrTimeStampServerFailed	 0x1002

// DRM errors
#define msoerrDrmFeatureUnknown         0x1010
#define msoerrDrmUserAbort              0x1011
#define msoerrDrmDisabled               0x1012
#define msoerrDrmNoPassportEveryone     0x1013
#define msoerrDrmNoPassportTemplate     0x1014
#define msoerrDrmRestrictedSite         0x1015
#define msoerrDrmOfficeServiceRetires    0x1016
#define msoerrDrmEnterpriseServiceRetires 0x1017
#define msoerrDrmSaveRightExpired       0x1018
#define msoerrDrmGICExpired             0x1019
#define msoerrDrmOfficeEnrollmentServiceRetires    0x101a
#define msoerrDrmEnterpriseEnrollmentServiceRetires 0x101b
#define msoerrDrmNeedsCredential        0x101c
#define msoerrDrmRemovePermissionFailed 0x101d
#define E_MSODRM_FEATUREUNKNOWN (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmFeatureUnknown))
#define E_MSODRM_USERABORT		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmUserAbort))
#define E_MSODRM_DISABLED		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmDisabled))
#define E_MSODRM_NOPASSPORTEVERYONE		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmNoPassportEveryone))
#define E_MSODRM_NOPASSPORTTEMPLATE		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmNoPassportTemplate))
#define E_MSODRM_RESTRICTEDSITE		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmRestrictedSite))
#define E_MSODRM_OFFICE_SERVICE_RETIRES	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmOfficeServiceRetires))
#define E_MSODRM_ENTERPRISE_SERVICE_RETIRES	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmEnterpriseServiceRetires))
#define E_MSODRM_SAVE_RIGHT_EXPIRED	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmSaveRightExpired))
#define E_MSODRM_GIC_EXPIRED	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmGICExpired))
#define E_MSODRM_OFFICE_ENROLLMENT_SERVICE_RETIRES	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmOfficeEnrollmentServiceRetires))
#define E_MSODRM_ENTERPRISE_ENROLLMENT_SERVICE_RETIRES	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmEnterpriseEnrollmentServiceRetires))
#define E_MSODRM_NEEDS_CREDENTIAL	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmNeedsCredential))
#define E_MSODRM_REMOVE_PERMISSION_FAILED	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDrmRemovePermissionFailed))

// XML parsing errors
#define msoerrXmlNotWellFormed			0x1020
#define msoerrXmlInvalid				0x1021
#define msoerrXmlInvalidSchema			0x1022
#define msoerrXmlInstanceStructureNotDefined	0x1023
#define msoerrXmlInvalidSelectionNamespace	0x1024
#define E_MSOXML_NOTWELLFORMED	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrXmlNotWellFormed))
#define E_MSOXML_INVALID		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrXmlInvalid))
#define E_MSOXML_INVALIDSCHEMA	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrXmlInvalidSchema))
#define E_MSOXML_INSTANCESTRUCTURENOTDEFINED	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrXmlInstanceStructureNotDefined))
#define E_MSOXML_INVALIDSELECTIONNAMESPACE	(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrXmlInvalidSelectionNamespace))

// SyncMan errors
#define msoerrSmNotRoaming           0x1030
#define msoerrSmNoPsm                0x1031
#define msoerrSmNoVersion            0x1032
#define msoerrSmNoConflict           0x1033
#define msoerrSmSuspended            0x1034
#define msoerrSmInvalidCompare       0x1035
#define msoerrSmNoDwsConnection    0x1036
#define E_MSOSM_NOTROAMING           (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmNotRoaming))
#define E_MSOSM_NOPSM                (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmNoPsm))
#define E_MSOSM_NOVERSION            (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmNoVersion))
#define E_MSOSM_NOCONFLICT           (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmNoConflict))
#define E_MSOSM_SUSPENDED            (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmSuspended))
#define E_MSOSM_INVALIDCOMPARE       (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmInvalidCompare))
#define E_MSOSM_NODWSCONNECTION   (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrSmNoDwsConnection))

// Versions errors
#define msoerrVNoVersions           			0x1050
#define msoerrVServerError                		0x1051
#define msoerrVInvalidVersion                    0x1052
#define msoerrVInvalidVersionIndex           0x1053
#define msoerrVServerTimeout                		0x1054
#define msoerrVDirtyDoc                              0x1055
#define msoerrVOffline                                 0x1056
#define msoerrVcvRestore                           0x1057
#define msoerrVcvDelete                             0x1058
#define E_MSOV_NOVERSIONS           (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVNoVersions ))
#define E_MSOV_SERVERERROR                (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVServerError))
#define E_MSOV_SERVERTIMEOUT                (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVServerTimeout))
#define E_MSOV_INVALIDVERSION            (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVInvalidVersion))
#define E_MSOV_INVALIDVERSIONINDEX           (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVInvalidVersionIndex))
#define E_MSOV_DIRTYDOC                        (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVDirtyDoc))
#define E_MSOV_OFFLINE                          (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVOffline))
#define E_MSOV_CVRESTORE                     (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVcvRestore))
#define E_MSOV_CVDELETE                       (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrVcvDelete))

// MSO Network Status errors
#define msoerrNetStatusNetworkUnavailable  0x1100  // No network is available.
#define msoerrNetStatusUserOffline         0x1101  // The user is currently offline.
#define msoerrNetStatusUserOfflineAlerted  0x1102  // The user is currently offline, and has been alerted of this.
#define E_NETSTATUS_NETWORKUNAVAILABLE    (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrNetStatusNetworkUnavailable))
#define E_NETSTATUS_USEROFFLINE           (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrNetStatusUserOffline))
#define E_NETSTATUS_USEROFFLINEALERTED    (MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrNetStatusUserOfflineAlerted))

// DWS Errors
#define msoerrDWSWorkspaceNotExist		0x1200
#define msoerrDWSWorkspaceExist		0x1201
#define msoerrDWSFileUnsaved			0x1202
#define msoerrDWSLongFileName			0x1203
#define msoerrDWSServerProblem			0x1204
#define msoerrDWSNoData				0x1205
#define msoerrDWSFileNotExist			0x1206
#define msoerrDWSFileExist				0x1207
#define msoerrDWSDeleteCurDoc			0x1208
#define msoerrDWSStampFailed			0x1209
#define msoerrDWSBadChar				0x1210
#define E_MSODWS_WORKSPACENOTEXIST		(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSWorkspaceNotExist))
#define E_MSODWS_WORKSPACEEXIST			(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSWorkspaceExist))
#define E_MSODWS_FILEUNSAVED				(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSFileUnsaved))
#define E_MSODWS_LONGFILENAME			(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSLongFileName))
#define E_MSODWS_SERVERPROBLEM			(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSServerProblem))
#define E_MSODWS_NODATA					(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSNoData))
#define E_MSODWS_FILENOTEXIST				(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSFileNotExist))
#define E_MSODWS_FILEEXIST					(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSFileExist))
#define E_MSODWS_DELETECURDOC			(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSDeleteCurDoc))
#define E_MSODWS_STAMPFAILED				(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSStampFailed))
#define E_MSODWS_BADCHAR					(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, msoerrDWSBadChar))


/*************************************************************************
	Error reporting routines
*************************************************************************/

/* *************** MSO callers ONLY ***************** */

#if 0 //$[VSMSO]
/* Displays an alert string szMsg to the user, using the icons and buttons
	specified by mb and "extending" the alert if iAlertID is registered.
	This version is for MSO callers since it will automatically look up the
	Watsonized Alert ID in the MSO table. */
int AlertSzIdMso(const CHAR* szMsg, int mb, int iAlertID);

/* Displays an alert string wtzMsg to the user, using the icons and buttons
	specified by mb.  This version is for MSO callers since it will automatically
	look up the Watsonized Alert ID in the MSO table.*/
int AlertWtzMso(const WCHAR* wtzMsg, int mb, int iAlertID);
#endif //$[VSMSO]


/* *************** end MSO ONLY     ***************** */


/*	Report an out of memory error to the user */
MSOAPIX_(void) MsoOutOfMemory(void);

#if 0 //$[VSMSO]
/*	Displays a standard Office error alert based on the current error
	value stored in GetLastError.  Returns 0 if we couldn't figure out
	what to do with the error. */
MSOAPI_(BOOL) MsoFStdError(void);
#endif //$[VSMSO]

/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	The string ids is assumed to have an insertion mark in it, which will
	be replaced by the sz0 string. */
MSOAPIX_(int) MsoAlertIdsSz1(HINSTANCE hinst, int ids, const CHAR* sz0, int mb);

/* Displays an alert string wtzMsg to the user, using the icons and buttons
	specified by mb.  See MsoAlertWtzWA for a preferred version including the Watsonized
	Alert ID. */
MSOAPI_(int) MsoAlertWtz(const WCHAR* wtzMsg, int mb, int iAlertID);

/* Displays an alert string wtzMsg to the user, using the icons and buttons
	specified by mb.  Also passes in the Watsonized Alert ID.
	This function is preferred to MsoAlertWtz.
	*/
MSOAPI_(int) MsoAlertWtzWA(const WCHAR* wtzMsg, int mb, int iAlertID, int iWAlertID);

#if 0 //$[VSMSO]
/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst (And expands env variables).
	Alert icons and buttons are specified by mb. */
MSOAPIX_(int) MsoAlertExpandIds(HINSTANCE hinst, int ids, int mb);
#endif //$[VSMSO]

/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb. */
MSOAPI_(int) MsoAlertIds(HINSTANCE hinst, int ids, int mb);

/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	Takes the Watsonized Alert ID as a parameter as well. */
MSOAPI_(int) MsoAlertIdsWA(HINSTANCE hinst, int ids, int mb, int iWAlertID);

#if 0 //$[VSMSO]
/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	The string ids is assumed to have an insertion mark in it, which will
	be replaced by the sz0 string. */
MSOAPIX_(int) MsoAlertIdsSz1(HINSTANCE hinst, int ids, const CHAR* sz0, int mb);
#endif //$[VSMSO]

/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	The string ids is assumed to have an insertion mark in it, which will
	be replaced by the wtz0 string. */
MSOAPIX_(int) MsoAlertIdsWtz1(HINSTANCE hinst, int ids, const WCHAR* wtz0, int mb);
MSOAPIX_(int) MsoAlertIdsWz1(HINSTANCE hinst, int ids, const WCHAR* wz0, int mb);

#if 0 //$[VSMSO]
/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	The string ids is assumed to have an two insertion marks in it, which
	will be replaced by the sz0 and sz1 strings. */
MSOAPIX_(int) MsoAlertIdsSz2(HINSTANCE hinst, int ids, const CHAR* sz0, const CHAR* sz1, int mb);
#endif //$[VSMSO]

/* Displays an alert string szMsg to the user, using the icons and buttons
	specified by mb */
MSOAPI_(int) MsoAlertSz(const CHAR* szMsg, int mb);

/* Displays an alert string szMsg to the user, using the icons and buttons
	specified by mb and "extending" the alert if iAlertID is registered*/
MSOAPIX_(int) MsoAlertSzId(const CHAR* szMsg, int mb, int iAlertID);

#if 0 //$[VSMSO]
/* Displays an alert string szMsg to the user, using the icons and buttons
	specified by mb and "extending" the alert if iAlertID is registered.
	Also passes in the Watsonized Alert ID. This function is preferred to MsoAlertWtz.
	*/
MSOAPIX_(int) MsoAlertSzIdWA(const CHAR* szMsg, int mb, int iAlertID, int iWAlertID);
#endif //$[VSMSO]

/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	The string ids is assumed to have an two insertion marks in it, which
	will be replaced by the wtz0 and wtz1 strings. */
MSOAPIX_(int) MsoAlertIdsWtz2(HINSTANCE hinst, int ids, const WCHAR* wtz0, const WCHAR* wtz1, int mb);
MSOAPIX_(int) MsoAlertIdsWz2(HINSTANCE hinst, int ids, const WCHAR* wz0, const WCHAR* wz1, int mb);

/*	Displays an alert string to the user, pulling a string ids from the
	resource file hinst.  Alert icons and buttons are specified by mb.
	The string ids is assumed to have an three insertion marks in it, which
	will be replaced by the wz0, wz1, and wz2 strings. */
MSOAPIX_(int) MsoAlertIdsWz3(HINSTANCE hinst, int ids, const WCHAR* wz0, const WCHAR* wz1, const WCHAR* wz2, int mb);

#if 0 //$[VSMSO]
/*	Displays a standard alert string to the user, using an error string
	corresponding to err in a message box with options mb.  Returns the
	MessageBox return value */
MSOAPIX_(int) MsoAlertErr(int err, int mb);

/* Returns the error string corresponding to err into the string buffer
	rgch, which has size cch.  Returns the actual length of the string */
MSOAPI_(int) MsoLoadSzError(int err, CHAR* rgch, int cch);

/* Returns the error string corresponding to err into the string buffer
	wtz, which has room for cch wide characters.  Returns the actual length
	of the string */
MSOAPI_(int) MsoLoadWtzError(int err, WCHAR* wtz, int cch);
#endif //$[VSMSO]

#endif /* MSOERR_H */
