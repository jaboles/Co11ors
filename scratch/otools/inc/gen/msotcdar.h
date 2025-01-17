/******************************************************************************
 * Gimme Header File
 * Automatically generated by %OTOOLS%\bin\GimmeBuild.bat
 *
 * Owner: JJames and Office TCO
 * (C) Copyright 1997 - 1999, Microsoft Corporation
 ******************************************************************************/

#ifndef TCDARO_H
#define TCDARO_H

#define msodwGimmeTableVersion 2

enum {
	msofidNil = -1,
	msoqfidNil = -1,
	msofidMsoDll           = 0x00000,	// (0) mso$d.dll
	msofidRiched20Dll      = 0x00001,	// (1) riched20.dll
	msofidUser32Dll        = 0x00002,	// (2) USER32.dll
	msofidGdi32Dll         = 0x00003,	// (3) GDI32.dll
	msofidWinnlsDll        = 0x00004,	// (4) WINNLS.dll
	msofidShell32Dll       = 0x00005,	// (5) SHELL32.dll
	msofidComctl32Dll      = 0x00006,	// (6) COMCTL32.dll
	msofidOleaut32Dll      = 0x00007,	// (7) OLEAUT32.dll
	msofidComDlg32Dll      = 0x00008,	// (8) COMDLG32.dll
	msofidKernel32Dll      = 0x00009,	// (9) KERNEL32.dll
	msofidVersionDll       = 0x0000a,	// (10) VERSION.dll
	msofidWinmmDll         = 0x0000b,	// (11) WINMM.dll
	msofidMapi32Dll        = 0x0000c,	// (12) MAPI32.dll
	msofidHlinkDll         = 0x0000d,	// (13) HLINK.dll
	msofidUrlmonDll        = 0x0000e,	// (14) URLMON.dll
	msofidOleAccDll        = 0x0000f,	// (15) OLEACC.dll
	msofidWsock32Dll       = 0x00010,	// (16) WSOCK32.dll
	msofidMprDll           = 0x00011,	// (17) MPR.dll
	msofidOdma32Dll        = 0x00012,	// (18) ODMA32.dll
	msofidWininetDll       = 0x00013,	// (19) wininet.dll
	msofidRpcrt4Dll        = 0x00014,	// (20) RPCRT4.dll
	msofidImm32Dll         = 0x00015,	// (21) imm32.dll
	msofidShlwapiDll       = 0x00016,	// (22) shlwapi.dll
	msofidMscat32Dll       = 0x00017,	// (23) mscat32.dll
	msofidWsmEngDll        = 0x00018,	// (24) wsmeng.dll
	msofidSetupapiDll      = 0x00019,	// (25) setupapi.dll
	msofidDbgHelpDll       = 0x0001a,	// (26) dbghelp.dll
	msofidSecur32Dll       = 0x0001b,	// (27) secur32.dll
	msofidUxThemeDll       = 0x0001c,	// (28) uxtheme.dll
	msofidMsimg32Dll       = 0x0001d,	// (29) msimg32.dll
	msofidCreduiDll        = 0x0001e,	// (30) credui.dll
	msofidSensapiDll       = 0x0001f,	// (31) sensapi.dll
	msofidWinsock2Dll      = 0x00020,	// (32) ws2_32.dll
	msofidRasapi32Dll      = 0x00021,	// (33) rasapi32.dll
	msofidRasdlgDll        = 0x00022,	// (34) rasdlg.dll
	msofidUcscribeDll      = 0x00023,	// (35) UCSCRIBE.DLL
	msofidImeShareDll      = 0x00024,	// (36) MSOSTYLE.dll
	msofidAWJIMECLPath     = 0x00025,	// (37) imejpcl.aw
	msofidAWJIMESMPath     = 0x00026,	// (38) imejpsm.aw
	msofidUDateDll         = 0x00027,	// (39) IntlDate.DLL
	msofidUspDll           = 0x00028,	// (40) usp10.dll
	msofidBidi32Dll        = 0x00029,	// (41) BIDI32.DLL
	msofidSelfRegDll       = 0x0002a,	// (42) SelfReg$D.dll
	msofidT2EmbedDll       = 0x0002b,	// (43) T2EMBED.DLL
	msofidOdbc32Dll        = 0x0002c,	// (44) ODBC32.DLL
	msofidOdbccp32Dll      = 0x0002d,	// (45) ODBCCP32.dll
	msofidMLangDll         = 0x0002e,	// (46) mlang.dll
	msofidWinspoolDll      = 0x0002f,	// (47) winspool.drv
	msofidUnlockDll        = 0x00030,	// (48) unlock.dll
	msofidmsoDll           = 0x00031,	// (49) mso$d#h.dll
	msofidMsJet40Dll       = 0x00032,	// (50) MSJET40.dll
	msoqfidMsoIntlDll      = 0x00033,	// (51) msointl$d#h.dll
	msoqfidMsoACL          = 0x00034,	// (52) mso.acl
	msofidavifil32Dll      = 0x00035,	// (53) avifil32.dll
	msofidmsvfw32Dll       = 0x00036,	// (54) msvfw32.dll
	msofidMsacm32Dll       = 0x00037,	// (55) msacm32.dll
	msofidIppWFFDll        = 0x00038,	// (56) ippwff.dll
	msofidNetApi32Dll      = 0x00039,	// (57) netapi32.dll
	msofidVbeDll           = 0x0003a,	// (58) vbe6.dll
	msoqfidMsoHelp         = 0x0003b,	// (59) msohelp$d.exe
	msoqfidTOCManifestWord = 0x0003c,	// (60) mf_wdtoc.xml
	msoqfidTOCManifestExcel = 0x0003d,	// (61) mf_xltoc.xml
	msoqfidTOCManifestGraph = 0x0003e,	// (62) mf_xgtoc.xml
	msoqfidTOCManifestPowerPoint = 0x0003f,	// (63) mf_pptoc.xml
	msoqfidTOCManifestAccess = 0x00040,	// (64) mf_actoc.xml
	msoqfidTOCManifestOutlook = 0x00041,	// (65) mf_oltoc.xml
	msoqfidTOCManifestFrontPage = 0x00042,	// (66) mf_fptoc.xml
	msoqfidTOCManifestCag  = 0x00043,	// (67) mf_catoc.xml
	msoqfidTOCManifestJot  = 0x00044,	// (68) mf_ontoc.xml
	msoqfidTOCManifestOis  = 0x00045,	// (69) mf_oitoc.xml
	msoqfidTOCManifestPub  = 0x00046,	// (70) mf_pbtoc.xml
	msoqfidTOCManifestRms  = 0x00047,	// (71) mf_ictoc.xml
	msoqfidTOCManifestVisioPro = 0x00048,	// (72) mf_vptoc.xml
	msoqfidTOCManifestVisioStd = 0x00049,	// (73) mf_votoc.xml
	msoqfidTOCManifestInfoPath = 0x0004a,	// (74) mf_intoc.xml
	msoqfidTOCManifestProjectPro = 0x0004b,	// (75) mf_ptocp.xml
	msoqfidTOCManifestProjectStd = 0x0004c,	// (76) mf_pjtoc.xml
	msoqfidOfficePssHelp   = 0x0004d,	// (77) pss10r.chm
	msoqfidOfficePssOEMHelp = 0x0004e,	// (78) pss10o.chm
	msoqfidOfficeMainHelp  = 0x0004f,	// (79) ofmain11.chm
	msoqfidDaoHelp         = 0x00050,	// (80) dao360.chm
	msoqfidVbaHelp         = 0x00051,	// (81) vblr6.chm
	msoqfidAccessHelpVBTOC = 0x00052,	// (82) vb_actoc.xml
	msoqfidExcelHelpVBTOC  = 0x00053,	// (83) vb_xltoc.xml
	msoqfidFrontPageHelpVBTOC = 0x00054,	// (84) vb_fptoc.xml
	msoqfidMediaStoreHelpTOC = 0x00055,	// (85) catoc.xml
	msoqfidOutlookHelpVBTOC = 0x00056,	// (86) vb_oltoc.xml
	msoqfidPowerPointHelpVBTOC = 0x00057,	// (87) vb_pptoc.xml
	msoqfidVBEHelpTOC      = 0x00058,	// (88) vbetoc.xml
	msoqfidVBSEHelpTOC     = 0x00059,	// (89) vbsetoc.xml
	msoqfidWordHelpVBTOC   = 0x0005a,	// (90) vb_wdtoc.xml
	msoqfidOFCTOCManifest  = 0x0005b,	// (91) mf_catoc.xml
	msoqfidRMCTOCManifest  = 0x0005c,	// (92) mf_ictoc.xml
	msofidNLGSpellGlue     = 0x0005d,	// (93) csapi3t1.dll
	msofidNLGThesaurusGlue = 0x0005e,	// (94) ctapi3t2.dll
	msofidNLGHyphenatorGlue = 0x0005f,	// (95) chapi3t1.dll
	msofidHlinkPrxDll      = 0x00060,	// (96) hlinkprx.dll
	msofidMSQuery          = 0x00061,	// (97) msqry32.exe
	msofidPubPrtf9         = 0x00062,	// (98) prtf9$d.dll
	msofidPubPtxt9         = 0x00063,	// (99) ptxt9$d.dll
	msofidMsInfoExe        = 0x00064,	// (100) msinfo32.exe
	msofidAccessExe        = 0x00065,	// (101) msaccess.exe
	msofidExcelExe         = 0x00066,	// (102) excel.exe
	msofidWordExe          = 0x00067,	// (103) winword$d#h.exe
	msofidPowerPointExe    = 0x00068,	// (104) powerpnt.exe
	msofidOutlookExe       = 0x00069,	// (105) outlook.exe
	msofidBinderExe        = 0x0006a,	// (106) binder$d#h.exe
	msofidExcelLiteExe     = 0x0006b,	// (107) excellt.exe
	msofidWordLiteExe      = 0x0006c,	// (108) wword$d#hlt.exe
	msofidPowerPointLiteExe = 0x0006d,	// (109) powpntlt.exe
	msofidWord9Exe         = 0x0006e,	// (110) winword.exe
	msofidFrontPageExe     = 0x0006f,	// (111) frontpg.exe
	msofidPublisherExe     = 0x00070,	// (112) mspub.exe
	msofidProjectExe       = 0x00071,	// (113) winproj.exe
	msofidFoxProExe        = 0x00072,	// (114) vfp.exe
	msofidOsaExe           = 0x00073,	// (115) osa$d.exe
	msofidPhotoDrawExe     = 0x00074,	// (116) photodrw.exe
	msofidVisioExe         = 0x00075,	// (117) visio.exe
	msofidThemesFontMap    = 0x00076,	// (118) themes.inf
	msofidMsLid            = 0x00077,	// (119) mslid.dll
	msofidCrypt32Dll       = 0x00078,	// (120) crypt32.dll
	msofidWintrustDll      = 0x00079,	// (121) wintrust.dll
	msofidCryptdlgDll      = 0x0007a,	// (122) cryptdlg.dll
	msofidSoftpubDll       = 0x0007b,	// (123) softpub.dll
	msofidAdvapi32Dll      = 0x0007c,	// (124) advapi32.dll
	msofidTypeLibStd2      = 0x0007d,	// (125) stdole2.tlb
	msofidTypeLibVba       = 0x0007e,	// (126) vbaen32.olb
	msofidMsowcDll         = 0x0007f,	// (127) owc11$d.dll
	msofidAnswerWizard     = 0x00080,	// (128) aw$d.dll
	msofidFindFastExe      = 0x00081,	// (129) findfast.exe|ffastd.exe
	msofidMSE10            = 0x00082,	// (130) mse7.exe
	msofidWWordBreakerJ    = 0x00083,	// (131) msgr3jp.dll
	msofidSCWrdBkrDll      = 0x00084,	// (132) msgr3sc.dll
	msofidTCWrdBkrDll      = 0x00085,	// (133) wdbrkcht.dll
	msoqfidVIM             = 0x00086,	// (134) mapivi32.dll
	msofidMsoHevDll        = 0x00087,	// (135) msohev.dll
	msofidMsoXevDll        = 0x00088,	// (136) msoxev.dll
	msofidOISPip           = 0x00089,	// (137) ois.pip
	msofidXDocsPip         = 0x0008a,	// (138) infopath.pip
	msofidStemmerDll       = 0x0008b,	// (139) msstko32.dll
	msoqfidEnvelopeIntlDll = 0x0008c,	// (140) envelopr.dll
	msofidMssign32Dll      = 0x0008d,	// (141) mssign32.dll
	msoqfidOBalloon        = 0x0008e,	// (142) oballoon.dll
	msofidOlepro32Dll      = 0x0008f,	// (143) OLEPRO32.DLL
	msofidCryptuiDll       = 0x00090,	// (144) cryptui.dll
	msoqfidSpeechGrammars  = 0x00091,	// (145) srintl.dll
	msoqfidSpeechTrainingVideo = 0x00092,	// (146) video.mht
	msoqfidSpeechTrainingText = 0x00093,	// (147) offtext.txt
	msofidCenturyFont      = 0x00094,	// (148) century.ttf
	msofidMyPictures       = 0x00095,	// (149) shfolder.dll
	msofidSAExtDll         = 0x00096,	// (150) saext.dll
	msofidCSSeqChkDll      = 0x00097,	// (151) seqchk10.dll
	msofidWdbImpDll        = 0x00098,	// (152) wdbimp.dll
	msofidMsoSvDll         = 0x00099,	// (153) msosv$d.dll
	msoqfidMsoSvIntlDll    = 0x0009a,	// (154) msosvint.dll
	msoqfidMsoICCEnRGBCMYK = 0x0009b,	// (155) rswop.icm
	msoqfidMsoPantoneDLL   = 0x0009c,	// (156) poce98.dll
	msofidCtryInfo         = 0x0009d,	// (157) ctryinfo.txt
	msofidLicenseStore     = 0x0009e,	// (158) opa11.bak
	msoqfidmstoreIntlDLL   = 0x0009f,	// (159) mstintl.dll|MSTIntlD.dll
	msofidmstoreSearchDLL  = 0x000a0,	// (160) mstores.dll|MStoreSD.dll
	msofidActivedsDll      = 0x000a1,	// (161) activeds.dll
	msofidTypeLibFactoid   = 0x000a2,	// (162) mstag.tlb
	msofidViniutilDll      = 0x000a3,	// (163) fpcutl.dll
	msoqfidFPUtlIntlDll    = 0x000a4,	// (164) fputlsat.dll
	msoqfidFPPageKey       = 0x000a5,	// (165) normal.htm
	msoqfidFPThemeKey      = 0x000a6,	// (166) normal.elm
	msoqfidFPSmartPageKey  = 0x000a7,	// (167) spstd1.inf
	msoqfidFPWebPackageKey = 0x000a8,	// (168) weblog.inf
	msoqfidFPFrameKey      = 0x000a9,	// (169) bantoc.htm
	msoqfidFPCSSKey        = 0x000aa,	// (170) normal.inf_0002
	msoqfidFPWebKey        = 0x000ab,	// (171) empty.inf
	msoqfidFPLayoutTmplKey = 0x000ac,	// (172) nl.htm
	msoqfidFPDHTMLActions  = 0x000ad,	// (173) CallJS.js
	msofidSchemasHtml      = 0x000ae,	// (174) ie5_0dom.tlb
	msoqfidAccessibilityRulesets = 0x000af,	// (175) WCAG1_0.xml
	msofidDocumentImagingExe = 0x000b0,	// (176) mspview.exe
	msoqfidNavBarStyles    = 0x000b1,	// (177) navbars.ini
	msoqfidListViewStyles  = 0x000b2,	// (178) lstviews.ini
	msofidVC7Runtime       = 0x000b3,	// (179) FL_msvcr70$d_dll_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidVC7CPPRuntime    = 0x000b4,	// (180) FL_msvcp70$d_dll_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidVC71Runtime      = 0x000b5,	// (181) FL_msvcr71$d_dll_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidVC71CPPRuntime   = 0x000b6,	// (182) FL_msvcp71$d_dll_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidVSAssert         = 0x000b7,	// (183) Vsassert_dll_2_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidATL70Unicode     = 0x000b8,	// (184) FL_atl70_dll_1_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidATL70Ansi        = 0x000b9,	// (185) FL_atl70_dll_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msofidATL70Debug       = 0x000ba,	// (186) atl70dbg_dll_5_____X86.3643236F_FC70_11D3_A536_0090278A1BB8
	msoqfidmstoreMainHelp  = 0x000bb,	// (187) mstore10.chm
	msoqfidmstoreCagCat10  = 0x000bc,	// (188) cagcat10.mml
	msoqfidmstoreOffice10  = 0x000bd,	// (189) office10.mml
	msofidWtsapi32Dll      = 0x000be,	// (190) wtsapi32.dll
	msofidMediaGallery     = 0x000bf,	// (191) mstore$d.exe
	msofidConverterWord97  = 0x000c0,	// (192) mswrd832.cnv
	msofidConverterCore    = 0x000c1,	// (193) msconv97.dll
	msofidNUIQueryDll      = 0x000c2,	// (194) NUIQuery.dll
	msoqfidWordDrmTemplate = 0x000c3,	// (195) prottpln.doc
	msoqfidExcelDrmTemplate = 0x000c4,	// (196) prottpln.xls
	msoqfidPptDrmTemplate  = 0x000c5,	// (197) prottpln.ppt
	msoqfidWordDrmTemplateV = 0x000c6,	// (198) prottplv.doc
	msoqfidExcelDrmTemplateV = 0x000c7,	// (199) prottplv.xls
	msoqfidPptDrmTemplateV = 0x000c8,	// (200) prottplv.ppt
	msofidGDIPLUS          = 0x000c9,	// (201) _3C144D0D917C41E981E59D9C18E43E88.40D5CE2532074296B6DD2138D9286013
	msofidOISExe           = 0x000ca,	// (202) ois.exe
	msoqfidOISIntlDll      = 0x000cb,	// (203) oisintl.dll
	msoqfidAlrtIntlDll     = 0x000cc,	// (204) alrtintl.dll
	msofidMsftEdDll        = 0x000cd,	// (205) msftedit.dll
	msofidOneNote          = 0x000ce,	// (206) onenote.exe
	msofidOneNoteDllIntl   = 0x000cf,	// (207) onintl.dll
	msofidIeawsdc          = 0x000d0,	// (208) ieawsdc.dll
	msofidXDocs            = 0x000d1,	// (209) infopath.exe
	msofidDocImaging       = 0x000d2,	// (210) mspview.exe
	msofidDocScanning      = 0x000d3,	// (211) mspscan.exe
	msofidRMS              = 0x000d4,	// (212) rms.exe
};

enum {
	msocidNonSetup = -2,
	msocidNil = -1,
	msoqcidNil = -1,
	msocidOffice           = 0x00000,	// (0) mso$d#h.dll
	msoqcidMsoIntlDll      = 0x00001,	// (1) msointl$d#h.dll
	msoqcidHelpFiles       = 0x00002,	// (2) msohelp$d.exe
	msoqcidAWWord          = 0x00003,	// (3) wdmain10.aw
	msoqcidAWExcel         = 0x00004,	// (4) xlmain10.aw
	msoqcidAWExcelFunctions = 0x00005,	// (5) workfunc.aw
	msoqcidAWGraph         = 0x00006,	// (6) graph10.aw
	msoqcidAWPowerPoint    = 0x00007,	// (7) ppmain10.aw
	msoqcidAWAccess        = 0x00008,	// (8) acmain10.aw
	msoqcidAWOutlook       = 0x00009,	// (9) olmain10.aw
	msoqcidAWFrontPage     = 0x0000a,	// (10) fpmain10.aw
	msoqcidAWCag           = 0x0000b,	// (11) mstore10.aw
	msoqcidAWJot           = 0x0000c,	// (12) onmain.aw
	msoqcidAWOis           = 0x0000d,	// (13) oismain.aw
	msoqcidAWPub           = 0x0000e,	// (14) pbmain10.aw
	msoqcidAWRms           = 0x0000f,	// (15) rmmain.aw
	msoqcidAWVisioPro      = 0x00010,	// (16) pw.aw
	msoqcidAWVisioStd      = 0x00011,	// (17) visio.aw
	msoqcidAWInfoPath      = 0x00012,	// (18) infmain.aw
	msoqcidAWProjectPro    = 0x00013,	// (19) pjmnpro.aw
	msoqcidAWProjectStd    = 0x00014,	// (20) pjmnstd.aw
	msoqcidNLGSpellingV1   = 0x00015,	// (21) _
	msoqcidNLGSpellingV3   = 0x00016,	// (22) msspell3.dll
	msoqcidNLGGrammarV3    = 0x00017,	// (23) msgr3en.dll
	msoqcidNLGGrammarV2    = 0x00018,	// (24) _
	msoqcidNLGGrammarV1    = 0x00019,	// (25) _
	msoqcidNLGThesaurusV3  = 0x0001a,	// (26) msthes3.dll
	msoqcidNLGThesaurusV1  = 0x0001b,	// (27) _
	msoqcidNLGHyphenatorV2 = 0x0001c,	// (28) mshyph2.dll
	msoqcidNLGHyphenatorV1 = 0x0001d,	// (29) _
	msoqcidNLGDictionary   = 0x0001e,	// (30) msdi_enh.dll
	msoqcidNLGSpellingDataV3 = 0x0001f,	// (31) mssp3en.lex
	msoqcidNLGSpellingDataV1 = 0x00020,	// (32) _
	msoqcidNLGGrammarDataV3 = 0x00021,	// (33) msgr3en.lex
	msoqcidNLGGrammarDataV2 = 0x00022,	// (34) _
	msoqcidNLGGrammarDataV1 = 0x00023,	// (35) _
	msoqcidNLGThesaurusDataV3 = 0x00024,	// (36) msth3am.lex
	msoqcidNLGThesaurusDataV1 = 0x00025,	// (37) _
	msoqcidNLGHyphenatorDataV2 = 0x00026,	// (38) mshy2_en.lex
	msoqcidNLGHyphenatorDataV1 = 0x00027,	// (39) _
	msoqcidNLGDictionaryData = 0x00028,	// (40) msdi_hbe.lex
	msoqcidNLGMorphoLex    = 0x00029,	// (41) mswds_en.lex
	msocidProfWiz          = 0x0002a,	// (42) proflwiz.exe
	msocidOsa              = 0x0002b,	// (43) osa$d.exe
	msocidMsInfoOcx        = 0x0002c,	// (44) oinfo11.ocx
	msoqcidTemplates       = 0x0002d,	// (45) Global_PowerPoint_DesignTemplateMaple
	msocidExcelQueries     = 0x0002e,	// (46) Global_Excel_WebQueries
	msoqcidOfficeDataServices = 0x0002f,	// (47) Global_Office_DataServices
	msoqcidOutlookStationery = 0x00030,	// (48) _
	msoqcidOutlookStationeryBG = 0x00031,	// (49) _
	msoqcidThemes          = 0x00032,	// (50) boldstri.inf
	msoqcidSharedImportConverter = 0x00033,	// (51) recovr32.cnv
	msoqcidSharedExportConverter = 0x00034,	// (52) wrd6er32.cnv
	msocidActorPreviews    = 0x00035,	// (53) blnmgr.dll
	msoqcidFullActors      = 0x00036,	// (54) rocky.acs
	msoqcidImportGraphicFilter = 0x00037,	// (55) _
	msoqcidExportGraphicFilter = 0x00038,	// (56) _
	msocidWTCSCTranslator  = 0x00039,	// (57) _
	msoqcidHHC32Dll        = 0x0003a,	// (58) hhc32.dll
	msoqcidHHCData         = 0x0003b,	// (59) hhc.lex
	msoqcidVbaIntlDLL      = 0x0003c,	// (60) Global_VBA_Intl
	msoqcidMSE_VSPkgsIntl  = 0x0003d,	// (61) msenvui.dll
	msoqcidMSE_UA_MSECore  = 0x0003e,	// (62) Global_UserAssistance_MSECore
	msoqcidMSE_AW_MSEDefault = 0x0003f,	// (63) Global_AnswerWiz_MSEDefault
	msoqcidMSE_DebuggerIntl = 0x00040,	// (64) Global_Vse_DebuggerIntl
	msoqcidServerByClsid   = 0x00041,	// (65) -
	msoqcidVBAHelp         = 0x00042,	// (66) {CC4932ED-A4EC-11d2-B8E2-@0$1#3000F806A77F}
	msocidDpcFile          = 0x00043,	// (67) -
	msoqcidFPCodeSnippets  = 0x00044,	// (68) Global_FrontPageCore_CodeSnippets
	msocidMediaStoreCore   = 0x00045,	// (69) mstore$d.exe
	msoqcidWebComponents   = 0x00046,	// (70) Global_FrontPageCore_PageComponentsIntl
	msocidmstoreThemesLines = 0x00047,	// (71) Global_MediaStore_Office10ClipsLines
	msocidmstoreThemesBullets = 0x00048,	// (72) Global_MediaStore_Office10ClipsBullets
	msoqcidTsDictionary    = 0x00049,	// (73) msb1enfr.its
	msoqcidCAGMMW          = 0x0004a,	// (74) _
	msoqcidCAGMML          = 0x0004b,	// (75) _
	msoqcidCAGAssets       = 0x0004c,	// (76) _
	msocidCheckForOSUpgrade = 0x0004d,	// (77) Global_Office_CheckForOSUpgrade
	msocidDssm             = 0x0004e,	// (78) Global_DocumentServices_dssm
	msocidOIS              = 0x0004f,	// (79) ois.exe
	msoqcidOISIntl         = 0x00050,	// (80) oisintl.dll
	msocidMSXML5           = 0x00051,	// (81) msxml5.dll
	msoqcidMSXML5R         = 0x00052,	// (82) -
	msocidMSSOAP3          = 0x00053,	// (83) mssoap30.dll
	msoqcidMSSOAP3R        = 0x00054,	// (84) -
};

enum {
	msoftidNil = -1,
	msoftidVBAHelp         = 0x00000,	// (0) VBAHelpFiles
	msoftidFFMorph         = 0x00001,	// (1) mswds_en.lex
	msoftidProfWiz         = 0x00002,	// (2) proflwiz.exe
	msoftidThemes          = 0x00003,	// (3) THEMESFiles
	msoftidMSAgent         = 0x00004,	// (4) ASSISTANTFiles
	msoftidMSGraph         = 0x00005,	// (5) graph.exe
	msoftidMSEquation      = 0x00006,	// (6) eqnedt32.exe
	msoftidMSQuery         = 0x00007,	// (7) msqry32.exe
	msoftidWebDrive        = 0x00008,	// (8) WebDriveFiles
	msoftidWebDriveUserData = 0x00009,	// (9) WebDriveUserData
	msoftidOfficeUserData  = 0x0000a,	// (10) OfficeUserData
	msoftidMsoCore         = 0x0000b,	// (11) Global_Office_Core
	msoftidHTMLSourceEditing = 0x0000c,	// (12) HTMLSourceEditing
	msoftidWebScripting    = 0x0000d,	// (13) WebScripting
	msoftidSpeechFeature   = 0x0000e,	// (14) SpeechFiles
	msoftidCiceroFeature   = 0x0000f,	// (15) CiceroFiles
	msoftidOWSWebDiscussions = 0x00010,	// (16) owsclt.dll
	msoftidMediaStore      = 0x00011,	// (17) CAGFiles
	msoftidCAGCategoryFiles = 0x00012,	// (18) CAGCategoryFiles
	msoftidmstoreOfficeThemes = 0x00013,	// (19) CAGOffice10
	msoftidVBAFiles        = 0x00014,	// (20) VBAFiles
	msoftidWebDebugging    = 0x00015,	// (21) WebDebugging
	msoftidCAGCat10        = 0x00016,	// (22) CAGCat10
	msoftidCAGCat10Clips   = 0x00017,	// (23) CAGCat10Clips
	msoftidCAGOffice10Clips = 0x00018,	// (24) CAGOffice10Clips
	msoftidTablet          = 0x00019,	// (25) WISPFiles
};

// Remapped identifiers
#define CVT_CORE_VERSION_LS 0x14B90000
#define CVT_CORE_VERSION_MS 0x07D3044C
#define CVT_CORE_VERSION_STRING L"2003.1100.5305.0"
#define CVT_WORD97_VERSION_LS 0x14B90000
#define CVT_WORD97_VERSION_MS 0x07D3044C
#define CVT_WORD97_VERSION_STRING L"2003.1100.5305.0"
#define VBA_VERSION_LS 0x0061003B
#define VBA_VERSION_MS 0x00060004
#define VBA_VERSION_STRING L"6.4.97.59"
#define msofidMSE msofidMSE10
#define msoqfidVbaErrorHelp msoqfidVbaHelp

// GUID defines
#ifndef DEBUG
#define ACCESS_CORE_GUID L"{F2D782F8-6B14-4FA4-8FBA-565CDDB9B2A8}"
#define BI_CORE_GUID L"{246A505C-24A7-479F-9342-3C5D5C5B0F9B}"
#define EXCEL_CORE_GUID L"{A2B280D4-20FB-4720-99F7-40C09FBCE10A}"
#define FRONTPAGE_CORE_GUID L"{81E9830C-5A6B-436A-BEC9-4FB759282DE3}"
#define OIS_CORE_GUID L"{5B5BE2F2-49E4-4125-8CBB-7805ACDB87E6}"
#define ONENOTE_CORE_GUID L"{D2C0E18B-C463-4E90-92AC-CA94EBEC26CE}"
#define OUTLOOK_CORE_GUID L"{3CE26368-6322-4ABF-B11B-458F5C450D0F}"
#define POWERPNT_CORE_GUID L"{C86C0B92-63C0-4E35-8605-281275C21F97}"
#define PUBLISHER_CORE_GUID L"{8F71499B-7879-41CB-90F9-A64C5719CB6F}"
#define RMS_CORE_GUID L"{02BA4BD0-2066-40A0-B161-31122F2CDBCB}"
#define WORD_CORE_GUID L"{1EBDE4BC-9A51-4630-B541-2561FA45CCC5}"
#define XDOCS_CORE_GUID L"{1A66B512-C4BE-4347-9F0C-8638F8D1E6E4}"
#else
#define ACCESS_CORE_GUID L"{F2D782F8-6B14-4FA4-8FBA-165CDDB9B2A8}"
#define BI_CORE_GUID L"{246A505C-24A7-479F-9342-1C5D5C5B0F9B}"
#define EXCEL_CORE_GUID L"{A2B280D4-20FB-4720-99F7-10C09FBCE10A}"
#define FRONTPAGE_CORE_GUID L"{81E9830C-5A6B-436A-BEC9-1FB759282DE3}"
#define OIS_CORE_GUID L"{5B5BE2F2-49E4-4125-8CBB-1805ACDB87E6}"
#define ONENOTE_CORE_GUID L"{D2C0E18B-C463-4E90-92AC-1A94EBEC26CE}"
#define OUTLOOK_CORE_GUID L"{3CE26368-6322-4ABF-B11B-158F5C450D0F}"
#define POWERPNT_CORE_GUID L"{C86C0B92-63C0-4E35-8605-181275C21F97}"
#define PUBLISHER_CORE_GUID L"{8F71499B-7879-41CB-90F9-164C5719CB6F}"
#define RMS_CORE_GUID L"{02BA4BD0-2066-40A0-B161-11122F2CDBCB}"
#define WORD_CORE_GUID L"{1EBDE4BC-9A51-4630-B541-1561FA45CCC5}"
#define XDOCS_CORE_GUID L"{1A66B512-C4BE-4347-9F0C-1638F8D1E6E4}"
#endif

#endif // TCDARO_H
