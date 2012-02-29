/****************************************************************************
*                                                                           *
* Copyright (c) 1999 - 2002, Microsoft Corporation.  All rights reserved.   *
*                                                                           *
* File       - WsmMeta.h                                                    *
*                                                                           *
* Purpose    - Contains text strings used in compiled meta data.            *
*                                                                           *
* History    -     Date        Alias         Comment                        *
*              26-Dec-2001     chrisab       Created                        *
*                                                                           *
****************************************************************************/

#pragma once

#ifndef _WSMMETA_H_
#define _WSMMETA_H_

//____________________________________________________________________________
//
// Standard namespace properties, as defined in state manifest.
// Accessible from property namespace obtained by GetNamespace(0)
//____________________________________________________________________________

//
// Standard namespace properties
//

#define WSMNP_DisplayName        ((const UCHAR *) "DisplayName")       // string
#define WSMNP_Manufacturer       ((const UCHAR *) "manufacturer")      // string
#define WSMNP_Description        ((const UCHAR *) "description")       // string
#define WSMNP_Language           ((const UCHAR *) "language")          // integer
#define WSMNP_MajorVersion       ((const UCHAR *) "majorVersion")      // integer
#define WSMNP_MinorVersion       ((const UCHAR *) "minorVersion")      // integer
#define WSMNP_NamespaceId        ((const UCHAR *) "Id")                // binary
#define WSMNP_ProductName        ((const UCHAR *) "ProductName")       // string
#define WSMNP_Roam               ((const UCHAR *) "roam")              // boolean
#define WSMNP_Schema             ((const UCHAR *) "schema")            // string
#define WSMNP_StoreFlags         ((const UCHAR *) "StoreFlags")        // integer
#define WSMNP_ItemEstimate       ((const UCHAR *) "ItemEstimate")      // integer
#define WSMNP_MaxStoreSize       ((const UCHAR *) "maxStoreSize")      // integer, in KB
#define WSMNP_MinStoreSize       ((const UCHAR *) "minStoreSize")      // integer, in KB
#define WSMNP_MaxStoreDays       ((const UCHAR *) "maxStoreDays")      // integer, in Days
#define WSMNP_MinStoreDays       ((const UCHAR *) "minStoreDays")      // integer, in Days
#define WSMNP_ArchiveInterval    ((const UCHAR *) "archiveInterval")   // integer

//_____________________________________________________________________________
//
// Compiled manifest metadata schema definitions (Alphabetical order)
//_____________________________________________________________________________

#define szCompositePrefix   ((const UCHAR*)"$C?")
#define szDependencyPrefix  ((const UCHAR*)"$D|")
#define szExpressionPrefix  ((const UCHAR*)"$E|")
#define szExpressionIndex   ((const UCHAR*)"$E?")
#define szGroupPrefix       ((const UCHAR*)"$G|")  // followed by group name, '|', and attribute name
#define szGroupIndex        ((const UCHAR*)"$G?")
#define szHandlerPrefix     ((const UCHAR*)"$H|")
#define szHandlerIndex      ((const UCHAR*)"$H?")
#define szSetupDetectPrefix ((const UCHAR*)"$I|")
#define szScriptCodePrefix  ((const UCHAR*)"$J|")
#define szPropertyPrefix    ((const UCHAR*)"$P|")  // followed by property name
#define szPropertyIndex     ((const UCHAR*)"$P?")  // followed by 1-based sequnce number
#define szScriptPrefix      ((const UCHAR*)"$R|")   
#define szScriptIndex       ((const UCHAR*)"$R?")
#define szSettingPrefix     ((const UCHAR*)"$S|")  // followed by setting name, '|', and attribute name
#define szSettingIndex      ((const UCHAR*)"$S?")  // followed by 1-based sequence number
#define szTranslatePrefix   ((const UCHAR*)"$T|")
#define szTranslateIndex    ((const UCHAR*)"$T?")
#define szExecDetectPrefix  ((const UCHAR*)"$X|")

#define szSettingGroup          ((const UCHAR*)"$Group")
#define szSettingFlags          ((const UCHAR*)"$Flags")
#define szSettingElement        ((const UCHAR*)"$Element")
#define szSettingRange          ((const UCHAR*)"$Range")
#define szSettingQualifier      ((const UCHAR*)"$Qualifier")
#define szSettingMruBase        ((const UCHAR*)"$MruBase")
#define szSettingMruSort        ((const UCHAR*)"$MruSort")
#define szSettingComment        ((const UCHAR*)"$Comment")
#define szSettingComposite      ((const UCHAR*)"$Composite")
#define szSettingBlob           ((const UCHAR*)"$Blob")
#define szSettingDescription    ((const UCHAR*)"$Description")
#define szSettingLength         ((const UCHAR*)"$Length")
#define szSettingPad            ((const UCHAR*)"$Pad")
#define szSettingCount          ((const UCHAR*)"$Count")
#define szSettingBitPad         ((const UCHAR*)"$BitPad")
#define szSettingBitLength      ((const UCHAR*)"$BitLength")
#define szSettingPrefixLength   ((const UCHAR*)"$PrefixLength")
#define szSettingExplicit       ((const UCHAR*)"$Explicit")

#define szSettingWhen           ((const UCHAR*)"$When")
#define szSettingStart          ((const UCHAR*)"$Start")
#define szSettingEnd            ((const UCHAR*)"$End")

#define szGroupTarget           ((const UCHAR*)"$Target")
#define szGroupHandler          ((const UCHAR*)"$Handler")
#define szGroupFlags            ((const UCHAR*)"$Flags")

#define szDetectPath            ((const UCHAR*)"$Path")
#define szDetectType            ((const UCHAR*)"$Type")
#define szDetectPattern         ((const UCHAR*)"$Pattern")

#define szHandlerID             ((const UCHAR*)"$ID")


#endif // _WSMMETA_H_
