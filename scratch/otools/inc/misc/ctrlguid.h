//+---------------------------------------------------------------------------
//
//  Microsoft Forms
//  Copyright (C) Microsoft Corporation, 1992 - 1995.
//
//  File:       ctrlguid.h
//
//  Contents:   extern references for forms guids
//
//----------------------------------------------------------------------------

// Please check \forms3\src\site\include\siteguid.h for information about GUID

#ifndef __CTRLGUID_H__
#define __CTRLGUID_H__

// Use PUBLIC_GUID for GUIDs used outside FORMS3.DLL.
// Use PRIVATE_GUID for all other GUIDS.

#ifndef PUBLIC_GUID
#define PUBLIC_GUID(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8) DEFINE_GUID(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8);
#endif

#ifndef PRIVATE_GUID
#define PRIVATE_GUID(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8) DEFINE_GUID(name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8);
#endif

#ifndef MIDL_COMPILER
#ifndef PRODUCT_96
EXTERN_C const GUID CLSID_CCDGenericPropertyPage;
EXTERN_C const GUID CLSID_CCDFontPropertyPage;
EXTERN_C const GUID CLSID_CCDColorPropertyPage;
#endif
#endif

// Classes
#ifdef PRODUCT_97
PUBLIC_GUID(CLSID_CCommandButton95,         0x3050f180, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CComboBox95,              0x3050f181, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CTextBox95,               0x3050f182, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CCheckBox95,              0x3050f183, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_COptionButton95,          0x3050f184, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CListBox95,               0x3050f185, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CGroupBox95,              0x3050f186, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CHorizontalScrollBar95,   0x3050f187, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CVerticalScrollBar95,     0x3050f188, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CPropTabs,                0x3050f189, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CLabel95,                 0x3050f18a, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
#endif
#ifdef PRODUCT_96
PUBLIC_GUID(CLSID_CImage,                   0x4C599241,0x6926,0x101B,0x99,0x92,0x00,0x00,0x0B,0x65,0xC6,0xF9)
PUBLIC_GUID(CLSID_CRichTextObject,          0x00020CF5,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46)
PUBLIC_GUID(CLSID_CLabelControl,            0x978C9E23,0xD4B0,0x11ce,0xBF,0x2D,0x00,0xAA,0x00,0x3F,0x40,0xD0)
PUBLIC_GUID(CLSID_CCommandButton,           0xD7053240,0xCE69,0x11CD,0xA7,0x77,0x00,0xDD,0x01,0x14,0x3C,0x57)
PUBLIC_GUID(CLSID_CMorphDataControl,        0x3bcfb910,0x370d,0x11ce,0x8e,0xa0,0x00,0xaa,0x00,0x4b,0xa6,0xae)
PUBLIC_GUID(CLSID_CMdcText,                0x8bd21d10,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(CLSID_CMdcList,                0x8bd21d20,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(CLSID_CMdcCombo,               0x8bd21d30,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(CLSID_CMdcCheckButton,         0x8bd21d40,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(CLSID_CMdcOptionButton,        0x8bd21d50,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(CLSID_CMdcToggleButton,        0x8bd21d60,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(CLSID_CScrollbar,               0XDFD181E0,0X5E2F,0X11CE,0XA4,0X49,0X00,0XAA,0X00,0X4A,0X80,0X3D)
PUBLIC_GUID(CLSID_CSpinbutton,              0x79176fb0,0xb7f2,0x11ce,0x97,0xef,0x00,0xaa,0x00,0x6d,0x27,0x76)
PUBLIC_GUID(CLSID_CTabStrip,                0xeae50eb0,0x4a62,0x11ce,0xbe,0xd6,0x00,0xaa,0x00,0x61,0x10,0x80)
PUBLIC_GUID(CLSID_CTab,                     0xa38bffc0,0xa5a0,0x11ce,0x81,0x07,0x00,0xaa,0x00,0x61,0x10,0x80)
PUBLIC_GUID(CLSID_CTabs,                    0x944acf90,0xa1e6,0x11ce,0x81,0x04,0x00,0xaa,0x00,0x61,0x10,0x80)
PRIVATE_GUID(CLSID_CTreeViewControl,        0x038ec050,0x4f44,0x11ce,0x8e,0x59,0x00,0xaa,0x00,0x60,0xf3,0x06)
PRIVATE_GUID(CLSID_CCDControlSelector,      0x7cbbabf0,0x36b9,0x11ce,0xbf,0x0d,0x00,0xaa,0x00,0x44,0xbb,0x60)
PRIVATE_GUID(CLSID_WHTMLSubmit,             0x5512d110,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLImage,              0x5512d112,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLReset,              0x5512d114,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLCheckbox,           0x5512d116,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLOption,             0x5512d118,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLText,               0x5512d11a,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLHidden,             0x5512d11c,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLPassword,           0x5512d11e,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLSelect,             0x5512d122,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
PRIVATE_GUID(CLSID_WHTMLTextArea,           0x5512d124,0x5cc6,0x11cf,0x8d,0x67,0x00,0xaa,0x00,0xbd,0xce,0x1d)
#else
PUBLIC_GUID(CLSID_CImage,              0x3050f18b, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CRichTextObject,     0x3050f18c, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CLabelControl,       0x3050f18d, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CCommandButton,      0x3050f18e, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMorphDataControl,   0x3050f18f, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMdcText,            0x3050f190, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMdcList,            0x3050f191, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMdcCombo,           0x3050f192, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMdcCheckButton,     0x3050f193, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMdcOptionButton,    0x3050f194, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CMdcToggleButton,    0x3050f195, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CScrollbar,          0x3050f196, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CSpinbutton,         0x3050f197, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CTabStrip,           0x3050f198, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CTab,                0x3050f199, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CTabs,               0x3050f19a, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CTreeViewControl,   0x3050f19b, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDControlSelector, 0x3050f19c, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PUBLIC_GUID(CLSID_CDummyControl,       0x3050f20d, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)

// CLSID_WHTML... uses exclusively in Forms3 96
#endif

// Property pages
#ifdef PRODUCT_97
PRIVATE_GUID(CLSID_CCDLabelPropertyPage,         0x3050f1a7, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDCheckBoxPropertyPage,      0x3050f1a8, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDTextBoxPropertyPage,       0x3050f1a9, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDOptionButtonPropertyPage,  0x3050f1aa, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDListBoxPropertyPage,       0x3050f1ab, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDCommandButtonPropertyPage, 0x3050f1ac, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDComboBoxPropertyPage,      0x3050f1ad, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
//PRIVATE_GUID(CLSID_CCDScrollBarPropertyPage,   0x3050f1ae, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDGroupBoxPropertyPage,      0x3050f1af, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDImagePropertyPage,         0x3050f1b0, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDHorzScrollBarPropertyPage, 0x3050f1b1, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDVertScrollBarPropertyPage, 0x3050f1b2, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDLabel2PropertyPage,        0x3050f1b3, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDTextBox2PropertyPage,      0x3050f1b4, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDComboBox2PropertyPage,     0x3050f1b5, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDListBox2PropertyPage,      0x3050f1b6, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDCommandButton2PropertyPage,0x3050f1b7, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDImage2PropertyPage,        0x3050f1b8, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDCtrlSelPropertyPage,       0x3050f1b9, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CCDCtrlSelItemPropertyPage,   0x3050f1ba, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
#endif // PRODUCT_97

// Misc
#ifdef PRODUCT_96
PRIVATE_GUID(CLSID_CtrlSelTemplate,      0x0274ef50,0x555a,0x11ce,0xbf,0x10,0x00,0xaa,0x00,0x44,0xbb,0x60)
PRIVATE_GUID(CLSID_FormGroup,            0x379a2500,0x556f,0x11ce,0xbf,0x10,0x00,0xaa,0x00,0x44,0xbb,0x60)
PRIVATE_GUID(CLSID_CFontNew,             0xafc20920,0xda4e,0x11ce,0xb9,0x43,0x00,0xaa,0x00,0x68,0x87,0xb4)
#else
PRIVATE_GUID(CLSID_CtrlSelTemplate,      0x3050f1bb, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_FormGroup,            0x3050f1bc, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
PRIVATE_GUID(CLSID_CFontNew,             0x3050f1bd, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b)
#endif

// Dispinterfaces
#ifdef PRODUCT_97
PUBLIC_GUID(IID_CommandButton95Events,      0xce5c3540,0x3c99,0x101b,0x97,0x8c,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_ListBox95Events,            0xce5c3608,0x3c99,0x101b,0x97,0x8c,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_ComboBox95Events,           0xd3bef6cc,0x3ca7,0x101b,0x97,0x8d,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_TextBox95Events,            0xce5c3734,0x3c99,0x101b,0x97,0x8c,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_CheckBox95Events,           0xce5c3798,0x3c99,0x101b,0x97,0x8c,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_OptionButton95Events,       0xce5c38c4,0x3c99,0x101b,0x97,0x8c,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_GroupBox95Events,           0xd3bef604,0x3ca7,0x101b,0x97,0x8d,0x00,0x00,0x0b,0x65,0xc6,0x85)
PUBLIC_GUID(IID_HorizontalScrollBar95Events,0xc9f4ff40,0x49e2,0x101b,0x97,0x9f,0x00,0x00,0x0b,0x65,0xc0,0x8b)
PUBLIC_GUID(IID_VerticalScrollBar95Events,  0xff7c5988,0x49e7,0x101b,0x97,0xa0,0x00,0x00,0x0b,0x65,0xc0,0x8b)
PRIVATE_GUID(IID_PropTabsEvents,            0x2ea9e390,0xbbaa,0x11cd,0xba,0x41,0x00,0xaa,0x00,0x4a,0xe2,0x5a)
PRIVATE_GUID(IID_TreeViewControl,           0x038ec051,0x4f44,0x11ce,0x8e,0x59,0x00,0xaa,0x00,0x60,0xf3,0x06)
PRIVATE_GUID(IID_TreeViewEvents,            0x7b020ec3,0xaf6c,0x11ce,0x9f,0x46,0x00,0xaa,0x00,0x57,0x4a,0x4f)
PUBLIC_GUID(IID_Label95Events,              0xce5c37fc,0x3c99,0x101b,0x97,0x8c,0x00,0x00,0x0b,0x65,0xc6,0x85)
#endif

PUBLIC_GUID(IID_ImageEvents,                0x4c5992a5,0x6926,0x101b,0x99,0x92,0x00,0x00,0x0b,0x65,0xc6,0xf9)
PUBLIC_GUID(IID_LabelControlEvents,         0x978C9E22,0xD4B0,0x11ce,0xBF,0x2D,0x00,0xAA,0x00,0x3F,0x40,0xD0)
PUBLIC_GUID(IID_CommandButtonEvents,        0x7b020ec1,0xaf6c,0x11ce,0x9f,0x46,0x00,0xaa,0x00,0x57,0x4a,0x4f)
PUBLIC_GUID(IID_MorphDataControlEvents,     0x7b020ec4,0xaf6c,0x11ce,0x9f,0x46,0x00,0xaa,0x00,0x57,0x4a,0x4f)
PUBLIC_GUID(IID_MdcTextEvents,              0x8bd21d12,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(IID_MdcListEvents,              0x8bd21d22,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(IID_MdcComboEvents,             0x8bd21d32,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(IID_MdcCheckBoxEvents,          0x8bd21d42,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(IID_MdcOptionButtonEvents,      0x8bd21d52,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(IID_MdcToggleButtonEvents,      0x8bd21d62,0xec42,0x11ce,0x9e,0x0d,0x00,0xaa,0x00,0x60,0x02,0xf3)
PUBLIC_GUID(IID_TabStripEvents,             0x7b020ec7,0xaf6c,0x11ce,0x9f,0x46,0x00,0xaa,0x00,0x57,0x4a,0x4f)
PUBLIC_GUID(IID_ScrollbarEvents,            0x7b020ec2,0xaf6c,0x11ce,0x9f,0x46,0x00,0xaa,0x00,0x57,0x4a,0x4f)
PUBLIC_GUID(IID_SpinbuttonEvents,           0x79176fb2,0xb7f2,0x11ce,0x97,0xef,0x00,0xaa,0x00,0x6d,0x27,0x76)
PUBLIC_GUID(IID_WHTMLControlEvents,         0x796ed650,0x5fe9,0x11cf,0x8d,0x68,0x00,0xaa,0x00,0xbd,0xce,0x1d)

// GUID for the IRowset property of a databinding control
PUBLIC_GUID(IID_PIROWSET,                0x0c733a00,0x2a1c,0x11ce,0xad,0xe5,0x00,0xaa,0x00,0x44,0x77,0x3d)

// GUID for the WordIA special control interface
PUBLIC_GUID(IID_SimpleWebControl,       0x7e554e90,0x5cd0,0x11cf,0xb1,0xd5,0x00,0xaa,0x00,0xbd,0xce,0x0b)

#endif
