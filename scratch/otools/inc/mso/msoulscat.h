#pragma once

/****************************************************************************
	MsoULSId.h

	Owner: TCoon
	Copyright (c) 2001 Microsoft Corporation

	This file contains developer definable Logging categories and event ids. 
	Please add new categories to the end of the enum. 

	New categories must be set to a value of the form 'XX' where XX is a two digit number
	Alphabetic characters are not allowed in the value.	
****************************************************************************/

#ifndef MSOULSCAT_H
#define MSOULSCAT_H

// Developer defineable Category Ids
typedef enum _msoulscategoryid
{
	msoulscatTbInput = 28,
	msoulscat_faxconfig_cs = 29,
	msoulscat_faxlistenerservice_cs = 30,
	msoulscat_spredirector_cs = 31,
	msoulscat_imagecontrols_cs = 32,
	msoulscat_ups_cs = 33,
	msoulscatSDMDrawing = 34,
	msoulscatUnknown = 35,
	msoulscat_faxhandler_cs = 36,
	msoulscat_faxnotifier_cs = 37,
	msoulscat_myprofile_core_cs = 38,
	msoulscat_faxsender_cs = 39,
	msoulscat_dmuplbar_cpp = 40,
	msoulscat_dmuoldoc_cpp = 41,
	msoulscat_dmuopen_cpp = 42,
	msoulscat_faxlistener_cs = 43,
	msoulscat_formcontrols_cs = 44,
	msoulscat_miscellaneouscontrols_cs = 45,
	msoulscat_faxservicebase_cs = 46,
	msoulscat_faxjobs_cs = 47,
	msoulscat_faxhandlerservice_cs = 48,
	msoulscatTbDrawing = 49,
	msoulscat_titleareacontrols_cs = 50,
	msoulscat_anchorcontrols_cs = 51,
	msoulscat_faxservice_cs = 52,
	msoulscat_faxwizardfakeselectdocs_cs = 53,
	msoulscatMemoryAlloc = 54,
	msoulscat_tablecontrols_cs = 55,
	msoulscat_panelcontrols_cs = 56,
	msoulscat_validationcontrols_cs = 57,
	msoulscat_hailstorm_cs = 58,
	msoulscat_faxresourcemanager_cs = 59,
	msoulscat_actionpp_cpp = 60,
	msoulscat_oaentry_cpp = 61,
	msoulscat_toolbarcontrols_cs = 62,
	msoulscat_ratemessage_cs = 63,
	msoulscat_calendar_blom_cs = 64,
	msoulscat_postui_cs = 65,
	msoulscat_discussiongroups_cs = 66,
	msoulscat_profileedit_cs = 67,
	msoulscat_msgget_cs = 68,
	msoulscat_awspage_cs = 69,
	msoulscat_contacts_blom_cs = 70,
	msoulscat_calendar_cs = 71,
	msoulscat_profile_blom_cs = 72,
	msoulscat_dgmessagelist_cs = 73,
	msoulscat_profileview_cs = 74,
	msoulscat_contacts_cs = 75,
	msoulscat_navdata_cs = 76,
	msoulscat_postsend_cs = 77,
	msoulscat_postsetscreenname_cs = 78,
	msoulscat_uls_cs = 79,
	msoulscat_commoncontrols_scs = 80,
	msoulscat_myofficelanguagesettings_cs = 81,
	msoulscat_webpartzone_cs = 82,
	msoulscat_scriptwebpart_cs = 83,
	msoulscat_enumpropertyrow_cs = 84,
	msoulscat_partcache_cs = 85,
	msoulscat_smartpage_cs = 86,
	msoulscat_webpart_cs = 87,
	msoulscat_contentwebpart_cs = 88,
	msoulscat_threadnotify_cs = 89,
	msoulscat_notify_cs = 90,
	msoulscat_msgread_cs = 91,
	msoulscat_frameworktoolpart_cs = 92,
	msoulscat_contentframetoolpart_cs = 93,
	msoulscat_imagetoolpart_cs = 94,
	msoulscat_contenttoolpart_cs = 95,
	msoulscat_permit_cs = 96,
	msoulscat_htmltrloadbalancer_cs = 97,
	msoulscat_htmltranslation_cs = 98,
	msoulscat_ossmresourcemanager_cs_pp = 99,
	msoulscat_faxwizardsubmitfax_scs = 100,
	msoulscat_cachedobjs_cs = 101,
	msoulscat_menucontrols_cs = 102,
	msoulscat_faxsend_scs = 103,
	msoulscat_commoncontrols_cs_pp = 104,
	msoulscat_vendorfax_cs = 105,
	msoulscat_htmltrlauncher_cs = 106,
	msoulscat_querylogging_cs = 107,
	msoulscat_htmltrcache_cs = 108,
	msoulscat_faxwizardsubmitconfirmation_scs = 109,
	msoulscat_request_cs = 110,
	msoulscat_workitem_cs = 111,
	msoulscat_worker_cs = 112,
	msoulscat_manager_cs = 113,
	msoulscat_assetmapping_cs = 114,
	msoulscat_searchaws_cs = 115,
	msoulscat_propertysheetctl_cs = 116,
	msoulscat_toolpane_cs = 117,
	msoulscat_toolpart_cs = 118,
	msoulscat_searchresulttable_cs = 119,
	msoulscat_importwebpart_cs = 120,
	msoulscat_importtoolpart_cs = 121,
	msoulscat_gallery_cs = 122,
	msoulscat_htmltrloadbalancerservice_cs = 123,
	msoulscat_htmltrlauncherservice_cs = 124,
	msoulscat_oxml_cs = 125,
	msoulscat_util_cs = 126,
	msoulscat_errorpage_cs_pp = 127,
	msoulscat_htmltranslate_cs = 128,
	msoulscat_myprofile_ui_cs = 129,
	UseSDA = 130,
	msoulscat_bdk_cs = 131,
	msoulscat_connectiondesign_cs = 132,
	msoulscat_faxcost_scs = 133,
	msoulscat_faxpreview_scs = 134,
	msoulscat_smartpageswebservice_cs = 135,
	msoulscat_httpmodule_cs = 136,
	msoulscat_faxrecipients_scs = 137,
	msoulscat_settingsgroups_cs = 138,
	msoulscat_utilities_cs = 139,
	msoulscat_contentrating_cs = 140,
	msoulscat_awsquerythreadpool_cs = 141,
	msoulscat_msgprint_cs = 142,
	msoulscat_searchcontrol_cs_pp = 143,
	msoulscat_imagecontrols_cs_pp = 144,
	msoulscat_trailnav_cs_pp = 145,
	msoulscat_results_cs = 146,
	msoulscat_messagebody_cs = 147,
	msoulscat_portalpreview_cs = 148,
	msoulscat_templatedownload_cs = 149,
	msoulscat_redir_cs = 150,
	msoulscat_frutil_cs = 151,
	msoulscat_trainingbase_cs = 152,
	msoulscat_mediadownload_cs = 153,
	msoulscat_ttp_cs = 154,
	msoulscat_frqueue_cs = 155,
	msoulscat_filerepair_aspx_cs = 156,
	msoulscat_portalproviderpreview_cs = 157,
	msoulscat_helper_cs = 158,
	msoulscat_filemetadata_cs = 159,
	msoulscat_repairstate_cs = 160,
	msoulscat_partnernew_cs = 161,
	msoulscat_toc_cs = 162,
	msoulscat_ratingcontrol_cs = 163,
	msoulscat_clipartdownload_cs = 164,
	msoulscat_htmltrservicebase_cs = 165,
	msoulscat_peoplerating_cs = 166,
	msoulscat_xformstatus_cs = 167,
	msoulscat_repairstore_cs = 168,
	msoulscat_discussiongroup_cs = 169,
	msoulscat_cddg_cs = 170,
	msoulscat_download_aspx_cs = 171,
	msoulscat_messagelists_cs = 172,
	msoulscat_srp_cs = 173,
	msoulscat_partnereditlisting_cs = 174,
	msoulscat_filerepair_saspx = 175,
	msoulscat_default_aspx_cs = 176,
	msoulscat_userinputsettings_cs = 177,
	msoulscat_partnermain_cs = 178,
	msoulscat_router_cs = 179,
	msoulscat_feedback_cs = 180,
	msoulscat_provisionservices_asmx_cs = 181,
	msoulscat_default_cs = 182,
	msoulscat_currentrepair_cs = 183,
	msoulscat_clientsector_asmx_cs = 184,
	msoulscat_dbwrapper_cs = 185,
	msoulscat_cddgsearch_cs = 186,
	msoulscat_admin_cs = 187,
	msoulscat_survey_aspx_cs = 188,
	msoulscat_common_cs = 189,
	msoulscat_sessionmgr_cs = 190,
	msoulscat_postrating_cs = 191,
	msoulscat_partnereditcontact_cs = 192,
	msoulscat_dglxasset_cs = 193,
	msoulscat_dglx_cs = 194,
	msoulscat_userprofile_cs = 195,
	msoulscat_portalcategory_cs = 196,
	msoulscat_test_sidmapping_cs = 197,
	msoulscat_fsauthzpermitupdate_cpp = 198,
	msoulscat_basket_cs = 199,
	msoulscat_advsearchcontrol_cs = 200,
	msoulscat_preview_cs = 201,
	msoulscat_category_cs = 202,
	msoulscat_spotlightcontrol_cs = 203,
	msoulscat_awslogging_cs = 204,
	msoulscat_UseSDA = 205,
	msoulscat_applogin_aspx_cs = 206,
	msoulscat_default_aspx_cs_pp = 207,
	msoulscat_newmws_aspx_cs = 208,
	msoulscat_search_cs = 209,
	msoulscat_xmlreader_cpp = 210,
	msoulscat_multitrieindexservice_cpp = 211,
	msoulscat_queryserviceimpl_cpp = 212,
	msoulscat_rrisapi_cpp = 213,
	msoulscat_rrsc_cpp = 214,
	msoulscat_bookarticleinfo_cpp = 215,
	msoulscat_callback_cpp = 216,
	msoulscat_courselist_cs = 217,
	msoulscat_urlparser_cpp = 218,
	msoulscat_rrcharmap_cpp = 219,
	msoulscat_searchroutingserviceimpl_cpp = 220,
	msoulscat_rrutil_cpp = 221,
	msoulscat_editorialassets_cs = 222,
	msoulscat_articleinfoserviceimpl_cpp = 223,
	msoulscat_rrcharmaps_cpp = 224,
	msoulscat_serviceproxy_cs = 225,
	msoulscat_testulsexception_saspx = 226,
	msoulscat_register_cs = 227,
	msoulscat_genresxids_xsl = 228,
	msoulscat_test_saspx = 229,
	msoulscat_formreg_cs = 230,
	msoulscat_global_sasax = 231,
	msoulscat_smtp_cs = 232,
	msoulscat_msnsearch_cs = 233,
	msoulscat_rutil_cpp = 234,
	msoulscat_testfile_cs = 235,
	msoulscat_xmlwebpart_cs = 236,
	msoulscat_xmltoolpart_cs = 237,
	msoulscat_testloadbalancer_cs = 238,
	msoulscat_urlassign_saspx = 239,
	msoulscat_usernew_saspx = 240,
	msoulscat_userrequestaccess_saspx = 241,
	msoulscat_urlgroupnew_saspx = 242,
	msoulscat_userwaitinglist_saspx = 243,
	msoulscat_urluserassociate_saspx = 244,
	msoulscat_usergroupdelete_saspx = 245,
	msoulscat_userdelete_saspx = 246,
	msoulscat_urlassignowner_saspx = 247,
	msoulscat_urlgroupdelete_saspx = 248,
	msoulscat_previewpagingbar_cs = 249,
	msoulscat_userassign_saspx = 250,
	msoulscat_usergroupnew_saspx = 251,
	msoulscat_worldwide_cs_pp = 252,
	msoulscat_operfdll_cpp = 253,
	msoulscat_factfinder_cs = 254,
	msoulscat_threadpool_cpp = 255,
	msoulscat_export_cpp = 256,
	msoulscat_moneycommon_cs = 257,
	msoulscat_moneycentral_cs = 258,
	msoulscat_formcontrols_cs_pp = 259,
	msoulscat_utility_cs = 260,
	msoulscat_taskwatcher_cs = 261,
	msoulscat_slmanager_cs = 262,
	msoulscat_suggestions_cs = 263,
	msoulscat_contactus_cs = 264,
	msoulscat_main_cs = 265,
	msoulscat_slcache_cs = 266,
	msoulscat_partnerpage_cs = 267,
	msoulscat_ofsauth_cpp = 268,
	msoulscat_awsjukebox_cs = 269,
	msoulscat_internaltoolspage_cs = 270,
	msoulscat_varyrankedparams_cs = 271,
	msoulscat_keysprotection_cs = 272,
	msoulscat_default_saspx = 273,
	msoulscat_csa_cs = 274,
	msoulscat_dcresults_cs = 275,
	msoulscat_topcategory_cs = 276,
	msoulscat_omplite_cs = 277,
	msoulscat_mputil_cs = 278,
	msoulscat_office_cs = 279,
	msoulscat_ui_cs = 280,
	msoulscat_uls_cpp = 281,
	msoulscat_rankingdictionary_cs = 282,
	msoulscat_fremail_cs = 283,
	msoulscat_tr_cs = 284,
	msoulscat_searchloadbalancing_cs = 285,
	msoulscat_wa_cs = 286,
	msoulscat_test_ipo = 287,
	msoulscat_quiz_cs = 288,
	msoulscat_base_cs = 289,
	msoulscat_hfwssrch_cs = 290,
	msoulscat_hfws_cs = 291,
	msoulscat_downloads_cs = 292,
	msoulscat_r_cpp = 293,
	msoulscat_safesql_cs = 294,
	msoulscat_fileimporter_cs = 295,
	msoulscat_iposmfauthupdate_cs = 296,
} MSOULSdCATEGORYID;

// Developer defineable Event Ids
// NOTE: You probably don't need to add anything here EventTagging will occur
//       automatically like assert tagging


#endif
