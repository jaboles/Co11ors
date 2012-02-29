/*
	Prove that we build only nice neat WIN32 stuff
*/

// Mso's batch.lst sets
//  -DSTRICT -DWIN32 -DX86 -DDEBUG -DDM_UNICODE -DOFFICE_BUILD -DSTATICASST  


#ifdef WIN16
#error
#endif

#ifdef MAC
#error
#endif

#ifdef MAC_WLM
#error
#endif

// Why do some mso files barf on this, while others don't?
//  drsp, wgnewdlg, wgautcor, wgpstdoc, wgprpdlg
// They're all compiled at the same time, via batch.*
// Do only these #include odef.h?
//#ifndef WIN
//#error
//#endif

#ifndef WIN32
#error
#endif

#ifdef KANAACCEL
#error
#endif

// batch.lst doesn't set -DCC
//#ifndef DM96	// dm.mak doesn't turn on -DCC for some reason
//#ifndef CC
//#error
//#endif
//#endif

// This is defined in mso\sdm_.h
//#ifdef	CSTD_WITH_WINDOWS
//#error
//#endif

#ifdef	CSTD_WITH_OS2
#error
#endif

#ifdef	CSTD_WITH_CWINDOWS
#error
#endif
