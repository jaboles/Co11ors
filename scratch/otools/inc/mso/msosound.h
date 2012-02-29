#pragma once

/****************************************************************************
	msosound.h

	Owner: JBelt
 	Copyright (c) 1995 Microsoft Corporation

	Office sound routines.

	ALL APPS must call us to play their sounds. This is the only way to
	ensure that Office will mix the sounds properly. (Exceptions, you know
	who you are - if you're not an exception and think you want to be one,
	talk to me first).
****************************************************************************/
#ifndef MSOSOUND_H
#define MSOSOUND_H


// Sound id's are in there
#ifndef STATIC_LIB_DEF
#include "msosndxx.h"
#endif

// compatibility
#define msosndOpen			msosndProcessComplete
#define msosndSave			msosndProcessComplete
#define msosndPrint			msosndProcessComplete
#define msosndBoot			msosndProcessComplete
#define msosndInsertObject	msosndProcessComplete
#define msosndAutoFill		msosndProcessComplete

/* Plays an Office-registered sound, snd is one of msosndxxx contants above
	These sounds are customizable by the user in Control Panel\Sounds. Office handles
	all the registry keys and sound filenames.
	If this returns an error value (FAILED(hr)), then no sound was produced (sound
	is off, no soundcard on the computer, the audio device is busy, etc). You may
	want to call MessageBeep or something like that directly in this case if you
	really want audio feedback to be provided to the user. Note that Office does this
	automatically for msosndAlert only, and will return S_OK. */
MSOAPI_(HRESULT) MsoPlaySound(int snd);

//	Plays the sound sample stored in the file. On the Mac, plays the first
//	sound resource in the file. Returns FALSE if unable to find the file.
//	if fIgnoreOffSnd is TRUE, we will play the sound regardless of the state
//	of office Sounds.
MSOAPIXX_(BOOL) MsoFPlaySoundFile(const WCHAR *wzFile, BOOL fIgnoreOffSnd);

/* Sets Office sound on or off. */
MSOAPI_(void) MsoSetSoundState(BOOL fOn);

/*	Sets sound on or off. When turning sound on, Office will try to detect a
	a sound device, and on failure, warns the user. The user then has the option
	to turn sounds on anyway or keep them off. Recommended for use with the
	"Provide feedback with sound" option, and make sure to call MsoFGetSoundState
	right afterwards to see what the final sound state is. */
MSOAPI_(void) MsoSetSoundStateAndPrompt(BOOL fOn);

/* Returns TRUE iff sound is on. */
MSOAPI_(BOOL) MsoFGetSoundState();

#if 0 //$[VSMSO]
/* Register sounds */
MSOAPIX_(HRESULT) MsoHrRegisterSounds();

/* UnRegister sounds */
MSOAPIX_(HRESULT) MsoHrUnRegisterSounds();
#endif //$[VSMSO]

#endif // MSOSOUND_H
