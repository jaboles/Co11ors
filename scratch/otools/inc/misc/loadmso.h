#include "delayimp.h"

#if 0 //$[VSMSO]
#if defined(__cplusplus)
extern "C"
{
#endif

#ifdef MSOTC_H
extern MSOTCFCF vfcf;
#define HinstLoadMso(cid) HinstLoadMsoCid(cid,&vfcf)
HINSTANCE __stdcall HinstLoadMsoCid(msocidT, MSOTCFCF*);
#endif

// flags for HinstLoadMsoWzEx
#define LOADMSO_ADDREF_ALWAYS		((ULONG) 0x0001)

HINSTANCE __stdcall HinstLoadMsoWz(const WCHAR *wzCoreComponentGUID);
HINSTANCE __stdcall HinstLoadMsoWzEx(const WCHAR *wzCoreComponentGUID, BOOL fUseDarwin);
HINSTANCE __stdcall HinstLoadMsoWzEx2(const WCHAR *wzCoreComponentGUID, BOOL fUseDarwin, ULONG ulFlags);
HINSTANCE __stdcall MsoLoadRichEdit();
HINSTANCE __stdcall MsoLoadUniscribe();
HINSTANCE __stdcall HinstLoadMsoAbsPath(HINSTANCE hinstModule);

void __stdcall DisplayAppNotConfiguredMsg();
FARPROC WINAPI DelayLoadHook (unsigned dliNotify, PDelayLoadInfo pdli);
FARPROC WINAPI MsoDelayLoadHook (unsigned dliNotify, PDelayLoadInfo pdli);

#if defined(__cplusplus)
}
#endif
#endif //$[VSMSO]
