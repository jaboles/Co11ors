/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Wed Sep 13 22:33:23 2000
 */
/* Compiler settings for msuimw32.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_IAImeProfile = {0xB2AA53DF,0x21AB,0x40f2,{0xB3,0x86,0xED,0x04,0x8C,0xFC,0x1C,0x9D}};


const IID IID_IAImeContext = {0x5F5B4ACB,0xD55D,0x492c,{0xB5,0x96,0xF6,0x39,0x0E,0x1A,0xD7,0x98}};


#ifdef __cplusplus
}
#endif

