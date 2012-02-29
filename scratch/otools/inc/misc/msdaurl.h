/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0158 */
/* at Mon Jul 13 12:59:58 1998
 */
/* Compiler settings for msdaurl.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data , no_format_optimization
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __msdaurl_h__
#define __msdaurl_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __RootBinder_FWD_DEFINED__
#define __RootBinder_FWD_DEFINED__

#ifdef __cplusplus
typedef class RootBinder RootBinder;
#else
typedef struct RootBinder RootBinder;
#endif /* __cplusplus */

#endif 	/* __RootBinder_FWD_DEFINED__ */


/* header files for imported files */
#include "oledb.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t size);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_msdaurl_0000 */
/* [local] */ 

#ifndef __msdasc_h__
typedef 
enum tagEBindInfoOptions
    {	BIO_BINDER	= 0x1
    }	EBindInfoOptions;
#endif // __msdasc_h__

#define STGM_CREATE_COLLECTION			0x00002000L
#define STGM_CREATECOLLECTION			STGM_CREATE_COLLECTION
#define STGM_COLLECTION					0x00002000L
#define STGM_SOURCE	0x00008000L
#define STGM_CREATE_STRUCTURED_DOCUMENT	STGM_SOURCE
#define STGM_CREATE_DIRECTORY			STGM_CREATE_COLLECTION
#define STGM_CREATE_INTERMEDIATE			0x00004000L
#define STGM_OPENALWAYS					0x00800000L
#define STGM_OPEN						0x80000000L
#define STGM_OVERWRITE					0x00400000L
#define STGM_SHARE_RECURSIVE				0x01000000L
#define STGM_RECURSIVE				    0x01000000L
#define STGM_RESERVE						0x10000000L
#define STGM_TEMP_CHECKOUT				STGM_RESERVE
#define STGM_KEEP_CHECKEDOUT				0x20000000L
#define STGM_ENCODEURL					0x02000000L
#define STGM_STRICTOPEN					0x40000000L


extern RPC_IF_HANDLE __MIDL_itf_msdaurl_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_msdaurl_0000_v0_0_s_ifspec;


#ifndef __MSDAURLLib_LIBRARY_DEFINED__
#define __MSDAURLLib_LIBRARY_DEFINED__

/* library MSDAURLLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_MSDAURLLib;

EXTERN_C const CLSID CLSID_RootBinder;

#ifdef __cplusplus

class DECLSPEC_UUID("FF151822-B0BF-11D1-A80D-000000000000")
RootBinder;
#endif
#endif /* __MSDAURLLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
