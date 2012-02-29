#pragma once
/*****************************************************************************

   Module  : OACR
   Owner   : HannesR

******************************************************************************

   Definitions and #defines for OACR support

*****************************************************************************/

#if defined(__cplusplus)

extern "C++"
{

#endif

#define __noreturn __declspec( noreturn )

#if( OACR )
// helper macro, do not use directly
#define OACRSTR( x ) #x
#endif

// OACR keywords:
// annotations undertstood by PREfast plugins

#ifndef __rpc
// use __rpc for functions that used as remote procedure calls, the keyword will silence
// various OACR checks on formal parameters of the marked functions
#if( OACR )
#define __rpc __declspec("__rpc")
#else
#define __rpc
#endif
#endif // __rpc


#ifndef __fallthrough
// use __fallthrough in switch statement labels where you want the desired behaviour to
// distinguish your code from forgotten break statements.
#if( OACR )
void __FallThrough();
#define __fallthrough __FallThrough();
#else
#define __fallthrough
#endif
#endif // __fallthrough


#ifndef __genericfunctype
// use __genericfunctype for function typedefs used for arrays of functions of different function types.
// if the typedef is marked as __genericfunctype, OACR will not generate DANGEROUS_FUNCTIONCAST (5418) warnings
// e.g. typedef __genericfunctype void (*FUNCPTR)();
#if( OACR )
#define __genericfunctype __declspec("__genericfunctype")
#else
#define __genericfunctype
#endif
#endif // __genericfunctype


#ifndef __checkreturn
// use __checkreturn for function return types, will cause a warning if function is toplevel
// e.g __checkreturn HRESULT HrFoo();
// void Bar()
// {
//    HrFoo();           // cause OACR warning
//    (void)HrFoo();     // does not issue OACR warning
// }
#if( OACR )
#define __checkreturn __declspec("__checkreturn")
#else
#define __checkreturn
#endif
#endif // __checkreturn


#ifndef __memberinitializer
// use __memberinitializer for init functions that initialize all members of a class
// e.g.:
// class X
// {
//    int m_i:
//    int m_j:
//    __memberinitializer void Init(){ m_i = m_j = 0; }
// public:
//    X(){ Init(); }
// };
#if( OACR )
#define __memberinitializer __declspec("__memberinitializer")
#else
#define __memberinitializer
#endif
#endif // __memberinitializer


#ifndef __noheap
// use __noheap classes that should not be instantiated on the heap
// e.g.:
// __noheap class CriticalSection
// {
// public:
//    CriticalSection();
//    ~CriticalSection();
// };
#if( OACR )
#define __noheap __declspec("__noheap")
#else
#define __noheap
#endif
#endif // __noheap


#ifndef __maybe_null
// use __maybe_null for function that return pointers that can be NULL
// often used for memory allocators 
// e.g 	__maybe_null void* MsoPvAlloc(int cb);
// corresponds to the pfxmodel.xml <IsSetToNull value="sometimes"/> tag
#if( OACR )
#define __maybe_null __declspec("SAL_maybenull")
#else
#define __maybe_null
#endif
#endif // __maybe_null


#ifndef __unsafe_string_api
// use __unsafe_string_api to phase out functions that pass unbound writable buffers
// e.g.
// __unsafe_string_api void MyStrCpy( char* szTo, const char* szFrom );
#if( OACR )
#define __unsafe_string_api __declspec("__unsafe_string_api")
#else
#define __unsafe_string_api
#endif
#endif // __unsafe_string_api


#ifndef __min_function
// use __min_function to annotate functions that compute the minimum of two numbers and follow the example's signature
// e.g.
// __min_function int Min( int iLeft, int iRight );
#if( OACR )
#define __min_function __declspec("__min_function")
#else
#define __min_function
#endif
#endif // __min_function


#ifndef __max_function
// use __max_function to annotate functions that compute the maximum of two numbers and follow the example's signature
// e.g.
// __max_function float Max( float iLeft, float iRight );
#if( OACR )
#define __max_function __declspec("__max_function")
#else
#define __max_function
#endif
#endif // __max_function


#ifndef __const
// Use __const where you would otherwise use const, but in situations
//	where you don't want to affect the signature. This is useful in
//	situations where you're using an external API that wasn't correctly
//	const-decorated.
#if( OACR )
#define __const	const
#else
#define	__const
#endif
#endif // __const


// OACR keywords for buffer sizes

#ifndef __recount
// keyword for the number of elements in a buffer than can be read because they have valid contents.
// Never larger than an ecount for the same buffer, obviously.
#if( OACR )
#define __recount(var)     __declspec("__rsize_is("OACRSTR(var)")")
#else
#define __recount(var)
#endif
#endif // __recount

#ifndef __rbcount
// keyword for specifying the readable number of bytes in a buffer.
// Only use for void*, try to use __recount for everything else.
// Never larger than an ecount for the same buffer, obviously.
// e.g CopyBytes( __rbcount( cbTo) const void* pvFrom, __rbcount( cbTo ) void* pvTo, DWORD cbTo );
#if( OACR )
#define __rbcount(var)     __declspec("__rbyte_count("OACRSTR(var)")")
#else
#define __rbcount(var)
#endif
#endif // __rbcount


#ifndef __tcount
// A parameter so marked can assume values between 0 and ONE LESS THAN the size of the buffer (in elements), inclusive.
// (Of course, will usually be the maximal value of buffersize - 1.)
// Like __ecount, but the tcount is the size of the buffer less 1, because the count of readable chars is put in the zeroth element.
// For tcounts, the largest writable index is equal to the tcount limit.  Characters cannot be put in the zeroth element.
#if( OACR )
#define __tcount(var)     __declspec("__tcount("OACRSTR(var)")")
#else
#define __tcount(var)
#endif
#endif // __tcount

#ifndef __rtcount
// The value in the zeroth element of a __tbuffer.  One less than the number of readable elements in the buffer because
// the number of readable elements includes this count.
// For rtcounts, the largest readable index is equal to the tcount limit.  Characters cannot be put in the zeroth element.
#if( OACR )
#define __rtcount(var)     __declspec("__rtcount("OACRSTR(var)")")
#else
#define __rtcount(var)
#endif
#endif // __rtcount


#ifndef __zcount
// A parameter so marked can assume values between 0 and ONE LESS THAN the size of the buffer (in elements), inclusive.
// (Of course, will usually be the maximal value of buffersize - 1.)
// Like __ecount, but it is the size of the buffer less 1, because the last element is reserved for the NULL terminator.
// If the last element contains a valid character, it must be the terminator (of course, if the string is not long
// enough to fill the buffer, the last element won't be used).
// The highest valid index is equal to the zcount, and when used must be the terminator.
#if( OACR )
#define __zcount(var)     __declspec("__zcount("OACRSTR(var)")")
#else
#define __zcount(var)
#endif
#endif // __zcount

#ifndef __rzcount
// Like __zcount, but the number of readable elements in the buffer, rather writable.  There is actually one more
// readable element than that specified, this element is the terminator for the buffer's contents (usually 0).
// .e.g, int MsoCchWzLen(__rzcount(return) const WCHAR* wz)  because it returns the number of chars not counting the terminator.
#if( OACR )
#define __rzcount(var)     __declspec("__rzcount("OACRSTR(var)")")
#else
#define __rzcount(var)
#endif
#endif // __rzcount

#ifndef __tzcount
// A parameter so marked can assume values between 0 and TWO LESS THAN the size of the buffer (in elements), inclusive.
// (Of course, will usually be the maximal value of buffersize - 2.)
// This is a combination of __tcount and __zcount.  The zeroth elements holds the string size, and the last element
// is reserved for the string terminator.  The highest index for the buffer is the tzcount + 1, which, when used,
// must contain the string terminator.

#if( OACR )
#define __tzcount(var)     __declspec("__tzcount("OACRSTR(var)")")
#else
#define __tzcount(var)
#endif
#endif // __tzcount

#ifndef __rtzcount
// The value in the zeroth element of a __tzbuffer.  Two less than the number of readable elements in the buffer because
// the number of readable elements includes this count and because the terminating 0 is not in the count.
// For rtzcounts, the largest readable index is equal to the rtcount limit plus 1 (which is where the terminating 0 goes).

#if( OACR )
#define __rtzcount(var)     __declspec("__rtzcount("OACRSTR(var)")")
#else
#define __rtzcount(var)
#endif
#endif // __rtzcount

#ifndef __endptr
// keyword for specifying the end ptr of a start\end ptr pair
// e.g. FillChars( __endptr( pwchEnd ) WCHAR* pwchStart, const WCHAR* pwchEnd ); 
#if( OACR )
#define __endptr(var)     __declspec("__endptr("OACRSTR(var)")")
#else
#define __endptr(var)
#endif
#endif // __endptr


#ifndef __no_ecount
// use more general __no_count
#if( OACR )
#define __no_ecount       __declspec("__no_ecount")
#else
#define __no_ecount
#endif
#endif // __no_ecount

#ifndef __no_count
// keyword for formal parameters that represent writable buffers but do not need an element count being passed as well
#if( OACR )
// declspec argument will be changed to __no_count when plugin supports it
#define __no_count        __declspec("__no_ecount")
#else
#define __no_count
#endif
#endif // __no_count

// OACR keywords for resource tracking

#ifndef __oblg
#if( OACR )
#define __oblg(token)     __declspec("oblg"OACRSTR(token))
#else
#define __oblg(token) 
#endif
#endif // __oblg

#ifndef __oblgthis
#if( OACR )
#define __oblgthis(token) __declspec("thisoblg"OACRSTR(token))
#else
#define __oblgthis(token)
#endif
#endif // __oblgthis


#ifndef __poblg
#if( OACR )
#define __poblg(token)    __declspec("pointeeoblg"OACRSTR(token))
#else
#define __poblg(token)
#endif
#endif // __poblg


#ifndef __roblg
#if( OACR )
#define __roblg(token)    __declspec("pointeeoblg"OACRSTR(token))
#else
#define __roblg(token)
#endif
#endif // __roblg


#if defined(__cplusplus)

#ifndef __stripOblg
#if( OACR )
template< class C> C __StripOblgFn( C c ){ return c; }
#define __stripOblg(x)            __StripOblgFn(x)
#else
#define __stripOblg(x) x
#endif
#endif // __stripOblg


#ifndef __injectOblg
#if( OACR )
template< class C> C __InjectOblgFn( char *tokenstring, C c ){ return c; }
#define __injectOblg(token, exp)  __InjectOblgFn("oblg$"OACRSTR(token) , exp)
#else
#define __injectOblg(token, exp) exp
#endif
#endif // __injectOblg


#ifndef __castOblg
#if( OACR )
template< class C> C __CastOblgFn( char *tokenstring, C c ){ return c; }
#define __castOblg(token, exp)    __CastOblgFn("oblg$"OACRSTR(token) , exp)
#else
#define __castOblg(token, exp) exp
#endif
#endif // __castOblg


#ifndef __shiftOblg
#if( OACR )
template< class C> C __ShiftOblgFn(void *c1, C c2){ c1; return c2; }
#define __shiftOblg(x, y)         __ShiftOblgFn((void *) x, y)
#else
#define __shiftOblg(x, y) y
#endif
#endif // __shiftOblg


#ifndef __ignoreOblg
#if( OACR )
#define __ignoreOblg __declspec("__ignoreOblg")
#else
#define __ignoreOblg
#endif
#endif // __ignoreOblg


#endif // __cplusplus

// OACR helper macros to suppress particular warnings

#ifndef OACR_USE_PTR
// use to suprress constness and related warnings (5403, 5404, 5405 5432, 5433)
#if( OACR )
void OACRUsePtr( void* p );
#define OACR_USE_PTR( p ) OACRUsePtr( p )
#elif defined(DEBUG) && !defined(__cplusplus)
// for C code #define OACR_USE_PTR as a noop statement causing a compile error if used before declarations
#define OACR_USE_PTR( p ) do{}while(0)
#else
#define OACR_USE_PTR( p )
#endif
#endif // OACR_USE_PTR


#ifndef OACR_OWN_PTR
// can be used for objects that attach themselves to an owner
// in their constructors
#if( OACR )
void OACROwnPtr( const void* p );
#define OACR_OWN_PTR( p ) OACROwnPtr( p )
#elif defined(DEBUG) && !defined(__cplusplus)
// for C code #define OACR_OWN_PTR as a noop statement causing a compile error if used before declarations
#define OACR_OWN_PTR( p ) do{}while(0)
#else
#define OACR_OWN_PTR( p )
#endif
#endif // OACR_OWN_PTR


#ifndef OACR_PTR_NOT_NULL
// tells OACR that a pointer is not null at this point
#if( OACR )
void OACRPtrNotNull( void** pp );
#define OACR_PTR_NOT_NULL( p ) OACRPtrNotNull( (void**)&p)
#elif defined(DEBUG) && !defined(__cplusplus)
// for C code #define OACR_PTR_NOT_NULL as a noop statement causing a compile error if used before declarations
#define OACR_PTR_NOT_NULL( p ) do{}while(0)
#else
#define OACR_PTR_NOT_NULL( p )
#endif
#endif // OACR_PTR_NOT_NULL


#ifndef OACR_ASSERT
#if( OACR && defined (DEBUG) )
#if defined(__cplusplus)
bool __FAssert( bool fCondition );
#else
int __FAssert( int fCondition );
#endif
#define OACR_ASSERT( fCondition ) __FAssert( fCondition )
#else
#define OACR_ASSERT( fCondition )
#endif
#endif


#ifndef OACR_ASSERTDO
#if( OACR )
#if defined(__cplusplus)
bool __FAssertDo( bool fCondition );
#else
int __FAssertDo( int fCondition );
#endif
#define OACR_ASSERTDO( fCondition ) __FAssertDo( fCondition )
#else
#define OACR_ASSERTDO( fCondition ) (fCondition)
#endif
#endif


#ifndef OACR_NOT_IMPLEMENTED_MEMBER
#if( OACR )
#define OACR_NOT_IMPLEMENTED_MEMBER OACR_USE_PTR( (void*)this )
#else
#define OACR_NOT_IMPLEMENTED_MEMBER
#endif
#endif // OACR_NOT_IMPLEMENTED_MEMBER


#ifndef OACR_DECLARE_FILLER
#if( OACR )
#define OACR_DECLARE_FILLER( type, inst ) type __filler##inst;
#else
#define OACR_DECLARE_FILLER( type, inst )
#endif
#endif // OACR_DECLARE_FILLER


// use this macro once you have inspected warnings WARNING_FUNCTION_NEEDS_REVIEW (5428)
#ifndef OACR_REVIEWED_CALL
#if( OACR )
void __OACRReviewedCall();
#define OACR_REVIEWED_CALL( reviewer, functionCall ) ( __OACRReviewedCall(), functionCall )
#else
#define OACR_REVIEWED_CALL( reviewer, functionCall ) functionCall
#endif
#endif // OACR_REVIEWED_CALL


// use this macro once you have inspected warnings WARNING_URL_NEEDS_TO_BE_REVIEWED (5485)
#ifndef OACR_REVIEWED_URL
#if( OACR )
void __OACRReviewedUrl();
#define OACR_REVIEWED_URL( reviewer, reviewedUrl ) ( __OACRReviewedUrl(), reviewedUrl )
#else
#define OACR_REVIEWED_URL( reviewer, reviewedUrl ) reviewedUrl
#endif
#endif // OACR_REVIEWED_URL

#if( OACR && defined(_WINDEF_) && 0 )

// redefine FALSE & TRUE for better HRESULT<->BOOL conversion detection
#ifdef FALSE
#undef FALSE
#define FALSE ((BOOL)0)
#endif

#ifdef TRUE
#undef TRUE
#define TRUE ((BOOL)1)
#endif

#endif

// macro to tell OACR to not issue a specific warning for the following line of code
// use to suppress false positives from OACR
// e.g.
// if( fPointerNotNull )
//    OACR_SUPRESSWARNING( 11, "pointer access is guarded by 'fPointerNotNull'" )
//    p->Foo();
// this will work with C7
#if( OACR )
#ifndef OACR_SUPPRESSWARNING
#define OACR_SUPPRESSWARNING( cWarning, comment ) __pragma ( prefast(suppress: cWarning, comment) )
#endif
#else
#ifndef OACR_SUPPRESSWARNING
#define OACR_SUPPRESSWARNING( cWarning, comment )
#endif
#endif
 
#if defined(__cplusplus)
} // extern "c++" {
#endif
