//
// VSPWeb labels and macros
//
// 11/30/99
//

// Ignore unreferenced labels.
#pragma warning (disable : 4102)

//
// Macros for defining arrays of pointers to protected data.
//
// NOTE:
//	Vulcan will not allow enumeration of relocations with a data
//  unless it has been accessed. The macros TOUCH_SCP_DATA()
//  and TOUCH_DATA can be used to access the data without
//  altering it.
//

#define SCP_DATA(SegNum1, SegNum2, CryptMethod, Significance) \
void const *ScpProtectedData_##SegNum1##_##SegNum2##_##CryptMethod##_##Significance##_00_00

#define SCP_DATA_ARRAY(SegNum1, SegNum2, CryptMethod, Significance) \
ScpProtectedData_##SegNum1##_##SegNum2##_##CryptMethod##_##Significance##_00_00

#define	SCP_PROTECTED_DATA_LIST	"ScpProtectedData"

#define TOUCH_SCP_DATA(SegNum1, SegNum2, CryptMethod, Significance) { \
_asm	push	EAX \
_asm	lea		EAX, ScpProtectedData_##SegNum1##_##SegNum2##_##CryptMethod##_##Significance##_00_00 \
_asm	pop		EAX }

#define	TOUCH_DATA( x ) {	\
_asm	push	EAX			\
_asm	lea		EAX,x		\
_asm	pop		EAX	}

//
// Macros for defining verification call placement.
//

#define SCP_VERIFY_CALL_PLACEMENT(MarkerId) \
SCP_VERIFY_CALL_PLACEMENT_##MarkerId##:

//
// VSPWeb macros for verified segments outside functions
//

#define BEGIN_VSPWEB_SCP_SEGMENT(SegNum1, SegNum2, CryptMethod, Significance, Proximity, Redundance) \
_declspec(naked) void Begin_Vspweb_Scp_Segment_##SegNum1##_##SegNum2() \
{ \
__asm { \
	__asm nop \
	} \
BEGIN_SCP_SEGMENT_##SegNum1##_##SegNum2##_##CryptMethod##_##Significance##_##Proximity##_##Redundance: \
__asm { \
	__asm nop \
	} \
} 

#define END_VSPWEB_SCP_SEGMENT(SegNum1, SegNum2) \
_declspec(naked) void End_Vspweb_Scp_Segment_##SegNum1##_##SegNum2() \
{ \
__asm { \
	__asm nop \
	} \
END_SCP_SEGMENT_##SegNum1##_##SegNum2: \
__asm { \
	__asm nop \
	} \
} 

//
// some hacks to make sure some VSP library functions are included in the
// application to be protected (to avoid these functions from being dynamically
// linked)
//

// Put this in some source file.
#define VSPWEB_DECL_SETUP \
extern "C" { \
extern void __cdecl DoComparison01(void *, void *, int); \
extern void __cdecl DoVerification01(void *, void *, void *, void *, BYTE, BYTE, BYTE, int);  \
extern void __cdecl DoVerification02(BYTE, void *, void *, void *, void *, BYTE, BYTE, int); \
extern void __cdecl DoVerification03(void *, void *, void *, BYTE, BYTE, BYTE, int, void *); \
extern void __cdecl DoVerification04(void *, void *, void *, BYTE, BYTE, BYTE, int, void *); \
extern void __cdecl DoVerification05(void *, void *, void *, void *, BYTE, BYTE, BYTE, int); \
}

// Put this in one live function.
#define VSPWEB_SETUP \
{ \
	__asm cmp eax, DoComparison01 \
	__asm cmp ebx, DoVerification01 \
	__asm cmp ebx, DoVerification02 \
	__asm cmp ebx, DoVerification03 \
	__asm cmp ebx, DoVerification04 \
	__asm cmp ebx, DoVerification05 \
}


/*-----------------------------------------------------------------------------
	And now the macros actually used in Office
------------------------------------------------------------------- VadimC ----*/
//
// Plain labels in the code are too succeptible to typos.
// Besides, we may need to add some nops around the labels.
//

// VSP (code verification)
#define BEGIN_VSP_SEGMENT(SegNum1, SegNum2, CryptMethod, Significance, Proximity, Redundance) \
{ \
	__asm { __asm nop } \
	BEGIN_SCP_SEGMENT_##SegNum1##_##SegNum2##_##CryptMethod##_##Significance##_##Proximity##_##Redundance: \
	__asm { __asm nop } \
}

#define END_VSP_SEGMENT(SegNum1, SegNum2) \
{ \
	__asm { __asm nop } \
	END_SCP_SEGMENT_##SegNum1##_##SegNum2: \
	__asm { __asm nop } \
}

//
// Same, but with hardcoded segment parameters (because they will be overriden in the config file anyways)
//
#ifndef _WIN64
#define BEGIN_VERIFIED(SegNum1, SegNum2) BEGIN_VSP_SEGMENT(SegNum1,SegNum2,0,10,0,0)
#define END_VERIFIED(SegNum1, SegNum2) END_VSP_SEGMENT(SegNum1,SegNum2)

#define BEGIN_VERIFIED_SECTION(SegNum1, SegNum2) \
	__declspec(naked) void _Begin_Verified_Section_##SegNum1##_##SegNum2 () \
	{	__asm { xor eax, SegNum1 }\
		BEGIN_VERIFIED(SegNum1,SegNum2);	 \
		__asm { xor edx, SegNum2 }\
		__asm { ret }\
	}
	
#define END_VERIFIED_SECTION(SegNum1, SegNum2) \
	__declspec(naked) void _End_Verified_Section_##SegNum1##_##SegNum2 () \
	{	__asm { add eax, SegNum1 }\
		END_VERIFIED(SegNum1,SegNum2);	 \
		__asm { sub edx, SegNum2 }\
		__asm { ret }\
	}
	
#define REFERENCE_VERIFIED_SECTION(SegNum1,SegNum2) { \
	extern void _Begin_Verified_Section_##SegNum1##_##SegNum2 (); \
	extern void _End_Verified_Section_##SegNum1##_##SegNum2 (); \
		__asm { lea eax, _Begin_Verified_Section_##SegNum1##_##SegNum2 }\
		__asm { lea eax, _End_Verified_Section_##SegNum1##_##SegNum2 }\
	}
	
#define CHECK_SEGMENT_HERE(id1,id2) CHECK_SEGMENT_HERE_##id1##_##id2: __asm { __asm nop }

// SCP (code obfuscation)
#define BEGIN_OBFUSCATED() \
	BEGIN_SCP_SECTION:

#define OBFUSCATED(BlockNum) \
{ \
	__asm { __asm nop } \
	SCP_BLOCK_##BlockNum: \
	__asm { __asm nop } \
}
	
#define END_OBFUSCATED() \
	END_SCP_SECTION:
	
#else//_WIN64
#define BEGIN_VERIFIED(SegNum1, SegNum2)
#define END_VERIFIED(SegNum1, SegNum2)
#define BEGIN_VERIFIED_SECTION(SegNum1, SegNum2)
#define END_VERIFIED_SECTION(SegNum1, SegNum2)
#define REFERENCE_VERIFIED_SECTION(SegNum1,SegNum2)
#define CHECK_SEGMENT_HERE(id1,id2)
#define BEGIN_OBFUSCATED()
#define OBFUSCATED(BlockNum)
#define END_OBFUSCATED()
#endif//_WIN64

//
// Data protection
//

// Declare protected data
#define DECLARE_VERIFIED_DATA(SegNum1, SegNum2) \
	SCP_DATA(SegNum1, SegNum2, 0, 10)

// Reference protected data
#define REFERENCE_VERIFIED_DATA(SegNum1, SegNum2) \
	extern SCP_DATA(SegNum1, SegNum2, 0, 10)[]; \
	TOUCH_SCP_DATA(SegNum1, SegNum2, 0, 10)

