/*----------------------------------------------------------------------------
	%%File: FEDICT.H
	%%Unit: COM
	%%Contact: seijia
	%%Date: Jan. 24, 1996

	defines IME dictionary utility interfaces in COM format
----------------------------------------------------------------------------*/

#ifndef __FEDICT_H__
#define __FEDICT_H__

#include <objbase.h>
#include "outpos.h"
#include "actdict.h"

// Part Of Speach
#define IFED_POS_NONE					0x00000000
#define IFED_POS_NOUN					0x00000001
#define IFED_POS_VERB					0x00000002
#define IFED_POS_ADJECTIVE				0x00000004
#define IFED_POS_ADJECTIVE_VERB			0x00000008
#define IFED_POS_ADVERB					0x00000010
#define IFED_POS_ADNOUN					0x00000020
#define IFED_POS_CONJUNCTION			0x00000040
#define IFED_POS_INTERJECTION			0x00000080
#define IFED_POS_INDEPENDENT			0x000000ff
#define IFED_POS_INFLECTIONALSUFFIX		0x00000100
#define IFED_POS_PREFIX					0x00000200
#define IFED_POS_SUFFIX					0x00000400
#define IFED_POS_AFFIX					0x00000600
#define IFED_POS_TANKANJI				0x00000800
#define IFED_POS_IDIOMS					0x00001000
#define IFED_POS_SYMBOLS				0x00002000
#define IFED_POS_PARTICLE				0x00004000
#define IFED_POS_AUXILIARY_VERB			0x00008000
#define IFED_POS_SUB_VERB				0x00010000
#define IFED_POS_DEPENDENT				0x0001c000
#define IFED_POS_ALL					0x0001ffff

// GetWord Selection Type
#define IFED_SELECT_NONE				0x00000000
#define IFED_SELECT_READING				0x00000001
#define IFED_SELECT_DISPLAY				0x00000002
#define IFED_SELECT_POS					0x00000004
#define IFED_SELECT_COMMENT				0x00000008
#define IFED_SELECT_ALL					0x0000000f

// Registered Word/Dependency Type
#define IFED_REG_NONE					0x00000000
#define IFED_REG_USER					0x00000001
#define IFED_REG_AUTO					0x00000002
#define IFED_REG_GRAMMAR				0x00000004
#define IFED_REG_ALL					0x00000007

// Dictionary Type
#define IFED_TYPE_NONE					0x00000000
#define IFED_TYPE_GENERAL				0x00000001
#define	IFED_TYPE_NAMEPLACE				0x00000002
#define IFED_TYPE_SPEECH				0x00000004
#define IFED_TYPE_REVERSE				0x00000008
#define	IFED_TYPE_ENGLISH				0x00000010
#define IFED_TYPE_ALL					0x0000001f

// HRESULTS for IFEDictionary interface

//no more entries in the dictionary
#define IFED_S_MORE_ENTRIES				MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x7200)
//dictionary is empty, no header information is returned
#define IFED_S_EMPTY_DICTIONARY			MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x7201)
//word already exists in dictionary
#define IFED_S_WORD_EXISTS				MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x7202)

//dictionary is not found
#define IFED_E_NOT_FOUND				MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7300)
//invalid dictionary format
#define IFED_E_INVALID_FORMAT			MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7301)
//failed to open file
#define IFED_E_OPEN_FAILED				MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7302)
//failed to write to file
#define IFED_E_WRITE_FAILED				MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7303)
//no entry found in dictionary
#define IFED_E_NO_ENTRY					MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7304)
//this routines doesn't support the current dictionary
#define IFED_E_REGISTER_FAILED			MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7305)
//not a user dictionary
#define IFED_E_NOT_USER_DIC				MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7306)
//not supported
#define IFED_E_NOT_SUPPORTED			MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7307)
//failed to insert user comment			
#define IFED_E_USER_COMMENT				MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x7308)

#define cbCommentMax			256

//Private Unicode Character
#define wchPrivate1				0xE000

// Where to place registring word
typedef enum 
{
	IFED_REG_HEAD,
	IFED_REG_TAIL,
	IFED_REG_DEL,
} IMEREG;

// Relation type
typedef enum
{
	IFED_REL_NONE,
	IFED_REL_NO,
	IFED_REL_GA,
	IFED_REL_WO,
	IFED_REL_NI,
	IFED_REL_DE,
	IFED_REL_YORI,
	IFED_REL_KARA,
	IFED_REL_MADE,
	IFED_REL_HE,
	IFED_REL_TO,
	IFED_REL_IDEOM,
	IFED_REL_FUKU_YOUGEN,		//p2_1
	IFED_REL_KEIYOU_YOUGEN,		//p2_2
	IFED_REL_KEIDOU1_YOUGEN,	//p2_3
	IFED_REL_KEIDOU2_YOUGEN,	//p2_4
	IFED_REL_TAIGEN,			//p2_5
	IFED_REL_YOUGEN,			//p2_6
	IFED_REL_RENTAI_MEI,		//p3_1
	IFED_REL_RENSOU,			//p3_2
	IFED_REL_KEIYOU_TO_YOUGEN,	//p3_3
	IFED_REL_KEIYOU_TARU_YOUGEN,//p3_4
	IFED_REL_UNKNOWN1,			//p3_5
	IFED_REL_UNKNOWN2,			//p3_6
	IFED_REL_ALL,				//any type
} IMEREL;

// Type of IME dictionary
typedef enum
{
	IFED_UNKNOWN,
	IFED_MSIME2_BIN_SYSTEM,
	IFED_MSIME2_BIN_USER,
	IFED_MSIME2_TEXT_USER,
	IFED_MSIME95_BIN_SYSTEM,
	IFED_MSIME95_BIN_USER,
	IFED_MSIME95_TEXT_USER,
	IFED_MSIME97_BIN_SYSTEM,
	IFED_MSIME97_BIN_USER,
	IFED_MSIME97_TEXT_USER,
	IFED_MSIME_BIN_SYSTEM,
	IFED_MSIME_BIN_USER,
	IFED_MSIME_TEXT_USER,
	IFED_MSIME_SYSTEM_CE,
	IFED_ACTIVE_DICT,
	IFED_ATOK9,
	IFED_ATOK10,
	IFED_NEC_AI_,
	IFED_WX_II,
	IFED_WX_III,
	IFED_VJE_20,
} IMEFMT;

// Type of User Comment
typedef enum
{
	IFED_UCT_NONE,
	IFED_UCT_STRING_SJIS,
	IFED_UCT_STRING_UNICODE,
	IFED_UCT_USER_DEFINED,
	IFED_UCT_MAX,
} IMEUCT;

// WoRD found in a dictionary
typedef struct _IMEWRD
{
	WCHAR		*pwchReading;
	WCHAR		*pwchDisplay;
	union {
		ULONG ulPos;	
		struct {
			WORD		nPos1;		//hinshi
			WORD		nPos2;		//extended hinshi
		};
	};
	ULONG		rgulAttrs[2];	 	//attributes
	INT			cbComment;			//size of user comment
	IMEUCT		uct;				//type of user comment
	VOID		*pvComment;			//user comment
} IMEWRD, *PIMEWRD;

// DePendency Info
typedef struct _IMEDP
{
	IMEWRD		wrdModifier;	//kakari-go
	IMEWRD		wrdModifiee;	//uke-go
	IMEREL		relID;
} IMEDP, *PIMEDP;

typedef BOOL (*PFNLOG)(IMEDP *, HRESULT);

//////////////////////////////
// The IFEDictionary interface
//////////////////////////////

#undef  INTERFACE
#define INTERFACE   IFEDictionary

DECLARE_INTERFACE_(IFEDictionary, IUnknown)
{
	// IUnknown members
    STDMETHOD(QueryInterface)(THIS_ REFIID refiid, VOID **ppv) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;

	// IFEDictionary members
    STDMETHOD(Open)			(THIS_
							_In_z_ CHAR *pchDictPath,	 		//(in) dictionary path
							IMESHF *pshf				//(out) dictionary header
							) PURE;
    STDMETHOD(Close)		(THIS) PURE;
    STDMETHOD(GetHeader)	(THIS_
							_In_z_ CHAR *pchDictPath, 			//(in) dictionary path
							IMESHF *pshf,				//(out) dictionary header
							IMEFMT *pjfmt,				//(out) dictionary format
							ULONG *pulType				//(out) dictionary type
							) PURE;
    STDMETHOD(DisplayProperty) (THIS_
							HWND hwnd 					//(in) parent window handle
							) PURE;
    STDMETHOD(GetPosTable) 	(THIS_
    						POSTBL **prgPosTbl,			//(out) pos table pointer
    						int *pcPosTbl				//(out) pos table count pointer
    						) PURE;
    STDMETHOD(GetWords)		(THIS_
							_In_z_ WCHAR *pwchFirst, 			//(in) starting range of reading
							_In_z_ WCHAR *pwchLast, 			//(in) ending range of reading
							_In_z_ WCHAR *pwchDisplay,			//(in) display
							ULONG ulPos,				//(in) part of speech (IFED_POS_...)
							ULONG ulSelect,				//(in) output selection
							ULONG ulWordSrc,			//(in) user or auto registered word?
							UCHAR *pchBuffer,			//(in/out) buffer for storing array of IMEWRD
							ULONG cbBuffer,				//(in) size of buffer in bytes
							ULONG *pcWrd				//(out) count of IMEWRD's returned
							) PURE;
    STDMETHOD(NextWords)	(THIS_
							UCHAR *pchBuffer,			//(in/out) buffer for storing array of IMEWRD
							ULONG cbBuffer,				//(in) size of buffer in bytes
							ULONG *pcWrd				//(out) count of IMEWRD's returned
							) PURE;
    STDMETHOD(Create)		(THIS_
							_In_z_ CHAR *pchDictPath,			//(in) path for the new dictionary
							IMESHF *pshf				//(in) dictionary header
							) PURE;
    STDMETHOD(SetHeader)	(THIS_
							IMESHF *pshf				//(in) dictionary header
							) PURE;
    STDMETHOD(ExistWord)	(THIS_
							IMEWRD *pwrd				//(in) word to check
							) PURE;
    STDMETHOD(ExistDependency) (THIS_
							IMEDP *pdp					//(in) see if dependency pair exist in dictionary
							) PURE;
    STDMETHOD(RegisterWord) (THIS_
							IMEREG reg,					//(in) type of operation to perform on IMEWRD
							IMEWRD *pwrd				//(in) word to be registered or deleted
							) PURE;
    STDMETHOD(RegisterDependency) (THIS_
							IMEREG reg,					//(in) type of operation to perform on IMEDP
							IMEDP *pdp					//(in) dependency to be registered or deleted
							) PURE;
    STDMETHOD(GetDependencies) (THIS_
							_In_z_ WCHAR *pwchKakariReading, 	//(in) reading for kakarigo
							_In_z_ WCHAR *pwchKakariDisplay,	//(in) display for kakarigo
							ULONG ulKakariPos,			//(in) part of speech for kakarigo
							_In_z_ WCHAR *pwchUkeReading, 		//(in) reading for ukego
							_In_z_ WCHAR *pwchUkeDisplay, 		//(in) display for ukego
							ULONG ulUkePos,				//(in) part of speech for ukego
							IMEREL  jrel,				//(in) relation type
							ULONG ulWordSrc,			//(in) registration type
							UCHAR *pchBuffer,			//(in/out) buffer for storing array of IMEDP
							ULONG cbBuffer,				//(in) size of buffer in bytes
							ULONG *pcdp					//(out) count of IMEDP's returned
							) PURE;
    STDMETHOD(NextDependencies) (THIS_
							UCHAR *pchBuffer,			//(in/out) buffer for storing array of IMEDP
							ULONG cbBuffer,				//(in) size of buffer in bytes
							ULONG *pcDp					//(out) count of IMEDP's retrieved
							) PURE;
#ifndef PEGASUSP
    STDMETHOD(ConvertFromOldMSIME) (THIS_
							_In_z_ CHAR *pchDic,				//(in) old ime user dictionary path
							PFNLOG pfnLog,				//(in) pointer to log function
							IMEREG reg					//(in) word registration info
							) PURE;
    STDMETHOD(ConvertFromUserToSys) (THIS) PURE;
#endif // PEGASUSP
};

#ifdef __cplusplus
extern "C" {
#endif

// The following API replaces CoCreateInstance(), when CLSID is not used.
HRESULT WINAPI CreateIFEDictionaryInstance(VOID **ppvObj);
typedef HRESULT (WINAPI *fpCreateIFEDictionaryInstanceType)(VOID **ppvObj);

#ifdef __cplusplus
} /* end of 'extern "C" {' */
#endif

#endif //__FEDICT_H__
