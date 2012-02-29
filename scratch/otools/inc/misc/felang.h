/*----------------------------------------------------------------------------
	%%File: FELang.H
	%%Unit: COM
	%%Contact: HiroakiK   v-takasu
	%%Date: 1/22, 1997    4/23

	defines IFELanguage interfaces in COM format
----------------------------------------------------------------------------*/

#ifndef __FELang_H__
#define __FELang_H__

#include <ole2.h>

#ifdef __cplusplus
extern "C" {
#endif


///////////////////////////////////
// The IFELanguage constants
///////////////////////////////////

#pragma pack(1)

// Word Descriptor
typedef struct tagWDD{
	WORD		wDispPos;	// Offset of Output string
	union {
		WORD	wReadPos;	// Offset of Reading string
		WORD	wCompPos;
	};
	
	WORD  		cchDisp;  	//number of ptchDisp
	union {
		WORD	cchRead;  	//number of ptchRead
		WORD	cchComp;
	};
	
	DWORD		nWordCost;	//a word cost based on Unigram
	
	WORD	 	nPos;    	//part of speech
	
							// implementation-defined
	WORD  		fPhrase : 1;//start of phrase
	WORD  		fAutoCorrect: 1;//auto-corrected
	WORD  		fNumericPrefix: 1;//kansu-shi expansion(JPN)
	WORD  		fUserRegistered: 1;//from user dictionary
	WORD  		fUnknown: 1;//unknown word (duplicated information with nPos.)
	WORD		fRecentUsed: 1;	//used recently flag
	WORD		:10;		//
	
	VOID    	*pReserved;	//points directly to WORDITEM
} WDD;

#pragma warning(push)
#pragma warning(disable:4200) // zero-size array in structure

typedef struct tagMORRSLT {
	DWORD		dwSize;				// total size of this block.
	WCHAR		*pwchOutput;		// conversion result string.
	WORD		cchOutput;			// lengh of result string.
	union {
		WCHAR	*pwchRead;			// reading string
		WCHAR	*pwchComp;
	};
	union {
		WORD	cchRead;			// length of reading string.
		WORD	cchComp;
	};
	WORD		*pchInputPos;		// index array of reading to input character.
	WORD		*pchOutputIdxWDD;	// index array of output character to WDD
	union {
		WORD	*pchReadIdxWDD;		// index array of reading character to WDD
		WORD	*pchCompIdxWDD;
	};	
	WORD		*paMonoRubyPos;		// array of position of monoruby
	WDD			*pWDD;				// pointer to array of WDD
	INT			cWDD;				// number of WDD
	VOID		*pPrivate;			// pointer of private data area
	WCHAR		BLKBuff[];			// area for stored above members.
									//	WCHAR	wchOutput[cchOutput];
									//	WCHAR	wchRead[cchRead];
									//	CHAR	chInputIdx[cwchInput];
									//	CHAR	chOutputIdx[cchOutput];
									//	CHAR	chReadIndx[cchRead];
									//  ????    Private
									//	WDD		WDDBlk[cWDD];
}MORRSLT;

#pragma warning(pop)

#pragma pack()


// request for conversion (dwRequest)
#define	FELANG_REQ_CONV			0x00010000
#define	FELANG_REQ_RECONV		0x00020000
#define	FELANG_REQ_REV			0x00030000


// Conversion mode (dwCMode)

#define	FELANG_CMODE_MONORUBY		0x00000002	//モノルビ
#define	FELANG_CMODE_NOPRUNING		0x00000004	//枝狩りを行わない
#define	FELANG_CMODE_KATAKANAOUT	0x00000008	//カタカナ出力
#define	FELANG_CMODE_HIRAGANAOUT	0x00000000	//default output
#define	FELANG_CMODE_HALFWIDTHOUT	0x00000010	//半角出力
#define	FELANG_CMODE_FULLWIDTHOUT	0x00000020	//全角出力
#define	FELANG_CMODE_BOPOMOFO		0x00000040	//
#define	FELANG_CMODE_HANGUL			0x00000080	//
#define	FELANG_CMODE_PINYIN			0x00000100	//
#define	FELANG_CMODE_PRECONV		0x00000200	// 以下の変換前の変換をする
												// ・ローマ字かな変換     [ａ]→[あ]
												// ・変換前オートコレクト [ー]→[－]
												// ・句読点、括弧類		  [，]→[、]
#define	FELANG_CMODE_RADICAL		0x00000400	//
#define	FELANG_CMODE_UNKNOWNREADING	0x00000800	//
#define	FELANG_CMODE_MERGECAND		0x00001000	// 表記の同じ候補をmargeする
#define	FELANG_CMODE_ROMAN			0x00002000	// 
#define	FELANG_CMODE_BESTFIRST		0x00004000	// BEST 1st だけ作る。次候補はない。
#define	FELANG_CMODE_USENOREVWORDS  0x00008000	// REV/RECONVで逆変換に使わない単語も引く。

#define FELANG_CMODE_NONE			0x01000000	// IME_SMODE_NONE
#define FELANG_CMODE_PLAURALCLAUSE	0x02000000	// IME_SMODE_PLAURALCLAUSE
#define FELANG_CMODE_SINGLECONVERT	0x04000000	// IME_SMODE_SINGLECONVERT
#define FELANG_CMODE_AUTOMATIC		0x08000000	// IME_SMODE_AUTOMATIC
#define FELANG_CMODE_PHRASEPREDICT	0x10000000	// IME_SMODE_PHRASEPREDICT
#define FELANG_CMODE_CONVERSATION	0x20000000	// IME_SMODE_CONVERSATION
#define FELANG_CMODE_NAME			FELANG_CMODE_PHRASEPREDICT	// Name mode (MSKKIME)
#define FELANG_CMODE_NOINVISIBLECHAR    0x40000000      // remove invisible chars (e.g. tone mark)


// Error message
#define E_NOCAND			0x30	//候補が要求された数だけ無い
#define E_NOTENOUGH_BUFFER	0x31	//文字列バッファが足りない
#define E_NOTENOUGH_WDD		0x32	//WDD格納バッファが足りない
#define E_LARGEINPUT		0x33	//WDD格納バッファが足りない


//Morphology Info
#define FELANG_CLMN_WBREAK		0x01
#define FELANG_CLMN_NOWBREAK	0x02
#define FELANG_CLMN_PBREAK		0x04
#define FELANG_CLMN_NOPBREAK	0x08
#define FELANG_CLMN_FIXR		0x10
#define FELANG_CLMN_FIXD		0x20	// 単語表記固定

#define	FELANG_INVALD_PO		0xFFFF	// 入力文字未対応桁


///////////////////////////
// The IFELanguage global export function
///////////////////////////

HRESULT WINAPI CreateIFELanguageInstance(REFCLSID clsid, VOID **ppvObj);
typedef HRESULT (WINAPI *fpCreateIFELanguageInstanceType)(REFCLSID clsid, VOID **ppvObj);




///////////////////////////
// The IFELanguage interface
///////////////////////////

#undef INTERFACE
#define INTERFACE		IFELanguage

//IFELanguage template
DECLARE_INTERFACE_(IFELanguage,IUnknown)
{
	// IUnknown members
    STDMETHOD(QueryInterface)(THIS_ REFIID refiid, VOID **ppv) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;

	// Ijconv members.  must be virtual functions
    STDMETHOD(Open)(THIS) PURE;
    STDMETHOD(Close)(THIS) PURE;
	

//	・ppResultはサーバー側でCoTaskMemAllocで確保する。
//	  クライアントは使用後にCoTaskMemFreeで開放すること。
//	・pfCInfoはNULLもOK
	STDMETHOD(GetJMorphResult)(THIS_
						DWORD   dwRequest,			// [in]
						DWORD	dwCMode,			// [in]
						INT     cwchInput,      	// [in]
						_In_count_(cwchInput) WCHAR   *pwchInput,     	// [in, size_is(cwchInput)]
						DWORD   *pfCInfo,       	// [in, size_is(cwchInput)]
						MORRSLT **ppResult ) PURE;  // [out]

	STDMETHOD(GetConversionModeCaps)(THIS_ DWORD *pdwCaps) PURE;
	
	STDMETHOD(GetPhonetic)(THIS_
						BSTR	string,				// [in]
						LONG	start,				// [in]
						LONG	length,				// [in]
						BSTR *	phonetic ) PURE;	// [out, retval]

	STDMETHOD(GetConversion)(THIS_
						BSTR	string,				// [in]
						LONG	start,				// [in]
						LONG	length,				// [in]
						BSTR *	result ) PURE;		// [out, retval]

};

#ifdef __cplusplus
} /* end of 'extern "C" {' */
#endif


///========================================================================================
//	
//	ロードとアンロードの例
//	
//	/********************************************************************/
//  static HMODULE hLib = NULL;
//  
//  // IJMorphologyのロードとアンロード
//  HRESULT 
//  LoadIFELanguage( IFELanguage **ppjMorph ) 
//  {
//  
//  	HRESULT rc = S_FALSE;
//  							//ライブラリロード
//  	if (!hLib) {
//  		if ( Debug_Mode == TRUE ) {
//  			hLib = LoadLibrary( "debug98k.dll" );
//  		}
//  		else {
//  			hLib = LoadLibrary( "msime98k.dll" );
//  		}
//  		if(!hLib) {
//  			return rc;
//  		}
//  	}
//  
//  	//関数アドレス取得
//  	// cdeclデフォルトだとできない。
//  	HRESULT (*lpfnCreateIFELanguageInstance)( CLSID *, VOID **);
//  	lpfnCreateIFELanguageInstance = 
//  		(long (__cdecl *)(struct _GUID *,void ** ))
//  			GetProcAddress(hLib, "CreateIFELanguageInstance");
//  
//  	if(!lpfnCreateIFELanguageInstance) {	//ファンクションが取れない
//  		FreeLibrary( hLib );
//  		hLib = NULL;
//  		return rc;
//  	}
//  
//  	//IJKKCポインタ取得
//  	rc = (*lpfnCreateIFELanguageInstance)
//  		( NULL/*clsid*/, (VOID**)ppjMorph );
//  
//  	if (rc == S_OK ) {
//  		rc = (*ppjMorph)->Open();
//  	}
//  
//  	return rc;
//  }
//  
//  VOID 
//  UnloadIFELanguage( IFELanguage *pjMorph ) 
//  {
//  	pjMorph->Close();
//  	pjMorph->Release();
//  
//  	if ( hLib ) {
//  		FreeLibrary( hLib );
//  		hLib = NULL;
//  	}
//  }
//
//===========================================================================================


#endif //__FELang_H__

