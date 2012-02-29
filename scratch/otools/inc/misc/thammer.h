//+--------------------------------------------------------------------------
//
//  Copyright (C) 1994, 1995, 1996 Microsoft Corporation.  All Rights Reserved.
//
//  File:       thammer.h
//
//  This include file defines 3 exported APIs and their callbacks that export
//  word-breaking functionality for non-spaced Asian languages (Japanese, Chinese)
//
//  Summary of exports:
//      EnumSelectionOffsets - This function returns the offsets for the
//          selection chunks as specified in the Selection Profile (set at compile-time)
//      EnumSummarizationOffsets - This function returns the offsets for the
//          prefix (if any), the stem, and bound morphemes (fuzokugo).
//
//  History:                    pathal          Created.
//              25-Jun-97       pathal          Add TH_ERROR_INIT_FAILED
//              05-Jul-97        pathal          Add EnumSentenceOffsets, etc.
//              17-Sep-97       pathal          Removed private interfaces
//              23-Sep-97       pathal          Add EnumSelectionOffsetsEx for
//                                              multi-locale support (PRC/SC)
//              ""-"""-""       pathal          Move IN from prototype to headers
//              22-May-98       pathal          Remove EnumSelectionOffsetsEx
//---------------------------------------------------------------------------

// Return errors: the following error codes can be returned from any of
// T-Hammer's exported APIs (EnumSelectionOffsets, and EnumSummarizationOffsets.
//

#define TH_ERROR_SUCCESS 0
#define TH_ERROR_NOHPBS 1
#define TH_ERROR_NO_SENTENCES 1
#define TH_ERROR_INVALID_INPUT 2
#define TH_ERROR_INVALID_CALLBACK 3
#define TH_ERROR_INIT_FAILED 4
#define TH_ERROR_NOT_IMPLEMENTED 5
#define TH_ERROR_OUT_OF_MEMORY 6

// Offset delimiter: the following code is used to delimit the end of a list of
// token offsets returned to one of the Enum* callback routines.  This is not
// an error code.

#define TH_SELECTION_INVALID_OFFSET 0xFFFFFFFF

// TOKENIZE_MODE: Begin and End HPB Modes
//
// Begin and End HPB modes signify that a hard phrase break comes before the
// first character in the string and/or follows after the last character in the string
// If these flags are not set, then the default behavior of EnumTokens is to start
// enumerating tokens to the right of the leftmost HPB, which probably won't
// be at the first character (unless it is a punctuation symbol) and to conclude
// enumeration at the rightmost HPB, which likely will not be the true end of the
// string.  So, these flags in affect force HPBs at the 0th and nth offsets, where
// n is the number of characters in the input buffer
//
// WARNNIG: Since Tokenize operates in batch mode, it assumes that the
// start and end of the input buffer are HPBs. These flags are only used for
// EnumTokens
//
// CAUTION: Following two declarations are OBSOLETE as of 2/10/98.
// Please use ...BEGIN_SENTENCE and ...END_SENTENCE in place of them.
//
#define TOKENIZE_MODE_BEGIN_HPB             0x00000001
#define TOKENIZE_MODE_END_HPB             0x00000002

// Note on HPBs:  HPB = hard phrase break.
// HPBs are statistically determined from analyzing a tagged corpora.
// Roughly, they cor-respond to places where you csn break with 100%
// precision (=confidence). Mostly this is around punctuation characters
// and certain conspicuous [case markers | character type] bigrams.

#define TOKENIZE_MODE_BEGIN_SENTENCE             0x00000001
#define TOKENIZE_MODE_END_SENTENCE             0x00000002

// When the Hide Punctuation mode is set in the tokenize flag parameter
// T-Hammer strips punctuation out of the Stem Offsets and Summarization Offsets
// callback
//
#define TOKENIZE_MODE_HIDE_PUNCTUATION    0x00000004

//+--------------------------------------------------------------------------
//  Routine:    EnumSelectionOffsetsCallback
//
//  Synopsis: client-side callback that receives a list of offsets for selection chunks
//
//  Parameters:
//  IN  pichOffsets - pointer to first element in an array of offsets into client
//          text buffer. NOTE: callback is not allowed to stash pichChunks for
//          later processing.  pichChunks will not persist between successive
//          callbacks.  If the callback wants to use the data pointed to by pich
//          it must copy it to its own store
//  IN  cOffsets - number of offsets passed to client (always > 1)
//  IN lpData - client defined data
//
// Return:
//      TRUE - to abort token enumeration
//      FALSE - to continue
//---------------------------------------------------------------------------
typedef BOOL (CALLBACK * ENUM_SELECTION_OFFSETS_CALLBACK)(
    CONST DWORD *pichOffsets,
    CONST DWORD cOffsets,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumSelectionOffsets
//
//  Synopsis:  This is the main entry point for tokenizing text.  Sends tokens,
//  which can either be offsets or zero delimited strings to callback.
//
//  Parameters:
//  IN  pwszText - pointer to wide-character text buffer to be tokenized,
//  IN  cchText - count of characters in text buffer,
//  IN  fBeginEndHPBMode - flag describing the callback mode  (see above),
//  IN  pfnEnumSelectionOffsets - pointer to callback procedure handling token
//          enumeration,
//  IN lpData - client defined data
//
//  Returns:
//      TH_ERROR_SUCCESS - if the call completed successfully
//      TH_ERROR_NOHPBS - if there were no HPBs
//      TH_ERROR_INVALID_INPUT - if the input buffer was bad
//      TH_ERROR_INVALID_CALLBACK - if the input callback was bad
//---------------------------------------------------------------------------
INT
APIENTRY
EnumSelectionOffsets(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fBeginEndHPBMode,
    ENUM_SELECTION_OFFSETS_CALLBACK pfnEnumSelectionOffsets,
    LPARAM lpData);

typedef INT (APIENTRY *LP_ENUM_SELECTION_OFFSETS)(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fBeginEndHPBMode,
    ENUM_SELECTION_OFFSETS_CALLBACK pfnEnumSelectionOffsets,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumSummarizationOffsetsCallback
//
//  Synopsis: client-side callback that receives a list of offsets for each stem
//      in the free morpheme (jiritsugo) phrase.   Last offset is always contains
//      the complete string of bound morphemes (fuzokugo).  For example,
//      for "kaisan shite nai sou desu", offsets are returned for "kaisan" and
//      "shite nai sou desu".  So, counting the first initial offset, there are three
//      offsets.
//
//  Parameters:
//  IN  pichOffsets - pointer to first element in an array of offsets into client
//          text buffer. NOTE: callback is not allowed to stash pichOffsets for
//          later processing.  pichOffsets will not persist between successive
//          callbacks.  If the callback wants to use the data pointed to by pich
//          it must copy it to its own store
//  IN  cOffsets - number of offsets passed to client (always > 1)
//  IN lpData - client defined data
//
// Return:
//  TRUE - to abort token enumeration
//  FALSE - to continue
//---------------------------------------------------------------------------
typedef BOOL (CALLBACK * ENUM_SUMMARIZATION_OFFSETS_CALLBACK)(
    CONST DWORD *pichOffsets,
    CONST DWORD cOffsets,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumSummarizationOffsets
//
//  Synopsis:  This is the entry point for returning offsets for tokens used
//  in summarization.   These tokens correspond to stems and bound morphemes
//  (fuzokugo) in the text.  A list of offsets (and a count) is sent to the
//  EnumSummarizationOffsets callback (see above)
//
//  Parameters:
//  IN pwszText - pointer to wide-character text buffer to be tokenized,
//  IN cchText - count of characters in text buffer,
//  IN fTokenizeMode - flag describing the callback mode  (see above),
//  IN pEnumSummarizationOffsets - pointer to callback procedure handling
//          token enumeration,
//  IN lpData - client defined data
//
//  Returns:
//      TH_ERROR_SUCCESS - if the call completed successfully
//      TH_ERROR_NOHPBS - if there were no HPBs
//      TH_ERROR_INVALID_INPUT - if the input buffer was bad
//      TH_ERROR_INVALID_CALLBACK - if the input callback was bad
//---------------------------------------------------------------------------
INT
APIENTRY
EnumSummarizationOffsets(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fBeginEndHPBMode,
    ENUM_SUMMARIZATION_OFFSETS_CALLBACK pfnEnumSummarizationOffsets,
    LPARAM lpData);

typedef INT (APIENTRY *LP_ENUM_SUMMARIZATION_OFFSETS)(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fBeginEndHPBMode,
    ENUM_SUMMARIZATION_OFFSETS_CALLBACK pfnEnumSummarizationOffsets,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumStemInfoCallback
//
//  Synopsis: client-side callback that receives offsets and stem information
//      CAUTION: Clients should not rely on the language specific pwszMCat
//      since the MCats will change in the future.  Please use pwszPOS for
//      language independent part-of-speech information.
//
//  Parameters:
//  IN  ichOffset - offset to first character in stem
//  IN  cchLen - length of the stem
//  IN  pwszPOS - string containing POS info
//  IN  pwszMCat - string containing MCat info (this is private and may change)
//  IN  pwszDictionaryForm - string containing Dictionary Form
//  IN OUT  lpData - client defined data
//
// Return:
//      TRUE - to abort token enumeration
//      FALSE - to continue
//---------------------------------------------------------------------------
typedef BOOL (CALLBACK * ENUM_STEM_INFO_CALLBACK)(
    CONST DWORD ichOffset,
    CONST DWORD cchLen,
    PCWSTR pwszPOS,
    PCWSTR pwszMCat,
    PCWSTR pwszDictionaryForm,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumStemInfo
//
//  Synopsis:  Call this routine to get information about stems.
//  For example, if you want the dictionary form, part-of-speech or
//  MCat information for a stem, then this is the API for you
//
//  Parameters:
//  IN  pwszText - pointer to wide-character text buffer to be tokenized,
//  IN  cchText - count of characters in text buffer,
//  IN  fTokenizeMode - flag describing the callback mode  (see above),
//  IN  pfnEnumStemInfo - pointer to callback procedure handling stem info
//          enumeration
//  IN  lpData - client defined data
//
//  Returns:
//      TH_ERROR_SUCCESS - if the call completed successfully
//      TH_ERROR_NOHPBS - if there were no HPBs
//      TH_ERROR_INVALID_INPUT - if the input buffer was bad
//      TH_ERROR_INVALID_CALLBACK - if the input callback was bad
//---------------------------------------------------------------------------
INT
APIENTRY
EnumStemInfo(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fBeginEndHPBMode,
    ENUM_STEM_INFO_CALLBACK pfnEnumStemInfo,
    OUT DWORD *pcchTextProcessed,
    LPARAM lpData);

typedef INT (APIENTRY *LP_ENUM_STEM_INFO)(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fBeginEndHPBMode,
    ENUM_STEM_INFO_CALLBACK pfnEnumStemInfo,
    OUT DWORD *pcchTextProcessed,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumSentenceOffsetsCallback
//
//  Synopsis: client-side callback that receives a list of offsets for sentence breaks
//
//  Parameters:
//  IN  ichOffsetStart - offset to start of sentence
//  IN  ichOffsetEnd - offset to end of sentence (includes terminating punctuation)
//  IN  lpData - client defined data
//
// Return:
//      TRUE - to abort token enumeration
//      FALSE - to continue
//---------------------------------------------------------------------------
// BOOL
// EnumSentenceOffsetsCallback (
//    DWORD ichOffsetStart,
//    DWORD ichOffsetEnd,
//    LPARAM lpData);

typedef BOOL (CALLBACK * ENUM_SENTENCE_OFFSETS_CALLBACK)(
    DWORD ichOffsetStart,
    DWORD ichOffsetEnd,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    EnumSentenceOffsets
//
//  Synopsis:  This is the main entry point for breaking sentences.
//      Sends offsets delimiting sentences to the callback.
//
//  Parameters:
//  IN  pwszText - pointer to wide-character text buffer to be tokenized,
//  IN  cchText - count of characters in text buffer,
//  IN  fTokenizeMode - not used.  later this will be used to control how
//          partial sentences are handled.
//  IN  pEnumSentenceOffsetsCallback - pointer to callback procedure handling offsets
//  IN  lpData - client defined data
//
//  Returns:
//      TH_ERROR_SUCCESS - if the call completed successfully
//      TH_ERROR_NOHPBS - if there were no HPBs
//      TH_ERROR_INVALID_INPUT - if the input buffer was bad
//      TH_ERROR_INVALID_CALLBACK - if the input callback was bad
//---------------------------------------------------------------------------
INT
APIENTRY
EnumSentenceOffsets(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fTokenizeMode,
    ENUM_SENTENCE_OFFSETS_CALLBACK pfnEnumSentenceOffsets,
    LPARAM lpData);

typedef INT (APIENTRY *LP_ENUM_SENTENCE_OFFSETS)(
    PCWSTR pwszText,
    DWORD cchText,
    DWORD fTokenizeMode,
    ENUM_SENTENCE_OFFSETS_CALLBACK pfnEnumSentenceOffsets,
    LPARAM lpData);

//+--------------------------------------------------------------------------
//  Routine:    WordBreakInit
//
//  Synopsis:   Initializes Spruce by loading the external resource file.
//              Supposed to be called before using any word breaking feature.
//              (Otherwise default path will be searched for the resource).
//
//  Parameters: IN pwszMdt - path to the external resource file.
//
//  Returns:    True if successful, else False.
//---------------------------------------------------------------------------
BOOL
APIENTRY
WordBreakInit(
    _In_z_ WCHAR *pwszMdt);

typedef BOOL (APIENTRY *LP_WORDBREAK_INIT)(
    WCHAR *pwszMdt);


//+--------------------------------------------------------------------------
//  Routine:    EnumAppWordsCallback
//
//  Synopsis:
//  Callback for enumerating application's private words
//  so they can used in wordbreaker.  For example, Answer
//  Wizard can provide special app-specific terms to improve
//  tokenization and indexing accuracy for queries specific to the
//  app.
//
//  Parameters:
//  IN wzWordBuffer - ptr to zero terminated buffer for holding a word
//      can be either zero-terminated for one word; or,
//      use double termination for sublists
//  IN cwcWordBuffer - count of wide chars (i.e. length of pwzWordsBuffer)
//     IN fSingleWord - FALSE means return double zero terminated list.
//                              TRUE means return single zero terminated word.
//  IN lpData - client defined data
//
//  Returns:
//  TRUE when buffer contains data. Otherwise, FALSE.
//
//--------------------------------------------------------------------------

typedef BOOL (CALLBACK *LP_ENUM_APP_WORDS_CALLBACK)(
    WCHAR  * wzWordBuffer,
    DWORD cwcWordBuffer,
    BOOL fSingleWord,
    LPARAM lpData
);


//+--------------------------------------------------------------------------
//  Routine:    WordBreakInitEx
//
//  Synopsis:
//  Extended initialization entry point.  This API
//  supports callback to the callee to enumerate lexical items.
//  Can be used to provide a supplemental secondary word list.
//  Examples: AW word list for Chinese wordbreaker.
//  Notes for WordBreakInit apply:
//        Supposed to be called before using any word breaking feature.
//       (Otherwise default path will be searched for the resource).
//
//  Parameters:
//  IN pfnEnumAppWords - ptr to callback for enumerating app words
//  IN lpData - client defined data used for callback
//
//  Returns:
//  TRUE if successful; otherwise, FALSE.
//---------------------------------------------------------------------------
BOOL
APIENTRY
WordBreakInitEx(
    _In_z_ WCHAR *pwszMdt,
    LP_ENUM_APP_WORDS_CALLBACK pfnEnumAppWords,
    LPARAM lpData);

typedef BOOL (APIENTRY *LP_WORDBREAK_INIT_EX)(
    WCHAR *pwszMdt,
    LP_ENUM_APP_WORDS_CALLBACK pfnEnumAppWords,
    LPARAM lpData);

