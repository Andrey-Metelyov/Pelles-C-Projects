/****************************************************************************
 *                                                                          *
 * File    : cppfile.c                                                      *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates a C++ syntax parser,           *
 *           a C++ dependency scanner, and some (incomplete) build rules.   *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <shlwapi.h>
#include <addin.h>
#include <wchar.h>
#include <wctype.h>
#include <stdlib.h>

#define CCHMAXLINE  4096

#define NELEMS(a)  (sizeof(a) / sizeof(a[0]))

// Flags for usCookie.
#define FLAG_COMMENT       0x01
#define FLAG_PREPROCESSOR  0x02
#define FLAG_EXT_COMMENT   0x04
#define FLAG_STRING        0x08
#define FLAG_CHAR          0x10
#define FLAG_UNUSED4       0x20    // Free for any use.
#define FLAG_UNUSED5       0x40    // Free for any use.
#define FLAG_UNUSED6       0x80    // Free for any use.

// Minimum and maximum folding level.
#define MIN_FOLDLEVEL  0
#define MAX_FOLDLEVEL  255

// Keyword list - must be sorted for IsKeyword().
static PCWSTR apcszKeywords[] = {
    L"alignas",  /* C++11 */
    L"alignof",  /* C++11 */
    L"and",
    L"and_eq",
    L"asm",
    L"auto",
    L"bitand",
    L"bitor",
    L"bool",
    L"break",
    L"case",
    L"catch",
    L"char",
    L"char16_t",  /* C++11 */
    L"char32_t",  /* C++11 */
    L"class",
    L"compl",
    L"concept",  /* C++20 */
    L"const",
    L"constexpr",  /* C++11 */
    L"const_cast",
    L"continue",
    L"decltype",  /* C++11 */
    L"default",
    L"delete",
    L"do",
    L"double",
    L"dynamic_cast",
    L"else",
    L"enum",
    L"explicit",
    L"export",
    L"extern",
    L"false",
    L"float",
    L"for",
    L"friend",
    L"goto",
    L"if",
    L"inline",
    L"int",
    L"long",
    L"mutable",
    L"namespace",
    L"new",
    L"noexcept",  /* C++11 */
    L"not",
    L"not_eq",
    L"nullptr",  /* C++11 */
    L"operator",
    L"or",
    L"or_eq",
    L"private",
    L"protected",
    L"public",
    L"register",
    L"reinterpret_cast",
    L"requires",  /* C++20 */
    L"return",
    L"short",
    L"signed",
    L"sizeof",
    L"static",
    L"static_assert",  /* C++11 */
    L"static_cast",
    L"struct",
    L"switch",
    L"template",
    L"this",
    L"thread_local",  /* C++11 */
    L"throw",
    L"true",
    L"try",
    L"typedef",
    L"typeid",
    L"typename",
    L"union",
    L"unsigned",
    L"using",
    L"virtual",
    L"void",
    L"volatile",
    L"wchar_t",
    L"while",
    L"xor",
    L"xor_eq"
};
/* +overide, +final - in some contexts */

// Helper macro for assigning color to column position.
#define DEFINE_BLOCK(pch,color) \
    do { \
        if (pPoints != NULL) { \
            if (*pcPoints == 0 || (pPoints - 1)->iColor != (color)) { \
                pPoints->iChar = (UINT)((pch) - pchText); \
                pPoints->iColor = (color); \
                pPoints++; \
                (*pcPoints)++; \
            } \
        } \
    } while (0)

// Known Unicode byte-order marks (little endian).
#define BOM_UTF8      0xBFBBEF
#define BOM_UTF16_LE  0xFEFF
// This sample ignores the following BOMs:
/* #define BOM_UTF16_BE  0xFFFE */
/* #define BOM_UTF32_LE  0x0000FEFF */
/* #define BOM_UTF32_BE  0xFFFE0000 */

// Text file encoding.
typedef enum Encodings {
    ENCODING_UNKNOWN,
    ENCODING_ANSI,
    ENCODING_UTF8,
    ENCODING_UTF16_LE,
    /* TODO: add more... */
} ENCODING;

#define CHAR_EOF  0x1A

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;

// Function prototypes.
static BOOL CALLBACK EnumProjFileCallback(LPCWSTR, LPVOID);
static BOOL IsCppFile(PCWSTR);
static USHORT CALLBACK Parser(USHORT, PCWSTR, int, ADDIN_PARSE_POINT [], PINT);
static size_t IsKeyword(PCWSTR [], size_t, PCWSTR, size_t);
static PCWSTR ParseNumber(PCWSTR, PCWSTR);
static BOOL CALLBACK Scanner(LPCWSTR, BOOL (CALLBACK *)(LPCWSTR, LPCVOID), LPCVOID);
static BOOL GetInputLine(HANDLE, PWSTR *, ENCODING *);
static ENCODING ReadTextFileEncoding(HANDLE);
static DWORD GetTextFileEncoding(DWORD, ENCODING *);
static BOOL ReadTextFileLine(HANDLE, ENCODING, PWSTR, DWORD);
static BOOL SeekFile(HANDLE, DWORD, DWORD);
static DWORD TellFile(HANDLE);
static PWSTR NoUnixSlash(PWSTR);
static PWSTR SearchIncludeFile(PCWSTR, PCWSTR);
static BOOL IsExistingFile(PCWSTR);

// Inline functions.
static inline PWSTR SkipWhiteSpace(PWSTR psz)
{
    while (*psz == L' ' || *psz == L'\t') ++psz;
    return psz;
}

static inline BOOL SeekFile(HANDLE hf, DWORD dwOffset, DWORD method)
{
    return SetFilePointer(hf, dwOffset, NULL, FILE_BEGIN) != INVALID_SET_FILE_POINTER;
}

static inline DWORD TellFile(HANDLE hf)
{
    return SetFilePointer(hf, 0, NULL, FILE_CURRENT);
}

/****************************************************************************
 *                                                                          *
 * Function: DllMain                                                        *
 *                                                                          *
 * Purpose : DLL entry and exit procedure.                                  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

BOOL WINAPI DllMain(HANDLE hDLL, DWORD dwReason, LPVOID lpReserved)
{
    switch (dwReason)
    {
        case DLL_PROCESS_ATTACH:
            g_hmod = hDLL;
            return TRUE;

        case DLL_PROCESS_DETACH:
            return TRUE;

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: AddInMain                                                      *
 *                                                                          *
 * Purpose : Add-in entry and exit procedure.                               *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

ADDINAPI BOOL WINAPI AddInMain(HWND hwnd, ADDIN_EVENT eEvent)
{
    switch (eEvent)
    {
        case AIE_APP_CREATE:
        {
            ADDIN_ADD_FILE_TYPE AddFile = {0};

            // Define a new file type in the IDE (.cpp).
            AddFile.cbSize = sizeof(AddFile);
            AddFile.pszDescription = L"C++ file";
            AddFile.pszExtension = L"cpp";  /* support *.cpp files */
            AddFile.pfnParser = Parser;  /* syntax color parser */
            AddFile.pszShells =
                L"$(CPP) $(CPPFLAGS) \"$!\" -Fo\"$@\"\0";  /* command #1 */
                L"\0";  /* terminate list of commands */
            AddFile.pfnScanner = Scanner;  /* dependency scanner */
            if (!AddIn_AddFileType(hwnd, &AddFile))
                return FALSE;

            // Save handle of the main IDE window.
            g_hwndMain = hwnd;

            return TRUE;
        }

        case AIE_PRJ_SAVE:  /* after significant changes, like adding or deleting project files */
        {
            // AddIn_SetProjectSymbol() will trigger a new AIE_PRJ_SAVE event.
            static int cRecurse = 0;
            if (!cRecurse++)
            {
                ADDIN_ENUM_PROJECT_FILES Enum = { .fuFlags = EPFF_TARGET_FILES };
                BOOL fGotCpp = FALSE;

                // Enumerate current project files - look for file with .cpp extension.
                Enum.cbSize = sizeof(Enum);
                Enum.pfnCallback = EnumProjFileCallback;
                Enum.fuFlags = EPFF_DEPENDENT_FILES;
                Enum.pvData = &fGotCpp;
                if (AddIn_EnumProjectFiles(hwnd, &Enum) && fGotCpp)
                {
                    WCHAR ach[4];

                    // Found file with .cpp extension in project - look for our CPP macro.
                    if (AddIn_GetProjectSymbol(hwnd, L"CPP", ach, NELEMS(ach)) == 0)
                    {
                        // The macro CPP wasn't found (or is empty) - set some defaults.

                        // TODO: CL will need a path in Tools -> Options -> Folders.
                        (void)AddIn_SetProjectSymbol(hwnd, L"CPP", L"cl.exe");

                        // TODO: change INCLUDE path's to something better!
                        (void)AddIn_SetProjectSymbol(hwnd, L"CPPFLAGS", L"/c /nologo /O2 /X"
                            L" /I\"C:\\Program Files\\Microsoft Visual Studio\\ATL\\Include\""
                            L" /I\"C:\\Program Files\\Microsoft Visual Studio\\Include\""
                            L" /I\"C:\\Program Files\\Microsoft Visual Studio\\MFC\\Include\"");
                    }
                }
            }
            --cRecurse;
            return TRUE;
        }

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: EnumProjFileCallback                                           *
 *                                                                          *
 * Purpose : AddIn_EnumProjectFiles() callback procedure.                   *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL CALLBACK EnumProjFileCallback(LPCWSTR pcszName, LPVOID pvData)
{
    // Check for .cpp extension.
    if (IsCppFile(pcszName))
    {
        // Found .cpp extension - no need to continue.
        *(BOOL *)pvData = TRUE;
        return FALSE;
    }

    return TRUE;
}

/****************************************************************************
 *                                                                          *
 * Function: AddInHelp                                                      *
 *                                                                          *
 * Purpose : Add-in help event procedure.                                   *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

ADDINAPI BOOL WINAPI AddInHelp(HWND hwnd, ADDIN_HELPEVENT eEvent, LPCVOID pcvData)
{
    // Use AIHE_SRC_KEYWORD_FIRST to avoid getting help for C keywords with the same name.
    if (eEvent == AIHE_SRC_KEYWORD_FIRST && IsKeyword(apcszKeywords,
        NELEMS(apcszKeywords), (LPCWSTR)pcvData, wcslen((LPCWSTR)pcvData)) != 0)
    {
        HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
        if (hwndDoc)
        {
            ADDIN_DOCUMENT_INFO DocInfo = {0};

            DocInfo.cbSize = sizeof(DocInfo);
            if (AddIn_GetDocumentInfo(hwndDoc, &DocInfo) && DocInfo.nType == AID_SOURCE && IsCppFile(DocInfo.szFilename))
            {
                // Display F1 help for the keyword pointed to by pcvData.
                WCHAR szText[256];

                // TODO: replace with call to WinHelp() or HTMLHelp(), for example.
                wsprintf(szText, L"Room for helpful text about '%.128s'...", (LPCWSTR)pcvData);
                MessageBox(hwnd, szText, L"Wonderful C++ help system", MB_OK);
                // We "handled" this - so return TRUE.
                return TRUE;
            }
        }
    }

    return FALSE;
}

/****************************************************************************
 *                                                                          *
 * Function: IsCppFile                                                      *
 *                                                                          *
 * Purpose : Check for C++ file extension.                                  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL IsCppFile(PCWSTR pcszFileName)
{
    PCWSTR pcsz;

    // Check for .cpp extension.
    return ((pcsz = wcsrchr(pcszFileName, L'.')) != NULL && wcscmp(pcsz, L".cpp") == 0) ? TRUE : FALSE;
}

/****************************************************************************
 *                                                                          *
 * Function: Parser                                                         *
 *                                                                          *
 * Purpose : Parse C++ source code - for syntax color highlighting.         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static USHORT CALLBACK Parser(USHORT usCookie, PCWSTR pchText, int cchText, ADDIN_PARSE_POINT pPoints[], PINT pcPoints)
{
    PCWSTR pch = pchText, pchEnd = &pchText[cchText];
    BYTE cFoldLevel = ADDIN_GET_COOKIE_LEVEL(usCookie);
    BYTE bFlags = ADDIN_GET_COOKIE_FLAGS(usCookie);

    if (pch == pchEnd)
        ;
    // Inside comment or extended comment.
    else if (bFlags & (FLAG_COMMENT|FLAG_EXT_COMMENT))
    {
        DEFINE_BLOCK(pch, ADDIN_COLOR_COMMENT);
    }
    // Inside string or char constant.
    else if (bFlags & (FLAG_CHAR|FLAG_STRING))
    {
        DEFINE_BLOCK(pch, ADDIN_COLOR_STRING);
    }
    // Inside preprocessor directive.
    else if (bFlags & FLAG_PREPROCESSOR)
    {
        DEFINE_BLOCK(pch, ADDIN_COLOR_PREPROCESSOR);
    }
    else /* normal state */
    {
        while (pch < pchEnd && (*pch == L' ' || *pch == L'\t'))
            pch++;

        // Check for preprocessor directive: #...
        if (pch < pchEnd && *pch == L'#')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_PREPROCESSOR);
            bFlags |= FLAG_PREPROCESSOR;

            pch++;
            while (pch < pchEnd && (*pch == L' ' || *pch == L'\t'))
                pch++;

            // Check for #if, #ifdef, #ifndef block.
            if (&pch[1] < pchEnd &&
                pch[0] == L'i' &&
                pch[1] == L'f' &&
                cFoldLevel < MAX_FOLDLEVEL)
            {
                cFoldLevel++;
            }
            else if (&pch[4] < pchEnd &&
                pch[0] == L'e' &&
                pch[1] == L'n' &&
                pch[2] == L'd' &&
                pch[3] == L'i' &&
                pch[4] == L'f' &&
                cFoldLevel > MIN_FOLDLEVEL)
            {
                cFoldLevel--;
            }
        }
    }

    while (pch < pchEnd)
    {
        // Inside string constant: "...."
        if (bFlags & FLAG_STRING)
        {
            // Check for escape sequence.
            if (*pch == L'\\')
                pch++;
            // Check for end of string constant.
            else if (*pch == L'\"')
                bFlags &= ~FLAG_STRING;
            pch++;
            continue;
        }

        // Inside char constant: '.'
        if (bFlags & FLAG_CHAR)
        {
            // Check for escape sequence.
            if (*pch == L'\\')
                pch++;
            // Check for end of char constant.
            else if (*pch == L'\'')
                bFlags &= ~FLAG_CHAR;
            pch++;
            continue;
        }

        // Inside single line comment: //....
        if (bFlags & FLAG_COMMENT)
            break;  /* done! */

        // Inside extended comment: /*....*/
        if (bFlags & FLAG_EXT_COMMENT)
        {
            // Check for end of extended comment...
            if (*pch == L'*' && &pch[1] < pchEnd && pch[1] == L'/')
            {
                bFlags &= ~FLAG_EXT_COMMENT;
                pch++;

                // Can fold extended comments.
                if (cFoldLevel > MIN_FOLDLEVEL) cFoldLevel--;
            }
            pch++;
            continue;
        }

        // Inside preprocessor directive: #...
        if (bFlags & FLAG_PREPROCESSOR)
        {
            // Check for string constant: "...."
            if (*pch == L'\"')
            {
                bFlags |= FLAG_STRING;
            }
            // Check for char constant: '.'
            else if (*pch == L'\'')
            {
                bFlags |= FLAG_CHAR;
            }
            // Check for single line comment: //
            else if (*pch == L'/' && &pch[1] < pchEnd && pch[1] == L'/')
            {
                DEFINE_BLOCK(pch, ADDIN_COLOR_COMMENT);
                bFlags |= FLAG_COMMENT;
                pch++;
            }
            // Check for extended comment: /*....*/
            else if (*pch == L'/' && &pch[1] < pchEnd && pch[1] == L'*')
            {
                DEFINE_BLOCK(pch, ADDIN_COLOR_COMMENT);
                bFlags |= FLAG_EXT_COMMENT;
                pch++;

                // Can fold extended comments.
                if (cFoldLevel < MAX_FOLDLEVEL) cFoldLevel++;
            }
            else
            {
                // Default case.
                DEFINE_BLOCK(pch, ADDIN_COLOR_PREPROCESSOR);
            }
            pch++;
            continue;
        }

        // Update curly brace level.
        if (*pch == L'{' && cFoldLevel < MAX_FOLDLEVEL)
            cFoldLevel++;
        else if (*pch == L'}' && cFoldLevel > MIN_FOLDLEVEL)
            cFoldLevel--;

        // Check for single line comment: //
        if (*pch == L'/' && &pch[1] < pchEnd && pch[1] == L'/')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_COMMENT);
            bFlags |= FLAG_COMMENT;
            break;  /* done! */
        }

        // Check for extended comment: /*....*/
        if (*pch == L'/' && &pch[1] < pchEnd && pch[1] == L'*')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_COMMENT);
            bFlags |= FLAG_EXT_COMMENT;
            pch += 2;

            // Can fold extended comments.
            if (cFoldLevel < MAX_FOLDLEVEL) cFoldLevel++;
            continue;
        }

        // Check for string constant: "...."
        if (*pch == L'\"')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_STRING);
            bFlags |= FLAG_STRING;
            pch++;
            continue;
        }

        // Check for char constant: '.'
        if (*pch == L'\'')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_STRING);
            bFlags |= FLAG_CHAR;
            pch++;
            continue;
        }

        // Check for operator.
        if (&pch[2] < pchEnd && (
            (pch[0] == L'<' && pch[1] == L'<' && pch[2] == L'=') ||  /* <<= */
            (pch[0] == L'>' && pch[1] == L'>' && pch[2] == L'=') ||  /* >>= */
            (pch[0] == L'.' && pch[1] == L'.' && pch[2] == L'.') ||  /* ... */
            (pch[0] == L'-' && pch[1] == L'>' && pch[2] == L'*') ||  /* ->* */
            (pch[0] == L'<' && pch[1] == L'=' && pch[2] == L'>')))   /* <=> */
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_OPERATOR);
            pch += 3;
            continue;
        }
        else if (&pch[1] < pchEnd && (
            (pch[0] == L'+' && pch[1] == L'+') ||  /* ++ */
            (pch[0] == L'-' && pch[1] == L'-') ||  /* -- */
            (pch[0] == L'-' && pch[1] == L'>') ||  /* -> */
            (pch[0] == L'<' && pch[1] == L'<') ||  /* << */
            (pch[0] == L'>' && pch[1] == L'>') ||  /* >> */
            (pch[0] == L'<' && pch[1] == L'=') ||  /* <= */
            (pch[0] == L'>' && pch[1] == L'=') ||  /* >= */
            (pch[0] == L'=' && pch[1] == L'=') ||  /* == */
            (pch[0] == L'!' && pch[1] == L'=') ||  /* != */
            (pch[0] == L'&' && pch[1] == L'&') ||  /* && */
            (pch[0] == L'|' && pch[1] == L'|') ||  /* || */
            (pch[0] == L'*' && pch[1] == L'=') ||  /* *= */
            (pch[0] == L'/' && pch[1] == L'=') ||  /* /= */
            (pch[0] == L'%' && pch[1] == L'=') ||  /* %= */
            (pch[0] == L'+' && pch[1] == L'=') ||  /* += */
            (pch[0] == L'-' && pch[1] == L'=') ||  /* -= */
            (pch[0] == L'&' && pch[1] == L'=') ||  /* &= */
            (pch[0] == L'^' && pch[1] == L'=') ||  /* ^= */
            (pch[0] == L'|' && pch[1] == L'=') ||  /* |= */
            (pch[0] == L'.' && pch[1] == L'*') ||  /* .* */
            (pch[0] == L':' && pch[1] == L':')))   /* :: */
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_OPERATOR);
            pch += 2;
            continue;
        }
        else if (
            *pch == L',' ||
            *pch == L'*' ||  /* indirection or multiply */
            *pch == L'(' ||
            *pch == L')' ||
            *pch == L'{' ||
            *pch == L'}' ||
            *pch == L'[' ||
            *pch == L']' ||
            *pch == L'=' ||
            *pch == L'&' ||  /* address of, or bitwise and */
            *pch == L'!' ||
            *pch == L'+' ||
            *pch == L'-' ||
            *pch == L'.' ||
            *pch == L'<' ||
            *pch == L'>' ||
            *pch == L'/' ||
            *pch == L'%' ||
            *pch == L'^' ||
            *pch == L'|' ||
            *pch == L'?' ||  /* ?: */
            *pch == L':' ||  /* ?: or bitfield */
            *pch == L'~')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_OPERATOR);
            pch++;
            continue;
        }

        // Check for number.
        if (iswdigit(*pch))
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_NUMBER);
            pch = ParseNumber(pch, pchEnd);
            continue;
        }

        // Check for identifier.
        if (*pch == L'_' || IsCharAlphaW(*pch))
        {
            PCWSTR pchIdent = pch;
            do
                pch++;
            while (pch < pchEnd && (*pch == L'_' || IsCharAlphaNumericW(*pch)));

            if (IsKeyword(apcszKeywords, NELEMS(apcszKeywords), pchIdent, pch - pchIdent))
            {
                DEFINE_BLOCK(pchIdent, ADDIN_COLOR_KEYWORD);
            }
            else
            {
                DEFINE_BLOCK(pchIdent, ADDIN_COLOR_TEXT);
            }
            continue;
        }

        // Default case.
        DEFINE_BLOCK(pch, ADDIN_COLOR_TEXT);
        pch++;

        // Skip useless white-space.
        while (pch < pchEnd && (*pch == L' ' || *pch == L'\t'))
            pch++;
    }

    // If no line continuation ('\'), clear *most* flags for next line.
    if (cchText < 2 || pchText[cchText-1] != L'\n' || pchText[cchText-2] != L'\\')
        bFlags &= FLAG_EXT_COMMENT;

    return ADDIN_MAKE_COOKIE(bFlags, cFoldLevel);
}

/****************************************************************************
 *                                                                          *
 * Function: IsKeyword                                                      *
 *                                                                          *
 * Purpose : Return non-zero index, if the given string is a keyword.       *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static size_t IsKeyword(PCWSTR apcszKeywords[], size_t cKeywords, PCWSTR pchText, size_t cchText)
{
    PCWSTR *ppcszKeywords = apcszKeywords;

    while (cKeywords > 0)
    {
        size_t pivot = cKeywords >> 1;
        int cond = wcsncmp(ppcszKeywords[pivot], pchText, cchText);
        if (cond > 0)
        {
            // Search below pivot.
            cKeywords = pivot;
            continue;
        }
        else if (cond < 0)
        {
            // Search above pivot.
            ppcszKeywords += pivot + 1;
            cKeywords -= pivot + 1;
            continue;
        }
        else
        {
            // Got a match, but is it exact?
            if (ppcszKeywords[pivot][cchText] == L'\0')
                return 1 + (ppcszKeywords - apcszKeywords) + pivot;

            // Search below pivot (for a shorter keyword).
            cKeywords = pivot;
            continue;
        }
    }
    return 0;
}

/****************************************************************************
 *                                                                          *
 * Function: ParseNumber                                                    *
 *                                                                          *
 * Purpose : Parse a numeric value.                                         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static PCWSTR ParseNumber(PCWSTR pch, PCWSTR pchEnd)
{
    BOOL fFloat = FALSE;

    // Parse hexadecimal number.
    if (pch[0] == L'0' && &pch[1] < pchEnd && (pch[1] == L'x' || pch[1] == L'X'))
    {
        pch += 2;

        while (pch < pchEnd && iswxdigit(*pch))
            pch++;
    }
    // Parse binary number (C++14).
    if (pch[0] == L'0' && &pch[1] < pchEnd && (pch[1] == L'b' || pch[1] == L'B'))
    {
        pch += 2;

        while (pch < pchEnd && (*pch == L'0' || *pch == L'1'))
            pch++;
    }
    // Parse decimal number.
    else
    {
        pch++;

        while (pch < pchEnd && iswdigit(*pch))
            pch++;

        if (pch < pchEnd && *pch == L'.')
        {
            pch++;

            while (pch < pchEnd && iswdigit(*pch))
                pch++;

            fFloat = TRUE;
        }

        if (pch < pchEnd && (*pch == L'e' || *pch == L'E'))
        {
            pch++;

            if (pch < pchEnd && (*pch == L'-' || *pch == L'+'))
                pch++;

            while (pch < pchEnd && iswdigit(*pch))
                pch++;

            fFloat = TRUE;
        }
    }

    // Parse suffix.
    if (fFloat)
    {
        if (pch < pchEnd && (*pch == L'f' || *pch == L'F' || *pch == L'l' || *pch == L'L'))
            pch++;
    }
    else
    {
        /* 'ULL' */
        if (&pch[2] < pchEnd && (pch[0] == L'u' || pch[0] == L'U') && (pch[1] == L'l' || pch[1] == L'L') && (pch[2] == L'l' || pch[2] == L'L'))
            pch += 3;
        /* 'LLU' */
        else if (&pch[2] < pchEnd && (pch[0] == L'l' || pch[0] == L'L') && (pch[1] == L'l' || pch[1] == L'L') && (pch[2] == L'u' || pch[2] == L'U'))
            pch += 3;
        /* 'LL' */
        else if (&pch[1] < pchEnd && (pch[0] == L'l' || pch[0] == L'L') && (pch[1] == L'l' || pch[1] == L'L'))
            pch += 2;
        /* 'UL' */
        else if (&pch[1] < pchEnd && (pch[0] == L'u' || pch[0] == L'U') && (pch[1] == L'l' || pch[1] == L'L'))
            pch += 2;
        /* 'LU' */
        else if (&pch[1] < pchEnd && (pch[0] == L'l' || pch[0] == L'L') && (pch[1] == L'u' || pch[1] == L'U'))
            pch += 2;
        /* 'U' */
        else if (pch < pchEnd && (*pch == L'u' || *pch == L'U'))
            pch++;
        /* 'L' */
        else if (pch < pchEnd && (*pch == L'l' || *pch == L'L'))
            pch++;
    }

    return pch;
}

/****************************************************************************
 *                                                                          *
 * Function: Scanner                                                        *
 *                                                                          *
 * Purpose : Scan a C++ file for dependencies.                              *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define INCLUDESTRING  L"include"

static BOOL CALLBACK Scanner(LPCWSTR pcszFileName, BOOL (CALLBACK *pfnAddDepFile)(LPCWSTR pcszDepFileName, LPCVOID pvCookie), LPCVOID pvCookie)
{
    HANDLE hf;
    PWSTR pszInput;
    BOOL fComment = FALSE;
    ENCODING eEncoding;
    BOOL fOK = TRUE;

    // Open the file.
    hf = CreateFile(pcszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, 0);
    if (hf == INVALID_HANDLE_VALUE)
        return FALSE;

    // Read the file, line-by-line.
    for (pszInput = NULL; fOK && GetInputLine(hf, &pszInput, &eEncoding); )
    {
        PWSTR psz = pszInput;

restart:
        if (fComment)
        {
            // In traditional comment, check for terminator.
            while (*psz != L'\0' && (*psz != L'*' || psz[1] != L'/'))
                psz++;

            if (*psz == L'\0')
                continue;

            psz += 2;
            fComment = FALSE;
        }

        psz = SkipWhiteSpace(psz);

        // Check for traditional comment.
        if (psz[0] == L'/' && psz[1] == L'*')
        {
            psz += 2;
            fComment = TRUE;
            goto restart;
        }

        // Parse preprocessor lines.
        if (*psz++ == L'#')
        {
            psz = SkipWhiteSpace(psz);

            // Look for #include preprocessor line.
            if (wcsncmp(psz, INCLUDESTRING, wcslen(INCLUDESTRING)) == 0)
            {
                psz += wcslen(INCLUDESTRING);

                psz = SkipWhiteSpace(psz);

                if (*psz++ == L'"')  /* ignore "standard places" (#include <name>) */
                {
                    PWSTR pszName;

                    // Scan for end of "name".
                    for (pszName = psz; *psz != L'\0' && *psz != L'"'; psz++)
                        ;

                    if (*psz == L'"')
                    {
                        PWSTR pszDepFileName;

                        // Insert sneaky terminator.
                        *psz = L'\0';

                        // Search for the included file.
                        pszDepFileName = SearchIncludeFile(NoUnixSlash(pszName), pcszFileName);
                        if (pszDepFileName != NULL)
                        {
                            // Found included file: add as new dependency... */
                            if (!pfnAddDepFile(pszDepFileName, pvCookie))
                                fOK = FALSE;
                            // ...and then recurse, to scan the new file.
                            else if (!Scanner(pszDepFileName, pfnAddDepFile, pvCookie))
                                fOK = FALSE;
                        }

                        // Clean up.
                        free(pszDepFileName);
                    }
                }
            }
        }
    }

    free(pszInput);
    CloseHandle(hf);

    return fOK;
}

#undef INCLUDESTRING

/****************************************************************************
 *                                                                          *
 * Function: GetInputLine                                                   *
 *                                                                          *
 * Purpose : Read a line of text into a dynamically allocated buffer.       *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL GetInputLine(HANDLE hf, PWSTR *ppszInput, ENCODING *peEncoding)
{
    // First call for this file?
    if (!*ppszInput)
    {
        // Allocate a line buffer.
        *ppszInput = malloc(CCHMAXLINE * sizeof(WCHAR));
        if (!*ppszInput) return FALSE;

        // Determine the file encoding.
        *peEncoding = ReadTextFileEncoding(hf);
    }

    // Fill the line buffer.
    return ReadTextFileLine(hf, *peEncoding, *ppszInput, CCHMAXLINE);
}

/****************************************************************************
 *                                                                          *
 * Function: ReadTextFileEncoding                                           *
 *                                                                          *
 * Purpose : Check for byte-order mark at the beginning of the file.        *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static ENCODING ReadTextFileEncoding(HANDLE hf)
{
    ENCODING eEncoding;
    DWORD bom;
    DWORD cbBom;

    bom = 0;
    if (!ReadFile(hf, &bom, sizeof(bom), &cbBom, NULL))
        return ENCODING_UNKNOWN;

    if (SetFilePointer(hf, GetTextFileEncoding(bom, &eEncoding), NULL, FILE_BEGIN) == INVALID_SET_FILE_POINTER)
        return ENCODING_UNKNOWN;

    return eEncoding;
}

/****************************************************************************
 *                                                                          *
 * Function: GetTextFileEncoding                                            *
 *                                                                          *
 * Purpose : Read byte-order mark from the given dword, return length.      *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define BYTEMASK(x,n)  ((x) & ~((~0UL)<<((n)*8)))

static DWORD GetTextFileEncoding(DWORD bom, ENCODING *peEncoding)
{
    // Start with the longest byte-order mark (BOM).
    if (BYTEMASK(bom,3) == BOM_UTF8)
    {
        // UTF-8.
        *peEncoding = ENCODING_UTF8;
        return 3;
    }
    else if (BYTEMASK(bom,2) == BOM_UTF16_LE)
    {
        // UTF-16/UCS-2 little endian.
        *peEncoding = ENCODING_UTF16_LE;
        return 2;
    }

    /* ASCII/ANSI */
    *peEncoding = ENCODING_ANSI;
    return 0;
}

#undef BYTEMASK

/****************************************************************************
 *                                                                          *
 * Function: ReadTextFileLine                                               *
 *                                                                          *
 * Purpose : Read a line from a text file using given encoding.             *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define RAWBUFSIZE  8192

static BOOL ReadTextFileLine(HANDLE hf, ENCODING eEncoding, PWSTR pchBuf, DWORD cchBufMax)
{
    /* Remember the current file position */
    DWORD dwCurrentOffset = TellFile(hf);
    DWORD cbRead;
    PVOID pvRaw;

    if (cchBufMax == 0)
        return FALSE;

    // Allocate a work buffer.
    if ((pvRaw = malloc(RAWBUFSIZE)) == NULL)
        return FALSE;

    __try
    {
        // Read a chunk from the file.
        if (!ReadFile(hf, pvRaw, RAWBUFSIZE, &cbRead, NULL))
            return FALSE;

        if (cbRead == 0)
            return FALSE;

        // Check for CTRL+Z from last read.
        if (*((PCHAR)pvRaw) == CHAR_EOF)
            return FALSE;

        // Handle UTF-16LE encoding.
        if (eEncoding == ENCODING_UTF16_LE)
        {
#define pwchRaw  ((PWCHAR)pvRaw)
            PWSTR pwch, pwchEnd;

            for (pwch = pwchRaw, pwchEnd = pwchRaw + cbRead/sizeof(WCHAR); pwch < pwchEnd; pwch++)
            {
                if (*pwch == L'\r' || *pwch == L'\n' || *pwch == CHAR_EOF)
                    break;

                if (--cchBufMax == 0)
                    return FALSE;

                *pchBuf++ = *pwch;
            }
            *pchBuf = L'\0';

            if (pwch == pwchEnd && cbRead == RAWBUFSIZE)
                return FALSE;

            if (pwch == pwchEnd)
                ;
            else if (*pwch == L'\n')
                ++pwch;
            else if (*pwch == L'\r' && *++pwch == L'\n')
                ++pwch;

            // Start from the correct file position next time.
            return SeekFile(hf, dwCurrentOffset + (DWORD)((pwch - pwchRaw) * sizeof(WCHAR)), FILE_BEGIN);
#undef pwchRaw
        }
        // Handle plain text encoding.
        else if (eEncoding == ENCODING_ANSI)
        {
#define pchRaw  ((char *)pvRaw)
            PCHAR pch, pchEnd;
            DWORD cch = 0;

            for (pch = pchRaw, pchEnd = pchRaw + cbRead; pch < pchEnd; pch++)
            {
                if (*pch == '\r' || *pch == '\n' || *pch == CHAR_EOF)
                    break;
            }

            if (pch == pchEnd && cbRead == RAWBUFSIZE)
                return FALSE;

            if ((DWORD)(pch - pchRaw) >= cchBufMax)
                return FALSE;

            if (pch > pchRaw && (cch = MultiByteToWideChar(CP_ACP,
                0, pchRaw, (int)(pch - pchRaw), pchBuf, cchBufMax-1)) == 0)
                return FALSE;

            pchBuf[cch] = L'\0';

            if (pch == pchEnd)
                ;
            else if (*pch == '\n')
                ++pch;
            else if (*pch == '\r' && *++pch == '\n')
                ++pch;

            // Start from the correct file position next time.
            return SeekFile(hf, dwCurrentOffset + (DWORD)(pch - pchRaw), FILE_BEGIN);
#undef pchRaw
        }
        // Handle UTF-8 encoding.
        else if (eEncoding == ENCODING_UTF8)
        {
#define pchRaw  ((char *)pvRaw)
            PCHAR pch, pchEnd;
            DWORD cch = 0;

            for (pch = pchRaw, pchEnd = pchRaw + cbRead; pch < pchEnd; pch++)
            {
                if (*pch == '\r' || *pch == '\n' || *pch == CHAR_EOF)
                    break;
            }

            if (pch == pchEnd && cbRead == RAWBUFSIZE)
                return FALSE;

            if (pch > pchRaw && (cch = MultiByteToWideChar(CP_UTF8, 0, pchRaw, (int)(pch - pchRaw), pchBuf, cchBufMax-1)) == 0)
                return FALSE;

            pchBuf[cch] = L'\0';

            if (pch == pchEnd)
                ;
            else if (*pch == '\n')
                ++pch;
            else if (*pch == '\r' && *++pch == '\n')
                ++pch;

            // Start from the correct file position next time.
            return SeekFile(hf, dwCurrentOffset + (DWORD)(pch - pchRaw), FILE_BEGIN);
#undef pchRaw
        }
        // Handle bad dog.
        else return FALSE;
    }
    __finally
    {
        free(pvRaw);
    }
    return TRUE;
}

#undef RAWBUFSIZE

/****************************************************************************
 *                                                                          *
 * Function: NoUnixSlash                                                    *
 *                                                                          *
 * Purpose : Normalize a pathname by changing any Unix-like slashes.        *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static PWSTR NoUnixSlash(PWSTR pszFileName)
{
    for (PWSTR psz = pszFileName; *psz != L'\0'; psz++)
        if (*psz == L'/') *psz = L'\\';

    return pszFileName;
}

/****************************************************************************
 *                                                                          *
 * Function: SearchIncludeFile                                              *
 *                                                                          *
 * Purpose : Get fully qualified name of the given include file.            *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static PWSTR SearchIncludeFile(PCWSTR pcszIncludeName, PCWSTR pcszRefFileName)
{
    WCHAR szPath[MAX_PATH];
    WCHAR szFileName[MAX_PATH];

    // Use path from the file containing the #include directive.
    lstrcpyn(szPath, pcszRefFileName, NELEMS(szPath));
    PathRemoveFileSpec(szPath);

    // Combine path with the included name.
    PathCombine(szFileName, szPath, pcszIncludeName);

    // Do we have a winner?
    if (IsExistingFile(szFileName))
        return _wcsdup(szFileName);

    // Apparently not.
    return NULL;
}

/****************************************************************************
 *                                                                          *
 * Function: IsExistingFile                                                 *
 *                                                                          *
 * Purpose : Return TRUE if the given file exists.                          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL IsExistingFile(PCWSTR pcszFileName)
{
    WIN32_FIND_DATA wfd;
    HANDLE hff;

    hff = FindFirstFile(pcszFileName, &wfd);
    if (hff == INVALID_HANDLE_VALUE)
        return FALSE;

    FindClose(hff);
    return TRUE;
}
