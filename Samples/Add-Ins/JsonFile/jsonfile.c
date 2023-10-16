/****************************************************************************
 *                                                                          *
 * File    : jsonfile.c                                                     *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates a JSON syntax parser.          *
 *                                                                          *
 * History : Date        Reason                                             *
 *           00/00/00   Created                                             *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <addin.h>
#include <wchar.h>
#include <wctype.h>

#define NELEMS(a)  (sizeof(a) / sizeof(a[0]))

// Flags for usCookie.
#define FLAG_STRING  0x01

// Limits for cFoldLevel.
#define MIN_FOLDLEVEL  0
#define MAX_FOLDLEVEL  255

// JSON keywords, sorted by wcsncmp().
static PCWSTR apcszKeywords[] = {
    L"false",
    L"null",
    L"true"
};

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

// Locals.
static HANDLE g_hmod = NULL;

// Function prototypes.
static USHORT Parser(USHORT, PCWSTR, int, ADDIN_PARSE_POINT [], PINT);
static size_t IsKeyword(PCWSTR [], size_t, PCWSTR, size_t);
static PCWSTR ParseNumber(PCWSTR, PCWSTR);

/****************************************************************************
 *                                                                          *
 * Function: DllMain                                                        *
 *                                                                          *
 * Purpose : DLL entry and exit procedure.                                  *
 *                                                                          *
 * History : Date        Reason                                             *
 *           00/00/00   Created                                             *
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
 * History : Date        Reason                                             *
 *           00/00/00   Created                                             *
 *                                                                          *
 ****************************************************************************/

ADDINAPI BOOL WINAPI AddInMain(HWND hwnd, ADDIN_EVENT eEvent)
{
    switch (eEvent)
    {
        case AIE_APP_CREATE:
        {
            ADDIN_ADD_FILE_TYPE AddFile = {0};

            /* Define a new file type in the IDE */
            AddFile.cbSize = sizeof(AddFile);
            AddFile.pszDescription = L"JSON file";
            AddFile.pszExtension = L"json";  /* support *.json files */
            AddFile.pfnParser = Parser;
            AddFile.pszShells = NULL;  /* not possible to build this type */
            AddFile.pfnScanner = NULL;  /* nothing to scan in this type */
            if (!AddIn_AddFileType(hwnd, &AddFile))
                return FALSE;

            return TRUE;
        }

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: Parser                                                         *
 *                                                                          *
 * Purpose : Parse source code - for syntax color highlighting.             *
 *                                                                          *
 * History : Date        Reason                                             *
 *           00/00/00   Created                                             *
 *                                                                          *
 ****************************************************************************/

/*
 * usCookie - (IN) saved state from previous line (zero for first line) - flags (1 byte), folding level (1 byte).
 * pszText  - (IN) pointer to text to parse.
 * cchText  - (IN) length of text to parse (can be zero, for empty lines!).
 * pPoints  - (OUT) array of ADDIN_PARSE_POINT that specifies color changes. Large enough for a color change
 *            on each character in the largest supported line length (4096 characters).
 * pcPoints - (OUT) number of entries stored in pPoints[].
 */
static USHORT Parser(USHORT usCookie, PCWSTR pchText, int cchText, ADDIN_PARSE_POINT pPoints[4096], PINT pcPoints)
{
    PCWSTR pch = pchText, pchEnd = &pchText[cchText];
    BYTE cFoldLevel = ADDIN_GET_COOKIE_LEVEL(usCookie);
    BYTE bFlags = ADDIN_GET_COOKIE_FLAGS(usCookie);

    if (pch == pchEnd)
        ;
    // Inside string.
    else if (bFlags & (FLAG_STRING))
    {
        DEFINE_BLOCK(pch, ADDIN_COLOR_STRING);
    }
    else /* normal state */
    {
        while (pch < pchEnd && (*pch == L' ' || *pch == L'\t'))
            pch++;
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

        // Check for string constant: "...."
        if (*pch == L'\"')
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_STRING);
            bFlags |= FLAG_STRING;
            pch++;
            continue;
        }

        // Update folding level.
        if ((*pch == L'[' || *pch == L'{') && cFoldLevel < MAX_FOLDLEVEL)
            cFoldLevel++;
        else if ((*pch == L']' || *pch == L'}') && cFoldLevel > MIN_FOLDLEVEL)
            cFoldLevel--;

        // Check for structural character.
        if (*pch == L'[' ||  /* begin-array */
            *pch == L']' ||  /* end-array */
            *pch == L'{' ||  /* begin-object */
            *pch == L'}' ||  /* end-object */
            *pch == L':' ||  /* name separator */
            *pch == L',')    /* value separator */
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_OPERATOR);
            pch++;
            continue;
        }

        // Check for number.
        if (*pch == L'-' || iswdigit(*pch))
        {
            DEFINE_BLOCK(pch, ADDIN_COLOR_NUMBER);
            pch = ParseNumber(pch, pchEnd);
            continue;
        }

        // Check for keyword.
        if (IsCharAlphaW(*pch))
        {
            PCWSTR pchIdent = pch;
            do
                pch++;
            while (pch < pchEnd && IsCharAlphaW(*pch));

            if (IsKeyword(apcszKeywords, NELEMS(apcszKeywords), pchIdent, (size_t)(pch - pchIdent)))
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

    return ADDIN_MAKE_COOKIE(usCookie, cFoldLevel);
}

/****************************************************************************
 *                                                                          *
 * Function: IsKeyword                                                      *
 *                                                                          *
 * Purpose : Return non-zero index, if the given string is a keyword.       *
 *                                                                          *
 * History : Date        Reason                                             *
 *           00/00/00   Created                                             *
 *                                                                          *
 ****************************************************************************/

static size_t IsKeyword(PCWSTR apcszKeywords[], size_t cKeywords, PCWSTR pchText, size_t cchText)
{
    PCWSTR *pcszKeywords = apcszKeywords;

    while (cKeywords > 0)
    {
        size_t pivot = cKeywords >> 1;
        int cond = wcsncmp(pcszKeywords[pivot], pchText, cchText);  /* case sensitive */
        if (cond > 0)
        {
            // Search below pivot.
            cKeywords = pivot;
            continue;
        }
        else if (cond < 0)
        {
            // Search above pivot.
            pcszKeywords += pivot + 1;
            cKeywords -= pivot + 1;
            continue;
        }
        else
        {
            // Got a match, but is it exact?
            if (pcszKeywords[pivot][cchText] == L'\0')
                return 1 + (pcszKeywords - apcszKeywords) + pivot;

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
 * History : Date        Reason                                             *
 *           00/00/00   Created                                             *
 *                                                                          *
 ****************************************************************************/

static PCWSTR ParseNumber(PCWSTR pch, PCWSTR pchEnd)
{
    if (*pch == L'-')
        ++pch;

    if (pch < pchEnd && *pch == L'0')
        ++pch;
    else while (pch < pchEnd && iswdigit(*pch))
        ++pch;

    if (pch < pchEnd && *pch == L'.')
    {
        ++pch;
        while (pch < pchEnd && iswdigit(*pch))
            ++pch;
    }

    if (pch < pchEnd && (*pch == L'e' || *pch == L'E'))
    {
        ++pch;
        if (pch < pchEnd && (*pch == L'-' || *pch == L'+'))
            ++pch;
        while (pch < pchEnd && iswdigit(*pch))
            ++pch;

    }

    return pch;
}
