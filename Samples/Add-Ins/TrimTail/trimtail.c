/****************************************************************************
 *                                                                          *
 * File    : trimtail.c                                                     *
 *                                                                          *
 * Purpose : Sample add-in DLL. Trim white-space at end of lines.           *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <addin.h>
#include <stdlib.h>
#include <stdbool.h>
#include "trimtail.h"

#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))

/* Private command identifiers */
#define ID_TRIMTAIL  1

/* Locals */
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;

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
            /* Remember module handle */
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
            ADDIN_ADD_COMMAND AddCmd = {0};
            WCHAR szText[80];

            /* Remember main window handle */
            g_hwndMain = hwnd;

            LoadString(g_hmod, IDS_MENUTEXT, szText, NELEMS(szText));

            /* Add command to source menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = szText;
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_ICON1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_TRIMTAIL;
            AddCmd.idMenu = AIM_MENU_SOURCE_CONVERT;
            return AddIn_AddCommand(hwnd, &AddCmd);
        }

        case AIE_APP_DESTROY:
            /* Forget main window handle */
            g_hwndMain = NULL;
            /* Remove from file menu */
            return AddIn_RemoveCommand(hwnd, ID_TRIMTAIL);

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: AddInCommandEx                                                 *
 *                                                                          *
 * Purpose : Add-in command handler.                                        *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define MAX_LINE_LENGTH  4096

ADDINAPI void WINAPI AddInCommandEx(int idCmd, LPCVOID pcvData)
{
    if (idCmd == ID_TRIMTAIL)
    {
        /* Get handle of the current document window (must be a source window!) */
        HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
        if (IsWindow(hwndDoc))
        {
            /* Allocate a reasonably large line buffer */
            WCHAR *pchLine = malloc(MAX_LINE_LENGTH * sizeof(WCHAR));
            if (pchLine != NULL)
            {
                ADDIN_RANGE Range;
                int iFirstLine;
                int iLastLine;

                /* Either we have a selection or we don't (no third option) */
                if (AddIn_GetSourceSel(hwndDoc, &Range) && Range.iStartPos != Range.iEndPos)
                {
                    iFirstLine = AddIn_SourceLineFromChar(hwndDoc, Range.iStartPos);
                    iLastLine = AddIn_SourceLineFromChar(hwndDoc, Range.iEndPos);
                }
                else /* no selection */
                {
                    iFirstLine = 0;
                    iLastLine = AddIn_GetSourceLineCount(hwndDoc) - 1;
                }

                /* Trim white-space at the end of affected lines */
                for (int iLine = iFirstLine; iLine <= iLastLine; iLine++)
                {
                    unsigned int cchLine = AddIn_GetSourceLineLength(hwndDoc, iLine);
                    if (cchLine > 0 && cchLine <= MAX_LINE_LENGTH)
                    {
                        /* Get text for current line, and sanity check the length */
                        if (cchLine == AddIn_GetSourceLine(hwndDoc, iLine, pchLine, MAX_LINE_LENGTH))
                        {
                            WCHAR *pch = &pchLine[cchLine];

                            /* Boycott any incomplete last line */
                            if (*--pch == L'\n')
                            {
                                bool fBother = false;

                                /* Trim white-space */
                                while (pch > pchLine && (pch[-1] == L' ' || pch[-1] == L'\t'))
                                    --pch, fBother = true;
                                *pch = L'\0';

                                /* Only bother updating when we need to */
                                if (fBother && !AddIn_SetSourceLine(hwndDoc, iLine, pchLine))
                                    break;  /* oops! */
                            }
                        }
                    }
                }

                /* Done with line buffer */
                free(pchLine);
            }
        }
    }
}
