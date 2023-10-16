/****************************************************************************
 *                                                                          *
 * File    : MakeRGB.c                                                      *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates RGB() macro helper.            *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <addin.h>
#include <commdlg.h>
#include <wchar.h>
#include "makergb.h"

#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))

/* Private command identifiers */
#define ID_RGBMAKER  1

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

            /* Remember main window handle */
            g_hwndMain = hwnd;

            /* Add command to source menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = L"Make RGB Color...";
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_ICON1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_RGBMAKER;
            AddCmd.idMenu = AIM_MENU_SOURCE;
            return AddIn_AddCommand(hwnd, &AddCmd);
        }

        case AIE_APP_DESTROY:
            /* Forget main window handle */
            g_hwndMain = NULL;
            /* Remove from file menu */
            return AddIn_RemoveCommand(hwnd, ID_RGBMAKER);

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

#define RGB_MACRO_FORMAT  L"RGB(%u,%u,%u)"

ADDINAPI void WINAPI AddInCommandEx(int idCmd, LPCVOID pcvData)
{
    if (idCmd == ID_RGBMAKER)
    {
        /* Get handle of the current document window (must be a source window) */
        HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
        if (IsWindow(hwndDoc))
        {
            static COLORREF aclrCustom[16] = {0};
            CHOOSECOLOR cc = {0};
            TCHAR szText[80];

            cc.lStructSize = sizeof(cc);
            cc.hwndOwner = g_hwndMain;
            cc.Flags = CC_ANYCOLOR|CC_FULLOPEN;
            cc.lpCustColors = aclrCustom;

            /* If the source window contains selected text, that looks like
             * a RGB macro, use it to initialize the color chooser dialog.
             */
            if (AddIn_GetSourceSelText(hwndDoc, szText, NELEMS(szText)) != 0)
            {
                UINT uRed, uGreen, uBlue;

                if (swscanf(szText, RGB_MACRO_FORMAT, &uRed, &uGreen, &uBlue) == 3)
                {
                    cc.rgbResult = RGB(uRed,uGreen,uBlue);
                    cc.Flags |= CC_RGBINIT;
                }
            }

            /* Time to display the color chooser dialog */
            if (ChooseColor(&cc))
            {
                ADDIN_RANGE Range;

                /* Format the selected color as a RGB macro */
                swprintf(szText, NELEMS(szText), RGB_MACRO_FORMAT,
                    GetRValue(cc.rgbResult),
                    GetGValue(cc.rgbResult),
                    GetBValue(cc.rgbResult));

                /* Replace the current selection, or insert new text */
                AddIn_ReplaceSourceSelText(hwndDoc, szText);

                /* Get the current selection */
                if (AddIn_GetSourceSel(hwndDoc, &Range))
                {
                    /* We have just added text, above, so the caret should be
                     * located right after the new text. Modify the starting
                     * position so that the text is selected.
                     */
                    Range.iStartPos -= wcslen(szText);
                    AddIn_SetSourceSel(hwndDoc, &Range);
                }
            }
        }
    }
}

#undef RGB_MACRO_FORMAT
