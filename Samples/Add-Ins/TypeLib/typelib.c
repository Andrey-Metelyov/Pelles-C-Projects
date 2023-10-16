/****************************************************************************
 *                                                                          *
 * File    : TypeLib.c                                                      *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates COM Type Library info.         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <commdlg.h>
#include <addin.h>
#include <wchar.h>
#include <ole2.h>
#include "typelib.h"

/* Private command ID */
#define ID_TYPELIB  1

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

            /* Save handle of main IDE window */
            g_hwndMain = hwnd;

            /* Load menu text from the resources */
            LoadString(g_hmod, IDS_MENUTEXT, szText, NELEMS(szText));

            /* Add command to file menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = szText;
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_ICON1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_TYPELIB;
            AddCmd.idMenu = AIM_MENU_FILE;  /* File menu */
            return AddIn_AddCommand(hwnd, &AddCmd);
        }

        case AIE_APP_DESTROY:
            /* Remove from file menu */
            return AddIn_RemoveCommand(hwnd, ID_TYPELIB);

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

ADDINAPI void WINAPI AddInCommandEx(int idCmd, LPCVOID pcvData)
{
    switch (idCmd)
    {
        case ID_TYPELIB:
        {
            /*
             * We need a filename. Why not ask the user?
             */
            OPENFILENAME ofn = {0};
            WCHAR szFilename[MAX_PATH];
            WCHAR szFilter[256];
            WCHAR szSystemPath[MAX_PATH];

            /* Load file filter string from the resources */
            LoadString(g_hmod, IDS_FILEFILTER, szFilter, NELEMS(szFilter));
            for (PWSTR psz = szFilter; *psz != L'\0'; psz++)
            {
                /* Replace 'magic' character with nul character */
                if (*psz == L'|') *psz = L'\0';
            }

            szFilename[0] = L'\0';

            /* Get location of the system folder */
            if (!GetSystemDirectory(szSystemPath, NELEMS(szSystemPath)))
                szSystemPath[0] = L'\0';

            /* Ask for a filename */
            ofn.lStructSize = sizeof(ofn);
            ofn.hwndOwner = g_hwndMain;
            ofn.lpstrFilter = szFilter;
            ofn.nFilterIndex = 1;
            ofn.lpstrFile = szFilename;
            ofn.nMaxFile = NELEMS(szFilename);
            ofn.lpstrInitialDir = szSystemPath;
            ofn.Flags = OFN_FILEMUSTEXIST|OFN_HIDEREADONLY|OFN_EXPLORER;
            if (GetOpenFileName(&ofn))
            {
                /*
                 * Bummer. The user managed to find a file.
                 * Pretend to be clever and actually do something.
                 */
                CoInitialize(0);
                PWSTR pszResult = DumpTypeLib(szFilename);
                CoUninitialize();

                /* Create a new source window in the IDE */
                HWND hwnd = AddIn_NewDocument(g_hwndMain, AID_SOURCE);
                if (hwnd != NULL)
                {
                    /* Set the text into the source window */
                    AddIn_SetSourceText(hwnd, pszResult != NULL ? pszResult : L"");
                }

                free(pszResult);
            }
            break;
        }
    }
}
