/****************************************************************************
 *                                                                          *
 * File    : NewGUID.c                                                      *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates generating new GUID.           *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <objbase.h>
#include <addin.h>
#include <wchar.h>
#include "newguid.h"

#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))

/* Private command identifiers */
#define ID_GUID_DEFINE  1
#define ID_GUID_STRUCT  2

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
            AddCmd.pszText = L"New DEFINE_GUID";
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_ICON1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_GUID_DEFINE;
            AddCmd.idMenu = AIM_MENU_SOURCE;
            if (!AddIn_AddCommand(hwnd, &AddCmd))
                return FALSE;

            /* Add command to source menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = L"New GUID struct";
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_ICON1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_GUID_STRUCT;
            AddCmd.idMenu = AIM_MENU_SOURCE;
            if (!AddIn_AddCommand(hwnd, &AddCmd))
                return FALSE;

            return TRUE;
        }

        case AIE_APP_DESTROY:
            /* Forget main window handle */
            g_hwndMain = NULL;
            /* Remove from file menu */
            return AddIn_RemoveCommand(hwnd, ID_GUID_DEFINE) &&
                   AddIn_RemoveCommand(hwnd, ID_GUID_STRUCT);

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
    if (idCmd == ID_GUID_DEFINE || idCmd == ID_GUID_STRUCT)
    {
        GUID guid;

        /* Let Windows generate a new UUID */
        if (FAILED(CoCreateGuid(&guid)))
        {
            /* Just as I expected - a complete failure... */
            MessageBox(g_hwndMain,
                L"Unable to create a new GUID!",
                L"NewGUID Add-In",
                MB_OK|MB_ICONSTOP);
            return;
        }

        /* Generate default formatting (comment) */
        WCHAR szGuid[64];
        swprintf(szGuid, NELEMS(szGuid), L"%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X",
            guid.Data1, guid.Data2, guid.Data3,
            guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
            guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);

        /* Generate info in selected format */
        WCHAR szBuffer[256];
        if (idCmd == ID_GUID_DEFINE)
        {
            swprintf(szBuffer, NELEMS(szBuffer),
                L"\n/* {%ls} */\n"
                L"DEFINE_GUID(%ls,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x);\n",
                szGuid,
                L"NONAME",
                guid.Data1, guid.Data2, guid.Data3, guid.Data4[0],
                guid.Data4[1], guid.Data4[2], guid.Data4[3], guid.Data4[4],
                guid.Data4[5], guid.Data4[6], guid.Data4[7]);
        }
        else /* ID_GUID_STRUCT */
        {
            swprintf(szBuffer, NELEMS(szBuffer),
                L"\n/* {%ls} */\n"
                L"const GUID %ls = {0x%x,0x%x,0x%x,{0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x,0x%x}};\n",
                szGuid,
                L"NONAME",
                guid.Data1, guid.Data2, guid.Data3, guid.Data4[0],
                guid.Data4[1], guid.Data4[2], guid.Data4[3], guid.Data4[4],
                guid.Data4[5], guid.Data4[6], guid.Data4[7]);
        }

        /* Get handle of the current document window - must be a source window */
        HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
        if (IsWindow(hwndDoc))
        {
            ADDIN_RANGE Range;

            /* Get the current selection */
            if (AddIn_GetSourceSel(hwndDoc, &Range))
            {
                /* Set the current selection (start=end, i.e. nothing selected) */
                Range.iStartPos = Range.iEndPos;
                if (AddIn_SetSourceSel(hwndDoc, &Range))
                {
                    /* Replace current selection (nothing), i.e. insert new text at caret position */
                    AddIn_ReplaceSourceSelText(hwndDoc, szBuffer);
                }
            }
        }
    }
}
