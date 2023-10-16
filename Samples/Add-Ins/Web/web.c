/****************************************************************************
 *                                                                          *
 * File    : web.c                                                          *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates web browser interface.         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE /* for Windows API */
#define _UNICODE    /* for C runtime */
#include <windows.h>
#include <windowsx.h>
#include <tchar.h>
#include <addin.h>
#include "web.h"

#define NELEMS(a)  (sizeof(a) / sizeof(a[0]))

// Private command IDs.
enum {
    ID_GO_SMORGASBORDET,
    ID_GO_MICROSOFT
};

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;

// Prototypes.
static BOOL OnAppCreate(HWND);
static BOOL OnAppDestroy(HWND);
static BOOL OnDocNavigate(HWND);

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
            return OnAppCreate(hwnd);

        case AIE_APP_DESTROY:
            return OnAppDestroy(hwnd);

        case AIE_DOC_NAVIGATE:
            return OnDocNavigate(hwnd);

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: OnAppCreate                                                    *
 *                                                                          *
 * Purpose : Handle the AIE_APP_CREATE event.                               *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL OnAppCreate(HWND hwndMain)
{
    /* Save handle of IDE main window */
    g_hwndMain = hwndMain;

    ADDIN_ADD_COMMAND AddCmd = {0};

    /* Add choice #1 to web menu */
    AddCmd.cbSize = sizeof(AddCmd);
    AddCmd.pszText = L"Go to www.smorgasbordet.com";
    AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_APP_ICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
    AddCmd.id = ID_GO_SMORGASBORDET;
    AddCmd.idMenu = AIM_MENU_WEB;
    if (!AddIn_AddCommand(hwndMain, &AddCmd))
        return FALSE;

    /* Add choice #2 to web menu */
    AddCmd.cbSize = sizeof(AddCmd);
    AddCmd.pszText = L"Go to www.microsoft.com";
    AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_APP_ICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
    AddCmd.id = ID_GO_MICROSOFT;
    AddCmd.idMenu = AIM_MENU_WEB;
    if (!AddIn_AddCommand(hwndMain, &AddCmd))
        return FALSE;

    return TRUE;
}

/****************************************************************************
 *                                                                          *
 * Function: OnAppDestroy                                                   *
 *                                                                          *
 * Purpose : Handle the AIE_APP_DESTROY event.                              *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL OnAppDestroy(HWND hwndMain)
{
    /* Remove choice #1 and #2 from web menu */
    return AddIn_RemoveCommand(hwndMain, ID_GO_SMORGASBORDET) &&
           AddIn_RemoveCommand(hwndMain, ID_GO_MICROSOFT);
}

/****************************************************************************
 *                                                                          *
 * Function: OnDocNavigate                                                  *
 *                                                                          *
 * Purpose : Handle the AIE_DOC_NAVIGATE event.                             *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL OnDocNavigate(HWND hwndDoc)
{
    ADDIN_WEB_INFO WebInfo = { .cbSize = sizeof(WebInfo) };

    if (AddIn_GetWebInfo(hwndDoc, &WebInfo))
    {
        AddIn_WriteOutput(g_hwndMain, L"Navigating to a new URL:");
        AddIn_WriteOutput(g_hwndMain, WebInfo.pszURL);
    }
    else
    {
        AddIn_WriteOutput(g_hwndMain, L"Failed to get web document information. Bummer.");
    }

    return TRUE;
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
        case ID_GO_SMORGASBORDET:
        {
            /* Get window handle of the active MDI document */
            HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
            if (IsWindow(hwndDoc))
            {
                ADDIN_WEB_INFO WebInfo = {0};

                WebInfo.cbSize = sizeof(WebInfo);
                WebInfo.eEncryption = ADDIN_WEB_ENCRYPT_UNKNOWN;
                WebInfo.pszURL = L"http://www.smorgasbordet.com/pellesc";
                WebInfo.fWelcome = FALSE;
                AddIn_SetWebInfo(hwndDoc, &WebInfo);
            }
            break;
        }

        case ID_GO_MICROSOFT:
        {
            /* Get window handle of the active MDI document */
            HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
            if (IsWindow(hwndDoc))
            {
                ADDIN_WEB_INFO WebInfo = {0};

                WebInfo.cbSize = sizeof(WebInfo);
                WebInfo.eEncryption = ADDIN_WEB_ENCRYPT_UNKNOWN;
                WebInfo.pszURL = L"http://www.microsoft.com";
                WebInfo.fWelcome = FALSE;
                AddIn_SetWebInfo(hwndDoc, &WebInfo);
            }
            break;
        }
    }
}
