/****************************************************************************
 *                                                                          *
 * File    : newrsrc.c                                                      *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates resource manipulation.         *
 *                                                                          *
 * History : Date        Reason                                             *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

/*
 * This sample demonstrates adding new resources to the IDE,
 * and also how to enumerate through all resources.
 */

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <addin.h>
#include <wchar.h>
#include "newrsrc.h"

#define NELEMS(a)  (sizeof(a) / sizeof(a[0]))

// Private command ID - any number will do.
#define ID_NEWICON  1
#define ID_NEWDIALOG  2
#define ID_ENUMLIST  3

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;
static BOOL EnumCallback(const ADDIN_RESOURCE * const, LPVOID);

/****************************************************************************
 *                                                                          *
 * Function: DllMain                                                        *
 *                                                                          *
 * Purpose : DLL entry and exit procedure.                                  *
 *                                                                          *
 * History : Date        Reason                                             *
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
 * History : Date        Reason                                             *
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

            /* Save handle of main IDE window */
            g_hwndMain = hwnd;

            /* Add choice to New Resource menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = L"New icon sample";
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_APPICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_NEWICON;
            AddCmd.idMenu = AIM_MENU_RESOURCE_NEW;
            if (!AddIn_AddCommand(hwnd, &AddCmd))
                return FALSE;

            /* Add choice to New Resource menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = L"New dialog sample";
            AddCmd.hIcon = LoadImage(g_hmod, MAKEINTRESOURCE(IDR_APPICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR|LR_SHARED);
            AddCmd.id = ID_NEWDIALOG;
            AddCmd.idMenu = AIM_MENU_RESOURCE_NEW;
            if (!AddIn_AddCommand(hwnd, &AddCmd))
                return FALSE;

            /* Add choice to Resource menu */
            AddCmd.cbSize = sizeof(AddCmd);
            AddCmd.pszText = L"Enumerate resources";
            AddCmd.hIcon = NULL;
            AddCmd.id = ID_ENUMLIST;
            AddCmd.idMenu = AIM_MENU_RESOURCE;
            if (!AddIn_AddCommand(hwnd, &AddCmd))
                return FALSE;

            return TRUE;
        }

        case AIE_APP_DESTROY:
            /* Remove choice from resource menu */
            return AddIn_RemoveCommand(hwnd, ID_NEWICON) &&
                AddIn_RemoveCommand(hwnd, ID_NEWDIALOG) &&
                AddIn_RemoveCommand(hwnd, ID_ENUMLIST);

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
        case ID_NEWICON:
        {
            /* Get window handle of the active MDI document */
            HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
            if (IsWindow(hwndDoc))
            {
                HRSRC hrsrc;
                HGLOBAL hgRes;
                PVOID pvRes;

                // Find the icon template resource. Note! The IDE wants an icon file for
                // this resource type, so it's stored as RCDATA, not an ICON type.
                if ((hrsrc = FindResource(g_hmod, MAKEINTRESOURCE(IDR_NEWRSRC), RT_RCDATA)) != NULL &&
                    (hgRes = LoadResource(g_hmod, hrsrc)) != NULL &&
                    (pvRes = LockResource(hgRes)) != NULL)
                {
                    ADDIN_RESOURCE Resource;

                    // Add a new icon to the current resource window. For simplicity,
                    // we always call the new resource "NewIcon". This should be changed.
                    Resource.cbSize = sizeof(Resource);
                    Resource.pwchType = MakeResourceOrdinalW(ADDIN_RSRC_ICON);
                    Resource.pwchName = L"NewIcon";
                    Resource.pvData = pvRes;
                    Resource.cbData = SizeofResource(g_hmod, hrsrc);
                    Resource.wLanguageId = 0;  // Use default language.
                    if (!AddIn_AddResource(hwndDoc, &Resource))
                        MessageBox(NULL, L"Problem adding a new icon!", L"NewRsrc Add-In", MB_OK|MB_ICONSTOP);
                }
            }
            break;
        }

        case ID_NEWDIALOG:
        {
            /* Get window handle of the active MDI document */
            HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
            if (IsWindow(hwndDoc))
            {
                HRSRC hrsrc;
                HGLOBAL hgRes;
                PVOID pvRes;

                // Find the dialog template resource.
                if ((hrsrc = FindResource(g_hmod, MAKEINTRESOURCE(IDR_NEWDIALOG), RT_DIALOG)) != NULL &&
                    (hgRes = LoadResource(g_hmod, hrsrc)) != NULL &&
                    (pvRes = LockResource(hgRes)) != NULL)
                {
                    ADDIN_RESOURCE Resource;

                    // Add a new dialog to the current resource window. For simplicity, 
                    // we always call the new resource "NewDlg". This should be changed.
                    Resource.cbSize = sizeof(Resource);
                    Resource.pwchType = MakeResourceOrdinalW(ADDIN_RSRC_DIALOG);
                    Resource.pwchName = L"NewDlg";
                    Resource.pvData = pvRes;
                    Resource.cbData = SizeofResource(g_hmod, hrsrc);
                    Resource.wLanguageId = MAKELANGID(LANG_ENGLISH,SUBLANG_ENGLISH_US);  // Always U.S.
                    if (!AddIn_AddResource(hwndDoc, &Resource))
                        MessageBox(NULL, L"Problem adding a new dialog!", L"NewRsrc Add-In", MB_OK|MB_ICONSTOP);
                }
            }
            break;
        }

        case ID_ENUMLIST:
        {
            /* Get window handle of the active MDI document */
            HWND hwndDoc = AddIn_GetActiveDocument(g_hwndMain);
            if (IsWindow(hwndDoc))
            {
                // Enumerate all resources in the current document.
                ADDIN_ENUM_RESOURCES Enum;
                Enum.cbSize = sizeof(Enum);
                Enum.pfnCallback = EnumCallback;
                Enum.pvData = 0;
                AddIn_EnumResources(hwndDoc, &Enum);
            }
        }
    }
}

/****************************************************************************
 *                                                                          *
 * Function: EnumCallback                                                   *
 *                                                                          *
 * Purpose : Callback function for resource enumeration.                    *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL EnumCallback(const ADDIN_RESOURCE * const pResource, LPVOID pvData)
{
    TCHAR szName[80];
    TCHAR szType[80];
    TCHAR szBuf[256];

    // Build name of resource - either a string or an ordinal.
    if (IsResourceOrdinalW(pResource->pwchName))
        swprintf(szName, NELEMS(szName), L"#%u", GetResourceOrdinalW(pResource->pwchName));
    else
        swprintf(szName, NELEMS(szName), L"\"%ls\"", pResource->pwchName);

    // Build type of resource - either a string or an ordinal.
    if (IsResourceOrdinalW(pResource->pwchType))
    {
        // Check for predefined resource ordinals.
        switch (GetResourceOrdinalW(pResource->pwchType))
        {
            case ADDIN_RSRC_CURSOR: wcscpy(szType, L"Cursor"); break;
            case ADDIN_RSRC_BITMAP: wcscpy(szType, L"Bitmap"); break;
            case ADDIN_RSRC_ICON: wcscpy(szType, L"Icon"); break;
            case ADDIN_RSRC_MENU: wcscpy(szType, L"Menu"); break;
            case ADDIN_RSRC_DIALOG: wcscpy(szType, L"Dialog"); break;
            case ADDIN_RSRC_STRING:
            {
                // String tables are special...
                int id = (GetResourceOrdinalW(pResource->pwchName)-1) * 16;
                swprintf(szType, NELEMS(szType), L"String table (id %u - %u)", id, id + 15);
                szName[0] = L'\0';
                break;
            }
            case ADDIN_RSRC_ACCELERATOR: wcscpy(szType, L"Accelerator table"); break;
            case ADDIN_RSRC_RCDATA: wcscpy(szType, L"RCDATA"); break;
            case ADDIN_RSRC_MESSAGETABLE: wcscpy(szType, L"Message table"); break;
            case ADDIN_RSRC_VERSION: wcscpy(szType, L"Version"); break;
            case ADDIN_RSRC_ANICURSOR: wcscpy(szType, L"Animated cursor"); break;
            case ADDIN_RSRC_ANIICON: wcscpy(szType, L"Animated icon"); break;
            case ADDIN_RSRC_HTML: wcscpy(szType, L"HTML"); break;
            case ADDIN_RSRC_MANIFEST: wcscpy(szType, L"Manifest"); break;
            default: swprintf(szType, NELEMS(szType), L"#%u", GetResourceOrdinalW(pResource->pwchType));  /* huh? */
        }
    }
    else
    {
        // String type.
        swprintf(szType, NELEMS(szType), L"\"%ls\"", pResource->pwchType);
    }

    // Build output string and send to output window.
    swprintf(szBuf, NELEMS(szBuf), L"%ls %ls (language %X; default language %X)",
        szType, szName, pResource->wLanguageId, GetUserDefaultLangID());
    AddIn_WriteOutput(g_hwndMain, szBuf);

    // "Keep feeding me resources"
    return TRUE;
}
