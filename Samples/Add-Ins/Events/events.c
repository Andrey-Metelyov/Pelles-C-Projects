/****************************************************************************
 *                                                                          *
 * File    : events.c                                                       *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates IDE events.                    *
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
#include <wchar.h>

#define NELEMS(a)  (sizeof(a) / sizeof(a[0]))

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;

static void OutputFileInfo(HWND, PCWSTR);

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
            /* Save handle of main window */
            g_hwndMain = hwnd;
            return TRUE;

        case AIE_PRJ_CREATE:
            OutputFileInfo(hwnd, L"Project opened");
            return TRUE;

        case AIE_PRJ_SAVE:
            OutputFileInfo(hwnd, L"Project saved");
            return TRUE;

        case AIE_PRJ_STARTBUILD:
            AddIn_WriteOutput(g_hwndMain, L"Project build started");
            return TRUE;

        case AIE_PRJ_ENDBUILD:
            AddIn_WriteOutput(g_hwndMain, AddIn_DidProjectBuildFail(hwnd) > 0 ?
                L"Project build ended in complete failure" :
                L"Project build ended successfully");
            return TRUE;

        case AIE_PRJ_DESTROY:
            OutputFileInfo(hwnd, L"Project closed");
            return TRUE;

        case AIE_DOC_CREATE:
            OutputFileInfo(hwnd, L"Document opened");
            return TRUE;

        case AIE_DOC_SAVE:
            OutputFileInfo(hwnd, L"Document saved");
            return TRUE;

        case AIE_DOC_NAVIGATE:
            OutputFileInfo(hwnd, L"Document navigation");
            return TRUE;

        case AIE_DOC_DESTROY:
            OutputFileInfo(hwnd, L"Document closed");
            return TRUE;

        case AIE_PRF_START:
            OutputFileInfo(hwnd, L"Profiler starting");
            return TRUE;

        case AIE_PRF_END:
            OutputFileInfo(hwnd, L"Profiler ended");
            return TRUE;

        case AIE_DBG_START:
            OutputFileInfo(hwnd, L"Debugger starting");
            return TRUE;

        case AIE_DBG_BREAK:
            OutputFileInfo(hwnd, L"Debuggee paused");
            return TRUE;

        case AIE_DBG_END:
            OutputFileInfo(hwnd, L"Debugger ended");
            return TRUE;

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: OutputFileInfo                                                 *
 *                                                                          *
 * Purpose : Helper function - write event (and filename) to output tab.    *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void OutputFileInfo(HWND hwnd, PCWSTR pcszEvent)
{
    ADDIN_DOCUMENT_INFO DocInfo = {0};
    WCHAR achBuf[512];

    DocInfo.cbSize = sizeof(DocInfo);
    if (AddIn_GetDocumentInfo(hwnd, &DocInfo))
    {
        swprintf(achBuf, NELEMS(achBuf), L"%ls: %ls", pcszEvent, *DocInfo.szFilename ? DocInfo.szFilename : L"untitled");
        AddIn_WriteOutput(g_hwndMain, achBuf);
    }
    else /* no information available */
    {
        swprintf(achBuf, NELEMS(achBuf), L"%ls", pcszEvent);
        AddIn_WriteOutput(g_hwndMain, achBuf);
    }
}
