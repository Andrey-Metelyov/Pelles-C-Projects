/****************************************************************************
 *                                                                          *
 * File    : profdump.c                                                     *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates built-in profiler data.        *
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

#ifndef NELEMS
#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))
#endif /* NELEMS */

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;
static DWORD g_cTotalHits = 0;

// Static function prototypes.
static void DumpProfilerSrcLinesData(HWND);
static BOOL CALLBACK EnumFunctionsCallback(const ADDIN_PROFILER_FUNCTION_W * const, LPVOID);
static void DumpProfilerCallTreeData(HWND);
static BOOL CALLBACK EnumCallTreeCallback(const ADDIN_PROFILER_CALLTREE_W * const, LPVOID);

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
            /*
             * We need the handle of the main window later,
             * so save hwnd in global g_hwndMain now.
             */
            g_hwndMain = hwnd;
            return TRUE;

        case AIE_PRF_END:
        {
            /*
             * Keep it really simple in this example and emit
             * the data as soon as the profiler ends.
             * A better add-in could use AddIn_AddCommand(),
             * to add a new menu choice to the IDE.
             *
             * Pass on the handle of the profiler document
             * window to the worker function.
             */
            ADDIN_PROFILER_INFO ProfInfo = { .cbSize = sizeof(ProfInfo) };
            if (AddIn_GetProfilerInfo(g_hwndMain, &ProfInfo))
            {
                if (ProfInfo.eProfMode == ADDIN_PROFMODE_SRCLINES)
                    DumpProfilerSrcLinesData(hwnd);
                else /* ADDIN_PROFMODE_CALLTREE */
                    DumpProfilerCallTreeData(hwnd);
            }
            return TRUE;
        }

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: DumpProfilerSrcLinesData                                       *
 *                                                                          *
 * Purpose : Enumerate profiler data; write to the output tab.              *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpProfilerSrcLinesData(HWND hwndProf)
{
    ADDIN_ENUM_PROFILER_FUNCTIONS Enum = {0};

    /* Clear the output tab */
    AddIn_WriteOutput(g_hwndMain, NULL);

    /* Clear total hits */
    g_cTotalHits = 0;

    /* Start enumerating profiler data */
    Enum.cbSize = sizeof(Enum);                /* required size (version control) */
    Enum.pfnCallback = EnumFunctionsCallback;  /* pointer to callback function */
    if (AddIn_EnumProfilerFunctions(hwndProf, &Enum))
    {
        WCHAR szBuf[80];

        /* Enumeration was successful */
        if (swprintf(szBuf, NELEMS(szBuf), L"%u hit(s) in total", g_cTotalHits) > 0)
            AddIn_WriteOutput(g_hwndMain, szBuf);
    }
    else
    {
        /* Enumeration failed */
        AddIn_WriteOutput(g_hwndMain, L"Failed to complete enumeration!");
    }
}

/****************************************************************************
 *                                                                          *
 * Function: EnumFunctionsCallback                                          *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumProfilerFunctions().                    *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL CALLBACK EnumFunctionsCallback(const ADDIN_PROFILER_FUNCTION * const pFunction, LPVOID pvData)
{
    WCHAR szBuf[40+256+256];

    /* Remember total hits */
    g_cTotalHits = pFunction->cSessionHits;

    /* Process basic information */
    if (swprintf(szBuf, NELEMS(szBuf), L"%u hit(s) in function %.255ls, module %.255ls",
        pFunction->cFunctionHits, pFunction->pszFunctionName, pFunction->pszModuleName) > 0)
    {
        AddIn_WriteOutput(g_hwndMain, szBuf);

        /* Process any source lines */
        for (size_t i = 0; i < pFunction->cLines; i++)
        {
            const ADDIN_PROFILER_FUNCTION_LINE * const pLine = &pFunction->pLines[i];

            /*
             * Trick: if we start a line as "filename(lineno)", double-clicking on such
             *        a line in the output tab should open the file in the source editor...
             */
            if (swprintf(szBuf, NELEMS(szBuf), L"  %.255ls(%u): %u hit(s)",
                pFunction->pszSourceName, pLine->nLineNo, pLine->cHits) > 0)
            {
                AddIn_WriteOutput(g_hwndMain, szBuf);
            }
        }
    }

    return TRUE;
}

/****************************************************************************
 *                                                                          *
 * Function: DumpProfilerCallTreeData                                       *
 *                                                                          *
 * Purpose : Enumerate profiler data; write to the output tab.              *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpProfilerCallTreeData(HWND hwndProf)
{
    ADDIN_ENUM_PROFILER_CALLTREE Enum = {0};

    /* Clear the output tab */
    AddIn_WriteOutput(g_hwndMain, NULL);

    /* Clear total hits */
    g_cTotalHits = 0;

    /* Start enumerating profiler data */
    Enum.cbSize = sizeof(Enum);               /* required size (version control) */
    Enum.pfnCallback = EnumCallTreeCallback;  /* pointer to callback function */
    if (AddIn_EnumProfilerCallTree(hwndProf, &Enum))
    {
        WCHAR szBuf[80];

        /* Enumeration was successful */
        if (swprintf(szBuf, NELEMS(szBuf), L"%u hit(s) in total", g_cTotalHits) > 0)
            AddIn_WriteOutput(g_hwndMain, szBuf);
    }
    else
    {
        /* Enumeration failed */
        AddIn_WriteOutput(g_hwndMain, L"Failed to complete enumeration!");
    }
}

/****************************************************************************
 *                                                                          *
 * Function: EnumFunctionsCallback                                          *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumProfilerCallTree().                     *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static BOOL CALLBACK EnumCallTreeCallback(const ADDIN_PROFILER_CALLTREE * const pCallTree, LPVOID pvData)
{
    WCHAR szBuf[40+40+256+256];
    PWSTR pszStart;

    /* Remember total hits */
    g_cTotalHits = pCallTree->cSessionHits;

    /* Process basic information */
    pszStart = &szBuf[40];
    if (swprintf(pszStart, (&szBuf[NELEMS(szBuf)] - pszStart), L"%u hit(s) in function %.255ls, module %.255ls",
        pCallTree->cFunctionHits, pCallTree->pszFunctionName, pCallTree->pszModuleName) > 0)
    {
        /* Handle indentation */
        int nLevel = pCallTree->nIndentLevel;
        while (nLevel > 0 && pszStart > szBuf)
        {
            *--pszStart = L'.';
            *--pszStart = L'.';
            --nLevel;
        }
        AddIn_WriteOutput(g_hwndMain, pszStart);
    }

    return TRUE;
}
