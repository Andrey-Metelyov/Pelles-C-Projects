/****************************************************************************
 *                                                                          *
 * File    : dbginfo.c                                                      *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates debugger interfaces.           *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

/*
 * Comment out any of the following WANT_xxx if you don't need this info.
 */
#define WANT_SETTINGS_INFO          /* Debugger settings */
#define WANT_PROCESS_INFO           /* Debuggee precess info */
#define WANT_THREADS_INFO           /* Debuggee threads info */
#define WANT_MODULE_INFO            /* Debuggee modules info */
#define WANT_MODULE_FUNCTIONS_INFO  /* Debuggee functions info (by module) */
#define WANT_FUNCTION_LOCALS_INFO   /* Debuggee locals as info (by function) */
#define WANT_FUNCTION_LOCALS_TEXT   /* Debuggee locals as text (by function) */
#define WANT_MODULE_GLOBALS_INFO    /* Debuggee globals as info (by module) */
#define WANT_MODULE_GLOBALS_TEXT    /* Debuggee globals as text (by module) */
#define WANT_MODULE_SOURCE_INFO     /* Debuggee source lines (by module) */
#define WANT_BREAKPOINTS            /* Debuggee breakpoints */

/* There can be many functions, so set a limit for this sample */
#define MAX_FUNCTIONS_TO_LIST  10

/* There can be many source lines, so set a limit for this sample */
#define MAXE_SOURCE_LINES_TO_LIST  10

/* There can be many global symbols, so set a limit for this sample */
#define MAX_GLOBALS_TO_LIST  10

/* There can be many local symbols, so set a limit for this sample */
#define MAX_LOCALS_TO_LIST  10

/****************************************************************************/

#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define __STDC_WANT_LIB_EXT2__  1  /* for aswprintf() */
#include <windows.h>
#include <windowsx.h>
#include <addin.h>
#include <assert.h>
#include <stddef.h>
#include <stdlib.h>
#include <stdio.h>
#include <time.h>
#include <wchar.h>
#include "dbginfo.h"

#define TITLE_MACHINE(x)  ((x) == ADDIN_MACHINE_X86 ? L"X86" : (x) == ADDIN_MACHINE_X64 ? L"X64" : L"Unknown")

#undef verify
#ifdef NDEBUG
#define verify(e)  (void)(e)
#else /* !NDEBUG */
#define verify(e)  assert(e)
#endif /* !NDEBUG */

#ifndef NELEMS
#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))
#endif /* NELEMS */

#ifndef CPTRPLUS
#define CPTRPLUS(p,n)  (const void *)((const char *)(p) + (n))
#endif /* CPTRPLUS */

#ifndef CPTRDIFF
#define CPTRDIFF(p1,p2)  (size_t)(((const char *)(p1) - (const char *)(p2)) + 1)
#endif /* CPTRDIFF */

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;
static HWND g_hwndDbg = NULL;
static HWND g_hwndLB = NULL;
static ADDIN_DEBUG_PROCESS g_Proc = {
    .cbSize = sizeof(g_Proc),
    .eMachine = ADDIN_MACHINE_UNKNOWN
};

static void RefreshTabWindow(void);

#if defined(WANT_THREADS_INFO)
static BOOL CALLBACK EnumThreadsProc(const ADDIN_DEBUG_THREAD * const, LPVOID);
#endif /* WANT_THREADS_INFO */

#if defined(WANT_MODULE_FUNCTIONS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO) || defined(WANT_FUNCTION_LOCALS_TEXT) || \
    defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_MODULE_GLOBALS_TEXT) || defined(WANT_MODULE_SOURCE_INFO)
static BOOL CALLBACK EnumModulesProc(const ADDIN_DEBUG_MODULE * const, LPVOID);
#endif /* WANT_MODULE_FUNCTIONS_INFO || WANT_FUNCTION_LOCALS_INFO || WANT_FUNCTION_LOCALS_TEXT || WANT_MODULE_GLOBALS_INFO || WANT_MODULE_GLOBALS_TEXT || WANT_MODULE_SOURCE_INFO */

#if defined(WANT_MODULE_INFO)
static void DisasmCode(HWND, LPCVOID, DWORD);
static void DumpRawBytes(HWND, LPCVOID, DWORD);
#endif /* WANT_MODULE_INFO */

#if defined(WANT_MODULE_FUNCTIONS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO) || defined(WANT_FUNCTION_LOCALS_TEXT)
static BOOL CALLBACK EnumFunctionsProc(const ADDIN_DEBUG_FUNCTION * const, LPVOID);
#endif /* WANT_MODULE_FUNCTIONS_INFO || WANT_FUNCTION_LOCALS_INFO || WANT_FUNCTION_LOCALS_TEXT */

#if defined(WANT_MODULE_SOURCE_INFO)
static BOOL CALLBACK EnumSourceProc(const ADDIN_DEBUG_SOURCE * const, LPVOID);
#endif /* WANT_MODULE_SOURCE_INFO */

#if defined(WANT_BREAKPOINTS)
static BOOL CALLBACK EnumBreakpointsProc(const ADDIN_DEBUG_BREAKPOINT * const, LPVOID);
#endif /* WANT_BREAKPOINTS */

#if defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO)
static BOOL CALLBACK EnumSymbolsProc(const ADDIN_DEBUG_SYMBOL * const, LPVOID);
static LPCWSTR GetSymbolTypeString(const ADDIN_DEBUG_SYMBOL *);
static LPCWSTR GetSymbolSizeString(const ADDIN_DEBUG_SYMBOL *);
static LPCWSTR GetSymbolModifierString(const ADDIN_DEBUG_SYMBOL *);
#endif /* WANT_MODULE_GLOBALS_INFO || WANT_FUNCTION_LOCALS_INFO */

#if defined(WANT_MODULE_GLOBALS_TEXT) || defined(WANT_FUNCTION_LOCALS_TEXT)
static BOOL CALLBACK EnumSymbolsTextProc(LPCWSTR, LPCWSTR, LPCWSTR, LPCVOID, int, LPVOID);
#endif /* WANT_MODULE_GLOBALS_TEXT || WANT_FUNCTION_LOCALS_TEXT */


#define ListBox_AddFormatString(hwnd, fmt, ...) \
do { \
    PWSTR psz = NULL; \
    if (aswprintf(&psz, (fmt), __VA_ARGS__) >= 0) \
        (void)ListBox_AddString((hwnd), psz); \
    free(psz); \
} while (0)

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
    (void)lpReserved;
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
            // Save handle of the main IDE window.
            g_hwndMain = hwnd;
            return TRUE;

        case AIE_DBG_START:
            // Save handle of the debugger window.
            g_hwndDbg = hwnd;

            // Create new tab page for the debugger (fill the page with a listbox).
            g_hwndLB = CreateWindowEx(WS_EX_NOPARENTNOTIFY, L"listbox",
                NULL, WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_HASSTRINGS|LBS_NOINTEGRALHEIGHT,
                0, 0, 0, 0, g_hwndMain, NULL, g_hmod, NULL);
            if (g_hwndLB != NULL)
            {
                SetWindowFont(g_hwndLB, GetStockFont(ANSI_FIXED_FONT /*DEFAULT_GUI_FONT*/), TRUE);
                if (!AddIn_AddDebugTabPage(hwnd, L"DEBUGGER DEMO", g_hwndLB))
                    return FALSE;

                // Select the new tab page.
                (void)AddIn_SelectDebugTabPage(hwnd, g_hwndLB);
            }

            // Get basic debuggee info.
            if (!AddIn_GetDebugProcess(hwnd, &g_Proc))
                return FALSE;

            return TRUE;

        case AIE_APP_DESTROY:
        case AIE_DBG_END:
            // Shutting down (app or debugger); destroy tab page for the debugger.
            if (g_hwndLB != NULL)
            {
                if (AddIn_RemoveDebugTabPage(hwnd, g_hwndLB))
                {
                    (void)DestroyWindow(g_hwndLB);
                    g_hwndLB = NULL;
                }
            }
            g_hwndDbg = NULL;
            return TRUE;

        case AIE_DBG_BREAK:
            // Debuggee halted: refresh the tab page.
            RefreshTabWindow();
            return TRUE;

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: RefreshTabWindow                                               *
 *                                                                          *
 * Purpose : The debuggee is halted, so time to refresh our window.         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void RefreshTabWindow(void)
{
    HCURSOR hcurOld = SetCursor(LoadCursor(NULL, IDC_WAIT));
    (void)ShowCursor(TRUE);

    (void)ListBox_ResetContent(g_hwndLB);

#if defined(WANT_SETTINGS_INFO)
    //
    // Settings info.
    //
    ADDIN_DEBUG_INFO Info = {
        .cbSize = sizeof(Info)
    };
    verify(AddIn_GetDebugInfo(g_hwndDbg, &Info));

    ListBox_AddFormatString(g_hwndLB, L"- - - - - - - - - - GENERAL SETTINGS - - - - - - - - - -", 0);
    ListBox_AddFormatString(g_hwndLB, L"Maximum array elements: %u", Info.cMaxArrayElems);
    ListBox_AddFormatString(g_hwndLB, L"Display values in hexadecimal: %s", Info.fHexValues ? "yes" : "no");
    /* ListBox_AddFormatString(g_hwndLB, L"Check timestamp of source files: %s", Info.fCheckSourceTime ? "yes" : "no"); */
    ListBox_AddFormatString(g_hwndLB, L"Edit CPU flags through dialogs: %s", Info.fFlagsDialog ? "yes" : "no");
    ListBox_AddFormatString(g_hwndLB, L"Always break on entry-point: %s", Info.fBreakOnEntry ? "yes" : "no");
#endif /* WANT_SETTINGS_INFO */

#if defined(WANT_PROCESS_INFO)
    //
    // Process info.
    //
    ListBox_AddFormatString(g_hwndLB, L"- - - - - - - - - - PROCESS %lX - - - - - - - - - -", g_Proc.dwProcessId);
    ListBox_AddFormatString(g_hwndLB, L"Process type: \"%ls\"", TITLE_MACHINE(g_Proc.eMachine));
    ListBox_AddFormatString(g_hwndLB, L"Process ID: %lX", g_Proc.dwProcessId);
    ListBox_AddFormatString(g_hwndLB, L"Process handle: %P", g_Proc.hProcess);
#endif /* WANT_PROCESS_INFO */

#if defined(WANT_THREADS_INFO)
    //
    // Enumerate the current list of threads in the debuggee.
    //
    ADDIN_ENUM_DEBUG_THREADS EnumThreads = {
        .cbSize = sizeof(EnumThreads),
        .pfnCallback = EnumThreadsProc
    };
    verify(AddIn_EnumDebugThreads(g_hwndDbg, &EnumThreads));
#endif /* WANT_THREADS_INFO */

#if defined(WANT_MODULE_FUNCTIONS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO) || defined(WANT_FUNCTION_LOCALS_TEXT) || \
    defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_MODULE_GLOBALS_TEXT) || defined(WANT_MODULE_SOURCE_INFO)
    //
    // Enumerate the current list of loaded modules in the debuggee.
    //
    PVOID pvEntry = NULL;
    ADDIN_ENUM_DEBUG_MODULES EnumModules = {
        .cbSize = sizeof(EnumModules),
        .pfnCallback = EnumModulesProc,
        .pvData = &pvEntry,
    };
    verify(AddIn_EnumDebugModules(g_hwndDbg, &EnumModules));
#endif /* WANT_MODULE_FUNCTIONS_INFO || WANT_FUNCTION_LOCALS_INFO || WANT_FUNCTION_LOCALS_TEXT || WANT_MODULE_GLOBALS_INFO || WANT_MODULE_GLOBALS_TEXT || WANT_MODULE_SOURCE_INFO */

#if defined(WANT_BREAKPOINTS)
    // Set breakpoint at entry-point to make sure we have at least one breakpoint to enumerate...
    if (pvEntry != NULL && AddIn_SetDebugBreakpoint(g_hwndDbg, pvEntry, TRUE))
    {
        //
        // Enumerate the current list of breakpoints in the debuggee.
        //
        ADDIN_ENUM_DEBUG_BREAKPOINTS EnumBreakpoints = {
            .cbSize = sizeof(EnumBreakpoints),
            .pfnCallback = EnumBreakpointsProc
        };
        verify(AddIn_EnumDebugBreakpoints(g_hwndDbg, &EnumBreakpoints));
        verify(AddIn_RemoveDebugBreakpoint(g_hwndDbg, pvEntry));
    }
#endif /* WANT_BREAKPOINTS */

    (void)SetCursor(hcurOld);
    (void)ShowCursor(FALSE);
}

/****************************************************************************
 *                                                                          *
 * Function: EnumThreadsProc                                                *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugThreads().                         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_THREADS_INFO)
static BOOL CALLBACK EnumThreadsProc(const ADDIN_DEBUG_THREAD * const pThread, LPVOID pvData)
{
    (void)pvData;  /* shut up */

    //
    // Make sure we are compiled with the correct version!
    //
    if (pThread->cbSize != sizeof(*pThread))
        return FALSE;

    //
    // Thread info.
    //
    ListBox_AddFormatString(g_hwndLB, L"- - - - - - - - - - THREAD %lX - - - - - - - - - -", pThread->dwThreadId);
    ListBox_AddFormatString(g_hwndLB, L"Thread ID: %lX", pThread->dwThreadId);
    ListBox_AddFormatString(g_hwndLB, L"Thread handle: %P", pThread->hThread);

    //
    // Read some machine registers (depending on current machine).
    //
    if (g_Proc.eMachine == ADDIN_MACHINE_X86)
    {
        ADDIN_READ_DEBUG_REGISTER Read;

        Read.cbSize = sizeof(Read);
        Read.dwThreadId = pThread->dwThreadId;

        Read.eRegister = ADDIN_REG_X86_EAX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[EAX]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_EBX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[EBX]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_ECX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[ECX]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_EDX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[EDX]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_ESP;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[ESP]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_EBP;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[EBP]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_ESI;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[ESI]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_EDI;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[EDI]: %lX", Read.u.Reg32);

        Read.eRegister = ADDIN_REG_X86_EIP;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[EIP]: %lX", Read.u.Reg32);
    }
    else if (g_Proc.eMachine == ADDIN_MACHINE_X64)
    {
        ADDIN_READ_DEBUG_REGISTER Read;

        Read.cbSize = sizeof(Read);
        Read.dwThreadId = pThread->dwThreadId;

        Read.eRegister = ADDIN_REG_X64_RAX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RAX]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RBX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RBX]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RCX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RCX]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RDX;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RDX]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RSP;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RSP]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RBP;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RBP]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RSI;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RSI]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RDI;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RDI]: %llX", Read.u.Reg64);

        Read.eRegister = ADDIN_REG_X64_RIP;
        if (AddIn_ReadDebugRegister(g_hwndDbg, &Read))
            ListBox_AddFormatString(g_hwndLB, L"[RIP]: %llX", Read.u.Reg64);
    }

    return TRUE;  /* OK, feed me another thread... */
}
#endif /* WANT_THREADS_INFO */

/****************************************************************************
 *                                                                          *
 * Function: EnumModulesProc                                                *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugModules().                         *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_FUNCTIONS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO) || defined(WANT_FUNCTION_LOCALS_TEXT) || \
    defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_MODULE_GLOBALS_TEXT) || defined(WANT_MODULE_SOURCE_INFO)
static BOOL CALLBACK EnumModulesProc(const ADDIN_DEBUG_MODULE * const pModule, LPVOID pvData)
{
#if defined(WANT_MODULE_INFO)
    LPCVOID *ppvEntry = pvData;
#else /* !WANT_MODULE_INFO */
    (void)pvData;  /* shut up */
#endif /* !WANT_MODULE_INFO */

    //
    // Make sure we are compiled with the correct version!
    //
    if (pModule->cbSize != sizeof(*pModule))
        return FALSE;

#if defined(WANT_MODULE_INFO)
    //
    // Module info.
    //
    ListBox_AddFormatString(g_hwndLB, L"- - - - - - - - - - MODULE %lX - - - - - - - - - -", pModule->dwModuleId);
    ListBox_AddFormatString(g_hwndLB, L"Module ID: %lX", pModule->dwModuleId);
    ListBox_AddFormatString(g_hwndLB, L"Timestamp: %lX (%.24s)", pModule->dwTimeStamp, ctime((const time_t *)&pModule->dwTimeStamp));
    ListBox_AddFormatString(g_hwndLB, L"Address range: %P - %P", pModule->RangeImage.pvStartAddress, pModule->RangeImage.pvEndAddress);
    ListBox_AddFormatString(g_hwndLB, L"Entry-point: %P", pModule->pvEntryPoint);
    ListBox_AddFormatString(g_hwndLB, L"Display name: \"%ls\"", pModule->pszDispName);
    ListBox_AddFormatString(g_hwndLB, L"File name: \"%ls\"", pModule->pszFileName);

    if (pModule->pvEntryPoint != NULL)
    {
        // Disassemble a few instructions, just as a demonstration */
        DisasmCode(g_hwndLB, pModule->pvEntryPoint, 8);

        // Dump a few odd bytes, just as a demonstration */
        DumpRawBytes(g_hwndLB, pModule->pvEntryPoint, 100);

        if (*ppvEntry == NULL)
            *ppvEntry = pModule->pvEntryPoint;
    }
#endif /* WANT_MODULE_INFO */

#if defined(WANT_MODULE_FUNCTIONS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO) || defined(WANT_FUNCTION_LOCALS_TEXT)
    //
    // Enumerate module functions.
    //
    int cFuncsLimit = MAX_FUNCTIONS_TO_LIST;
    ADDIN_ENUM_DEBUG_FUNCTIONS EnumFunctions = {
        .cbSize = sizeof(EnumFunctions),
        .dwModuleId = pModule->dwModuleId,
        .pfnCallback = EnumFunctionsProc,
        .fuFlags = EDFF_SORT_BY_NAME,
        .pvData = &cFuncsLimit
    };
    verify(AddIn_EnumDebugFunctions(g_hwndDbg, &EnumFunctions));
#endif /* WANT_MODULE_FUNCTIONS_INFO || WANT_FUNCTION_LOCALS_INFO || WANT_FUNCTION_LOCALS_TEXT */

#if defined(WANT_MODULE_SOURCE_INFO)
    //
    // Enumerate module source lines.
    //
    int cLinesLimit = MAXE_SOURCE_LINES_TO_LIST;
    ADDIN_ENUM_DEBUG_SOURCE EnumSource = {
        .cbSize = sizeof(EnumSource),
        .dwModuleId = pModule->dwModuleId,
        .pfnCallback = EnumSourceProc,
        .pvData = &cLinesLimit
    };
    verify(AddIn_EnumDebugSourceW(g_hwndDbg, &EnumSource));
#endif /* WANT_MODULE_SOURCE_INFO */

#if defined(WANT_MODULE_GLOBALS_INFO)
    //
    // Enumerate module globals (as info).
    //
    int cSymsLimit = MAX_GLOBALS_TO_LIST;
    ADDIN_ENUM_DEBUG_GLOBALS EnumGlobals = {
        .cbSize = sizeof(EnumGlobals),
        .dwModuleId = pModule->dwModuleId,
        .pfnCallback = EnumSymbolsProc,
        .pvData = &cSymsLimit
    };
    verify(AddIn_EnumDebugGlobals(g_hwndDbg, &EnumGlobals));
#endif /* WANT_MODULE_GLOBALS_TEXT */

#if defined(WANT_MODULE_GLOBALS_TEXT)
    //
    // Enumerate module globals (as text).
    //
    int cSymsTextLimit = MAX_GLOBALS_TO_LIST;
    ADDIN_ENUM_DEBUG_GLOBALS_TEXT EnumGlobalsText = {
        .cbSize = sizeof(EnumGlobalsText),
        .dwModuleId = pModule->dwModuleId,
        .pfnCallback = EnumSymbolsTextProc,
        .pvData = &cSymsTextLimit,
        .fExpand = TRUE
    };
    verify(AddIn_EnumDebugGlobalsText(g_hwndDbg, &EnumGlobalsText));
#endif /* WANT_MODULE_GLOBALS_TEXT */

    return TRUE;  /* OK, feed me another module... */
}
#endif /* WANT_MODULE_FUNCTIONS_INFO || WANT_FUNCTION_LOCALS_INFO || WANT_FUNCTION_LOCALS_TEXT || WANT_MODULE_GLOBALS_INFO || WANT_MODULE_GLOBALS_TEXT || WANT_MODULE_SOURCE_INFO */

/****************************************************************************
 *                                                                          *
 * Function: DisasmCode                                                     *
 *                                                                          *
 * Purpose : Disassemble machine code starting at address <pvAddr>,         *
 *           display no more than <cInsns> number of instructions.          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_INFO)
static void DisasmCode(HWND hwndLB, LPCVOID pvAddr, DWORD cInsns)
{
    /* Some machines are more similar than others... */
    if (g_Proc.eMachine == ADDIN_MACHINE_X86 ||
        g_Proc.eMachine == ADDIN_MACHINE_X64)
    {
#define X86_MAX_INSN_LEN  15
        assert(cInsns > 0);
        DWORD cbCode = X86_MAX_INSN_LEN * cInsns;
        LPVOID pvCode = malloc(cbCode);
        if (pvCode != NULL)
        {
            ADDIN_READ_DEBUG_MEMORY Read = {
                .cbSize = sizeof(Read),
                .pvAddress = pvAddr,
                .pvData = pvCode,
                .cbData = cbCode,
            };
            if (AddIn_ReadDebugMemory(g_hwndDbg, &Read))
            {
                WCHAR szOpcode[40];
                WCHAR szOpands[80];
                ADDIN_DISASM_DEBUG_CODE Disasm = {
                    .cbSize = sizeof(Disasm),
                    /* .pvCode = p; */
                    /* .pvAddress = q; */
                    .pchOpcode = szOpcode,
                    .cchMaxOpcode = NELEMS(szOpcode),
                    .pchOperands = szOpands,
                    .cchMaxOperands = NELEMS(szOpands),
                    .fSymbols = TRUE,  /* Well, if you insist... */
                };
                for (DWORD iCode = 0, cbLen = 0; iCode < (cbCode - X86_MAX_INSN_LEN); iCode += cbLen)
                {
                    Disasm.pvCode = CPTRPLUS(pvCode, iCode);
                    Disasm.cbCode = X86_MAX_INSN_LEN;
                    Disasm.pvAddress = CPTRPLUS(pvAddr, iCode);
                    cbLen = AddIn_DisasmDebugCode(g_hwndDbg, &Disasm);

                    ListBox_AddFormatString(hwndLB, L"[%P + %04lX]: %ls %ls", pvAddr, iCode, szOpcode, szOpands);

                    if (--cInsns == 0) break;
                }
            }
            free(pvCode);
        }
#undef X86_MAX_INSN_LEN
    }
}
#endif /* WANT_MODULE_INFO */

/****************************************************************************
 *                                                                          *
 * Function: DumpRawBytes                                                   *
 *                                                                          *
 * Purpose : Dump raw bytes starting at address <pvAddr>,                   *
 *           display no more than <cBytes> number of bytes.                 *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_INFO)
static void DumpRawBytes(HWND hwndLB, LPCVOID pvAddr, DWORD cBytes)
{
    LPCVOID pvStop = CPTRPLUS(pvAddr, cBytes);
    while (pvAddr < pvStop)
    {
#define BYTES_PER_LINE  16
        BYTE abRawBytes[BYTES_PER_LINE];

        // Get the minimum of remaining bytes and BYTES_PER_LINE.
        DWORD64 cbRawBytes = CPTRDIFF(pvStop, pvAddr);
        if (cbRawBytes > BYTES_PER_LINE)
            cbRawBytes = BYTES_PER_LINE;

        // Attempt to read this many bytes from the debuggee.
        ADDIN_READ_DEBUG_MEMORY Read = {
            .cbSize = sizeof(Read),
            .pvAddress = pvAddr,
            .pvData = abRawBytes,
            .cbData = (DWORD)cbRawBytes
        };
        if (AddIn_ReadDebugMemory(g_hwndDbg, &Read))
        {
            // Format a line of bytes.
            WCHAR wchHexBytes[3 * BYTES_PER_LINE+1], *pwchHexBytes = wchHexBytes;
            WCHAR wchTxtBytes[BYTES_PER_LINE+1], *pwchTxtBytes = wchTxtBytes;

            for (size_t i = 0; i < BYTES_PER_LINE; i++)
            {
                if (i < cbRawBytes)
                {
                    BYTE b = abRawBytes[i];

                    *pwchHexBytes++ = L"0123456789ABCDEF"[b / 16];
                    *pwchHexBytes++ = L"0123456789ABCDEF"[b % 16];
                    *pwchHexBytes++ = L' ';

                    *pwchTxtBytes++ = (b >= L' ' ? b : L'.');
                }
                else /* unread */
                {
                    *pwchHexBytes++ = L' ';
                    *pwchHexBytes++ = L' ';
                    *pwchHexBytes++ = L' ';

                    *pwchTxtBytes++ = L' ';
                }
            }
            *pwchHexBytes = L'\0';
            *pwchTxtBytes = L'\0';

            // Emit a line of bytes.
            ListBox_AddFormatString(hwndLB, L"[%P]  %ls %ls", pvAddr, wchHexBytes, wchTxtBytes);
        }
        else
        {
            /* handle total failure... */
        }

        pvAddr = CPTRPLUS(pvAddr, cbRawBytes);
#undef BYTES_PER_LINE
    }
}
#endif /* WANT_MODULE_INFO */

/****************************************************************************
 *                                                                          *
 * Function: EnumFunctionsProc                                              *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugFunctions().                       *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_FUNCTIONS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO) || defined(WANT_FUNCTION_LOCALS_TEXT)
static BOOL CALLBACK EnumFunctionsProc(const ADDIN_DEBUG_FUNCTION * const pFunction, LPVOID pvData)
{
    int *pcFuncsLimit = pvData;  /* pick up the limit from EnumModulesProc() */
    if (*pcFuncsLimit > 0)
    {
        //
        // Make sure we are compiled with the correct version!
        //
        if (pFunction->cbSize != sizeof(*pFunction))
            return FALSE;

        //
        // Function info.
        //
        ListBox_AddFormatString(g_hwndLB, L"< function \"%ls\" at %P, %zX byte(s) >",
            pFunction->pszName,
            pFunction->RangeFunc.pvStartAddress,
            CPTRDIFF(pFunction->RangeFunc.pvEndAddress, pFunction->RangeFunc.pvStartAddress));

#if defined(WANT_FUNCTION_LOCALS_INFO)
        //
        // Enumerate function locals (as info).
        //
        int cSymsLimit = MAX_LOCALS_TO_LIST;
        ADDIN_ENUM_DEBUG_LOCALS EnumLocals = {
            .cbSize = sizeof(EnumLocals),
            .dwFunctionId = pFunction->dwFunctionId,
            .pfnCallback = EnumSymbolsProc,
            .pvData = &cSymsLimit
        };
        verify(AddIn_EnumDebugLocals(g_hwndDbg, &EnumLocals));
#endif /* WANT_FUNCTION_LOCALS_INFO */

#if defined(WANT_FUNCTION_LOCALS_TEXT)
        //
        // Enumerate function locals (as text).
        //
        int cSymsTextLimit = MAX_LOCALS_TO_LIST;
        ADDIN_ENUM_DEBUG_LOCALS_TEXT EnumLocalsText = {
            .cbSize = sizeof(EnumLocalsText),
            .dwFunctionId = pFunction->dwFunctionId,
            .pfnCallback = EnumSymbolsTextProc,
            .pvData = &cSymsTextLimit,
            .fExpand = TRUE
        };
        verify(AddIn_EnumDebugLocalsText(g_hwndDbg, &EnumLocalsText));
#endif /* WANT_FUNCTION_LOCALS_TEXT */

        --(*pcFuncsLimit);
        return TRUE;  /* OK, feed me another function... */
    }

    ListBox_AddFormatString(g_hwndLB, L"< function ... >", 0);
    return FALSE;  /* No thanks, I'm full... */
}
#endif /* WANT_MODULE_FUNCTIONS_INFO || WANT_FUNCTION_LOCALS_INFO || WANT_FUNCTION_LOCALS_TEXT */

/****************************************************************************
 *                                                                          *
 * Function: EnumSourceProc                                                 *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugSource().                          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_SOURCE_INFO)
static BOOL CALLBACK EnumSourceProc(const ADDIN_DEBUG_SOURCE * const pSource, LPVOID pvData)
{
    int *pcLinesLimit = pvData;  /* pick up the limit from EnumModulesProc() */
    if (*pcLinesLimit > 0)
    {
        //
        // Make sure we are compiled with the correct version!
        //
        if (pSource->cbSize != sizeof(*pSource))
            return FALSE;

        //
        // Source line info.
        //
        ListBox_AddFormatString(g_hwndLB, L"< source \"%ls\"(%u) at %P >",
            pSource->pszFileName,
            pSource->uLine,
            pSource->pvAddress);

        --(*pcLinesLimit);
        return TRUE;  /* OK, feed me another source chunk... */
    }

    ListBox_AddFormatString(g_hwndLB, L"< source ... >", 0);
    return FALSE;  /* No thanks, I'm full... */
}
#endif /* WANT_MODULE_SOURCE_INFO */

/****************************************************************************
 *                                                                          *
 * Function: EnumBreakpointsProc                                            *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugBreakpoints().                     *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_BREAKPOINTS)
static BOOL CALLBACK EnumBreakpointsProc(const ADDIN_DEBUG_BREAKPOINT * const pBrkpnt, LPVOID pvData)
{
    (void)pvData;  /* shut up */

    //
    // Make sure we are compiled with the correct version!
    //
    if (pBrkpnt->cbSize != sizeof(*pBrkpnt))
        return FALSE;

    if (g_Proc.eMachine == ADDIN_MACHINE_X86 ||
        g_Proc.eMachine == ADDIN_MACHINE_X64)
    {
        //
        // Breakpoint info.
        //
        ListBox_AddFormatString(g_hwndLB, L"- - - - - - - - - - BREAKPOINT - - - - - - - - - -", L"");
        ListBox_AddFormatString(g_hwndLB, L"Address: %p", pBrkpnt->pvAddress);
        ListBox_AddFormatString(g_hwndLB, L"Original byte: %hhX ( currently occupied by \"INT3\" (0xCC) )", pBrkpnt->Opcode);
    }

    return TRUE;  /* OK, feed me another breakpoint... */
}
#endif /* WANT_BREAKPOINTS */

/****************************************************************************
 *                                                                          *
 * Function: EnumSymbolsProc                                                *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugGlobals/Locals().                  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO)
static BOOL CALLBACK EnumSymbolsProc(const ADDIN_DEBUG_SYMBOL * const pSymbol, LPVOID pvData)
{
    int *pcSymsLimit = pvData;  /* pick up the limit from EnumModulesProc() */
    if (*pcSymsLimit > 0)
    {
        //
        // Make sure we are compiled with the correct version!
        //
        if (pSymbol->cbSize != sizeof(*pSymbol))
            return FALSE;

        DWORD dwType = ADDIN_SYMBOL_TYPE(pSymbol->fuKind);
        DWORD cbData = ADDIN_SYMBOL_SIZE(pSymbol->fuKind);

        WCHAR szData[32] = L"unknown";

        //
        // Attempt to fetch the value of plain integer symbols.
        //
        if ((dwType == ADDIN_SYMKIND_INT || dwType == ADDIN_SYMKIND_UINT || dwType == ADDIN_SYMKIND_ENUM) &&
            !ADDIN_SYMBOL_IS_ARRAY(pSymbol->fuKind) &&
            !ADDIN_SYMBOL_IS_POINTER(pSymbol->fuKind) &&
            !ADDIN_SYMBOL_IN_TLS(pSymbol->fuKind) &&
            cbData <= sizeof(DWORD64))
        {
            if (pSymbol->pvAddress != NULL)  /* global symbol, at fixed address */
            {
                DWORD64 qwData = 0;
                ADDIN_READ_DEBUG_MEMORY Read = {
                    .cbSize = sizeof(Read),
                    .pvAddress = pSymbol->pvAddress,
                    .pvData = &qwData,
                    .cbData = cbData
                };
                if (AddIn_ReadDebugMemory(g_hwndDbg, &Read))
                {
                    switch (cbData)
                    {
                        case 1:
                            (void)swprintf(szData, NELEMS(szData), L"%hhu", (BYTE)qwData);
                            break;
                        case 2:
                            (void)swprintf(szData, NELEMS(szData), L"%hu", (WORD)qwData);
                            break;
                        case 4:
                            (void)swprintf(szData, NELEMS(szData), L"%lu", (DWORD)qwData);
                            break;
                        case 8:
                            (void)swprintf(szData, NELEMS(szData), L"%llu", qwData);
                            break;
                        default:
                            (void)swprintf(szData, NELEMS(szData), L"?");
                            break;
                    }
                }
            }
            else  /* other kind of symbol */
            {
                // TODO: add code. much code.
            }
        }

        if (pSymbol->pvAddress != NULL)
        {
            ListBox_AddFormatString(g_hwndLB, L"< global symbol \"%ls\" at %P, %ls%ls%ls, value %ls >",
                pSymbol->pszName,
                pSymbol->pvAddress,
                GetSymbolModifierString(pSymbol), GetSymbolSizeString(pSymbol), GetSymbolTypeString(pSymbol),
                szData
                );
        }
        else /* local */
        {
            ListBox_AddFormatString(g_hwndLB, L"< local symbol \"%ls\", kind 0x%lX, size %lu, offset %ld, reg 0x%lX >",
                pSymbol->pszName,
                pSymbol->fuKind,
                pSymbol->cbTotalSize,
                pSymbol->dwOffsetValue,
                pSymbol->eRegister
                );
        }

        --(*pcSymsLimit);
        return TRUE;  /* OK, feed me another symbol... */
    }

    ListBox_AddFormatString(g_hwndLB, L"< symbol ... >", 0);
    return FALSE;  /* No thanks, I'm full... */
}
#endif /* WANT_MODULE_GLOBALS_INFO || WANT_FUNCTION_LOCALS_INFO */

/****************************************************************************
 *                                                                          *
 * Subfunction: GetSymbolTypeString                                         *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO)
static LPCWSTR GetSymbolTypeString(const ADDIN_DEBUG_SYMBOL *pSymbol)
{
    switch (ADDIN_SYMBOL_TYPE(pSymbol->fuKind))
    {
        case ADDIN_SYMKIND_VOID: return L"void";
        case ADDIN_SYMKIND_INT: return L"signed integer";
        case ADDIN_SYMKIND_UINT: return L"unsigned integer";
        case ADDIN_SYMKIND_REAL: return L"floating-point";
        case ADDIN_SYMKIND_ENUM: return L"enumerator";
        case ADDIN_SYMKIND_BITFIELD: return L"bit-field";
        case ADDIN_SYMKIND_STRUCT: return L"structure";
        case ADDIN_SYMKIND_UNION: return L"union";
        case ADDIN_SYMKIND_BITINT: return L"_BitInt(N)";            /* N == pSymbol->cbTotalSize */
        case ADDIN_SYMKIND_UBITINT: return L"unsigned _BitInt(N)";  /* N == pSymbol->cbTotalSize */
        default: return L"type?!";
    }
}
#endif /* WANT_MODULE_GLOBALS_INFO || WANT_FUNCTION_LOCALS_INFO */

/****************************************************************************
 *                                                                          *
 * Subfunction: GetSymbolSizeString                                         *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO)
static LPCWSTR GetSymbolSizeString(const ADDIN_DEBUG_SYMBOL *pSymbol)
{
    switch (ADDIN_SYMBOL_SIZE(pSymbol->fuKind))
    {
        case 0: return L"";
        case 1: return L"1-byte ";
        case 2: return L"2-byte ";
        case 4: return L"4-byte ";
        case 8: return L"8-byte ";
        case 16: return L"16-byte ";
        case 32: return L"32-byte ";
        default: return L"size?! ";
    }
}
#endif /* WANT_MODULE_GLOBALS_INFO || WANT_FUNCTION_LOCALS_INFO */

/****************************************************************************
 *                                                                          *
 * Subfunction: GetSymbolModifierString                                     *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_GLOBALS_INFO) || defined(WANT_FUNCTION_LOCALS_INFO)
static LPCWSTR GetSymbolModifierString(const ADDIN_DEBUG_SYMBOL *pSymbol)
{
    if (ADDIN_SYMBOL_IS_ARRAY(pSymbol->fuKind))
        return L"array of ";
    if (ADDIN_SYMBOL_IS_POINTER(pSymbol->fuKind))
        return L"pointer to ";
    return L"";
}
#endif /* WANT_MODULE_GLOBALS_INFO || WANT_FUNCTION_LOCALS_INFO */

/****************************************************************************
 *                                                                          *
 * Function: EnumSymbolsTextProc                                            *
 *                                                                          *
 * Purpose : Callback for AddIn_EnumDebugGlobalsText/LocalsText().          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

#if defined(WANT_MODULE_GLOBALS_TEXT) || defined(WANT_FUNCTION_LOCALS_TEXT)
static BOOL CALLBACK EnumSymbolsTextProc(LPCWSTR pszName, LPCWSTR pszValue, LPCWSTR pszType, LPCVOID pvAddress, int nLevel, LPVOID pvData)
{
    int *pcSymsLimit = pvData;  /* pick up the limit from EnumModulesProc() */
    if (*pcSymsLimit > 0)
    {
        ListBox_AddFormatString(g_hwndLB, L"%*s < text symbol \"%ls\", value \"%ls\", type \"%ls\", address %P >",
            4*nLevel, "",  /* indentation */
            pszName,
            pszValue,
            pszType,
            pvAddress
            );

        if (nLevel == 0)  /* top level */
            --(*pcSymsLimit);
        return TRUE;  /* OK, feed me another symbol... */
    }

    ListBox_AddFormatString(g_hwndLB, L"< text symbol ... >", 0);
    return FALSE;  /* No thanks, I'm full... */
}
#endif /* WANT_MODULE_GLOBALS_TEXT || WANT_FUNCTION_LOCALS_TEXT */
