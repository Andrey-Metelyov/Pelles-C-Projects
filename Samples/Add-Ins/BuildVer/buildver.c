/****************************************************************************
 *                                                                          *
 * File    : Buildver.c                                                     *
 *                                                                          *
 * Purpose : Sample add-in DLL. Demonstrates incrementing "build number"    *
 *           in a version resource.                                         *
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
#include <stdio.h>
#include <stdlib.h>

#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))

// Locals.
static HANDLE g_hmod = NULL;
static HWND g_hwndMain = NULL;

// Static function prototypes.
static void BumpBuildNumber(HWND);
static int ProcessResourceScript(PCWSTR, PWSTR, int);

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
            // Save module handle, in case we need it.
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
            // Remember main window handle.
            g_hwndMain = hwnd;
            return TRUE;

        case AIE_PRJ_STARTBUILD:
            // A new build in progress.
            BumpBuildNumber(hwnd);
            return TRUE;

        default:
            return TRUE;
    }
}

/****************************************************************************
 *                                                                          *
 * Function: BumpBuildNumber                                                *
 *                                                                          *
 * Purpose : Increment the version resource build number.                   *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void BumpBuildNumber(HWND hwndProj)
{
    if (hwndProj != NULL)
    {
        ADDIN_DOCUMENT_INFO DocInfo;

        // Get the path (and name) of the current project file.
        DocInfo.cbSize = sizeof(DocInfo);
        if (AddIn_GetDocumentInfo(hwndProj, &DocInfo))
        {
            PWSTR psz;
            WIN32_FIND_DATA wfd;
            HANDLE hff;

            // Replace name with search pattern (*.rc).
            for (psz = DocInfo.szFilename + wcslen(DocInfo.szFilename);
                 psz > DocInfo.szFilename && psz[-1] != L'\\' && psz[-1] != L'/';
                 psz--)
                ;
            lstrcpyn(psz, L"*.rc", (int)(DocInfo.szFilename + NELEMS(DocInfo.szFilename) - psz));

            // Search for matching files.
            hff = FindFirstFile(DocInfo.szFilename, &wfd);
            if (hff != INVALID_HANDLE_VALUE)
            {
                do
                {
                    WCHAR szVersion[32], szInfo[128];

                    // Replace name (again) with the actual name found.
                    lstrcpyn(psz, wfd.cFileName, (int)(DocInfo.szFilename + NELEMS(DocInfo.szFilename) - psz));

                    // After all this, maybe do something with the file?!
                    if (ProcessResourceScript(DocInfo.szFilename, szVersion, NELEMS(szVersion)))
                    {
                        swprintf(szInfo, NELEMS(szInfo), L"%ls set to version %ls", wfd.cFileName, szVersion);
                        AddIn_WriteOutput(g_hwndMain, szInfo);
                    }

                } while (FindNextFile(hff, &wfd));
                (void)FindClose(hff);
            }
        }
    }
}

/****************************************************************************
 *                                                                          *
 * Function: ProcessResourceScript                                          *
 *                                                                          *
 * Purpose : Read a resource script, look for a version resource to bump.   *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static int ProcessResourceScript(PCWSTR pcszFilename, PWSTR pchVersion, int cchMaxVersion)
{
    FILE *fin, *fout = 0;
    int cChanges = 0;
    WCHAR szTempfile[MAX_PATH];

    // Get a new temporary filename.
    if (GetTempFileName(L".", L"@@@", 0, szTempfile) == 0)
        szTempfile[0] = L'\0';

    // Open the input and output files.
    if ((fin = _wfopen(pcszFilename, L"r")) != NULL && (fout = _wfopen(szTempfile, L"w")) != NULL)
    {
        unsigned short v1, v2, v3, v4;
        int cBlock = 0;
        BOOL fVersion = FALSE;
        static WCHAR szLine[2048];

        // Read lines from input file until eof/error.
        while (fgetws(szLine, NELEMS(szLine), fin) != NULL)
        {
            PWSTR pszDup;

            // Make a copy, since _tcstok() will litter.
            if ((pszDup = _wcsdup(szLine)) != NULL)
            {
                PWSTR pszTok;
                PWSTR ptr;

                // Get first token on the line.
                if ((pszTok = wcstok(pszDup, L" \t\n", &ptr)) != NULL)
                {
                    // Seen a version resource yet?
                    if (!fVersion)
                    {
                        if (_wcsicmp(pszTok, L"VS_VERSION_INFO") == 0 || wcscmp(pszTok, L"1") == 0)
                        {
                            if ((pszTok = wcstok(NULL, L" \t\n", &ptr)) != NULL)
                            {
                                if (_wcsicmp(pszTok, L"VERSIONINFO") == 0)
                                {
                                    // Well, now we have...
                                    fVersion = TRUE;
                                    cBlock = 0;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (_wcsicmp(pszTok, L"BEGIN") == 0 || wcscmp(pszTok, L"{") == 0)
                        {
                            // Increment block nesting level.
                            ++cBlock;
                        }
                        else if (_wcsicmp(pszTok, L"END") == 0 || wcscmp(pszTok, L"}") == 0)
                        {
                            // Decrement block nesting level.
                            if (--cBlock <= 0)
                                fVersion = FALSE;
                        }
                        else if (cBlock == 0 && _wcsicmp(pszTok, L"FILEVERSION") == 0)
                        {
                            if ((pszTok = wcstok(NULL, L"\n", &ptr)) != NULL)
                            {
                                // Found binary version, get version number.
                                v1 = v2 = v3 = v4 = 0;
                                if (swscanf(pszTok, L"%hu, %hu, %hu, %hu", &v1, &v2, &v3, &v4) >= 2)
                                {
                                    ++v4;  // Pretend the lowest number is the build-number.
                                    swprintf(szLine, NELEMS(szLine), L"FILEVERSION %hu, %hu, %hu, %hu\n", v1, v2, v3, v4);
                                    cChanges++;
                                }
                            }
                        }
                        else if (cBlock > 0 && _wcsicmp(pszTok, L"VALUE") == 0)
                        {
                            if ((pszTok = wcstok(NULL, L", \t", &ptr)) != NULL)
                            {
                                if (_wcsicmp(pszTok, L"\"FileVersion\"") == 0 && cChanges == 1)
                                {
                                    // Found text version (and previously the binary version).
                                    swprintf(szLine, NELEMS(szLine), L"VALUE \"FileVersion\", \"%hu.%hu.%hu.%hu\\0\"\n", v1, v2, v3, v4);
                                    cChanges++;

                                    // Inform the caller about the new version.
                                    swprintf(pchVersion, cchMaxVersion, L"%hu.%hu.%hu.%hu", v1, v2, v3, v4);
                                }
                            }
                        }
                    }
                }

                free(pszDup);
            }

            // Write a line to the output file.
            if (fputws(szLine, fout) < 0)
            {
                cChanges = 0;
                break;
            }
        }
        if (ferror(fin))
            cChanges = 0;
    }

    if (fin)
    {
        // Close the input file.
        (void)fclose(fin);
    }
    if (fout)
    {
        // Close the output file.
        (void)fclose(fout);
        if (cChanges == 2)
        {
            // Replace original file with the modified file.
            _wremove(pcszFilename);
            _wrename(szTempfile, pcszFilename);
        }
        else
        {
            _wremove(szTempfile);
        }
    }

    return cChanges;
}
