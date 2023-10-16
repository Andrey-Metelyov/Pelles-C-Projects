/****************************************************************************
 *                                                                          *
 * File    : Worker.c                                                       *
 *                                                                          *
 * Purpose : Parse a COM Type Library and build type definitions.           *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

/*
 * This is a sample. Not all types are supported, and some of the algorithms
 * are very simplistic. The good news is that you got the source - so all
 * this can be fixed! ;-)
 */

#define WIN32_LEAN_AND_MEAN
#define UNICODE   /* for Windows API */
#define _UNICODE  /* for C runtime */
#include <windows.h>
#include <wchar.h>
#include <stdio.h>
#include "typelib.h"

#define COBJMACROS
#include <ole2.h>

#define MAXLINE  1024
#define BUFSTEP  (MAXLINE * 100)

typedef struct OUTPUT OUTPUT, *POUTPUT;
typedef struct NAME NAME, *PNAME;

/* Output buffer */
struct OUTPUT {
    PWSTR pchBuf;       /* Pointer to buffer, or NULL */
    DWORD cchBuf;       /* Current number of chars */
    DWORD cchMaxBuf;    /* Maximum number of chars */
};

/* Linked list for tag names */
struct NAME {
    PNAME pNext;        /* Pointer to next node, or NULL */
    PWSTR pszName;      /* Tag name */
};

/* Head of linked list */
static PNAME g_pNameHead = NULL;
static OUTPUT Fwd = {0};

/* Static function prototypes */
static void EnumTypeLib(POUTPUT, LPTYPELIB);
static void DumpTypeInfo(POUTPUT, LPTYPEINFO, BOOL);
static void DumpAliasType(POUTPUT, BSTR, LPTYPEATTR, LPTYPEINFO);
static void DumpClassType(POUTPUT, BSTR, LPTYPEATTR);
static void DumpEnumType(POUTPUT, BSTR, BSTR, LPTYPEATTR, LPTYPEINFO);
static void DumpRecordType(POUTPUT, BSTR, BSTR, LPTYPEATTR, LPTYPEINFO);
static void DumpInterfaceType(POUTPUT, BSTR, LPTYPEATTR, LPTYPEINFO, BOOL);
static void GetTypeName(PWSTR, size_t, TYPEDESC, BSTR, LPTYPEINFO);
static void StrCat(PWSTR, size_t, PCWSTR, ...);
static void BufCat(POUTPUT, PCWSTR, ...);
static PNAME AllocName(PCWSTR);
static PNAME LookupName(PCWSTR);

/****************************************************************************
 *                                                                          *
 * Function: DumpTypeLib                                                    *
 *                                                                          *
 * Purpose : Dump a type library to a buffer, and return it.                *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

PTSTR DumpTypeLib(PCWSTR pcszFilename)
{
    OUTPUT Out = {0};
    LPTYPELIB pITypeLib;

    memset(&Fwd, 0, sizeof(Fwd));

    __try
    {
        /* Load the given type library */
        HRESULT hr = LoadTypeLibEx(pcszFilename, REGKIND_NONE, &pITypeLib);
        if (hr == S_OK)
        {
            WCHAR szName[MAX_PATH];
            PCWSTR pcsz, pcszEnd;
            DWORD cchNow;

            /* Find beginning of name */
            for (pcsz = pcszFilename + wcslen(pcszFilename);
                 pcsz > pcszFilename && pcsz[-1] != L'\\' && pcsz[-1] != L'/' && pcsz[-1] != L':';
                 pcsz--)
                ;

            /* Find end of name */
            for (pcszEnd = pcsz; *pcszEnd != L'\0' && *pcszEnd != L'.'; pcszEnd++)
                ;

            /* Extract name */
            lstrcpyn(szName, pcsz, (int)(pcszEnd - pcsz + 1));
            CharUpper(szName);

            BufCat(&Out, L"#ifndef COM_NO_WINDOWS_H\n");
            BufCat(&Out, L"#include <windows.h>\n");
            BufCat(&Out, L"#include <ole2.h>\n");
            BufCat(&Out, L"#endif\n\n");

            BufCat(&Out, L"#ifndef H_%ls\n", szName);
            BufCat(&Out, L"#define H_%ls\n", szName);

            /* Remember buffer offset */
            cchNow = Out.cchBuf;

            /* Enumerate types in the type library */
            EnumTypeLib(&Out, pITypeLib);

            BufCat(&Out, L"\n#endif /* H_%ls */\n", szName);

            /* If we have interfaces, merge the forward declarations into the main buffer */
            if (Fwd.cchBuf != 0)
            {
                PWSTR pch = malloc((Fwd.cchBuf + Out.cchBuf + 1) * sizeof(WCHAR));
                if (pch != NULL)
                {
                    wmemcpy(pch, Out.pchBuf, cchNow);
                    wmemcpy(pch + cchNow, Fwd.pchBuf, Fwd.cchBuf);
                    wmemcpy(pch + cchNow + Fwd.cchBuf, Out.pchBuf + cchNow, Out.cchBuf - cchNow);
                    pch[cchNow + Fwd.cchBuf + Out.cchBuf] = L'\0';
                    free(Fwd.pchBuf); Fwd.pchBuf = NULL;
                    free(Out.pchBuf); Out.pchBuf = pch;
                }
            }

            ITypeLib_Release(pITypeLib);

            /* Free linked list of tag names */
            while (g_pNameHead != NULL)
            {
                PNAME pName = g_pNameHead;
                g_pNameHead = g_pNameHead->pNext;
                free(pName);
            }
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        free(Fwd.pchBuf);
        free(Out.pchBuf);
        Fwd.pchBuf = Out.pchBuf = NULL;
    }

    /* Return the buffer */
    return Out.pchBuf;
}

/****************************************************************************
 *                                                                          *
 * Function: EnumTypeLib                                                    *
 *                                                                          *
 * Purpose : Walk through a type library and dump it's types.               *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void EnumTypeLib(POUTPUT pOut, LPTYPELIB pITypeLib)
{
    UINT cTypes = ITypeLib_GetTypeInfoCount(pITypeLib);
    for (UINT i = 0; i < cTypes; i++)
    {
        LPTYPEINFO pITypeInfo;

        /* Dump this type, please */
        if (ITypeLib_GetTypeInfo(pITypeLib, i, &pITypeInfo) == S_OK)
        {
            DumpTypeInfo(pOut, pITypeInfo, FALSE);
            ITypeInfo_Release(pITypeInfo);
        }
    }
}

/****************************************************************************
 *                                                                          *
 * Function: DumpTypeInfo                                                   *
 *                                                                          *
 * Purpose : Dump information about a specific type.                        *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpTypeInfo(POUTPUT pOut, LPTYPEINFO pITypeInfo, BOOL fDispatch)
{
    BSTR bstrTypeName, bstrComment;

    if (ITypeInfo_GetDocumentation(pITypeInfo, MEMBERID_NIL, &bstrTypeName, &bstrComment, 0, 0) == S_OK)
    {
        TYPEATTR *pTypeAttr;

        if (ITypeInfo_GetTypeAttr(pITypeInfo, &pTypeAttr) == S_OK)
        {
            /*
             * Handle types we care about...
             */
            switch (pTypeAttr->typekind)
            {
                static TYPEKIND TKindPrev = TKIND_MAX;

                case TKIND_COCLASS:
                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"\n/* %ls */\n", bstrComment);
                    else if (TKindPrev != TKIND_COCLASS)
                        BufCat(pOut, L"\n");
                    DumpClassType(pOut, bstrTypeName, pTypeAttr);
                    TKindPrev = TKIND_COCLASS;
                    break;

                case TKIND_ENUM:
                    if (LookupName(bstrTypeName)) break;
                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"\n/* %ls */\n", bstrComment);
                    else
                        BufCat(pOut, L"\n");
                    DumpEnumType(pOut, bstrTypeName, bstrTypeName, pTypeAttr, pITypeInfo);
                    break;

                case TKIND_RECORD:
                case TKIND_UNION:
                    if (LookupName(bstrTypeName)) break;
                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"\n/* %ls */\n", bstrComment);
                    else
                        BufCat(pOut, L"\n");
                    DumpRecordType(pOut, bstrTypeName, bstrTypeName, pTypeAttr, pITypeInfo);
                    break;

                case TKIND_ALIAS:
                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"\n/* %ls */\n", bstrComment);
                    else
                        BufCat(pOut, L"\n");
                    DumpAliasType(pOut, bstrTypeName, pTypeAttr, pITypeInfo);
                    TKindPrev = TKIND_ALIAS;
                    break;

                case TKIND_DISPATCH:
                    if (pTypeAttr->wTypeFlags & TYPEFLAG_FDUAL)
                    {
                        HREFTYPE hreftype;
                        ITypeInfo_GetRefTypeOfImplType(pITypeInfo, -1U, &hreftype);
                        ITypeInfo_GetRefTypeInfo(pITypeInfo, hreftype, &pITypeInfo);
                        DumpTypeInfo(pOut, pITypeInfo, TRUE);
                    }
                    break;

                case TKIND_INTERFACE:
                    BufCat(pOut, L"\n");
                    BufCat(pOut, L"#ifndef __%ls_INTERFACE_DEFINED__\n", bstrTypeName);
                    BufCat(pOut, L"#define __%ls_INTERFACE_DEFINED__\n", bstrTypeName);
                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"\n/* %ls */\n", bstrComment);
                    DumpInterfaceType(pOut, bstrTypeName, pTypeAttr, pITypeInfo, fDispatch);
                    BufCat(pOut, L"#endif /* __%ls_INTERFACE_DEFINED__ */\n", bstrTypeName);
                    break;

                default:
                    break;
            }

            ITypeInfo_ReleaseTypeAttr(pITypeInfo, pTypeAttr);
        }

        SysFreeString(bstrTypeName);
        SysFreeString(bstrComment);
    }
}

/****************************************************************************
 *                                                                          *
 * Function: DumpAliasType                                                  *
 *                                                                          *
 * Purpose : Dump information for TKIND_ALIAS.                              *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpAliasType(POUTPUT pOut, BSTR bstrTypeName, LPTYPEATTR pTypeAttr, LPTYPEINFO pITypeInfo)
{
    BOOL fDone = FALSE;

    if ((pTypeAttr->tdescAlias.vt & VT_TYPEMASK) == VT_USERDEFINED)
    {
        LPTYPEINFO pITypeInfo2;

        if (ITypeInfo_GetRefTypeInfo(pITypeInfo, pTypeAttr->tdescAlias.hreftype, &pITypeInfo2) == S_OK)
        {
            TYPEATTR *pTypeAttr2;

            if (ITypeInfo_GetTypeAttr(pITypeInfo2, &pTypeAttr2) == S_OK)
            {
                if (pTypeAttr2->typekind == TKIND_RECORD ||
                    pTypeAttr2->typekind == TKIND_UNION ||
                    pTypeAttr2->typekind == TKIND_ENUM)
                {
                    BSTR bstrTypeName2;

                    if (ITypeInfo_GetDocumentation(pITypeInfo2, MEMBERID_NIL, &bstrTypeName2, 0, 0, 0) == S_OK)
                    {
                        if (pTypeAttr2->typekind == TKIND_ENUM)
                            DumpEnumType(pOut, bstrTypeName2, bstrTypeName, pTypeAttr2, pITypeInfo2);
                        else
                            DumpRecordType(pOut, bstrTypeName2, bstrTypeName, pTypeAttr2, pITypeInfo2);

                        SysFreeString(bstrTypeName2);
                    }

                    fDone = TRUE;
                }

                ITypeInfo_ReleaseTypeAttr(pITypeInfo2, pTypeAttr2);
            }

            ITypeInfo_Release(pITypeInfo2);
        }
    }

    if (!fDone)
    {
        WCHAR szType[256] = L"";

        GetTypeName(szType, NELEMS(szType), pTypeAttr->tdescAlias, NULL, pITypeInfo);
        BufCat(pOut, L"typedef %ls %ls;  /* ALIAS */\n", szType, bstrTypeName);
    }
}

/****************************************************************************
 *                                                                          *
 * Function: DumpClassType                                                  *
 *                                                                          *
 * Purpose : Dump information for TKIND_COCLASS.                            *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpClassType(POUTPUT pOut, BSTR bstrTypeName, LPTYPEATTR pTypeAttr)
{
    BufCat(pOut, L"DEFINE_GUID(CLSID_%ls,0x%08X,0x%04X,0x%04X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X);\n",
        bstrTypeName,
        pTypeAttr->guid.Data1,
        pTypeAttr->guid.Data2,
        pTypeAttr->guid.Data3,
        pTypeAttr->guid.Data4[0],
        pTypeAttr->guid.Data4[1],
        pTypeAttr->guid.Data4[2],
        pTypeAttr->guid.Data4[3],
        pTypeAttr->guid.Data4[4],
        pTypeAttr->guid.Data4[5],
        pTypeAttr->guid.Data4[6],
        pTypeAttr->guid.Data4[7]);
}

/****************************************************************************
 *                                                                          *
 * Function: DumpEnumType                                                   *
 *                                                                          *
 * Purpose : Dump information for TKIND_ENUM.                               *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpEnumType(POUTPUT pOut, BSTR bstrTagName, BSTR bstrTypeName, LPTYPEATTR pTypeAttr, LPTYPEINFO pITypeInfo)
{
    AllocName(bstrTagName);

    /* Check, just in case */
    if (pTypeAttr->cVars != 0)
    {
        BufCat(pOut, L"typedef enum %ls {\n", bstrTagName);

        /* Walk through all elements */
        for (int i = 0; i < pTypeAttr->cVars; i++)
        {
            VARDESC *pVarDesc;

            /* Get general information for a single element */
            if (ITypeInfo_GetVarDesc(pITypeInfo, i, &pVarDesc) == S_OK)
            {
                BSTR bstrVarName, bstrComment;

                /* Get name and description for a single element */
                if (ITypeInfo_GetDocumentation(pITypeInfo, pVarDesc->memid, &bstrVarName, &bstrComment, 0, 0) == S_OK)
                {
                    if (pVarDesc->varkind == VAR_CONST && V_VT(pVarDesc->lpvarValue) == VT_I4)
                        BufCat(pOut, L"\t%ls = %d,", bstrVarName, V_I4(pVarDesc->lpvarValue));
                    else
                        BufCat(pOut, L"\t%ls,", bstrVarName);

                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"  /* %ls */\n", bstrComment);
                    else
                        BufCat(pOut, L"\n");

                    SysFreeString(bstrVarName);
                    SysFreeString(bstrComment);
                }

                ITypeInfo_ReleaseVarDesc(pITypeInfo, pVarDesc);
            }
        }

        BufCat(pOut, L"} %ls;\n", bstrTypeName);
    }
}

/****************************************************************************
 *                                                                          *
 * Function: DumpRecordType                                                 *
 *                                                                          *
 * Purpose : Dump information for TKIND_RECORD or TKIND_UNION.              *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpRecordType(POUTPUT pOut, BSTR bstrTagName, BSTR bstrTypeName, LPTYPEATTR pTypeAttr, LPTYPEINFO pITypeInfo)
{
    AllocName(bstrTagName);

    /* Check, just in case */
    if (pTypeAttr->cVars != 0)
    {
        BufCat(pOut, L"typedef %ls %ls {\n", pTypeAttr->typekind == TKIND_UNION ? L"union" : L"struct", bstrTagName);

        /* Walk through all members */
        for (int i = 0; i < pTypeAttr->cVars; i++)
        {
            VARDESC *pVarDesc;

            /* Get general information for a single member */
            if (ITypeInfo_GetVarDesc(pITypeInfo, i, &pVarDesc) == S_OK)
            {
                BSTR bstrVarName, bstrComment;

                /* Get name and description for a single member */
                if (ITypeInfo_GetDocumentation(pITypeInfo, pVarDesc->memid, &bstrVarName, &bstrComment, 0, 0) == S_OK)
                {
                    WCHAR szType[256] = L"";

                    GetTypeName(szType, NELEMS(szType), pVarDesc->elemdescVar.tdesc, bstrVarName, pITypeInfo);
                    BufCat(pOut, L"\t%ls;", szType);

                    if (bstrComment != NULL && *bstrComment != L'\0')
                        BufCat(pOut, L"  /* %ls */\n", bstrComment);
                    else
                        BufCat(pOut, L"\n");

                    SysFreeString(bstrVarName);
                    SysFreeString(bstrComment);
                }

                ITypeInfo_ReleaseVarDesc(pITypeInfo, pVarDesc);
            }
        }

        BufCat(pOut, L"} %ls;\n", bstrTypeName);
    }
}

/****************************************************************************
 *                                                                          *
 * Function: DumpInterfaceType                                              *
 *                                                                          *
 * Purpose : Dump information for TKIND_INTERFACE.                          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void DumpInterfaceType(POUTPUT pOut, BSTR bstrTypeName, LPTYPEATTR pTypeAttr, LPTYPEINFO pITypeInfo, BOOL fDispatch)
{
    BufCat(pOut, L"DEFINE_GUID(IID_%ls,0x%08X,0x%04X,0x%04X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X,0x%02X);\n",
        bstrTypeName,
        pTypeAttr->guid.Data1,
        pTypeAttr->guid.Data2,
        pTypeAttr->guid.Data3,
        pTypeAttr->guid.Data4[0],
        pTypeAttr->guid.Data4[1],
        pTypeAttr->guid.Data4[2],
        pTypeAttr->guid.Data4[3],
        pTypeAttr->guid.Data4[4],
        pTypeAttr->guid.Data4[5],
        pTypeAttr->guid.Data4[6],
        pTypeAttr->guid.Data4[7]);

    // ? check for base interface (inheritance)
    BufCat(pOut, L"\n#undef INTERFACE\n");
    BufCat(pOut, L"#define INTERFACE  %ls\n", bstrTypeName);
    BufCat(pOut, L"DECLARE_INTERFACE(%ls) {\n", bstrTypeName);

    if (Fwd.cchBuf == 0) BufCat(&Fwd, L"\n/* Forward declarations */\n");
    BufCat(&Fwd, L"typedef interface %ls %ls;\n", bstrTypeName, bstrTypeName);

    BufCat(pOut, L"\t/* IUnknown methods */\n");
    BufCat(pOut, L"\tSTDMETHOD(QueryInterface)(THIS,REFIID,void**);\n");
    BufCat(pOut, L"\tSTDMETHOD_(ULONG,AddRef)(THIS);\n");
    BufCat(pOut, L"\tSTDMETHOD_(ULONG,Release)(THIS);\n");

    if (fDispatch)
    {
        BufCat(pOut, L"\t/* IDispatch methods */\n");
        BufCat(pOut, L"\tSTDMETHOD(GetTypeInfoCount)(THIS,UINT*);\n");
        BufCat(pOut, L"\tSTDMETHOD(GetTypeInfo)(THIS,UINT,LCID,ITypeInfo**);\n");
        BufCat(pOut, L"\tSTDMETHOD(GetIDsOfNames)(THIS,REFIID,LPOLESTR*,UINT,LCID,DISPID*);\n");
        BufCat(pOut, L"\tSTDMETHOD(Invoke)(THIS,DISPID,REFIID,LCID,WORD,DISPPARAMS*,VARIANT*,EXCEPINFO*,UINT*);\n");
    }

    BufCat(pOut, L"\t/* %ls methods */\n", bstrTypeName);
    for (int i = 0; i < pTypeAttr->cFuncs; i++)
    {
        FUNCDESC *pFuncDesc;

        if (ITypeInfo_GetFuncDesc(pITypeInfo, i, &pFuncDesc) == S_OK)
        {
            BSTR bstrFuncName = 0;

            if (ITypeInfo_GetDocumentation(pITypeInfo, pFuncDesc->memid, &bstrFuncName, 0, 0, 0) == S_OK)
            {
                WCHAR szFuncName[256];
                WCHAR szType[256];

                swprintf(szFuncName, NELEMS(szFuncName), L"%ls%ls", 
                    pFuncDesc->invkind == INVOKE_PROPERTYGET ? L"get_" :
                    pFuncDesc->invkind == INVOKE_PROPERTYPUT ? L"put_" :
                    pFuncDesc->invkind == INVOKE_PROPERTYPUTREF ? L"putref_" : L"",
                    bstrFuncName);

                /* Get description of return type */
                *szType = L'\0';
                GetTypeName(szType, NELEMS(szType), pFuncDesc->elemdescFunc.tdesc, NULL, pITypeInfo);

                if (wcscmp(szType, L"HRESULT") == 0)
                    BufCat(pOut, L"\tSTDMETHOD(%ls)", szFuncName);
                else
                    BufCat(pOut, L"\tSTDMETHOD_(%ls,%ls)", szType, szFuncName);

                if (pFuncDesc->cParams == 0)
                    BufCat(pOut, L"(THIS");
                else
                    BufCat(pOut, L"(THIS,");

                for (int k = 0; k < pFuncDesc->cParams; k++)
                {
                    *szType = L'\0';
                    GetTypeName(szType, NELEMS(szType), pFuncDesc->lprgelemdescParam[k].tdesc, NULL, pITypeInfo);
                    BufCat(pOut, L"%ls", szType);
                    if (k < pFuncDesc->cParams-1)
                        BufCat(pOut, L",");
                }

                BufCat(pOut, L");\n");

                SysFreeString(bstrFuncName);
            }

            ITypeInfo_ReleaseFuncDesc(pITypeInfo, pFuncDesc);
        }
    }

    BufCat(pOut, L"};\n\n");
    BufCat(pOut, L"#ifdef COBJMACROS\n");

    BufCat(pOut, L"#define %ls_QueryInterface(This,_1,_2)  (This)->lpVtbl->QueryInterface(This,_1,_2)\n", bstrTypeName);
    BufCat(pOut, L"#define %ls_AddRef(This)  (This)->lpVtbl->AddRef(This)\n", bstrTypeName);
    BufCat(pOut, L"#define %ls_Release(This)  (This)->lpVtbl->Release(This)\n", bstrTypeName);

    if (fDispatch)
    {
        BufCat(pOut, L"#define %ls_GetTypeInfoCount(This,_1)  (This)->lpVtbl->GetTypeInfoCount(This,_1)\n", bstrTypeName);
        BufCat(pOut, L"#define %ls_GetTypeInfo(This,_1,_2,_3)  (This)->lpVtbl->GetTypeInfo(This,_1,_2,_3)\n", bstrTypeName);
        BufCat(pOut, L"#define %ls_GetIDsOfNames(This,_1,_2,_3,_4,_5)  (This)->lpVtbl->GetIDsOfNames(This,_1,_2,_3,_4,_5)\n", bstrTypeName);
        BufCat(pOut, L"#define %ls_Invoke(This,_1,_2,_3,_4,_5,_6,_7,_8)  (This)->lpVtbl->Invoke(This,_1,_2,_3,_4,_5,_6,_7,_8)\n", bstrTypeName);
    }

    for (int i = 0; i < pTypeAttr->cFuncs; i++)
    {
        FUNCDESC *pFuncDesc;

        if (ITypeInfo_GetFuncDesc(pITypeInfo, i, &pFuncDesc) == S_OK)
        {
            BSTR bstrFuncName = 0;

            if (ITypeInfo_GetDocumentation(pITypeInfo, pFuncDesc->memid, &bstrFuncName, 0, 0, 0) == S_OK)
            {
                WCHAR szFuncName[256];

                swprintf(szFuncName, NELEMS(szFuncName), L"%ls%ls",
                    pFuncDesc->invkind == INVOKE_PROPERTYGET ? L"get_" :
                    pFuncDesc->invkind == INVOKE_PROPERTYPUT ? L"put_" :
                    pFuncDesc->invkind == INVOKE_PROPERTYPUTREF ? L"putref_" : L"",
                    bstrFuncName);

                BufCat(pOut, L"#define %ls_%ls(This", bstrTypeName, szFuncName);

                for (int k = 0; k < pFuncDesc->cParams; k++)
                    BufCat(pOut, L",_%d", k+1);

                BufCat(pOut, L")  (This)->lpVtbl->%ls(This", szFuncName);
                for (int k = 0; k < pFuncDesc->cParams; k++)
                    BufCat(pOut, L",_%d", k+1);

                BufCat(pOut, L")\n");

                SysFreeString(bstrFuncName);
            }

            ITypeInfo_ReleaseFuncDesc(pITypeInfo, pFuncDesc);
        }
    }

    BufCat(pOut, L"#endif /* COBJMACROS */\n\n");
}

/****************************************************************************
 *                                                                          *
 * Function: GetTypeName                                                    *
 *                                                                          *
 * Purpose : Get a C type from the given type description.                  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void GetTypeName(PTSTR pszBuf, size_t cchMaxBuf, TYPEDESC tdesc, BSTR bstrName, LPTYPEINFO pITypeInfo)
{
    if (tdesc.vt & VT_ARRAY)
        StrCat(pszBuf, cchMaxBuf, L"SAFEARRAY(");

    switch (tdesc.vt & VT_TYPEMASK)
    {
        case VT_I2: StrCat(pszBuf, cchMaxBuf, L"SHORT"); break;
        case VT_I4: StrCat(pszBuf, cchMaxBuf, L"LONG"); break;
        case VT_R4: StrCat(pszBuf, cchMaxBuf, L"float"); break;
        case VT_R8: StrCat(pszBuf, cchMaxBuf, L"double"); break;
        case VT_CY: StrCat(pszBuf, cchMaxBuf, L"CY"); break;
        case VT_DATE: StrCat(pszBuf, cchMaxBuf, L"DATE"); break;
        case VT_BSTR: StrCat(pszBuf, cchMaxBuf, L"BSTR"); break;
        case VT_DISPATCH: StrCat(pszBuf, cchMaxBuf, L"IDispatch*"); break;
        case VT_ERROR: StrCat(pszBuf, cchMaxBuf, L"SCODE"); break;
        case VT_BOOL: StrCat(pszBuf, cchMaxBuf, L"VARIANT_BOOL"); break;
        case VT_VARIANT: StrCat(pszBuf, cchMaxBuf, L"VARIANT"); break;
        case VT_UNKNOWN: StrCat(pszBuf, cchMaxBuf, L"IUnknown*"); break;
        case VT_DECIMAL: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {decimal} */"); break;
        case VT_I1: StrCat(pszBuf, cchMaxBuf, L"CHAR"); break;
        case VT_UI1: StrCat(pszBuf, cchMaxBuf, L"UCHAR"); break;
        case VT_UI2: StrCat(pszBuf, cchMaxBuf, L"USHORT"); break;
        case VT_UI4: StrCat(pszBuf, cchMaxBuf, L"ULONG"); break;
        case VT_I8: StrCat(pszBuf, cchMaxBuf, L"LONGLONG"); break;
        case VT_UI8: StrCat(pszBuf, cchMaxBuf, L"ULONGLONG"); break;
        case VT_INT: StrCat(pszBuf, cchMaxBuf, L"int"); break;
        case VT_UINT: StrCat(pszBuf, cchMaxBuf, L"UINT"); break;
        case VT_VOID: StrCat(pszBuf, cchMaxBuf, L"void"); break;
        case VT_HRESULT: StrCat(pszBuf, cchMaxBuf, L"HRESULT"); break;
        case VT_SAFEARRAY: StrCat(pszBuf, cchMaxBuf, L"SAFEARRAY*"); break;
        case VT_LPSTR: StrCat(pszBuf, cchMaxBuf, L"char*"); break;
        case VT_LPWSTR: StrCat(pszBuf, cchMaxBuf, L"WCHAR*"); break;
        case VT_RECORD: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {record} */"); break;
        case VT_FILETIME: StrCat(pszBuf, cchMaxBuf, L"FILETIME"); break;
        case VT_BLOB: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {blob} */"); break;
        case VT_STREAM: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {stream} */"); break;
        case VT_STORAGE: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {storage} */"); break;
        case VT_STREAMED_OBJECT: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {streamed object} */"); break;
        case VT_STORED_OBJECT: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {stored object} */"); break;
        case VT_BLOB_OBJECT: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {blob object} */"); break;
        case VT_CF: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {cf} */"); break;
        case VT_CLSID: StrCat(pszBuf, cchMaxBuf, L"/* NOT IMPLEMENTED: {CLSID} */"); break;
        case VT_CARRAY:
        {
            GetTypeName(pszBuf, cchMaxBuf, tdesc.lpadesc->tdescElem, NULL, pITypeInfo);
            StrCat(pszBuf, cchMaxBuf, L" %ls", bstrName); bstrName = NULL;
            for (int i = 0; i < tdesc.lpadesc->cDims; i++)
                StrCat(pszBuf, cchMaxBuf, L"[%u]", tdesc.lpadesc->rgbounds[i].cElements);
            break;
        }
        case VT_PTR:
        {
            GetTypeName(pszBuf, cchMaxBuf, *tdesc.lptdesc, NULL, pITypeInfo);
            StrCat(pszBuf, cchMaxBuf, L"*");
            break;
        }
        case VT_USERDEFINED:
        {
            if (ITypeInfo_GetRefTypeInfo(pITypeInfo, tdesc.hreftype, &pITypeInfo) == S_OK)
            {
                TYPEATTR *pTypeAttr;

                if (ITypeInfo_GetTypeAttr(pITypeInfo, &pTypeAttr) == S_OK)
                {
                    BSTR bstrTypeName;

                    if (ITypeInfo_GetDocumentation(pITypeInfo, MEMBERID_NIL, &bstrTypeName, 0, 0, 0) == S_OK)
                    {
                        if (pTypeAttr->typekind == TKIND_ENUM)
                            StrCat(pszBuf, cchMaxBuf, L"enum %ls", bstrTypeName);
                        else if (pTypeAttr->typekind == TKIND_RECORD)
                            StrCat(pszBuf, cchMaxBuf, L"struct %ls", bstrTypeName);
                        else if (pTypeAttr->typekind == TKIND_UNION)
                            StrCat(pszBuf, cchMaxBuf, L"union %ls", bstrTypeName);
                        else
                            StrCat(pszBuf, cchMaxBuf, L"%ls", bstrTypeName);

                        SysFreeString(bstrTypeName);
                    }

                    ITypeInfo_ReleaseTypeAttr(pITypeInfo, pTypeAttr);
                }

                ITypeInfo_Release(pITypeInfo);
            }
            break;
        }
    }

    if (bstrName != NULL)
        StrCat(pszBuf, cchMaxBuf, L" %ls", bstrName);

    if (tdesc.vt & VT_ARRAY)
        StrCat(pszBuf, cchMaxBuf, L")");
}

/****************************************************************************
 *                                                                          *
 * Function: StrCat                                                         *
 *                                                                          *
 * Purpose : Append text to a string (with basic size check).               *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void StrCat(PWSTR pszBuf, size_t cchMaxBuf, PCWSTR pcszFmt, ...)
{
    va_list args;

    va_start(args, pcszFmt);
    _vsnwprintf(pszBuf + wcslen(pszBuf), cchMaxBuf - wcslen(pszBuf), pcszFmt, args);
    va_end(args);
}

/****************************************************************************
 *                                                                          *
 * Function: BufCat                                                         *
 *                                                                          *
 * Purpose : Append text to a dynamic buffer.                               *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static void BufCat(POUTPUT pOut, PCWSTR pcszFmt, ...)
{
    static WCHAR achBuf[MAXLINE];
    va_list args;

    /* Format the text to append */
    va_start(args, pcszFmt);
    _vsnwprintf(achBuf, NELEMS(achBuf), pcszFmt, args);
    va_end(args);

    /* Increase the output buffer when needed */
    if (pOut->cchBuf + wcslen(achBuf) + 1 > pOut->cchMaxBuf)
    {
        PWSTR pchBuf = realloc(pOut->pchBuf, (pOut->cchMaxBuf + BUFSTEP) * sizeof(WCHAR));
        if (!pchBuf) return;  /* in case we are very unlucky... */
        pOut->pchBuf = pchBuf;
        pOut->cchMaxBuf += BUFSTEP;
    }

    /* Append the formatted text to the output buffer */
    wcscpy(pOut->pchBuf + pOut->cchBuf, achBuf);
    pOut->cchBuf += (DWORD)wcslen(achBuf);
}

/****************************************************************************
 *                                                                          *
 * Function: AllocName                                                      *
 *                                                                          *
 * Purpose : Add a new (tag) name to the linked list.                       *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static PNAME AllocName(const WCHAR *pcszName)
{
    PNAME pName;

    /* Already seen this name?! */
    if ((pName = LookupName(pcszName)) != NULL)
        return pName;

    /* Allocate a new node */
    pName = malloc(sizeof(*pName));
    if (!pName) return NULL;

    pName->pNext = NULL;
    pName->pszName = _wcsdup(pcszName);

    if (!g_pNameHead)
    {
        /* This is the first node. Start the list */
        g_pNameHead = pName;
    }
    else
    {
        PNAME pNameT;

        /* Find the end of the list and append the new node */
        for (pNameT = g_pNameHead; pNameT->pNext; pNameT = pNameT->pNext)
                ;

        pNameT->pNext = pName;
    }

    return pName;
}

/****************************************************************************
 *                                                                          *
 * Function: LookupName                                                     *
 *                                                                          *
 * Purpose : Search for a (tag) name in the linked list.                    *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static PNAME LookupName(const WCHAR *pcszName)
{
    PNAME pName;

    /* Search for the name in the linked list */
    for (pName = g_pNameHead; pName != NULL; pName = pName->pNext)
        if (wcscmp(pName->pszName, pcszName) == 0) return pName;

    /* Bah. Not found */
    return NULL;
}
