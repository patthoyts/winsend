/* winsend.c - Copyright (C) 2002 Pat Thoyts <Pat.Thoyts@bigfoot.com>
 */

/*
 * TODO:
 *   Put the ROT cookie into a client data structure.
 *   Use a C based COM object.
 *   Arrange to register the dispinterface ID?
 *   Check the tkWinSend.c file and implement a Tk send.
 */

static const char rcsid[] =
"$Id$";

#include <windows.h>
#include <ole2.h>
#include <tchar.h>
#include <tcl.h>

#include <initguid.h>
#include "WinSendCom.h"

#ifndef DECLSPEC_EXPORT
#define DECLSPEC_EXPORT __declspec(dllexport)
#endif /* ! DECLSPEC_EXPORT */

/* Should be defined in WTypes.h but mingw 1.0 is missing them */
#ifndef _ROTFLAGS_DEFINED
#define ROTFLAGS_REGISTRATIONKEEPSALIVE   0x01
#define ROTFLAGS_ALLOWANYCLIENT           0x02
#endif /* ! _ROTFLAGS_DEFINED */

#define WINSEND_PACKAGE_VERSION   "0.3"
#define WINSEND_PACKAGE_NAME      "winsend"
#define WINSEND_CLASS_NAME        "TclEval"

DWORD g_dwROTCookie;

static void Winsend_PkgDeleteProc _ANSI_ARGS_((ClientData clientData));
static int Winsend_CmdProc(ClientData clientData, Tcl_Interp *interp,
                           int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdInterps(ClientData clientData, Tcl_Interp *interp,
                              int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdSend(ClientData clientData, Tcl_Interp *interp,
                           int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdTest(ClientData clientData, Tcl_Interp *interp,
                           int objc, Tcl_Obj *CONST objv[]);
static int Winsend_ObjSendCmd(LPDISPATCH pdispInterp, Tcl_Interp *interp, 
                              int objc, Tcl_Obj *CONST objv[]);
static HRESULT BuildMoniker(LPCOLESTR name, LPMONIKER *pmk);
static Tcl_Obj* Winsend_Win32ErrorObj(HRESULT hrError);

// -------------------------------------------------------------------------
// DllMain
// -------------------------------------------------------------------------

EXTERN_C BOOL APIENTRY
DllMain(HANDLE hModule, DWORD dwReason, LPVOID lpReserved)
{
    switch(dwReason)
    {
        case DLL_PROCESS_ATTACH:
            break;
        case DLL_PROCESS_DETACH:
            break;
    }
    
    return TRUE;
}

// -------------------------------------------------------------------------
// Winsend_Init
// -------------------------------------------------------------------------

EXTERN_C int DECLSPEC_EXPORT
Winsend_Init(Tcl_Interp* interp)
{
    HRESULT hr = S_OK;
    int r = TCL_OK;
    IUnknown *pUnk = NULL;

#ifdef USE_TCL_STUBS
    Tcl_InitStubs(interp, "8.3", 0);
#endif

    // Initialize COM
    hr = CoInitialize(0);
    if (FAILED(hr)) {
        Tcl_SetResult(interp, "failed to initialize the " WINSEND_PACKAGE_NAME " package", TCL_STATIC);
        return TCL_ERROR;
    }

    // Create our registration object.
    hr = WinSendCom_CreateInstance(interp, &IID_IUnknown, (void**)&pUnk);

    if (SUCCEEDED(hr))
    {
        LPRUNNINGOBJECTTABLE pROT = NULL;
        Tcl_Obj *name = NULL;

        if (Tcl_Eval(interp, "file tail $::argv0") == TCL_OK)
            name = Tcl_GetObjResult(interp);
        if (name == NULL)
            name = Tcl_NewUnicodeObj(L"tcl", 3);
        
        hr = GetRunningObjectTable(0, &pROT);
        if (SUCCEEDED(hr))
        {
            int n = 1;
            LPMONIKER pmk = NULL;
            OLECHAR oleName[64];
            oleName[0] = 0;

            do
            {
                if (n > 1)
                    swprintf(oleName, L"%s #%u", Tcl_GetUnicode(name), n);
                else
                    wcscpy(oleName, Tcl_GetUnicode(name));
                n++;

                hr = BuildMoniker(oleName, &pmk);

                if (SUCCEEDED(hr)) {
                    hr = pROT->lpVtbl->Register(pROT,
                                                ROTFLAGS_REGISTRATIONKEEPSALIVE
                                                | ROTFLAGS_ALLOWANYCLIENT,
                                                pUnk, pmk, &g_dwROTCookie);
                    pmk->lpVtbl->Release(pmk);
                }

                /* If the moniker was registered, unregister the duplicate and
                 * try again.
                 */
                if (hr == MK_S_MONIKERALREADYREGISTERED)
                    pROT->lpVtbl->Revoke(pROT, g_dwROTCookie);

            } while (hr == MK_S_MONIKERALREADYREGISTERED);

            pROT->lpVtbl->Release(pROT);
        }

        pUnk->lpVtbl->Release(pUnk);

        // Create our winsend command
        if (SUCCEEDED(hr)) {
            Tcl_CreateObjCommand(interp, "winsend", Winsend_CmdProc, (ClientData)0, (Tcl_CmdDeleteProc*)0);
        }

        /* Create an exit procedure to handle unregistering when the
         * Tcl interpreter terminates.
         */
        Tcl_CreateExitHandler(Winsend_PkgDeleteProc, NULL);
        r = Tcl_PkgProvide(interp, WINSEND_PACKAGE_NAME, WINSEND_PACKAGE_VERSION);

    }
    if (FAILED(hr))
    {
        // TODO: better error handling - rip code from w32_exception.h
        Tcl_Obj *err = Winsend_Win32ErrorObj(hr);
        Tcl_SetObjResult(interp, err);
        r = TCL_ERROR;
    } else {
        Tcl_SetResult(interp, "", TCL_STATIC);
    }

    return r;
}

// -------------------------------------------------------------------------
// Winsend_SafeInit
// -------------------------------------------------------------------------

EXTERN_C int DECLSPEC_EXPORT
Winsend_SafeInit(Tcl_Interp* interp)
{
    Tcl_SetResult(interp, "not permitted in safe interp", TCL_STATIC);
    return TCL_ERROR;
}

// -------------------------------------------------------------------------
// WinsendExitProc
// -------------------------------------------------------------------------

static void
Winsend_PkgDeleteProc(ClientData clientData)
{
    LPRUNNINGOBJECTTABLE pROT = NULL;
    HRESULT hr = GetRunningObjectTable(0, &pROT);
    if (SUCCEEDED(hr))
    {
        hr = pROT->lpVtbl->Revoke(pROT, g_dwROTCookie);
        pROT->lpVtbl->Release(pROT);
    }
    //ASSERT
}

// -------------------------------------------------------------------------
// Winsend_CmdProc
// -------------------------------------------------------------------------

static int
Winsend_CmdProc(ClientData clientData, Tcl_Interp *interp,
                    int objc, Tcl_Obj *CONST objv[])
{
    enum {WINSEND_INTERPS, WINSEND_SEND, WINSEND_TEST};
    static char* cmds[] = { "interps", "send", "test", NULL };
    int index = 0, r = TCL_OK;

    if (objc < 2) {
        Tcl_WrongNumArgs(interp, 1, objv, "command ?args ...?");
        return TCL_ERROR;
    }

    r = Tcl_GetIndexFromObj(interp, objv[1], cmds, "command", 0, &index);
    if (r == TCL_OK)
    {
        switch (index)
        {
        case WINSEND_INTERPS:
            r = Winsend_CmdInterps(clientData, interp, objc, objv);
            break;
        case WINSEND_SEND:
            r = Winsend_CmdSend(clientData, interp, objc, objv);
            break;
        case WINSEND_TEST:
            r = Winsend_CmdTest(clientData, interp, objc, objv);
            break;
        }
    }
    return r;
}

// -------------------------------------------------------------------------
// Winsend_CmdInterps
// -------------------------------------------------------------------------

static int 
Winsend_CmdInterps(ClientData clientData, Tcl_Interp *interp,
                   int objc, Tcl_Obj *CONST objv[])
{
    LPRUNNINGOBJECTTABLE pROT = NULL;
    LPCOLESTR oleszStub = OLESTR("\\\\.\\TclInterp");
    HRESULT hr = S_OK;
    Tcl_Obj *objList;
    int r = TCL_OK;
    
    if (objc != 2) {
        Tcl_WrongNumArgs(interp, 1, objv, "interps");
        r = TCL_ERROR;
    } else {

        hr = GetRunningObjectTable(0, &pROT);
        if(SUCCEEDED(hr))
        {
            IBindCtx* pBindCtx = NULL;
            objList = Tcl_NewListObj(0, NULL);
            hr = CreateBindCtx(0, &pBindCtx);
            if (SUCCEEDED(hr))
            {

                IEnumMoniker* pEnum;
                hr = pROT->lpVtbl->EnumRunning(pROT, &pEnum);
                if(SUCCEEDED(hr))
                {
                    IMoniker* pmk = NULL;
                    while (pEnum->lpVtbl->Next(pEnum, 1, &pmk, (ULONG*)NULL) == S_OK) 
                    {
                        LPOLESTR olestr;
                        hr = pmk->lpVtbl->GetDisplayName(pmk, pBindCtx, NULL, &olestr);
                        if (SUCCEEDED(hr))
                        {
                            IMalloc *pMalloc = NULL;

                            if (wcsncmp(olestr, oleszStub, wcslen(oleszStub)) == 0)
                            {
                                LPOLESTR p = olestr + wcslen(oleszStub) + 1;
                                r = Tcl_ListObjAppendElement(interp, objList, Tcl_NewUnicodeObj(p, -1));
                            }

                            hr = CoGetMalloc(1, &pMalloc);
                            if (SUCCEEDED(hr))
                            {
                                pMalloc->lpVtbl->Free(pMalloc, (void*)olestr);
                                pMalloc->lpVtbl->Release(pMalloc);
                            }
                        }
                        pmk->lpVtbl->Release(pmk);
                    }
                    pEnum->lpVtbl->Release(pEnum);
                }
                pBindCtx->lpVtbl->Release(pBindCtx);
            }
            pROT->lpVtbl->Release(pROT);
        }
    }

    if (FAILED(hr)) {
        /* expire the list if set */
        if (objList != NULL) {
            Tcl_DecrRefCount(objList);
        }
        Tcl_SetObjResult(interp, Winsend_Win32ErrorObj(hr));
        r = TCL_ERROR;
    }

    if (r == TCL_OK)
        Tcl_SetObjResult(interp, objList);

    return r;
}

// -------------------------------------------------------------------------
// Winsend_CmdSend
// -------------------------------------------------------------------------

static int
Winsend_CmdSend(ClientData clientData, Tcl_Interp *interp,
                   int objc, Tcl_Obj *CONST objv[])
{
    int r = TCL_OK;
    HRESULT hr = S_OK;
    
    if (objc < 4) {

        Tcl_WrongNumArgs(interp, 1, objv, "app cmd ?arg arg arg?");
        r = TCL_ERROR;

    } else {

        LPRUNNINGOBJECTTABLE pROT = NULL;

        hr = GetRunningObjectTable(0, &pROT);
        if (SUCCEEDED(hr))
        {
            IBindCtx* pBindCtx = NULL;
            hr = CreateBindCtx(0, &pBindCtx);
            if (SUCCEEDED(hr))
            {
                LPMONIKER pmk = NULL;
                hr = BuildMoniker(Tcl_GetUnicode(objv[2]), &pmk);
                if (SUCCEEDED(hr))
                {
                    IUnknown* punkInterp = NULL;
                    IDispatch* pdispInterp = NULL;
                    hr = pROT->lpVtbl->IsRunning(pROT, pmk);
                    hr = pmk->lpVtbl->BindToObject(pmk, pBindCtx, NULL, &IID_IUnknown, (void**)&punkInterp);
                    //hr = pROT->lpVtbl->GetObject(pROT, pmk, &punkInterp);
                    if (SUCCEEDED(hr))
                        hr = punkInterp->lpVtbl->QueryInterface(punkInterp, &IID_IDispatch, (void**)&pdispInterp);
                    if (SUCCEEDED(hr))
                    {
                        r = Winsend_ObjSendCmd(pdispInterp, interp, objc, objv);
                        pdispInterp->lpVtbl->Release(pdispInterp);
                        punkInterp->lpVtbl->Release(punkInterp);
                    }
                    pmk->lpVtbl->Release(pmk);
                }
                pBindCtx->lpVtbl->Release(pBindCtx);
            }
            pROT->lpVtbl->Release(pROT);
        }
        if (FAILED(hr))
        {
            Tcl_SetObjResult(interp, Winsend_Win32ErrorObj(hr));
            r = TCL_ERROR;
        }
    }

    return r;
}

/* -------------------------------------------------------------------------
 * Winsend_CmdTest
 * -------------------------------------------------------------------------
 */
static int 
Winsend_CmdTest(ClientData clientData, Tcl_Interp *interp,
                 int objc, Tcl_Obj *CONST objv[])
{
    enum {WINSEND_TEST_ERROR};
    static char* cmds[] = { "error", NULL };
    int index = 0, r = TCL_OK;

    if (objc < 3) {
        Tcl_WrongNumArgs(interp, 2, objv, "command ?args ...?");
        return TCL_ERROR;
    }

    r = Tcl_GetIndexFromObj(interp, objv[2], cmds, "command", 0, &index);
    if (r == TCL_OK)
    {
        switch (index)
        {
        case WINSEND_TEST_ERROR:
            {
                Tcl_Obj *err = Winsend_Win32ErrorObj(E_INVALIDARG);
                Tcl_SetObjResult(interp, err);
                r = TCL_ERROR;
            }
            break;
        }
    }
    return r;
}

// -------------------------------------------------------------------------
// Helpers.
// -------------------------------------------------------------------------

// -------------------------------------------------------------------------
// Winsend_ObjSendCmd
// -------------------------------------------------------------------------
// Description:
//   
static int
Winsend_ObjSendCmd(LPDISPATCH pdispInterp, Tcl_Interp *interp, 
                   int objc, Tcl_Obj *CONST objv[])
{
    VARIANT vCmd, vResult;
    DISPPARAMS dp;
    EXCEPINFO ei;
    UINT uiErr = 0;
    HRESULT hr = S_OK;
    Tcl_Obj *cmd = NULL;

    cmd = Tcl_ConcatObj(objc - 3, &objv[3]);

    VariantInit(&vCmd);
    VariantInit(&vResult);
    memset(&dp, 0, sizeof(dp));
    memset(&ei, 0, sizeof(ei));

    vCmd.vt = VT_BSTR;
    vCmd.bstrVal = SysAllocString(Tcl_GetUnicode(cmd));

    dp.cArgs = 1;
    dp.rgvarg = &vCmd;

    hr = pdispInterp->lpVtbl->Invoke(pdispInterp, 1, &IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dp, &vResult, &ei, &uiErr);
    if (SUCCEEDED(hr))
    {
        hr = VariantChangeType(&vResult, &vResult, 0, VT_BSTR);
        if (SUCCEEDED(hr))
            Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(vResult.bstrVal, -1));
    }

    VariantClear(&vCmd);

    return (SUCCEEDED(hr) ? TCL_OK : TCL_ERROR);
}

// Construct a Tcl interpreter moniker from an interp name.
static HRESULT
BuildMoniker(LPCOLESTR name, LPMONIKER *pmk)
{
    LPCOLESTR oleszStub = OLESTR("\\\\.\\TclInterp");
    LPMONIKER pmkClass = NULL;
    HRESULT hr = CreateFileMoniker(oleszStub, &pmkClass);
    if (SUCCEEDED(hr)) {
        LPMONIKER pmkItem = NULL;
        hr = CreateFileMoniker(name, &pmkItem);
        if (SUCCEEDED(hr)) {
            LPMONIKER pmkJoint = NULL;
            hr = pmkClass->lpVtbl->ComposeWith(pmkClass, pmkItem, FALSE, &pmkJoint);
            if (SUCCEEDED(hr)) {
                *pmk = pmkJoint;
                (*pmk)->lpVtbl->AddRef(*pmk);
                pmkJoint->lpVtbl->Release(pmkJoint);
            }
            pmkItem->lpVtbl->Release(pmkItem);
        }
        pmkClass->lpVtbl->Release(pmkClass);
    }
    return hr;
}

static Tcl_Obj*
Winsend_Win32ErrorObj(HRESULT hrError)
{
    LPTSTR lpBuffer = NULL, p = NULL;
	TCHAR  sBuffer[30];
	Tcl_Obj* err_obj = NULL;

    FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
                  NULL, hrError, LANG_NEUTRAL, (LPTSTR)&lpBuffer, 0, NULL);

    if (lpBuffer == NULL) {
		lpBuffer = sBuffer;
        wsprintf(sBuffer, _T("Error Code: %08lX"), hrError);
    }

    if ((p = _tcsrchr(lpBuffer, _T('\r'))) != NULL)
        *p = _T('\0');

#ifdef _UNICODE
		err_obj = Tcl_NewUnicodeObj(lpBuffer, wcslen(lpBuffer));
#else
		err_obj = Tcl_NewStringObj(lpBuffer, strlen(lpBuffer));
#endif

	if (lpBuffer != sBuffer)    
        LocalFree((HLOCAL)lpBuffer);

	return err_obj;
}
