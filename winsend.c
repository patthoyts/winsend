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
#include "debug.h"

#ifndef DECLSPEC_EXPORT
#define DECLSPEC_EXPORT __declspec(dllexport)
#endif /* ! DECLSPEC_EXPORT */

/* Should be defined in WTypes.h but mingw 1.0 is missing them */
#ifndef _ROTFLAGS_DEFINED
#define _ROTFLAGS_DEFINED
#define ROTFLAGS_REGISTRATIONKEEPSALIVE   0x01
#define ROTFLAGS_ALLOWANYCLIENT           0x02
#endif /* ! _ROTFLAGS_DEFINED */

#define WINSEND_PACKAGE_VERSION   "0.4"
#define WINSEND_PACKAGE_NAME      "winsend"
#define WINSEND_CLASS_NAME        "TclEval"

/* Package information structure.
 * This is used to keep interpreter specific details for use when releasing
 * the package resources upon interpreter deletion or package removal.
 */
typedef struct WinsendPkg_t {
    Tcl_Obj *appname;           /* the registered application name */
    DWORD ROT_cookie;           /* ROT cookie returned on registration */
    LPUNKNOWN obj;              /* Interface for the registration object */
    Tcl_Command token;          /* Winsend command token */
} WinsendPkg;

static void Winsend_InterpDeleteProc (ClientData clientData, Tcl_Interp *interp);
static void Winsend_PkgDeleteProc (ClientData clientData);
static int Winsend_CmdProc(ClientData clientData, Tcl_Interp *interp,
                           int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdInterps(ClientData clientData, Tcl_Interp *interp,
                              int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdSend(ClientData clientData, Tcl_Interp *interp,
                           int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdAppname(ClientData clientData, Tcl_Interp *interp,
                              int objc, Tcl_Obj *CONST objv[]);
static int Winsend_CmdTest(ClientData clientData, Tcl_Interp *interp,
                           int objc, Tcl_Obj *CONST objv[]);
static int Winsend_ObjSendCmd(LPDISPATCH pdispInterp, Tcl_Interp *interp, 
                              int objc, Tcl_Obj *CONST objv[]);
static HRESULT BuildMoniker(LPCOLESTR name, LPMONIKER *pmk);
static Tcl_Obj* Winsend_Win32ErrorObj(HRESULT hrError);

/* -------------------------------------------------------------------------
 * DllMain
 * -------------------------------------------------------------------------
 */

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

/* -------------------------------------------------------------------------
 * Winsend_Init
 * -------------------------------------------------------------------------
 */

EXTERN_C int DECLSPEC_EXPORT
Winsend_Init(Tcl_Interp* interp)
{
    HRESULT hr = S_OK;
    int r = TCL_OK;
    WinsendPkg *pkg = NULL;

#ifdef USE_TCL_STUBS
    Tcl_InitStubs(interp, "8.3", 0);
#endif

    pkg = (WinsendPkg*)Tcl_Alloc(sizeof(WinsendPkg));
    if (pkg == NULL) {
        Tcl_SetResult(interp, "out of memory", TCL_STATIC);
        return TCL_ERROR;
    }
    memset(pkg, 0, sizeof(WinsendPkg));

    /* Initialize COM */
    hr = CoInitialize(0);
    if (FAILED(hr)) {
        Tcl_SetResult(interp, "failed to initialize the " WINSEND_PACKAGE_NAME " package", TCL_STATIC);
        return TCL_ERROR;
    }

    /* Create our registration object. */
    hr = WinSendCom_CreateInstance(interp, &IID_IUnknown, (void**)&pkg->obj);

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
                                                ROTFLAGS_REGISTRATIONKEEPSALIVE,
                                                pkg->obj,
                                                pmk,
                                                &pkg->ROT_cookie);
                    if (SUCCEEDED(hr)) {
                        pkg->appname = Tcl_NewUnicodeObj(oleName, -1);
                        Tcl_IncrRefCount(pkg->appname);
                    }
                    pmk->lpVtbl->Release(pmk);
                }

                /* If the moniker was registered, unregister the duplicate and
                 * try again.
                 */
                if (hr == MK_S_MONIKERALREADYREGISTERED)
                    pROT->lpVtbl->Revoke(pROT, pkg->ROT_cookie);

            } while (hr == MK_S_MONIKERALREADYREGISTERED);

            pROT->lpVtbl->Release(pROT);
        }

        /* Create our winsend command */
        if (SUCCEEDED(hr)) {
            pkg->token = Tcl_CreateObjCommand(interp,
                                              "winsend",
                                              Winsend_CmdProc,
                                              (ClientData)pkg,
                                              (Tcl_CmdDeleteProc*)NULL);
        }

        /* Create an exit procedure to handle unregistering when the
         * Tcl interpreter terminates.
         */
        Tcl_CallWhenDeleted(interp, Winsend_InterpDeleteProc, (ClientData)pkg);
        Tcl_CreateExitHandler(Winsend_PkgDeleteProc, (ClientData)pkg);
        r = Tcl_PkgProvide(interp, WINSEND_PACKAGE_NAME, WINSEND_PACKAGE_VERSION);

    }
    if (FAILED(hr))
    {
        Tcl_SetObjResult(interp, Winsend_Win32ErrorObj(hr));
        r = TCL_ERROR;
    } else {
        Tcl_SetResult(interp, "", TCL_STATIC);
    }

    return r;
}

/* -------------------------------------------------------------------------
 * Winsend_SafeInit
 * ------------------------------------------------------------------------- */

EXTERN_C int DECLSPEC_EXPORT
Winsend_SafeInit(Tcl_Interp* interp)
{
    Tcl_SetResult(interp, "not permitted in safe interp", TCL_STATIC);
    return TCL_ERROR;
}

/* -------------------------------------------------------------------------
 * Winsend_InterpDeleteProc
 * -------------------------------------------------------------------------
 * Description:
 *  Called when the interpreter is deleted, this procedure clean up the COM
 *  registration for us. We need to revoke our registered object, and release
 *  our own object reference. The WinSendCom object should delete itself now.
 */

static void 
Winsend_InterpDeleteProc (ClientData clientData, Tcl_Interp *interp)
{
    WinsendPkg *pkg = (WinsendPkg*)clientData;
    LPRUNNINGOBJECTTABLE pROT = NULL;
    HRESULT hr = S_OK;

    LTRACE(_T("Winsend_InterpDeleteProc( {ROT_cookie: 0x%08X, obj: 0x%08X})\n"),
           pkg->ROT_cookie, pkg->obj);

    /* Lock the package structure in memory */
    Tcl_Preserve((ClientData)pkg);

    if (pkg->ROT_cookie != 0) {
        hr = GetRunningObjectTable(0, &pROT);
        if (SUCCEEDED(hr))
        {
            hr = pROT->lpVtbl->Revoke(pROT, pkg->ROT_cookie);
            pROT->lpVtbl->Release(pROT);
            pkg->ROT_cookie = 0;
        }
        _ASSERTE(SUCCEEDED(hr));
    }

    /* Remove the winsend command and release the interp pointer */
    /* I get assertions from this from within tcl 8.3 - it's not 
     * really needed anyway
     */
    /* if (interp != NULL) {
        Tcl_DeleteCommandFromToken(interp, pkg->token);
        pkg->token = 0;
    }*/

    /* release the appname object */
    if (pkg->appname != NULL) {
        Tcl_DecrRefCount(pkg->appname);
        pkg->appname = NULL;
    }

    /* Release the registration object */
    pkg->obj->lpVtbl->Release(pkg->obj);
    pkg->obj = NULL;

    /* unlock the package data structure. */
    Tcl_Release((ClientData)pkg);
}

/* -------------------------------------------------------------------------
 * Winsend_PkgDeleteProc
 * -------------------------------------------------------------------------
 * Description:
 *  Called when the package is removed, we clean up any outstanding memory
 *  and release any handles.
 */

static void
Winsend_PkgDeleteProc(ClientData clientData)
{
    WinsendPkg *pkg = (WinsendPkg*)clientData;
    LTRACE(_T("Winsend_PkgDeleteProc( {ROT_cookie: 0x%08X, obj: 0x%08X})\n"),
           pkg->ROT_cookie, pkg->obj);

    if (pkg->obj != NULL) {
        Winsend_InterpDeleteProc(clientData, NULL);
    }

    _ASSERTE(pkg->obj == NULL);
    Tcl_Free((char*)pkg);

    CoUninitialize();
}

/* -------------------------------------------------------------------------
 * Winsend_CmdProc
 * -------------------------------------------------------------------------
 * Description:
 *  The Tcl winsend proc is first processed here and then passed on depending
 *  upong the command argument.
 */
static int
Winsend_CmdProc(ClientData clientData, Tcl_Interp *interp,
                    int objc, Tcl_Obj *CONST objv[])
{
    enum {WINSEND_INTERPS, WINSEND_SEND, WINSEND_APPNAME, WINSEND_TEST};
    static char* cmds[] = { "interps", "send", "appname", "test", NULL };
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
        case WINSEND_APPNAME:
            r = Winsend_CmdAppname(clientData, interp, objc, objv);
            break;
        case WINSEND_TEST:
            r = Winsend_CmdTest(clientData, interp, objc, objv);
            break;
        }
    }
    return r;
}

/* -------------------------------------------------------------------------
 * Winsend_CmdInterps
 * -------------------------------------------------------------------------
 * Description:
 *  Iterate over the running object table and identify all the Tcl registered
 *  objects. We build a list using the tail part of the file moniker and return
 *  this as the list of registered interpreters.
 */
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
                                LPOLESTR p = olestr + wcslen(oleszStub);
                                if (*p)
                                    r = Tcl_ListObjAppendElement(interp, objList, Tcl_NewUnicodeObj(p + 1, -1));
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

/* -------------------------------------------------------------------------
 * Winsend_CmdSend
 * -------------------------------------------------------------------------
 */
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
 * Winsend_CmdAppname
 * -------------------------------------------------------------------------
 * Description:
 *  If called with no additional parameters returns the registered application
 *  name. If a new name is given then we attempt to register the new app name
 *  If the name is already registered then we cancel the new registration and
 *  continue using the old appname.
 */
static int
Winsend_CmdAppname(ClientData clientData, Tcl_Interp *interp,
                   int objc, Tcl_Obj *CONST objv[])
{
    int r = TCL_OK;
    HRESULT hr = S_OK;
    WinsendPkg* pkg = (WinsendPkg*)clientData;
    
    if (objc == 2) {

        Tcl_SetObjResult(interp, Tcl_DuplicateObj(pkg->appname));

    } else if (objc > 3) {

        Tcl_WrongNumArgs(interp, 2, objv, "?appname?");
        r = TCL_ERROR;

    } else {

        LPRUNNINGOBJECTTABLE pROT = NULL;
        hr = GetRunningObjectTable(0, &pROT);
        if (SUCCEEDED(hr)) {
            LPMONIKER pmk;
            LPOLESTR szNewName;

            szNewName = Tcl_GetUnicode(objv[2]);

            /* construct a new moniker */
            hr = BuildMoniker(szNewName, &pmk);
            if (SUCCEEDED(hr)) {
                DWORD cookie = 0;

                /* register the new name */
                hr = pROT->lpVtbl->Register(pROT, ROTFLAGS_REGISTRATIONKEEPSALIVE, pkg->obj, pmk, &cookie);
                if (SUCCEEDED(hr)) {
                    if (hr == MK_S_MONIKERALREADYREGISTERED) {
                        pROT->lpVtbl->Revoke(pROT, cookie);
                        Tcl_SetObjResult(interp, Winsend_Win32ErrorObj(hr));
                        r = TCL_ERROR;
                    } else {

                        /* revoke the old name */
                        hr = pROT->lpVtbl->Revoke(pROT, pkg->ROT_cookie);
                        if (SUCCEEDED(hr)) {

                            /* update the package structure */
                            pkg->ROT_cookie = cookie;
                            Tcl_SetUnicodeObj(pkg->appname, szNewName, -1);
                            Tcl_SetObjResult(interp, pkg->appname);
                        }
                    }
                }
                pmk->lpVtbl->Release(pmk);
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

/* -------------------------------------------------------------------------
 * Helpers.
 * -------------------------------------------------------------------------
 */

/* -------------------------------------------------------------------------
 * Winsend_ObjSendCmd
 * -------------------------------------------------------------------------
 * Description:
 *   Perform an interface call to the server object. We convert the Tcl arguments
 *   into a BSTR using 'concat'. The result should be a BSTR that we can set as the
 *   interp's result string.
 */   
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

/* -------------------------------------------------------------------------
 * BuildMoniker
 * -------------------------------------------------------------------------
 * Description:
 *  Construct a moniker from the given name. This ensures that all our
 *  monikers have the same prefix.
 */
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

/* -------------------------------------------------------------------------
 * Winsend_Win32ErrorObj
 * -------------------------------------------------------------------------
 * Description:
 *  Convert COM or Win32 API errors into Tcl strings.
 */
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
