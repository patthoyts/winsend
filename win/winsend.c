/* winsend.c - Copyright (C) 2002 Pat Thoyts <Pat.Thoyts@bigfoot.com>
 */

/*
 * TODO:
 *   Arrange to register the dispinterface ID?
 *   Check the tkWinSend.c file and implement a Tk send then raise a TIP?
 */

static const char rcsid[] =
"$Id$";

#include <windows.h>
#include <process.h>
#include <ole2.h>
#include <tchar.h>
#include <tcl.h>

#include <initguid.h>
#include "WinSendCom.h"
#include "debug.h"

#ifdef BUILD_winsend
#undef TCL_STORAGE_CLASS
#define TCL_STORAGE_CLASS DLLEXPORT
#endif /* BUILD_winsend */

/* Tcl 8.4 CONST support */
#ifndef CONST84
#define CONST84
#endif

/* Should be defined in WTypes.h but mingw 1.0 is missing them */
#ifndef _ROTFLAGS_DEFINED
#define _ROTFLAGS_DEFINED
#define ROTFLAGS_REGISTRATIONKEEPSALIVE   0x01
#define ROTFLAGS_ALLOWANYCLIENT           0x02
#endif /* ! _ROTFLAGS_DEFINED */

#define WINSEND_PACKAGE_VERSION   VERSION
#define WINSEND_PACKAGE_NAME      "winsend"
#define WINSEND_CLASS_NAME        "TclEval"
#define WINSEND_REGISTRATION_BASE L"TclEval"

#define MK_E_MONIKERALREADYREGISTERED \
    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x02A1)

/* Package information structure.
 * This is used to keep interpreter specific details for use when releasing
 * the package resources upon interpreter deletion or package removal.
 */
typedef struct WinsendPkg_t {
    Tcl_Obj *appname;           /* the registered application name */
    DWORD ROT_cookie;           /* ROT cookie returned on registration */
    LPUNKNOWN obj;              /* Interface for the registration object */
    Tcl_Command token;          /* Winsend command token */
    IStream *stream;            /* stream to use for thread marshalling */
    HANDLE   handle;            /* handle to the asynchronous registration thread */
    DWORD    threadid;          /* the thread id of the async registration thread */
    HRESULT  reg_err;           /* the error code returned by the async registration */
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
static HRESULT RegisterName(LPCOLESTR szName, IUnknown *pUnknown, BOOL bForce,
                            LPOLESTR *pNewName, DWORD cbNewName, DWORD *pdwCookie);
static Tcl_Obj* Winsend_Win32ErrorObj(HRESULT hrError);

/* -------------------------------------------------------------------------
 * Winsend_Init
 * -------------------------------------------------------------------------
 */

EXTERN int
Winsend_Init(Tcl_Interp* interp)
{
    HRESULT hr = S_OK;
    int r = TCL_OK;
    Tcl_Obj *name = NULL;
    WinsendPkg *pkg = NULL;

#ifdef USE_TCL_STUBS
    Tcl_InitStubs(interp, "8.1", 0);
#endif

    /* allocate our package structure */
    pkg = (WinsendPkg*)Tcl_Alloc(sizeof(WinsendPkg));
    if (pkg == NULL) {
        Tcl_SetResult(interp, "out of memory", TCL_STATIC);
        return TCL_ERROR;
    }
    memset(pkg, 0, sizeof(WinsendPkg));
    pkg->handle = INVALID_HANDLE_VALUE;

    /* discover our root name - default to 'tcl'*/
    if (Tcl_Eval(interp, "file tail $::argv0") == TCL_OK)
        name = Tcl_GetObjResult(interp);
    if (name == NULL)
        name = Tcl_NewUnicodeObj(L"tcl", 3);

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
        LPOLESTR szNewName = NULL;
        
        szNewName = (LPOLESTR)malloc(sizeof(OLECHAR) * 64);
        if (szNewName == NULL)
            hr = E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            hr = RegisterName(Tcl_GetUnicode(name), pkg->obj, FALSE,
                              &szNewName, 64, &pkg->ROT_cookie);
            if (SUCCEEDED(hr)) {
                pkg->appname = Tcl_NewUnicodeObj(szNewName, -1);
                Tcl_IncrRefCount(pkg->appname);
            }
            free(szNewName);
        }
    }

    /* Create our winsend command */
    if (SUCCEEDED(hr))
    {
        pkg->token = Tcl_CreateObjCommand(interp,
                                          "winsend",
                                          Winsend_CmdProc,
                                          (ClientData)pkg,
                                          (Tcl_CmdDeleteProc*)NULL);

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

EXTERN int
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
    static CONST84 char* cmds[] = { "interps", "send", "appname", "test", NULL };
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
    LPCOLESTR oleszStub = WINSEND_REGISTRATION_BASE;
    HRESULT hr = S_OK;
    Tcl_Obj *objList = NULL;
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

    } else if (objc > 4) {

        Tcl_WrongNumArgs(interp, 2, objv, "?-exact? ?appname?");
        r = TCL_ERROR;

    } else {

        LPOLESTR szNewName = NULL;
        DWORD    dwCookie = 0;
        BOOL     bForce = TRUE;
        Tcl_Obj *objAppname = objv[2];
        enum {WINSEND_APPNAME_FORCE, WINSEND_APPNAME_INCREMENT};
        static CONST84 char *opts[] = {"-force", "-increment", NULL};

        if (objc == 4) {
            int index = 0;
            objAppname = objv[3];
            r = Tcl_GetIndexFromObj(interp, objv[2], opts, "option", 0, &index);
            if (r == TCL_OK) {
                switch (index) {
                case WINSEND_APPNAME_FORCE: bForce = TRUE; break;
                case WINSEND_APPNAME_INCREMENT: bForce = FALSE; break;
                }
            } else {
                return r;
            }
        }
        
        szNewName = (LPOLESTR)malloc(sizeof(OLECHAR) * 64);
        if (szNewName == NULL)
            hr = E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            hr = RegisterName(Tcl_GetUnicode(objAppname), pkg->obj, bForce,
                              &szNewName, 64, &dwCookie);
            if (SUCCEEDED(hr))
            {
                LPRUNNINGOBJECTTABLE pROT = NULL;
                hr = GetRunningObjectTable(0, &pROT);
                if (SUCCEEDED(hr))
                {
                    /* revoke the old name */
                    hr = pROT->lpVtbl->Revoke(pROT, pkg->ROT_cookie);
                    /* record the new registration details */
                    pkg->ROT_cookie = dwCookie;
                    Tcl_DecrRefCount(pkg->appname); /* release the old name */
                    pkg->appname = Tcl_NewUnicodeObj(szNewName, -1);
                    Tcl_IncrRefCount(pkg->appname); /* hang onto the new name */
                    Tcl_SetObjResult(interp, pkg->appname);

                    pROT->lpVtbl->Release(pROT);
                }
            }
            free(szNewName);
        }
        if (FAILED(hr))
        {
            Tcl_Obj *err = NULL;

            if (hr == MK_E_MONIKERALREADYREGISTERED)
                err = Tcl_NewStringObj("error: this name is already in use", -1);
            else 
                err = Winsend_Win32ErrorObj(hr);

            Tcl_SetObjResult(interp, err);
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
    enum {WINSEND_TEST_ERROR, WINSEND_TEST_OBJECT};
    static CONST84 char * cmds[] = { "error", "object", NULL };
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
        case WINSEND_TEST_OBJECT:
            {
                Tcl_Obj *obj = Tcl_NewLongObj((long)0xBB66AA55L);
                Tcl_IncrRefCount(obj);
                Tcl_SetObjResult(interp, obj);
                Tcl_DecrRefCount(obj);
                r = TCL_OK;
            }
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
    HRESULT hr = S_OK, ehr = S_OK;
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
    ehr = VariantChangeType(&vResult, &vResult, 0, VT_BSTR);
    if (SUCCEEDED(ehr))
        Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(vResult.bstrVal, -1));
    if (hr == DISP_E_EXCEPTION)
    {
        Tcl_Obj *opError, *opErrorCode, *opErrorInfo;

        if (ei.bstrSource != NULL)
        {
            int len;
            char * szErrorInfo;

            opError = Tcl_NewUnicodeObj(ei.bstrSource, -1);
            Tcl_ListObjIndex(interp, opError, 0, &opErrorCode);
            Tcl_SetObjErrorCode(interp, opErrorCode);

            Tcl_ListObjIndex(interp, opError, 1, &opErrorInfo);
            szErrorInfo = Tcl_GetStringFromObj(opErrorInfo, &len);
            Tcl_AddObjErrorInfo(interp, szErrorInfo, len);
        }
    }


    SysFreeString(ei.bstrDescription);
    SysFreeString(ei.bstrSource);
    SysFreeString(ei.bstrHelpFile);
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
BuildMoniker(LPCOLESTR name, LPMONIKER *ppmk)
{
    LPMONIKER pmkClass = NULL;
    HRESULT hr = CreateFileMoniker(WINSEND_REGISTRATION_BASE, &pmkClass);
    if (SUCCEEDED(hr)) {
        LPMONIKER pmkItem = NULL;
        
        hr = CreateFileMoniker(name, &pmkItem);
        if (SUCCEEDED(hr)) {
            hr = pmkClass->lpVtbl->ComposeWith(pmkClass, pmkItem, FALSE, ppmk);
            pmkItem->lpVtbl->Release(pmkItem);
        }
        pmkClass->lpVtbl->Release(pmkClass);
    }
    return hr;
}

/* -------------------------------------------------------------------------
 * RegisterObject
 * -------------------------------------------------------------------------
 * Description:
 */
static HRESULT
RegisterName(LPCOLESTR szName, IUnknown *pUnknown, BOOL bForce,
             LPOLESTR *pNewName, DWORD cbNewName, DWORD *pdwCookie)
{
    HRESULT hr = S_OK;
    LPRUNNINGOBJECTTABLE pROT = NULL;

    if (pNewName == NULL || cbNewName < 2)
        hr = E_INVALIDARG;
    if (SUCCEEDED(hr))
        hr = GetRunningObjectTable(0, &pROT);
    if (SUCCEEDED(hr))
    {
        int n = 1;
        LPMONIKER pmk = NULL;

        memset(*pNewName, 0, sizeof(OLECHAR) * cbNewName);

        do
        {
            if (n > 1)
                _snwprintf(*pNewName, cbNewName, L"%s #%u", szName, n);
            else
                wcsncpy(*pNewName, szName, cbNewName);
            n++;

            hr = BuildMoniker(*pNewName, &pmk);

            if (SUCCEEDED(hr)) {
                hr = pROT->lpVtbl->Register(pROT,
                                            ROTFLAGS_REGISTRATIONKEEPSALIVE,
                                            pUnknown,
                                            pmk,
                                            pdwCookie);
                pmk->lpVtbl->Release(pmk);
            }

            /* If the moniker was registered, unregister the duplicate and
             * try again.
             * If the force flag was set then convert this into an error code.
             */
            if (hr == MK_S_MONIKERALREADYREGISTERED) {
                pROT->lpVtbl->Revoke(pROT, *pdwCookie);
                *pdwCookie = (DWORD)-1;
                if (bForce)
                    hr = MK_E_MONIKERALREADYREGISTERED;
            }
        } while (hr == MK_S_MONIKERALREADYREGISTERED);

        pROT->lpVtbl->Release(pROT);
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
        _stprintf(sBuffer, _T("Error Code: %08lX"), hrError);
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
