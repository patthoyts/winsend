/* WinSendCom.c - Copyright (C) 2002 Pat Thoyts <Pat.Thoyts@bigfoot.com>
 *
 * Implement a COM class for use in registering Tcl interpreters with the
 * system's Running Object Table.
 * This class implements an IDispatch interface with the following method:
 *   Send(String cmd) As String 
 * In other words the Send methods takes a string and evaluates this in
 * the Tcl interprer. The result is returned as another string.
 */

static const char rcsid[] = 
"$Id$";

#include "WinSendCom.h"

/* ---------------------------------------------------------------------- */
/* Non-public prototypes.
/* ---------------------------------------------------------------------- */

static STDMETHODIMP WinSendCom_QueryInterface(IDispatch *This,
                                              REFIID riid, void **ppvObject);
static STDMETHODIMP_(ULONG) WinSendCom_AddRef(IDispatch *This);
static STDMETHODIMP_(ULONG) WinSendCom_Release(IDispatch *This);
static STDMETHODIMP WinSendCom_GetTypeInfoCount(IDispatch *This,
                                                UINT *pctinfo);
static STDMETHODIMP WinSendCom_GetTypeInfo(IDispatch *This, UINT iTInfo,
                                           LCID lcid, ITypeInfo **ppTI);
static STDMETHODIMP WinSendCom_GetIDsOfNames(IDispatch *This, REFIID riid,
                                             LPOLESTR *rgszNames,
                                             UINT cNames, LCID lcid,
                                             DISPID *rgDispId);
static STDMETHODIMP WinSendCom_Invoke(IDispatch *This, DISPID dispidMember,
                                      REFIID riid, LCID lcid, WORD wFlags,
                                      DISPPARAMS *pDispParams,
                                      VARIANT *pvarResult,
                                      EXCEPINFO *pExcepInfo,
                                      UINT *puArgErr);
static HRESULT Send(WinSendCom* obj, VARIANT vCmd, VARIANT* pvResult);

/* ---------------------------------------------------------------------- */
/* COM Class Helpers
/* ---------------------------------------------------------------------- */

/* Description:
 *   Create and initialises a new instance of the WinSend COM class and
 *   returns an interface pointer for you to use.
 */
HRESULT
WinSendCom_CreateInstance(Tcl_Interp *interp, REFIID riid, void **ppv)
{
    static IDispatchVtbl vtbl = {
        WinSendCom_QueryInterface,
        WinSendCom_AddRef,
        WinSendCom_Release,
        WinSendCom_GetTypeInfoCount,
        WinSendCom_GetTypeInfo,
        WinSendCom_GetIDsOfNames,
        WinSendCom_Invoke,
    };
    HRESULT hr = S_OK;
    WinSendCom *obj = NULL;

    obj = (WinSendCom*)malloc(sizeof(WinSendCom));
    if (obj == NULL) {
        *ppv = NULL;
        hr = E_OUTOFMEMORY;
    } else {
        obj->lpVtbl = &vtbl;
        obj->refcount = 0;
        obj->interp = interp;

        /* lock the interp? Tcl_AddRef/Retain? */

        hr = obj->lpVtbl->QueryInterface((IDispatch*)obj, riid, ppv);
    }

    return hr;
}

/* Description:
 *  Used to cleanly destroy instances of the CoClass.
 */
void
WinSendCom_Destroy(LPDISPATCH pdisp)
{
    free((void*)pdisp);
}

/* ---------------------------------------------------------------------- */

/*
 * WinSendCom IDispatch methods
 */

static STDMETHODIMP
WinSendCom_QueryInterface(IDispatch *This,
                          REFIID riid,
                          void **ppvObject)
{
    HRESULT hr = E_NOINTERFACE;
    WinSendCom *this = (WinSendCom*)This;
    *ppvObject = NULL;
    
    if (memcmp(riid, &IID_IUnknown, sizeof(IID)) == 0
        || memcmp(riid, &IID_IDispatch, sizeof(IID)) == 0) {
        *ppvObject = (void**)this;
        this->lpVtbl->AddRef(This);
        hr = S_OK;
    }
    return hr;
}

static STDMETHODIMP_(ULONG)
WinSendCom_AddRef(IDispatch *This)
{
    WinSendCom *this = (WinSendCom*)This;
    return InterlockedIncrement(&this->refcount);
}

static STDMETHODIMP_(ULONG)
WinSendCom_Release(IDispatch *This)
{
    long r = 0;
    WinSendCom *this = (WinSendCom*)This;
    if ((r = InterlockedDecrement(&this->refcount)) == 0) {
        WinSendCom_Destroy(This);
    }
    return r;
}

static STDMETHODIMP
WinSendCom_GetTypeInfoCount(IDispatch *This, UINT *pctinfo)
{
    HRESULT hr = E_POINTER;
    if (pctinfo != NULL) {
        *pctinfo = 0;
        hr = S_OK;
    }
    return hr;
}

static STDMETHODIMP 
WinSendCom_GetTypeInfo(IDispatch *This, UINT iTInfo,
                       LCID lcid, ITypeInfo **ppTI)
{
    HRESULT hr = E_POINTER;
    if (ppTI)
    {
        *ppTI = NULL;
        hr = E_NOTIMPL;
    }
    return hr;
}

static STDMETHODIMP 
WinSendCom_GetIDsOfNames(IDispatch *This, REFIID riid,
                         LPOLESTR *rgszNames,
                         UINT cNames, LCID lcid,
                         DISPID *rgDispId)
{
    HRESULT hr = E_POINTER;
    if (rgDispId)
    {
        hr = DISP_E_UNKNOWNNAME;
        if (_wcsicmp(*rgszNames, L"Send") == 0)
        {
            *rgDispId = 1;
            hr = S_OK;
        }
    }
    return hr;
}

static STDMETHODIMP 
WinSendCom_Invoke(IDispatch *This, DISPID dispidMember,
                  REFIID riid, LCID lcid, WORD wFlags,
                  DISPPARAMS *pDispParams,
                  VARIANT *pvarResult,
                  EXCEPINFO *pExcepInfo,
                  UINT *puArgErr)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    WinSendCom *this = (WinSendCom*)This;

    switch (dispidMember)
    {
    case 1: // Send
        if (wFlags | DISPATCH_METHOD)
        {
            if (pDispParams->cArgs != 1)
                hr = DISP_E_BADPARAMCOUNT;
            else
                hr = Send(this, pDispParams->rgvarg[0], pvarResult);
        }
    }
    return hr;
}

/* Description:
 *  Evaluates the string in the assigned interpreter. If the result
 *  is a valid address then set to the result returned by the evaluation.
 */
static HRESULT
Send(WinSendCom* obj, VARIANT vCmd, VARIANT* pvResult)
{
    HRESULT hr = S_OK;
    int r = TCL_OK;
    VARIANT v;

    VariantInit(&v);
    hr = VariantChangeType(&v, &vCmd, 0, VT_BSTR);
    if (SUCCEEDED(hr))
    {
        if (obj->interp)
        {
            Tcl_Obj *script = Tcl_NewUnicodeObj(v.bstrVal,
                                                SysStringLen(v.bstrVal));
            r = Tcl_EvalObjEx(obj->interp, script, TCL_EVAL_DIRECT);
            if (pvResult)
            {
                VariantInit(pvResult);
                pvResult->vt = VT_BSTR;
                pvResult->bstrVal = SysAllocString(Tcl_GetUnicode(Tcl_GetObjResult(obj->interp)));
            }
        }
        VariantClear(&v);
    }
    return hr;
}

/*
 * Local Variables:
 *  mode: c
 *  indent-tabs-mode: nil
 * End:
 */
