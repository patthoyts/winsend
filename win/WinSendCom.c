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

/* ----------------------------------------------------------------------
 * Non-public prototypes.
 * ----------------------------------------------------------------------
 */

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

static STDMETHODIMP ISupportErrorInfo_QueryInterface(ISupportErrorInfo *This,
                                                     REFIID riid, void **ppvObject);
static STDMETHODIMP_(ULONG) ISupportErrorInfo_AddRef(ISupportErrorInfo *This);
static STDMETHODIMP_(ULONG) ISupportErrorInfo_Release(ISupportErrorInfo *This);
static STDMETHODIMP ISupportErrorInfo_InterfaceSupportsErrorInfo(ISupportErrorInfo *This,
                                                                 REFIID riid);

static HRESULT Send(WinSendCom* obj, VARIANT vCmd, VARIANT* pvResult,
                    EXCEPINFO* pExcepInfo, UINT *puArgErr);

/* ----------------------------------------------------------------------
 * COM Class Helpers
 * ----------------------------------------------------------------------
 */

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

    static ISupportErrorInfoVtbl vtbl2 = {
        ISupportErrorInfo_QueryInterface,
        ISupportErrorInfo_AddRef,
        ISupportErrorInfo_Release,
        ISupportErrorInfo_InterfaceSupportsErrorInfo,
    };

    HRESULT hr = S_OK;
    WinSendCom *obj = NULL;

    obj = (WinSendCom*)malloc(sizeof(WinSendCom));
    if (obj == NULL) {
        *ppv = NULL;
        hr = E_OUTOFMEMORY;
    } else {
        obj->lpVtbl = &vtbl;
        obj->lpVtbl2 = &vtbl2;
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
    } else if (memcmp(riid, &IID_ISupportErrorInfo, sizeof(IID)) == 0) {
        *ppvObject = (void**)(this + 1);
        this->lpVtbl2->AddRef((ISupportErrorInfo*)(this + 1));
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
                hr = Send(this, pDispParams->rgvarg[0], pvarResult, pExcepInfo, puArgErr);
        }
    }
    return hr;
}

/* ---------------------------------------------------------------------- */

/*
 * WinSendCom ISupportErrorInfo methods
 */

static STDMETHODIMP 
ISupportErrorInfo_QueryInterface(ISupportErrorInfo *This,
                                 REFIID riid, void **ppvObject)
{
    WinSendCom *this = (WinSendCom*)(This - 1);
    return this->lpVtbl->QueryInterface((IDispatch*)this, riid, ppvObject);
}

static STDMETHODIMP_(ULONG)
ISupportErrorInfo_AddRef(ISupportErrorInfo *This)
{
    WinSendCom *this = (WinSendCom*)(This - 1);
    return InterlockedIncrement(&this->refcount);
}

static STDMETHODIMP_(ULONG)
ISupportErrorInfo_Release(ISupportErrorInfo *This)
{
    WinSendCom *this = (WinSendCom*)(This - 1);
    return this->lpVtbl->Release((IDispatch*)this);
}

static STDMETHODIMP
ISupportErrorInfo_InterfaceSupportsErrorInfo(ISupportErrorInfo *This, REFIID riid)
{
    WinSendCom *this = (WinSendCom*)(This - 1);
    return S_OK; /* or S_FALSE */
}

/* ---------------------------------------------------------------------- */


/* Description:
 *  Evaluates the string in the assigned interpreter. If the result
 *  is a valid address then set to the result returned by the evaluation.
 */
static HRESULT
Send(WinSendCom* obj, VARIANT vCmd, VARIANT* pvResult, EXCEPINFO* pExcepInfo, UINT *puArgErr)
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
            r = Tcl_EvalObjEx(obj->interp, script, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);
            if (pvResult)
            {
                VariantInit(pvResult);
                pvResult->vt = VT_BSTR;
                pvResult->bstrVal = SysAllocString(Tcl_GetUnicode(Tcl_GetObjResult(obj->interp)));
            }
            if (r == TCL_ERROR)
            {
                hr = DISP_E_EXCEPTION;
                if (pExcepInfo)
                {
                    Tcl_Obj *opError, *opErrorInfo, *opErrorCode;
                    ICreateErrorInfo *pCEI;
                    IErrorInfo *pEI;
                    HRESULT ehr;

                    opError = Tcl_GetObjResult(obj->interp);
		            opErrorInfo = Tcl_GetVar2Ex(obj->interp, "errorInfo", NULL, TCL_GLOBAL_ONLY);
                    opErrorCode = Tcl_GetVar2Ex(obj->interp, "errorCode", NULL, TCL_GLOBAL_ONLY);

                    Tcl_ListObjAppendElement(obj->interp, opErrorCode, opErrorInfo);

                    pExcepInfo->bstrDescription = SysAllocString(Tcl_GetUnicode(opError));
                    pExcepInfo->bstrSource = SysAllocString(Tcl_GetUnicode(opErrorCode));
                    pExcepInfo->scode = E_FAIL;

                    ehr = CreateErrorInfo(&pCEI);
                    if (SUCCEEDED(ehr))
                    {
                        ehr = pCEI->lpVtbl->SetGUID(pCEI, &IID_IDispatch);
                        ehr = pCEI->lpVtbl->SetDescription(pCEI, pExcepInfo->bstrDescription);
                        ehr = pCEI->lpVtbl->SetSource(pCEI, pExcepInfo->bstrSource);
                        ehr = pCEI->lpVtbl->QueryInterface(pCEI, &IID_IErrorInfo, (void**)&pEI);
                        if (SUCCEEDED(ehr))
                        {
                            SetErrorInfo(0, pEI);
                            pEI->lpVtbl->Release(pEI);
                        }
                        pCEI->lpVtbl->Release(pCEI);
                    }
                }
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
