/* WinSendCom.h - Copyright (C) 2002 Pat Thoyts <Pat.Thoyts@bigfoot.com>
 *
 * $Id$
 */

#include <windows.h>
#include <ole2.h>
#include <tchar.h>
#include <tcl.h>

/*
 * WinSendCom CoClass structure 
 */
typedef struct WinSendCom_t {
    IDispatchVtbl *lpVtbl;
    ISupportErrorInfoVtbl *lpVtbl2;
    long           refcount;
    Tcl_Interp     *interp;
} WinSendCom;

/*
 * WinSendCom public functions
 */
HRESULT WinSendCom_CreateInstance(Tcl_Interp *interp, REFIID riid, void **ppv);
void WinSendCom_Destroy(LPDISPATCH pdisp);

/*
 * Local Variables:
 *  mode: c
 *  indent-tabs-mode: nil
 * End:
 */
