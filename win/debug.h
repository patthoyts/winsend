/* debug.h - Copyright (C) 2002 Pat Thoyts <Pat.Thoyts@bigfoot.com>
 *
 * $Id$
 */

#ifndef _Debug_h_INCLUDE
#define _Debug_h_INCLUDE

#include <windows.h>
#include <assert.h>
#include <stdarg.h>
#include <tchar.h>

/* Debug related stuff */
#ifndef _DEBUG

#define _ASSERTE(x)  ((void)0)
#define LTRACE       1 ? ((void)0) : LocalTrace

#else /* _DEBUG */

#define _ASSERTE(x) if (!(x)) _assert(#x, __FILE__, __LINE__)
#define LTRACE LocalTrace

#endif /* _DEBUG */

/* Debug functions. */
void LocalTrace(LPCTSTR szFormat, ...);

#endif /* _Debug_h_INCLUDE */
