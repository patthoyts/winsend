#include "debug.h"
#include <stdio.h>

void
LocalTrace(LPCTSTR format, ...)
{
    int n;
    const int max = sizeof(TCHAR) * 512;
    TCHAR buffer[512];
    va_list args;
    va_start (args, format);
    
    n = _vsntprintf(buffer, max, format, args);
    _ASSERTE(n < max);
    OutputDebugString(buffer);
    va_end(args);
}
