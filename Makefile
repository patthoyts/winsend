# -*- Makefile -*- 
#
# This is a mingw gcc makefile for the winsend package.
#
# Currently: using mingw 1.0 (gcc version 2.95.3) this package
# will not compile due to errors in the headers provided by the
# compiler for the Running Object Table interface.
#
# @(#)$Id$

VER     =03
DBGX    =d
DBGFLAGS=-D_DEBUG

CC      =gcc -g
DLLWRAP =dllwrap
DLLTOOL =dlltool
RM      =rm -f
CFLAGS  =-Wall -I/opt/tcl/include -DUSE_TCL_STUBS $(DBGFLAGS)
LDFLAGS =-L/opt/tcl/lib
LIBS    =-ltclstub83${DBGX} -lole32 -loleaut32 -ladvapi32 -luuid

DLL     =winsend${VER}${DBGX}.dll
DEFFILE =winsend.def

WRAPFLAGS =--driver-name $(CC) --def $(DEFFILE)

CSRCS   =winsend.c WinSendCom.c
OBJS    =$(CSRCS:.c=.o)

$(DLL): $(OBJS)
	$(DLLWRAP) $(WRAPFLAGS) -o $@ $^ $(LDFLAGS) $(LIBS)

clean:
	$(RM) *.o core *~

%.o: %.c
	$(CC) $(CFLAGS) -c $< -o $@

WinSendCom.o: WinSendCom.c WinSendCom.h

.PHONY: clean

#
# Local variables:
#   mode: makefile
# End:
#
