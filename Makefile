# -*- Makefile -*- 
#
# This is a mingw gcc makefile for the winsend package.
#
# Currently: using mingw 1.0 (gcc version 2.95.3) this package
# will not compile due to errors in the headers provided by the
# compiler for the Running Object Table interface.
#
# @(#)$Id$

#TCLROOT ="c:/Program Files/Tcl"
TCLROOT =/opt/tcl

VER     =05
DBGX    =d
DBGFLAGS=-g -D_DEBUG

CC      =gcc
DLLWRAP =dllwrap
DLLTOOL =dlltool
RM      =rm -f
CFLAGS  =-Wall -I$(TCLROOT)/include -DUSE_TCL_STUBS $(DBGFLAGS)
LDFLAGS =-L$(TCLROOT)/lib
LIBS    =-ltclstub83${DBGX} -lole32 -loleaut32 -ladvapi32 -luuid

DLL     =winsend${VER}${DBGX}.dll
DEFFILE =winsend.def

WRAPFLAGS =--driver-name $(CC) --def $(DEFFILE)

CSRCS   =winsend.c WinSendCom.c debug.c
OBJS    =$(CSRCS:.c=.o)

$(DLL): $(OBJS)
	$(DLLWRAP) $(WRAPFLAGS) -o $@ $^ $(LDFLAGS) $(LIBS)

clean:
	$(RM) *.o core *~

%.o: %.c
	$(CC) $(CFLAGS) -c $< -o $@

winsend.o: WinSendCom.h debug.h
WinSendCom.o: WinSendCom.c WinSendCom.h
debug.o: debug.h

.PHONY: clean

#
# Local variables:
#   mode: makefile
# End:
#
