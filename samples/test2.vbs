' test2.vbs - Copyright (C) 2002-2004 Pat Thoyts <patthoyts@users.sourceforge.net>
'
' Demo accessing Tcl using the winsend package from the Visual Basic
' Scripting engine via the Windown Scripting Host (WSH).
'
' Run this using 'cscript test2.vbs' or 'wscript test2.vbs' when you
' have a wish session using the winsend appname wish.
'
' $Id$

Dim o
Set o = GetObject("TclEval\wish")

cmd =                   "package require Tk"
cmd = cmd & vbNewline & "toplevel .t"
cmd = cmd & vbNewline & "button .t.b -text OK -command {destroy .t}"
cmd = cmd & vbNewline & "pack .t.b -side top -fill both"
cmd = cmd & vbNewline & "tkwait window .t"
cmd = cmd & vbNewline & "info tcl"

WScript.Echo o.Send(cmd)
