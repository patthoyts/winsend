Dim o
Set o = GetObject("\\.\TclInterp\tkcon.tcl")
WScript.Echo o.Send("puts {Hello, from VB} ; tkcon master wm title .")

Dim p
Set p = GetObject("\\.\TclInterp\tkcon.tcl #2")
WScript.Echo  p.Send("puts {Hello, 2}; tkcon master wm title .")



