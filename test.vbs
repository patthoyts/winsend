Dim o
Set o = GetObject("\\.\TclInterp\tkchat")
WScript.Echo o.Send("puts {Hello, from VB} ; info tcl")
