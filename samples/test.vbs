Dim o
Set o = GetObject("TclEval\tkchat")
WScript.Echo o.Send("puts {Hello, from VB} ; info tcl")
