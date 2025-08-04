Dim ArgObj, var1, user

Set ArgObj = WScript.Arguments
user = ""
var1 = ArgObj(0) & " " & 1 & " " & user & " " & "saveas"

CreateObject("WScript.Shell").SendKeys "{F5}"

var2 = """C:\Program Files\Sigep Client Modules\SigepPrimavera\SIGEPPRIMAVERA.exe """  & " " & var1  

rem MsgBox var2

Set wshShell = WScript.CreateObject ("WSCript.shell")

wshshell.run var2,2

set wshshell = nothing

rem MsgBox "Publishing project " & var1 & " to SIGEP done.",0,"Information"
