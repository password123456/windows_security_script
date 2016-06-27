
CONFIG_OPTION = 0

If WScript.Arguments.Count >= 1 Then
	Command1 = WScript.Arguments.Item(0)
    if Command1 = "/h" then
    	WScript.Echo "argument: " & Command1        
        wscript.quit
    elseif Command1 = "/y" then
         'WScript.Echo "argument: " & Command1        
        CONFIG_OPTION = 1
    elseif wscript.arguments(0) = "/v" then
         'WScript.Echo "argument: " & Command1        
        CONFIG_OPTION = 2
      end if
End If

If WScript.Arguments.Named.Exists("elevated") = False Then
  CreateObject("Shell.Application").ShellExecute "cscript.exe", """" & _
    WScript.ScriptFullName & """ " & Command1 & " /elevated", "", "runas", 1
  WScript.Quit
End If

Function GET_PAUSE()
    dim strMessage
    strMessage = "Press the ENTER key to continue."
    Wscript.StdOut.Write strMessage

    Do While Not WScript.StdIn.AtEndOfLine
        Input = WScript.StdIn.Read(1)
    Loop
end Function

WScript.Echo Command1
wscript.echo CONFIG_OPTION 
GET_PAUSE()
