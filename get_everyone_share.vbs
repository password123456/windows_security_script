Option Explicit

dim objShell

if WScript.Arguments.length =0 Then
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument with leading blank space, say [ uac ]
    objShell.ShellExecute "wscript.exe", Chr(34) & _
    WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    
else

    dim a1
    a1= GET_EVERYONE_SHARE_FOLDER()
    
    Function GET_EVERYONE_SHARE_FOLDER()

        Const WBEM_RETURN_IMMEDIATELY = &H10
        Const WBEM_FORWARD_ONLY = &H20

        dim strComputer, strDomain 
        dim objACE, objSD
        dim objWMI, objSecuritySettings
        dim colItems, objItem
        dim strCheckEveryone, LIST, RESULT, AccessUser
        dim intCount,i

        strComputer = "."

        Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")

        ' Return all of the shares (Type = 0 means File Shares only, exclude
        ' Administrative share, Printer, etc)
        Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Share WHERE Type='0'", "WQL", _
                                         WBEM_RETURN_IMMEDIATELY + WBEM_FORWARD_ONLY)

        strCheckEveryone = "everyone"
        intCount = "0"

        For Each objItem in colItems
             Set objSecuritySettings = objWMI.Get("Win32_LogicalFileSecuritySetting='" & objItem.Path & "'")
            objSecuritySettings.GetSecurityDescriptor objSD

            strDomain = objSD.Owner.Domain
          
            If strDomain <> "" Then 
                strDomain = strDomain & ""
            end if

            For Each objACE in objSD.DACL
                strDomain = objAce.Trustee.Domain
                If strDomain <> "" Then strDomain = strDomain & ""
                'WScript.Echo "    Trustee: " & UCase(strDomain & objAce.Trustee.Name)
                AccessUser = AccessUser & vbcrlf & lCase(strDomain & "\" & objAce.Trustee.Name)

                If lcase(strDomain & objAce.Trustee.Name) = lcase(strCheckEveryone) then
                    intCount = intCount + 1
                    LIST = LIST & vbcrlf & vbcrlf & "* 공유폴더: " & objItem.Path & vbcrlf & _
                           "* 사용자목록: " & AccessUser 
                end if
            Next
        Next

        If intCount = 0 then
            RESULT = "양호" & vbcrlf & "- 내용: Everyone 공유폴더 [ " & intCOUNT & " ] 개 발견"
        else
            RESULT = "취약" & vbcrlf & "- 내용: Everyone 공유폴더 [ " & intCOUNT & " ] 개 발견" & vbcrlf & _
                     "- 목록: " & LIST
        end if
        wscript.echo RESULT
    End Function
end if
