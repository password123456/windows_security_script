'--------------------------------------------------------------------
' Windows security check script
' supported on : windows 7 over
' created by password123456 / 2015.08
' Freeware, but You always need to cite your sources.
' https://github.com/password123456
'---------------------------------------------------------------------

Option Explicit

'-------------------------------------------
' UAC 실행 삽입
'-------------------------------------------
dim objshell

if WScript.Arguments.length =0 Then
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument with leading blank space, say [ uac ]
    objShell.ShellExecute "wscript.exe", Chr(34) & _
    WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    
else

'-------------------------------------------
' 원격데스크톱 상태 & 포트번호
'-------------------------------------------

    Dim rdp
    rdp = REMOTE_DESK_TOP()

 
    Function REMOTE_DESK_TOP()

        Dim WSHShell, STATUS, PORT_NUMBER, RESULT

        On Error Resume Next
        Set WSHShell = CreateObject("WScript.Shell")

        STATUS = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\fDenyTSConnections")
        PORT_NUMBER = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\PortNumber")

        if STATUS = 1 then
            STATUS ="사용하지않음"
        else
            STATUS ="사용중"
        end if

        if PORT_NUMBER > 3389 then
            RESULT = "양호"
        else
            RESULT = "취약 / 포트번호 변경필요"
        end if

        wscript.echo "Remote Desktop : " & STATUS & vbcrlf & "Port Number : " & PORT_NUMBER & vbcrlf & "Result : " & RESULT 
     
        set WSHShell = nothing
    end function

end If