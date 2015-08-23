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
' 이벤트로그 저장 사이즈 확인
'-------------------------------------------

    dim EVENT_LOG_SIZE
    EVENT_LOG_SIZE = GET_EVENT_LOG_SET_SIZE()

 
    Function GET_EVENT_LOG_SET_SIZE()

        dim WSHShell
        dim VAL_SECURITY, VAL_APPLICATION, VAL_SYSTEM
        dim SECURITY_SIZE, APPLICATION_SIZE, SYSTEM_SIZE

        On Error Resume Next
        Set WSHShell = CreateObject("WScript.Shell")

        VAL_SECURITY = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\EventLog\Security\MaxSize")
        VAL_APPLICATION = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\EventLog\Application\MaxSize") 
        VAL_SYSTEM = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\EventLog\System\MaxSize")


        if VAL_SECURITY >= 204800 then
            SECURITY_SIZE = "양호 / " & VAL_SECURITY & " KB"
        else
            SECURITY_SIZE = "취약 / " & VAL_SECURITY & " KB / 200MB 이상 설정 필요"
        end if

        if VAL_APPLICATION >= 204800 then
            APPLICATION_SIZE = "양호 / " & VAL_APPLICATION & " KB"
        else
            APPLICATION_SIZE = "취약 / " & VAL_APPLICATION & " KB / 200MB 이상 설정 필요"
        end if

        if VAL_SYSTEM >= 204800 then
            SYSTEM_SIZE = "양호 / " & VAL_SYSTEM & " KB"
        else
            SYSTEM_SIZE = "취약 / " & VAL_SYSTEM  & " KB / 200MB 이상 설정 필요"
        end if

        wscript.echo "1.보안 로그 설정 : " & SECURITY_SIZE & vbcrlf & "2.어플리케이션 로그 설정  : " & APPLICATION_SIZE & vbcrlf & "3.시스템 로그 설정 : " & SYSTEM_SIZE 
     
        set WSHShell = nothing
    end function

end If