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
' 화면보호기 설정 확인
'-------------------------------------------

    dim SCREEN_SAVER
    SCREEN_SAVER = GET_SCREEN_SAVER()

 
    Function GET_SCREEN_SAVER()

        dim WSHShell
        dim VAL_SAVER_USE, VAL_SAVER_PROTECT, VAL_SAVER_TIME
        dim SAVER_USE, SAVER_PROTECT, SAVER_TIME

        On Error Resume Next
        Set WSHShell = CreateObject("WScript.Shell")

        VAL_SAVER_USE = WSHShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveActive")
        VAL_SAVER_PROTECT = WSHShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaverIsSecure") 
        VAL_SAVER_TIME = WSHShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveTimeOut")


        if VAL_SAVER_USE = 1 then
            SAVER_USE = "양호"
        else
            SAVER_USE = "취약"
        end if

        if VAL_SAVER_PROTECT = 1 then
            SAVER_PROTECT = "양호"
        else
            SAVER_PROTECT = "취약"
        end if

        if VAL_SAVER_TIME > 600 then
            SAVER_TIME = "양호" & VAL_SAVER_TIME & "초"
        else
            SAVER_TIME = "취약[ " & VAL_SAVER_TIME & " ] 초 / 10분 이상으로 설정 필요"
        end if

        wscript.echo "화면 보호기 사용 : " & SAVER_USE & vbcrlf & "화면 보호기 암호로 보호  : " & SAVER_PROTECT & vbcrlf & "화면 보호기 대기시간 : " & SAVER_TIME 
     
        set WSHShell = nothing
    end function

end If