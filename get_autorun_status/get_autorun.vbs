'--------------------------------------------------------------------
' Windows security check script
' supported on : windows 7 over
' created by password123456 / 2015.08
' Freeware, but You always need to cite your sources
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
' 자동실행 정책 확인
'-------------------------------------------

    dim AUTORUN
    AUTORUN = GET_AUTORUN_POLICY()

 
    Function GET_AUTORUN_POLICY()

        dim WSHShell
        dim VAL_NOAUTORUN, VAL_NODRIVETYPEAUTORUN
        dim AUTORUN_USE, NODRIVETYPEAUTORUN

        On Error Resume Next
        Set WSHShell = CreateObject("WScript.Shell")

        VAL_NOAUTORUN = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoAutorun")
        VAL_NODRIVETYPEAUTORUN = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun") 


        if VAL_NOAUTORUN = 1 then
            AUTORUN_USE = "자동실행사용안함 / 양호"
        else
            AUTORUN_USE = "설정되지않음 / 취약"
        end if

        select case VAL_NODRIVETYPEAUTORUN  
            Case "181"
                NODRIVETYPEAUTORUN = "이동식미디어만 자동실행 안함"
            Case "255"
                NODRIVETYPEAUTORUN = "모든 드라이브 자동실행 안함"
            Case else
                NODRIVETYPEAUTORUN = "설정된 내용 없음"
        end select 

        wscript.echo "자동 실행 기본동작 : " & AUTORUN_USE & vbcrlf & "자동 실행 사용안함  : " & NODRIVETYPEAUTORUN  
     
        set WSHShell = nothing
    end function

end If