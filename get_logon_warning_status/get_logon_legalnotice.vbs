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
' 로그온 시도시 경고 메시지 설정
'-------------------------------------------

    dim LOGON_NOTICE_POLICY
    LOGON_NOTICE_POLICY = GET_LOGON_NOTICE_POLICY()

 
    Function GET_LOGON_NOTICE_POLICY()

        dim WSHShell 
        dim VAL_NOTICECAPTION, VAL_NOTICETEXT
        dim CAPTION_USE, NOTICETEXT_USE

        On Error Resume Next
        Set WSHShell = CreateObject("WScript.Shell")
       
	VAL_NOTICECAPTION = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticecaption")
	VAL_NOTICETEXT = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticecaption")
	
        If VAL_NOTICECAPTION <> "" Then 
	    CAPTION_USE = "양호 / 설정내용:[ " & VAL_NOTICECAPTION & " ]"
        else
	    CAPTION_USE = "취약 / 메시지 제목 설정되지 않음"
        end if
               
        If VAL_NOTICETEXT <> "" Then 
	    NOTICETEXT_USE = "양호 / 설정내용:[ " & VAL_NOTICETEXT & " ]"
        else
            NOTICETEXT_USE = "취약 / 메시지 내용 설정되지 않음"
        end if
        
        wscript.echo "1.로그온 시도 경고 제목 " & vbcrlf & "- 결과: " & CAPTION_USE & vbcrlf & "2.로그온 시도 경고 내용 : " & vbcrlf & "- 결과: " & NOTICETEXT_USE 
        set WSHShell = nothing
    end function

end If