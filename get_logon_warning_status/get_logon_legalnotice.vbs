'--------------------------------------------------------------------
' Windows security check script
' supported on : windows 7 over
' created by password123456 / 2015.08
' Freeware, but You always need to cite your sources.
' https://github.com/password123456
'---------------------------------------------------------------------

Option Explicit

'-------------------------------------------
' UAC ���� ����
'-------------------------------------------
dim objshell

if WScript.Arguments.length =0 Then
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument with leading blank space, say [ uac ]
    objShell.ShellExecute "wscript.exe", Chr(34) & _
    WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    
else

'-------------------------------------------
' �α׿� �õ��� ��� �޽��� ����
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
	    CAPTION_USE = "��ȣ / ��������:[ " & VAL_NOTICECAPTION & " ]"
        else
	    CAPTION_USE = "��� / �޽��� ���� �������� ����"
        end if
               
        If VAL_NOTICETEXT <> "" Then 
	    NOTICETEXT_USE = "��ȣ / ��������:[ " & VAL_NOTICETEXT & " ]"
        else
            NOTICETEXT_USE = "��� / �޽��� ���� �������� ����"
        end if
        
        wscript.echo "1.�α׿� �õ� ��� ���� " & vbcrlf & "- ���: " & CAPTION_USE & vbcrlf & "2.�α׿� �õ� ��� ���� : " & vbcrlf & "- ���: " & NOTICETEXT_USE 
        set WSHShell = nothing
    end function

end If