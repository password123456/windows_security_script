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
' �̺�Ʈ�α� ���� ������ Ȯ��
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
            SECURITY_SIZE = "��ȣ / " & VAL_SECURITY & " KB"
        else
            SECURITY_SIZE = "��� / " & VAL_SECURITY & " KB / 200MB �̻� ���� �ʿ�"
        end if

        if VAL_APPLICATION >= 204800 then
            APPLICATION_SIZE = "��ȣ / " & VAL_APPLICATION & " KB"
        else
            APPLICATION_SIZE = "��� / " & VAL_APPLICATION & " KB / 200MB �̻� ���� �ʿ�"
        end if

        if VAL_SYSTEM >= 204800 then
            SYSTEM_SIZE = "��ȣ / " & VAL_SYSTEM & " KB"
        else
            SYSTEM_SIZE = "��� / " & VAL_SYSTEM  & " KB / 200MB �̻� ���� �ʿ�"
        end if

        wscript.echo "1.���� �α� ���� : " & SECURITY_SIZE & vbcrlf & "2.���ø����̼� �α� ����  : " & APPLICATION_SIZE & vbcrlf & "3.�ý��� �α� ���� : " & SYSTEM_SIZE 
     
        set WSHShell = nothing
    end function

end If