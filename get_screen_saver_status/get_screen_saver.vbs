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
' ȭ�麸ȣ�� ���� Ȯ��
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
            SAVER_USE = "��ȣ"
        else
            SAVER_USE = "���"
        end if

        if VAL_SAVER_PROTECT = 1 then
            SAVER_PROTECT = "��ȣ"
        else
            SAVER_PROTECT = "���"
        end if

        if VAL_SAVER_TIME > 600 then
            SAVER_TIME = "��ȣ" & VAL_SAVER_TIME & "��"
        else
            SAVER_TIME = "���[ " & VAL_SAVER_TIME & " ] �� / 10�� �̻����� ���� �ʿ�"
        end if

        wscript.echo "ȭ�� ��ȣ�� ��� : " & SAVER_USE & vbcrlf & "ȭ�� ��ȣ�� ��ȣ�� ��ȣ  : " & SAVER_PROTECT & vbcrlf & "ȭ�� ��ȣ�� ���ð� : " & SAVER_TIME 
     
        set WSHShell = nothing
    end function

end If