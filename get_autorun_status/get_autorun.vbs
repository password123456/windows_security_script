'--------------------------------------------------------------------
' Windows security check script
' supported on : windows 7 over
' created by password123456 / 2015.08
' Freeware, but You always need to cite your sources
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
' �ڵ����� ��å Ȯ��
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
            AUTORUN_USE = "�ڵ���������� / ��ȣ"
        else
            AUTORUN_USE = "������������ / ���"
        end if

        select case VAL_NODRIVETYPEAUTORUN  
            Case "181"
                NODRIVETYPEAUTORUN = "�̵��Ĺ̵� �ڵ����� ����"
            Case "255"
                NODRIVETYPEAUTORUN = "��� ����̺� �ڵ����� ����"
            Case else
                NODRIVETYPEAUTORUN = "������ ���� ����"
        end select 

        wscript.echo "�ڵ� ���� �⺻���� : " & AUTORUN_USE & vbcrlf & "�ڵ� ���� ������  : " & NODRIVETYPEAUTORUN  
     
        set WSHShell = nothing
    end function

end If