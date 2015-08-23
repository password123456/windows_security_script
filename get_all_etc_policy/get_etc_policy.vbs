
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
' ����� ������Ʈ�� ���ȼ��� ����
'-------------------------------------------

    dim ETC_SECURITY_POLICY
    ETC_SECURITY_POLICY = GET_ETC_SECURITY_POLICY()

 
    Function GET_ETC_SECURITY_POLICY()

        dim WSHShell
        dim VAL_SHUTDOWNWITHOUTLOGON, VAL_DONTDISPLAYLASTUSERNAME, VAL_CACHEDLOGONSCOUNT
        dim SHUTDOWNWITHOUTLOGON_USE, DONTDISPLAYLASTUSERNAME_USE, CACHEDLOGONSCOUNT_USE
        dim VAL_PASSWORDEXPIRYWARNING, VAL_TASKSCHEDULER_LOG
        dim PASSWORDEXPIRYWARNING_USE, TASKSCHEDULER_LOG_USE

        On Error Resume Next
        Set WSHShell = CreateObject("WScript.Shell")
	
	' �ý��� ����: �α׿����� �ʰ� �ý��� ���� ���
        VAL_NOTICECAPTION = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\shutdownwithoutlogon")

	' ��ȭ�� �α׿�: ������ ����� �̸� ǥ�� �� ��
        VAL_DONTDISPLAYLASTUSERNAME = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\dontdisplaylastusername")

	' ��ȭ�� �α׿�: ĳ���� �α׿��� Ƚ��(������ ��Ʈ�ѷ��� ��� �Ұ����� ���)
        VAL_CACHEDLOGONSCOUNT = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\CachedLogonsCount")

	' ��ȭ�� �α׿�: ��ȣ ���� ���� ��ȣ�� �����ϵ��� �˸�
        VAL_PASSWORDEXPIRYWARNING = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\PasswordExpiryWarning")
	
	' �۾������췯 (��� �۾� ��� ����) 
        VAL_TASKSCHEDULER_LOG = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-TaskScheduler/Operational\Enabled")

	If VAL_SHUTDOWNWITHOUTLOGON = 0  Then 
	   SHUTDOWNWITHOUTLOGON_USE = "��ȣ / ������"
        else
	   SHUTDOWNWITHOUTLOGON_USE = "��� / �����"
        end if
               
        If VAL_DONTDISPLAYLASTUSERNAME  = 0 Then 
            DONTDISPLAYLASTUSERNAME_USE = "��� / ������"
        else
            DONTDISPLAYLASTUSERNAME_USE = "��ȣ / �����"
        end if

        If VAL_CACHEDLOGONSCOUNT <= 10 Then 
            CACHEDLOGONSCOUNT_USE = "��� / ĳ���� �α׿� [" & VAL_CACHEDLOGONSCOUNT & "]"
        else
            CACHEDLOGONSCOUNT_USE = "��ȣ / ĳ���� �α׿� [" & VAL_CACHEDLOGONSCOUNT & "]"
        end if

        If VAL_PASSWORDEXPIRYWARNING <= 10 Then 
            PASSWORDEXPIRYWARNING_USE = "��� / [" & VAL_PASSWORDEXPIRYWARNING & "] ��"
        else
            PASSWORDEXPIRYWARNING_USE = "��ȣ / [" & VAL_PASSWORDEXPIRYWARNING & "] ��"
        end if

        If VAL_TASKSCHEDULER_LOG = 1 Then 
            TASKSCHEDULER_LOG_USE = "��ȣ / ��� �۾���� ����"
        else
            TASKSCHEDULER_LOG_USE = "��� / ��� �۾���� ������ ����"
        end if

        
        wscript.echo "1.�α׿����� �ʰ� �ý��� ���� ��� " & vbcrlf & _
                     "- ���: " & SHUTDOWNWITHOUTLOGON_USE  & vbcrlf & _
                     "2.������ ����� �̸� ǥ�� �� �� " & vbcrlf & _
                     "- ���: " & DONTDISPLAYLASTUSERNAME_USE & vbcrlf & _
                     "3.ĳ���� �α׿��� Ƚ��(������ ��Ʈ�ѷ��� ��� �Ұ����� ���) " & vbcrlf & _
		     "- ���: " & CACHEDLOGONSCOUNT_USE  & vbcrlf & _
                     "4.��ȭ�� �α׿�: ��ȣ ���� ���� ��ȣ�� �����ϵ��� �˸� " & vbcrlf & _
		     "- ���: " & PASSWORDEXPIRYWARNING_USE  & vbcrlf & _ 
                     "5.�۾� �����췯 (��� �۾������ ����) " & vbcrlf & _
		     "- ���: " & TASKSCHEDULER_LOG_USE & vbcrlf
    end function

end If