
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
' 잡다한 레지스트리 보안설정 모음
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
	
	' 시스템 종료: 로그온하지 않고 시스템 종료 허용
        VAL_NOTICECAPTION = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\shutdownwithoutlogon")

	' 대화형 로그온: 마지막 사용자 이름 표시 안 함
        VAL_DONTDISPLAYLASTUSERNAME = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\dontdisplaylastusername")

	' 대화형 로그온: 캐시할 로그온의 횟수(도메인 컨트롤러가 사용 불가능할 경우)
        VAL_CACHEDLOGONSCOUNT = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\CachedLogonsCount")

	' 대화형 로그온: 암호 만료 전에 암호를 변경하도록 알림
        VAL_PASSWORDEXPIRYWARNING = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\PasswordExpiryWarning")
	
	' 작업스케쥴러 (모든 작업 기록 남김) 
        VAL_TASKSCHEDULER_LOG = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\Microsoft-Windows-TaskScheduler/Operational\Enabled")

	If VAL_SHUTDOWNWITHOUTLOGON = 0  Then 
	   SHUTDOWNWITHOUTLOGON_USE = "양호 / 사용안함"
        else
	   SHUTDOWNWITHOUTLOGON_USE = "취약 / 사용함"
        end if
               
        If VAL_DONTDISPLAYLASTUSERNAME  = 0 Then 
            DONTDISPLAYLASTUSERNAME_USE = "취약 / 사용안함"
        else
            DONTDISPLAYLASTUSERNAME_USE = "양호 / 사용함"
        end if

        If VAL_CACHEDLOGONSCOUNT <= 10 Then 
            CACHEDLOGONSCOUNT_USE = "취약 / 캐시할 로그온 [" & VAL_CACHEDLOGONSCOUNT & "]"
        else
            CACHEDLOGONSCOUNT_USE = "양호 / 캐시할 로그온 [" & VAL_CACHEDLOGONSCOUNT & "]"
        end if

        If VAL_PASSWORDEXPIRYWARNING <= 10 Then 
            PASSWORDEXPIRYWARNING_USE = "취약 / [" & VAL_PASSWORDEXPIRYWARNING & "] 일"
        else
            PASSWORDEXPIRYWARNING_USE = "양호 / [" & VAL_PASSWORDEXPIRYWARNING & "] 일"
        end if

        If VAL_TASKSCHEDULER_LOG = 1 Then 
            TASKSCHEDULER_LOG_USE = "양호 / 모든 작업기록 남김"
        else
            TASKSCHEDULER_LOG_USE = "취약 / 모든 작업기록 남기지 않음"
        end if

        
        wscript.echo "1.로그온하지 않고 시스템 종료 허용 " & vbcrlf & _
                     "- 결과: " & SHUTDOWNWITHOUTLOGON_USE  & vbcrlf & _
                     "2.마지막 사용자 이름 표시 안 함 " & vbcrlf & _
                     "- 결과: " & DONTDISPLAYLASTUSERNAME_USE & vbcrlf & _
                     "3.캐시할 로그온의 횟수(도메인 컨트롤러가 사용 불가능할 경우) " & vbcrlf & _
		     "- 결과: " & CACHEDLOGONSCOUNT_USE  & vbcrlf & _
                     "4.대화형 로그온: 암호 만료 전에 암호를 변경하도록 알림 " & vbcrlf & _
		     "- 결과: " & PASSWORDEXPIRYWARNING_USE  & vbcrlf & _ 
                     "5.작업 스케쥴러 (모든 작업기록을 남김) " & vbcrlf & _
		     "- 결과: " & TASKSCHEDULER_LOG_USE & vbcrlf
    end function

end If