
Option Explicit

'-------------------------------------------
' UAC 실행 삽입
'-------------------------------------------
dim objShell

if WScript.Arguments.length =0 Then
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument with leading blank space, say [ uac ]
    objShell.ShellExecute "wscript.exe", Chr(34) & _
    WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    
else

    '-------------------------------------------
    ' 스크립트 실행 설정
    '-------------------------------------------

    dim a1, a2, a3, a6, a7, a8, a9, a10, a11, a12
    dim b1, b2, c1, c2, c4
    dim d1, d2
    dim e1, e2, e3, e4, e5, e6, e7
    dim f1
    dim g5, g6, g7, g8
    dim h1

    a1 = GET_NEWADMINACCOUNT_NAME()
    a2 = GET_GUESTACCOUT_STATUS()
    a3 = GET_ACCOUNT_LOCKOUT_POLICY()
    a6 = GET_MIN_PASSWORD_LENGTH()
    a7 = GET_MAX_PASSWORD_AGE()
    a8 = GET_MIN_PASSWORD_AGE()
    a9 = GET_PASSWORD_HISTORY()
    a10 = GET_PASSWORDCOMPLEXITY()
    a11 = GET_CLEAR_TEXT_PASSWORD()
    a12 = GET_UACCONTROL_POLICY()
    b1 = GET_DEFAULT_SHARE()
    b2 = GET_EVERYONE_SHARE_FOLDER()
    c1 = GET_UNUSE_WINDOWS_SERVICES()
    c2 = GET_REMOTE_DESKTOP_PORT()
    c4 = GET_SNMP_COMMUNITY_STRINGS()
    d1 = GET_LOCAL_AUDIT_POLICY()
    d2 = GET_EVENT_LOG_SET_SIZE()
    e1 = GET_SCREEN_SAVER()
    e2 = GET_DISPLAY_LAST_LOGON_USERNAME()
    e3 = GET_SHUTDOWNWITHOUTLOGON()
    e4 = GET_LOGON_NOTICE_POLICY()
    e5 = GET_NULLSESSIONL_POLICY()
    e6 = GET_AUTOLOGON_POLICY()
    e7 = GET_TASKSCHEDULER_SETLOG()
    f1 = GET_ANTIVIRUS_STATUS()
    g5 = GET_LANMANAGER_POLICY()
    g6 = GET_DONOT_SAVE_LMHASH_POLICY()
    g7 = GET_CACHED_LOGON_COUNT_POLICY()
    g8 = GET_INCLUDESANONYMOUS_POLICY()
    h1 = GET_WINDOWSUPDATE_STATUS()

    '-------------------------------------------
    ' 점검 결과 저장
    '-------------------------------------------

    Function LOG_FILE(TITLE,RESULT)

        dim objFSO, objEnv, objobjLOGFILE, objWSHShell  
        dim objSCRIPTDIR, objLOGFILE
        
        Set objWSHShell = CreateObject("WScript.Shell")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objEnv = objWSHShell.Environment("process")

        objSCRIPTDIR = objFSO.GetParentFolderName(WScript.ScriptFullName)
        objLOGFILE = objSCRIPTDIR & "\" & objEnv("COMPUTERNAME") & "_checkRESULT.txt"

        if objFSO.fileExists(objLOGFILE) then 
            '이미 화일이 존재한다면 화일을 추가쓰기(Append)로 연다.
            set objobjLOGFILE = objFSO.OpenTextFile(objLOGFILE,8,true)
        else
            '아직 화일이 존재하지 않는다면 새로만든다.
            set objobjLOGFILE = objFSO.CreateTextFile(objLOGFILE,True)
        end if

        objobjLOGFILE.writeline ("[항목]: " & TITLE & vbcrlf & _
                              "- 결과: " & RESULT & vbcrlf)
        Set objFSO = Nothing
        objobjLOGFILE.close

    end Function


    '-------------------------------------------
    ' 보안 정책 가져오기
    '-------------------------------------------

    Function GET_LOCAL_POLICY(search_section, search_key)

        dim objWSHShell, objFSO, objEnv, objstrDUMP_SECPOLICY_FILE
        dim objSCRIPTDIR, strCMD, strDUMP_SECPOLICY
        dim intEqualPos
        dim readini, strFilePath, strKey, strLeftString, strLine, strSection
        dim myFilePath

        Set objWSHShell = CreateObject("WScript.Shell")
        Set objFSO= CreateObject("Scripting.FileSystemObject")
        Set objEnv = objWSHShell.Environment("process")

        objSCRIPTDIR = objFSO.GetParentFolderName(WScript.ScriptFullName)
        strDUMP_SECPOLICY = objSCRIPTDIR & "\" & objEnv("COMPUTERNAME") & "_secedit_dump.txt"

        ' 로컬 보안 정책 파일이 존재하지 않을 경우 덤프 실행 
        If not objFSO.FileExists( strDUMP_SECPOLICY ) Then 
            strCMD = "%comspec% /c %SYSTEMROOT%\system32\"'strCMD header
            objWSHShell.run strCMD & "secedit /export /cfg " & """" & strDUMP_SECPOLICY & """", 0
            'objWSHShell.run strCMD & "secedit /export /cfg " & strDUMP_SECPOLICY, 0
            wscript.sleep 1000
        End if

        Const ForReading   = 1
        Const ForWriting   = 2
        Const ForAppending = 8

        readIni     = ""
        strSection  = Trim( search_section )
        strKey      = Trim( search_key )

        If objFSO.FileExists( strDUMP_SECPOLICY ) Then
            Set objstrDUMP_SECPOLICY_FILE = objFSO.OpenTextFile( strDUMP_SECPOLICY, ForReading, False, True )
            Do While objstrDUMP_SECPOLICY_FILE.AtEndOfStream = False
                strLine = objstrDUMP_SECPOLICY_FILE.ReadLine

                If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                    strLine = Trim( objstrDUMP_SECPOLICY_FILE.ReadLine )

                    Do While Left( strLine, 1 ) <> "["
                        ' Find position of equal sign in the line
                        intEqualPos = InStr( 1, strLine, "=", 1 )

                        If intEqualPos > 0 Then
                            strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )

                            If LCase( strLeftString ) = LCase( strKey ) Then                            
                                readIni = Trim( Mid( strLine, intEqualPos + 1 ) )

                                If readIni = "" Then
                                    readIni = "NULL"
                                End If

                                Exit Do
                            End If
                        End If

                        If objstrDUMP_SECPOLICY_FILE.AtEndOfStream Then Exit Do
                        strLine = Trim( objstrDUMP_SECPOLICY_FILE.ReadLine )
                    Loop
                Exit Do
                End If
            Loop
            objstrDUMP_SECPOLICY_FILE.Close
        Else
            WScript.Echo strDUMP_SECPOLICY & " doesn't exists. Exiting..."
            Wscript.Quit 1
        End If

    GET_LOCAL_POLICY = strLeftString & "," & readini

    End Function 

   '----------------------------------------------------
   ' 스크립트 시작 
   '----------------------------------------------------

   ' a1
    Function GET_NEWADMINACCOUNT_NAME()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim COMPARE_RESULT
        dim DEFAULT_NAME
        dim NEW_ADMINACCOUNT_NAME

        TITLE ="1.1 Administrator 계정 관리"
        search_section = "System Access"
        search_key = "NewAdministratorName"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        DEFAULT_NAME="""" & "Administrator" & """"
        DEFAULT_NAME = lcase(DEFAULT_NAME)
        value = lcase(value)

        COMPARE_RESULT = strcomp(value, DEFAULT_NAME, 1)

        if COMPARE_RESULT = 0 then
            NEW_ADMINACCOUNT_NAME = "취약" & vbcrlf & "- 계정이름: [ " & value & " ] 로 설정 됨"
        else
            NEW_ADMINACCOUNT_NAME = "양호" & vbcrlf & "- 계정이름: [ " & value & " ] 로 설정 됨"
        end if

        call LOG_FILE(TITLE,NEW_ADMINACCOUNT_NAME)
    End Function

    ' a2
    Function GET_GUESTACCOUT_STATUS()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim GUESTACCOUNT_USE

        TITLE ="1.2 GUEST 계정 관리"
        search_section = "System Access"
        search_key = "EnableGuestAccount"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        if value = 0 then
            GUESTACCOUNT_USE = "양호 [ Guest 계정 사용안함 ]"
        else
            GUESTACCOUNT_USE = "취약 [ Guest 계정 사용 ]"
        end if

        call LOG_FILE(TITLE,GUESTACCOUNT_USE)
    End Function

    ' a3
    Function GET_ACCOUNT_LOCKOUT_POLICY()

        dim TITLE
        dim search_section, search_key, search_key_1, search_key_2
        dim RESULT, RESULT_1, RESULT_2 
        dim SPLIT_RESULT, SPLIT_RESULT_1, SPLIT_RESULT_2

        dim keyname, value, value1, value2
        dim ACCOUNT_LOCKOUT_USE

        TITLE ="1.3 계정잠금 정책"
        search_section = "System Access"
        search_key = "LockoutBadCount"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        if value = 0 then
            ACCOUNT_LOCKOUT_USE = "취약 [ 계정 잠그지 않음 ]"
        else
            search_key_1 = "LockoutDuration"
            search_key_2 = "ResetLockoutCount"
         
            RESULT_1 = GET_LOCAL_POLICY(search_section, search_key_1)
            RESULT_2 = GET_LOCAL_POLICY(search_section, search_key_2)

            SPLIT_RESULT_1 = Split(RESULT_1,",")
            SPLIT_RESULT_2 = Split(RESULT_2,",")

            value1 = TRIM(SPLIT_RESULT_1(1))
            value2 = TRIM(SPLIT_RESULT_2(1))
            
            if value = 5 then
                ACCOUNT_LOCKOUT_USE = vbcrlf & "1.3 계정 잠금 회수: 양호" & _
                                      "[ " & value & " ] 번의 잘못된 로그온 시도"
            else
                ACCOUNT_LOCKOUT_USE = vbcrlf & "1.3 계정 잠금 회수: 취약" & _
                                      "[ " & value & " ] 번의 잘못된 로그온 시도"
            end if
            
            if value1 = 30 then
                ACCOUNT_LOCKOUT_USE = ACCOUNT_LOCKOUT_USE & vbcrlf & _
                                      "1.4 계정 잠금 기간: 양호 [ " & value1 & " ] 분 동안 계정 잠금" 
            else
                ACCOUNT_LOCKOUT_USE = ACCOUNT_LOCKOUT_USE & vbcrlf & _
                                      "1.4 계정 잠금 기간: 취약 [ " & value1 & " ] 분 동안 계정 잠금"
            end if

            if value2 = 30 then
                ACCOUNT_LOCKOUT_USE = ACCOUNT_LOCKOUT_USE & vbcrlf & _
                                      "1.5 다음시간 후 계정 잠금수를 원래대로 설정: 양호" & _
                                      "[ " & value2 & " ] 분"
            else
                ACCOUNT_LOCKOUT_USE = ACCOUNT_LOCKOUT_USE & vbcrlf & _ 
                                      "1.5 다음시간 후 계정 잠금수를 원래대로 설정: 취약" & _
                                      "[ " & value2 & " ] 분"
            end if
        end if

        call LOG_FILE(TITLE,ACCOUNT_LOCKOUT_USE)
    End Function

    ' a6 
    Function GET_MIN_PASSWORD_LENGTH()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim MIN_PASSWORD_LENGTH_USE

        TITLE = "1.6 패스워드 최소 길이 설정"
        search_section = "System Access"
        search_key = "MinimumPasswordLength"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        if value >= 8 then
            MIN_PASSWORD_LENGTH_USE = "양호" & vbcrlf & "- 내용: [ " & value & " ] 자리로 설정 됨"
        else
            MIN_PASSWORD_LENGTH_USE = "취약" & vbcrlf & "- 내용: [ " & value & " ] 자리로 설정 됨"
        end if

        call LOG_FILE(TITLE,MIN_PASSWORD_LENGTH_USE)
    End Function

    ' a7 
    Function GET_MAX_PASSWORD_AGE()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim MAX_PASSWORD_AGE_USE

        TITLE = "1.7 패스워드 최대 사용시간 설정"
        search_section = "System Access"
        search_key = "MaximumPasswordAge"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        
        if value >= 90 then
            MAX_PASSWORD_AGE_USE = "양호" & vbcrlf & "- 내용: [ " & value & " ] 일"
        else
            MAX_PASSWORD_AGE_USE = "취약" & vbcrlf & "- 내용: [ " & value & " ] 일"
        end if

        call LOG_FILE(TITLE,MAX_PASSWORD_AGE_USE)
    End Function

    ' a8
    Function GET_MIN_PASSWORD_AGE()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim MIN_PASSWORD_AGE_USE

        TITLE = "1.8 패스워드 최소 사용기간 설정"
        search_section = "System Access"
        search_key = "MinimumPasswordAge"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        if value >= 1 then
            MIN_PASSWORD_AGE_USE = "양호" & vbcrlf & "- 내용: [ " & value & " ] 일"
        else
            MIN_PASSWORD_AGE_USE = "취약" & vbcrlf & "- 내용: [ " & value & " ] 일"
        end if

        call LOG_FILE(TITLE,MIN_PASSWORD_AGE_USE)
    End Function


    ' a9
    Function GET_PASSWORD_HISTORY()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim PASSWORD_HISTORY_USE

        TITLE = "1.9 사용자 최근 암호 기억"
        search_section = "System Access"
        search_key = "PasswordHistorySize"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))
        
        if value >= 6 then
            PASSWORD_HISTORY_USE = "양호" & vbcrlf & "- 내용: [ " & value & " ] 개 암호 기억됨"
        else
            PASSWORD_HISTORY_USE = "취약" & vbcrlf & "- 내용: [ " & value & " ] 개 암호 기억됨"
        end if

        call LOG_FILE(TITLE,PASSWORD_HISTORY_USE)
    End Function

    ' a10
    Function GET_PASSWORDCOMPLEXITY()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim PASSWORD_COMPLEX_USE

        TITLE = "1.10 암호 복잡성 사용"
        search_section = "System Access"
        search_key = "PasswordComplexity"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        if value = 1 then
            PASSWORD_COMPLEX_USE = "양호"
        else
            PASSWORD_COMPLEX_USE = "취약"
        end if

        call LOG_FILE(TITLE,PASSWORD_COMPLEX_USE)

    End Function

    ' a11 
    Function GET_CLEAR_TEXT_PASSWORD()

        dim TITLE
        dim search_section, search_key
        dim RESULT, SPLIT_RESULT
        dim keyname, value
        dim CLEAR_TEXT_PASSWORD_USE

        TITLE = "1.11 해독가능한 암호화를 사용하여 암호 저장"
        search_section = "System Access"
        search_key = "ClearTextPassword"

        RESULT = GET_LOCAL_POLICY(search_section, search_key)

        SPLIT_RESULT = Split(RESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        if value = 0 then
            CLEAR_TEXT_PASSWORD_USE = "양호 [ 사용안함 ]"
        else
            CLEAR_TEXT_PASSWORD_USE = "취약 [ 사용 ]"
        end if

        call LOG_FILE(TITLE,CLEAR_TEXT_PASSWORD_USE)
    End Function

    ' a12
    Function GET_UACCONTROL_POLICY()

        dim TITLE 
        dim objWSHShell
        dim VAL_PROMPTBEHAVIORADMIN, VAL_PROMPTONSECUREDESKTOP
        dim intPROMPTBEHAVIORADMIN_CHECK, intPROMPTONSECUREDESKTOP_CHECK
        dim UACCONTROL_CHECK, UACCONTROL_USE
        dim RESULT_BEHAVIORADMIN, RESULT_SECUREDESKTOP

        TITLE ="1.12 사용자 계정 컨트롤(User Account control) 설정"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_PROMPTBEHAVIORADMIN = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
                                                      "Policies\System\ConsentPromptBehaviorAdmin")

        VAL_PROMPTONSECUREDESKTOP = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
                                                        "Policies\System\PromptOnSecureDesktop") 

        if VAL_PROMPTBEHAVIORADMIN = 5 then
            intPROMPTBEHAVIORADMIN_CHECK = 1
        else
            intPROMPTBEHAVIORADMIN_CHECK = 0
        end if

        if VAL_PROMPTONSECUREDESKTOP = 1 then
            intPROMPTONSECUREDESKTOP_CHECK = 1
        else
            intPROMPTONSECUREDESKTOP_CHECK = 0
        end if

        UACCONTROL_CHECK = intPROMPTBEHAVIORADMIN_CHECK + intPROMPTONSECUREDESKTOP_CHECK

        if UACCONTROL_CHECK  = 2 then
            UACCONTROL_USE = "양호 [ UAC 사용중 ]"
        else

            select case VAL_PROMPTBEHAVIORADMIN 
                Case "0"
                    RESULT_BEHAVIORADMIN = "* UAC 사용안함"
                Case "2"
                    RESULT_BEHAVIORADMIN = "* UAC 모든사항 항상 알림"
                Case "5"
                    RESULT_BEHAVIORADMIN = "* 프로그램에서 사용자 몰래 변경하는경우 알림"
            end select 

            select case VAL_PROMPTONSECUREDESKTOP 
                Case "0"
                    RESULT_SECUREDESKTOP = "* 바탕화면 흐리게 표시하지 않음"
                Case "1"
                    RESULT_SECUREDESKTOP = "* 바탕화면 흐리게 표시 함"
            end select 

            UACCONTROL_USE = "취약" & vbcrlf & "- 내용: " & vbcrlf & _
                             RESULT_BEHAVIORADMIN & vbcrlf & _
                             RESULT_SECUREDESKTOP 
        end if
        
        set objWSHShell = nothing
        call LOG_FILE(TITLE,UACCONTROL_USE)

    end Function

    ' b1
    Function GET_DEFAULT_SHARE()

        dim TITLE
        dim strComputer,i
        dim objWMIService,colItems,objitem
        dim arraySHARE_LIST,RESULT,intCOUNT,LIST

        TITLE = "2.1 하드 디스크 기본 공유 제거"
        strComputer = "." 
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
        Set colItems = objWMIService.ExecQuery( _
             "SELECT * FROM Win32_Share",,48)  

        arraySHARE_LIST=array("Admin$","C$","D$","E$","F$","G$")
        intCOUNT = 0

        for each objItem in colItems 
            for i=Lbound(arraySHARE_LIST) to Ubound(arraySHARE_LIST)
                if StrComp(lcase(arraySHARE_LIST(i)), lcase(objItem.Name), 1) = 0 Then
                    intCOUNT = intCOUNT + 1
                    LIST= LIST & " " & objItem.Name               
                end if
            next
        next
        
        if intCOUNT <= 0 then
            RESULT = "양호" & vbcrlf & "- 내용: [ 기본공유 사용안함 ]"
        else
            RESULT = "취약" & vbcrlf & "- 내용: [ " & LIST & " ] 사용 중"
        end if
            
        call LOG_FILE(TITLE,RESULT)  
    end Function

    ' b2
    Function GET_EVERYONE_SHARE_FOLDER()

        dim TITLE
        dim RESULT

        TITLE = "2.2 공유 권한 및 사용자 그룹 설정"
        RESULT = "!!스크립트 만들어 주세요~~~"

        call LOG_FILE(TITLE,RESULT)  
    end Function

    ' c1
    Function GET_UNUSE_WINDOWS_SERVICES()

        dim TITLE
        dim strComputer,i
        dim objWMIService,colItems,objitem
        dim arraySERVICE_LIST,RESULT,intCOUNT,LIST
  
        TITLE ="3.1 불필요한 서비스 제거"
        strComputer = "." 
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
        Set colItems = objWMIService.ExecQuery( _
            "SELECT * FROM Win32_Service WHERE State = 'Running'",,48) 

        arraySERVICE_LIST=array("Application Layer Gateway Service","Certificate Propagation","DFSR","Diagnostic Policy Service", _
        "Diagnostic Service Host","Diagnostic System Host","Disk Defragmenter","Encrypting File System (EFS)", _
        "Extensible Authentication Protocol","IKE and AuthIP IPsec Keying Modules","Interactive Services Detection", _
        "Internet Connection Sharing (ICS)","IP Helper","IsmServ","KtmRm for Distributed Transaction Coordinator", _
        "Link-Layer Topology Discovery Mapper","Multimedia Class Scheduler","Network Access Protection Agent", _
        "PnP-X IP Bus Enumerator","Portable Device Enumerator Service","Print Spooler", _
        "Problem Reports and Solutions Control Panel Support","Protected Storage","Remote Access Auto Connection Manager", _
        "Remote Access Connection Manager","Routing and Remote Access","Secure Socket Tunneling Protocol Service", _
        "Shell Hardware Detection","Smart Card","Smart Card Removal Policy","SNMP Trap", _
        "Special Administration Console Helper","SSDP Discovery","Telephony","Thread Ordering Server", _
        "TPM Base Services","UPnP Device Host","Windows Audio","Windows Audio Endpoint Builder", _
        "Windows CardSpace","Windows Error Reporting Service","WinHTTP Web Proxy Auto-Discovery Service", _
        "Wired AutoConfig","Application Identity","Windows Remote Management (WS-Management)","Software Protection")

        intCOUNT = 0

        for each objItem in colItems 
            for i=Lbound(arraySERVICE_LIST) to Ubound(arraySERVICE_LIST)
                if StrComp(lcase(arraySERVICE_LIST(i)), lcase(objItem.DisplayName), 1) = 0 Then
                    RESULT = TRUE
                        intCOUNT = intCOUNT+1
                    LIST= LIST & vbcrlf & objItem.DisplayName & "  | " & objItem.state
                    'wscript.echo "불필요서비스: "  &  LIST
                else
                    RESULT = FALSE
                end if
            next
        next

        if intCOUNT = 0 then
            RESULT = "양호" & vbcrlf & "- 내용: 불필요 서비스 [ " & intCOUNT & " ] 개 발견"
        else
            RESULT = "취약" & vbcrlf & "- 내용: 불필요 서비스 [ " & intCOUNT & " ] 개 발견" & vbcrlf & _
                     "- 목록: " & LIST
        end if

        call LOG_FILE(TITLE,RESULT)  
    end Function

    ' c2
    Function GET_REMOTE_DESKTOP_PORT()

        dim TITLE
        dim objWSHShell, VAL_STATUS, VAL_PORT_NUMBER, RESULT

        TITLE ="3.2 터미널 서비스 포트 변경"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_STATUS = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\" & _
                                         "Terminal Server\fDenyTSConnections")

        VAL_PORT_NUMBER = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\" & _ 
                                              "Terminal Server\WinStations\RDP-Tcp\PortNumber")

        if VAL_STATUS = 1 then
            RESULT = "양호" & vbcrlf & "- 내용: 터미널 서비스 사용안함"
        else
            if PORT_NUMBER = "13389" then
                RESULT = "양호" & vbcrlf & "- 내용: 포트번호 [ " & PORT_NUMBER & " ]"
            else
                RESULT = "취약" & vbcrlf & "- 내용: 포트번호 [ " & PORT_NUMBER & " ]"
            end if

        end if
        set objWSHShell = nothing

        call LOG_FILE(TITLE,RESULT)  
    end Function

    ' c4
    Function GET_SNMP_COMMUNITY_STRINGS()

        dim TITLE
        dim RESULT
        TITLE = "3.4 SNMP 서비스 커뮤니티 스트링의 복잡성 설정"
        RESULT ="!!스크립트 만들어 주세요~~~"

        call LOG_FILE(TITLE,RESULT)  
    end Function

    ' d1 
    Function GET_LOCAL_AUDIT_POLICY()
        dim TITLE 
        dim search_section
        dim search_key, search_key_1, search_key_2, search_key_3, search_key_4
        dim search_key_5, search_key_6, search_key_7, search_key_8
        dim keyname, value 
        dim RESULT, RESULT_1, RESULT_2, RESULT_3, RESULT_4
        dim RESULT_5, RESULT_6, RESULT_7, RESULT_8, RESULT_9 
        dim varRESULT, SPLIT_RESULT
        dim LOCAL_AUDIT_POLICY

        TITLE ="4.1 보안로그의 감사 기능 설정"
        search_section = "Event Audit"
        search_key = "AuditObjectAccess"      '개체 엑세스 감사
        search_key_1 = "AuditAccountManage"   '계정 관리 감사
        search_key_2 = "AuditLogonEvents"     '계정 로그온 이벤트 감사
        search_key_3 = "AuditPrivilegeUse"    '권한 사용 감사 
        search_key_4 = "AuditLogonEvents"     '로그온 이벤트 감사
        search_key_5 = "AuditSystemEvents"    '시스템 이벤트 감사
        search_key_6 = "AuditProcessTracking" '프로세스 추적 감사 
        search_key_7 = "AuditPolicyChange"    '정책 변경 감사
        search_key_8 = "AuditDSAccess"        '디렉토리 서비스 엑세스 감사

        RESULT = array(0,1,2,3,4,5,6,7,8)

        RESULT_1 = GET_LOCAL_POLICY(search_section, search_key)
        RESULT_2 = GET_LOCAL_POLICY(search_section, search_key_1)
        RESULT_3 = GET_LOCAL_POLICY(search_section, search_key_2)
        RESULT_4 = GET_LOCAL_POLICY(search_section, search_key_3)
        RESULT_5 = GET_LOCAL_POLICY(search_section, search_key_4)
        RESULT_6 = GET_LOCAL_POLICY(search_section, search_key_5)
        RESULT_7 = GET_LOCAL_POLICY(search_section, search_key_6)
        RESULT_8 = GET_LOCAL_POLICY(search_section, search_key_7)
        RESULT_9 = GET_LOCAL_POLICY(search_section, search_key_8)
            
        RESULT(0) = RESULT_1
        RESULT(1) = RESULT_2
        RESULT(2) = RESULT_3
        RESULT(3) = RESULT_4
        RESULT(4) = RESULT_5
        RESULT(5) = RESULT_6
        RESULT(6) = RESULT_7
        RESULT(7) = RESULT_8
        RESULT(8) = RESULT_9

        For each varRESULT in RESULT
            SPLIT_RESULT = Split(varRESULT,",")

            keyname = TRIM(SPLIT_RESULT(0))
            value = TRIM(SPLIT_RESULT(1))

            select case value
                case 0
                    value = "설정 안됨 [ 취약 ]"
                case 1
                    value = "성공 [ 취약 ]"
                case 2
                    value = "실패 [ 취약 ]"
                case 3
                    value = "성공,실패 [ 양호 ]"
            End select

            select case keyname
                case "AuditObjectAccess"
                    keyname = vbcrlf & "* 개체 엑세스 감사"
                case "AuditAccountManage"
                    keyname = vbcrlf & "* 계정 관리 감사"
                case "AuditLogonEvents"
                    keyname = vbcrlf & "* 계정 로그온 이벤트 감사"
                case "AuditPrivilegeUse"
                    keyname = vbcrlf & "* 권한 사용 감사"
                case "AuditLogonEvents"
                    keyname = vbcrlf & "* 로그온 이벤트 감사"
                case "AuditSystemEvents"  
                    keyname =vbcrlf & "* 시스템 이벤트 감사"
                case "AuditProcessTracking"
                    keyname =vbcrlf & "* 프로세스 추적 감사"
                case "AuditDSAccess"
                    keyname = vbcrlf & "* 디렉토리 서비스 엑세스 감사"
                case "AuditPolicyChange"
                    keyname = vbcrlf & "* 정책변경 감사"
            End select        

            LOCAL_AUDIT_POLICY = LOCAL_AUDIT_POLICY & keyname & " : " & value
            
        Next

        call LOG_FILE(TITLE,LOCAL_AUDIT_POLICY)
    End Function

    'd2
    Function GET_EVENT_LOG_SET_SIZE()

        dim TITLE
        dim objWSHShell
        dim VAL_SECURITY, VAL_APPLICATION, VAL_SYSTEM
        dim SECURITY_SIZE, APPLICATION_SIZE, SYSTEM_SIZE
        dim RESULT

        TITLE ="4.2 이벤트 뷰어 설정"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_SECURITY = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\" & _
                                           "Microsoft\Windows\EventLog\Security\MaxSize")

        VAL_APPLICATION = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\" & _ 
                                              "Microsoft\Windows\EventLog\Application\MaxSize") 

        VAL_SYSTEM = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\" & _
                                         "Microsoft\Windows\EventLog\System\MaxSize")


        if VAL_SECURITY >= "524288" then
            SECURITY_SIZE = "양호 [ " & VAL_SECURITY & " ] KB"
        else
            SECURITY_SIZE = "취약 [ " & VAL_SECURITY & " ] KB / 524,288KB 이상 설정 필요"
        end if

        if VAL_APPLICATION >= "524288" then
            APPLICATION_SIZE = "양호 [ " & VAL_APPLICATION & " ] KB"
        else
            APPLICATION_SIZE = "취약 [ " & VAL_APPLICATION & " ] KB / 524,288KB 이상 설정 필요"
        end if

        if VAL_SYSTEM >= "524288" then
            SYSTEM_SIZE = "양호 [ " & VAL_SYSTEM & " ] KB"
        else
            SYSTEM_SIZE = "취약 [ " & VAL_SYSTEM  & " ] KB / 524,288KB 이상 설정 필요"
        end if

        RESULT = vbcrlf &  "* 보안 로그 설정 : " & SECURITY_SIZE & vbcrlf & _
                           "* 어플리케이션 로그 설정  : " & APPLICATION_SIZE & vbcrlf & _
                           "* 시스템 로그 설정 : " & SYSTEM_SIZE 
        set objWSHShell = nothing

        call LOG_FILE(TITLE,RESULT)
    end Function

    'e1
    Function GET_SCREEN_SAVER()

        dim TITLE
        dim objWSHShell
        dim VAL_SAVER_USE, VAL_SAVER_PROTECT, VAL_SAVER_TIME
        dim SAVER_USE, SAVER_PROTECT, SAVER_TIME
        dim RESULT

        TITLE = "5.1 로그오프나 워크스테이션 잠김"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_SAVER_USE = objWSHShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveActive")
        VAL_SAVER_PROTECT = objWSHShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaverIsSecure") 
        VAL_SAVER_TIME = objWSHShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveTimeOut")

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
            SAVER_TIME = "취약 [ " & VAL_SAVER_TIME & " ] 초 / 10분 이상으로 설정 필요"
        end if
 
        RESULT = vbcrlf &  "* 화면보호기 사용 : " & SAVER_USE & vbcrlf & _
                           "* 화면 보호기 암호로 보호  : " & SAVER_PROTECT & vbcrlf & _
                           "* 화면 보호기 대기시간 : " & SAVER_TIME 

        set objWSHShell = nothing
        call LOG_FILE(TITLE,RESULT)
    end function

    'e2
    Function GET_DISPLAY_LAST_LOGON_USERNAME()

        dim TITLE
        dim objWSHShell
        dim VAL_DONTDISPLAYLASTUSERNAME
        dim DONTDISPLAYLASTUSERNAME_USE
     
        TITLE ="5.2 마지막 로그온 사용자 계정 숨김"
        Set objWSHShell = CreateObject("WScript.Shell")
    
        VAL_DONTDISPLAYLASTUSERNAME = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
                                                          "Policies\System\dontdisplaylastusername")
              
        If VAL_DONTDISPLAYLASTUSERNAME  = 0 Then 
            DONTDISPLAYLASTUSERNAME_USE = "취약 / 사용안함"
        else
            DONTDISPLAYLASTUSERNAME_USE = "양호 / 사용함"
        end if

        set objWSHShell = nothing
        call LOG_FILE(TITLE,DONTDISPLAYLASTUSERNAME_USE)
    end Function

    'e3 
    Function GET_SHUTDOWNWITHOUTLOGON()

        dim TITLE
        dim objWSHShell
        dim VAL_SHUTDOWNWITHOUTLOGON
        dim SHUTDOWNWITHOUTLOGON_USE
    
        TITLE = "5.3 로그온 하지 않은 사용자 시스템 종료 방지"
        Set objWSHShell = CreateObject("WScript.Shell")
    
        VAL_SHUTDOWNWITHOUTLOGON = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\" & _
                                                "CurrentVersion\Policies\System\shutdownwithoutlogon")

        If VAL_SHUTDOWNWITHOUTLOGON = 0  Then 
            SHUTDOWNWITHOUTLOGON_USE = "양호 / 사용안함"
        else
            SHUTDOWNWITHOUTLOGON_USE = "취약 / 사용함"
        end if
              
        set objWSHShell = nothing
        call LOG_FILE(TITLE,SHUTDOWNWITHOUTLOGON_USE)

    end Function    

    'e4 
    Function GET_LOGON_NOTICE_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_NOTICECAPTION, VAL_NOTICETEXT
        dim CAPTION_USE, NOTICETEXT_USE
        dim RESULT

        TITLE = "5.4 로그온 메시지 출력 진단"
        Set objWSHShell = CreateObject("WScript.Shell")
       
        VAL_NOTICECAPTION = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
                                                "Policies\System\legalnoticecaption")

        VAL_NOTICETEXT = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
                                             "Policies\System\legalnoticecaption")
    
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

        RESULT = vbcrlf &  "* 로그온 시도 경고 제목 : " & CAPTION_USE & vbcrlf & _
                           "* 로그온 시도 경고 내용  : " & NOTICETEXT_USE

        set objWSHShell = nothing
        call LOG_FILE(TITLE,RESULT)

    end Function   

    'e5
    Function GET_NULLSESSIONL_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_RESTRCITANOMYMOUS, VAL_RESTRCITANOMYMOUSSAM

        dim intRESTRCITANOMYMOUS_CHECK,intRESTRCITANOMYMOUSSAM_CHECK
        dim NULLSESSION_CHECK, NULLSESSIONL_USE
        dim RESULT_ANOMYMOUS, RESULT_ANOMYMOUSSAM

        TITLE = "5.5 Null Session 설정"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_RESTRCITANOMYMOUS = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\restrictanonymous")
        VAL_RESTRCITANOMYMOUSSAM = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\restrictanonymoussam") 

        if VAL_RESTRCITANOMYMOUS = 1 then
            intRESTRCITANOMYMOUS_CHECK = 1
        else
            intRESTRCITANOMYMOUS_CHECK = 0
        end if

        if VAL_RESTRCITANOMYMOUSSAM = 1 then
            intRESTRCITANOMYMOUSSAM_CHECK = 1
        else
            intRESTRCITANOMYMOUSSAM_CHECK = 0
        end if

        NULLSESSION_CHECK = intRESTRCITANOMYMOUS_CHECK + VAL_RESTRCITANOMYMOUSSAM

        if NULLSESSION_CHECK  = 2 then
            NULLSESSIONL_USE = "양호" & vbcrlf &  "* SAM 계정과 공유의 익명 열거 허용 안 함: 사용 " & vbcrlf & _
                                                  "* SAM 계정의 익명 열거 허용 안 함: 사용" 
        else

            select case VAL_RESTRCITANOMYMOUS
                Case "0"
                    RESULT_ANOMYMOUS = "취약 / SAM 계정과 공유의 익명 열거 허용 안 함"
                Case "1"
                    RESULT_ANOMYMOUS = "양호 / SAM 계정과 공유의 익명 열거 허용 안 함"
            end select 

            select case VAL_RESTRCITANOMYMOUSSAM
                Case "0"
                    RESULT_ANOMYMOUSSAM = "취약 / SAM 계정의 익명 열거 허용 안 함"
                Case "1"
                    RESULT_ANOMYMOUSSAM = "양호 / SAM 계정의 익명 열거 허용 안 함"
            end select 

            NULLSESSIONL_USE = vbcrlf &  "* SAM 계정과 공유의 익명 열거 허용 안 함 : " & RESULT_ANOMYMOUS & vbcrlf & _
                                         "* SAM 계정의 익명 열거 허용 안 함  : " & RESULT_ANOMYMOUSSAM
        end if
 
        set objWSHShell = nothing
        call LOG_FILE(TITLE,NULLSESSIONL_USE)
    end function  

    'e6 
    Function GET_AUTOLOGON_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_AUTOLOGON
        dim AUTOLOGON_USE

        TITLE = "5.6 Autologon 기능 제어"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_AUTOLOGON = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon")

        if VAL_AUTOLOGON = 0 then
            AUTOLOGON_USE = "양호 / 사용안함"
        else
            AUTOLOGON_USE = "취약 / 자동로그인 사용함"
        end if
       
        set objWSHShell = nothing
        call LOG_FILE(TITLE,AUTOLOGON_USE)

    end Function

    'e7
    Function GET_TASKSCHEDULER_SETLOG()

        dim TITLE
        dim objWSHShell
        dim VAL_TASKSCHEDULER_LOG
        dim TASKSCHEDULER_LOG_USE

        TITLE = "5.7 작업스케쥴러 로그 설정"
        Set objWSHShell = CreateObject("WScript.Shell")
  
        VAL_TASKSCHEDULER_LOG = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\" & _
                                                    "CurrentVersion\WINEVT\Channels\Microsoft-Windows-TaskScheduler/Operational\Enabled")

        If VAL_TASKSCHEDULER_LOG = 1 Then 
            TASKSCHEDULER_LOG_USE = "양호 [ 모든 작업기록 남김 ]"
        else
            TASKSCHEDULER_LOG_USE = "취약 [ 모든 작업기록 남기지 않음 ]"
        end if

        set objWSHShell = nothing
        call LOG_FILE(TITLE,TASKSCHEDULER_LOG_USE)
    end Function


    'f1
    Function GET_ANTIVIRUS_STATUS()

        dim TITLE
        dim strAVpath1, strAVpath2
        dim objWSHShell, objFSO
        dim ANTIVIRUS_STATUS_USE
    
        TITLE = "6.백신 프로그램 설치"

        '임시로 알약으로 대체 
        strAVpath1 = "C:\Program Files\ESTsoft\ALYac\AYLaunch.exe"
        strAVpath2 = "C:\Program Files (x86)\ESTsoft\ALYac\AYLaunch.exe"
        
        Set objWSHShell = CreateObject("WScript.Shell")
        Set objFSO = CreateObject("Scripting.FileSystemObject")

        If objFSO.FileExists(strAVpath1) or objFSO.FileExists(strAVpath1) Then
            ANTIVIRUS_STATUS_USE = "양호 / 백신 설치됨"
        else
            ANTIVIRUS_STATUS_USE = "취약 / 백신 설치여부 확인 필요"
        end if
    
        set objWSHShell = nothing
        call LOG_FILE(TITLE,ANTIVIRUS_STATUS_USE)

    end Function

    'g1 7.1 로컬 로그온 허용
    'g2 7.2 시스템 종료
    'g3 7.3 원격 시스템에서 강제로 시스템 종료 (sdb)

    'g5 
    Function GET_LANMANAGER_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_LANMANAGER
        dim LANMANAGER_USE, RESULT_LANMANAGER

        TITLE = "7.5 LAN Manager 인증 수준"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_LANMANAGER = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\LmCompatibilityLevel")

        if VAL_LANMANAGER = 5 then
            LANMANAGER_USE = "양호 / NTLMv2 응답만 보내기 및 LM 및 NTLM 거부"
        else
    
            select case VAL_LANMANAGER 
                Case "0"
                    RESULT_LANMANAGER = "LM 및 NTLM 응답보내기"
                Case "1"
                    RESULT_LANMANAGER = "LM 및 NTLM 응답보내기 - 협상되면 NTLMv2 세션보안 사용"
                Case "2"
                    RESULT_LANMANAGER = "NTLM 응답만 보내기"
                Case "3"
                    RESULT_LANMANAGER = "NTLMv2 응답만 보내기"
                Case "4"
                    RESULT_LANMANAGER = "NTLMv2 응답만 보내기 및 LM 거부 "
                Case "5"
                    RESULT_LANMANAGER = "NTLMv2 응답만 보내기 및 LM 및 NTLM 거부"
            end select

            LANMANAGER_USE = "취약" & vbcrlf & "- 내용: " & _
                             " [ " & RESULT_LANMANAGER & " ] " 
        end if
     
        set objWSHShell = nothing
        call LOG_FILE(TITLE,LANMANAGER_USE)

    end Function

    'g6 
     Function GET_DONOT_SAVE_LMHASH_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_DONOTSAVELMHASH, DONOTSAVELMHASH_USE

        TITLE = "7.6 다음 암호 변경 시 Lan Manager 해시 값 저장 안 함"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_DONOTSAVELMHASH = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\NoLmHash")

        if VAL_DONOTSAVELMHASH= 1 then
            DONOTSAVELMHASH_USE = "양호 / 사용함"
        else
            DONOTSAVELMHASH_USE = "취약 / 사용안함"
        end if
       
        set objWSHShell = nothing
        call LOG_FILE(TITLE,DONOTSAVELMHASH_USE)

    end Function

    'g7 
     Function GET_CACHED_LOGON_COUNT_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_CACHEDLOGONCOUNT, CACHEDLOGONCOUNT_USE

        TITLE = "7.7 캐시할 로그온의 횟수(도메인 컨트롤러가 사용 불가능할 경우)"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_CACHEDLOGONCOUNT = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\CachedLogonsCount")

        if VAL_CACHEDLOGONCOUNT <= 5 then
            CACHEDLOGONCOUNT_USE = "양호" & vbcrlf & "- 내용: 캐시할 개수 [ " & VAL_CACHEDLOGONCOUNT & " ]"
        else
            CACHEDLOGONCOUNT_USE = "취약" & vbcrlf & "- 내용: 캐시할 개수 [ " & VAL_CACHEDLOGONCOUNT & " ]"
        end if
       
        set objWSHShell = nothing
        call LOG_FILE(TITLE,CACHEDLOGONCOUNT_USE)

    end Function

    'g8 
     Function GET_INCLUDESANONYMOUS_POLICY()

        dim TITLE
        dim objWSHShell
        dim VAL_INCLUDESANONYMOUS
        dim INCLUDESANONYMOUS_USE

        TITLE = "7.8 Everyone 사용권한을 익명 사용자에게 적용 안함"
        Set objWSHShell = CreateObject("WScript.Shell")

        VAL_INCLUDESANONYMOUS = objWSHShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\everyoneincludesanonymous")

        if VAL_INCLUDESANONYMOUS= 0 then
            INCLUDESANONYMOUS_USE = "양호 / 사용안함"
        else
            INCLUDESANONYMOUS_USE = "취약 / 사용함"
        end if
       
        set objWSHShell = nothing
        call LOG_FILE(TITLE,INCLUDESANONYMOUS_USE)

    end Function

    'h1 
    Function GET_WINDOWSUPDATE_STATUS()

        dim TITLE
        dim objWSHShell, objFSO, i
        dim UpdateSession, UpdateServiceManager, UpdateService, UpdateSearcher, SearchResult, Updates, update
        dim STR_CURRENT_DITRECTORY, WSUS_DB, MISSING_UPDATE, RESULT, VAL_CHECKMETHOD

        TITLE = "8.보안 패치"
        Set objWSHShell = CreateObject("WScript.Shell")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        STR_CURRENT_DITRECTORY = objFSO.GetParentFolderName(WScript.ScriptFullName)
        WSUS_DB = STR_CURRENT_DITRECTORY & "\wsusscn2.cab"
                
        if (objFSO.fileexists(WSUS_DB)) then
            VAL_CHECKMETHOD = "wsusscn2.cab 파일로 보안업데이트 상태 확인함"

            wscript.echo "wsusscn2.cab 파일을 이용하여 보안업데이트를 체크합니다." & _
                         vbcrlf & "약간의 체크 시간이 소요 됩니다." & _ 
                         vbcrlf & "점검 완료시 점검 완료 팝업이 보여집니다. 이후 점검 파일을 확인해주세요."

            Set UpdateSession = CreateObject("Microsoft.Update.Session")
            Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
            Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", WSUS_DB)
            Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
            UpdateSearcher.ServerSelection = 3
            UpdateSearcher.ServiceID = UpdateService.ServiceID

        else 
            VAL_CHECKMETHOD = "로컬 마지막 윈도우 업데이트 상태로 확인함."

            wscript.echo "wsusscn2.cab 파일이 존재하지 않아" & _
                         "로컬의 마지막 윈도우 업데이트 상태로 보안 업데이트 상태를 체크합니다." & _ 
                         vbcrlf & "약간의 체크 시간이 소요 됩니다." & _
                         vbcrlf & "점검 완료시 점검 완료 팝업이 보여집니다. 이후 점검 파일을 확인해주세요."

            Set updateSession = CreateObject("Microsoft.Update.Session")
            Set updateSearcher = updateSession.CreateupdateSearcher()

        end If

        Set SearchResult = UpdateSearcher.Search("IsInstalled=0 and Type='software'")
        Set Updates = SearchResult.Updates

        if searchResult.Updates.Count <= 0 then
            RESULT = "양호" & vbcrlf & "- 내용: " & VAL_CHECKMETHOD
            wscript.quit

        else
            for i = 0 to searchResult.Updates.Count-1
                Set update = searchResult.Updates.Item(i)
                i = i+1
                MISSING_UPDATE = MISSING_UPDATE & vbcrlf & update.Title
            next

        end If
        
        RESULT = "취약" & vbcrlf & "- 내용: " & VAL_CHECKMETHOD & vbcrlf & _
                 "* 미설치 업데이트 개수 : [ " & i & " ] 개" & vbcrlf & _
                 "* 미설치 업데이트 내역 : " & MISSING_UPDATE
        
        Set objWSHShell = nothing
        call LOG_FILE(TITLE,RESULT)

        wscript.echo "보안점검이 완료 되었습니다. 점검 로그 파일을 확인 해주세요!"

    end Function

end If
