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
' 로컬보안정책 설정확인
'-------------------------------------------

dim password_complex
dim password_history
dim max_passwordage
dim min_password_length
dim min_passwordage
dim cleartext_password
dim account_lockout_policy
dim adminaccount_status
dim guestaccount_status
dim new_adminaccount_name
dim new_guestaccount_name
dim local_audit_policy

password_complex = GET_PASSWORDCOMPLEXITY()
password_history = GET_PASSWORD_HISTORY()
max_passwordage = GET_MAX_PASSWORD_AGE()
min_password_length = GET_MIN_PASSWORD_LENGTH()
min_passwordage = GET_MIN_PASSWORD_AGE()
cleartext_password = GET_CLEAR_TEXT_PASSWORD()
account_lockout_policy = GET_ACCOUNT_LOCKOUT_POLICY()
adminaccount_status = GET_ADMINACCOUT_STATUS()
guestaccount_status= GET_GUESTACCOUT_STATUS()
new_adminaccount_name = GET_NEWADMINACCOUNT_NAME()
new_guestaccount_name = GET_NEWGUESTACCOUNT_NAME()
local_audit_policy = GET_LOCAL_AUDIT_POLICY()

' 점검 결과 '
wscript.echo GET_PASSWORDCOMPLEXITY & vbcrlf & vbcrlf & _
             GET_PASSWORD_HISTORY & vbcrlf & vbcrlf & _
             GET_MAX_PASSWORD_AGE & vbcrlf & vbcrlf & _
             GET_MIN_PASSWORD_LENGTH & vbcrlf & vbcrlf & _
             GET_CLEAR_TEXT_PASSWORD & vbcrlf & vbcrlf & _
             GET_ACCOUNT_LOCKOUT_POLICY & vbcrlf & vbcrlf & _ 
             GET_ADMINACCOUT_STATUS & vbcrlf & vbcrlf & _
             GET_GUESTACCOUT_STATUS & vbcrlf & vbcrlf & _
             GET_NEWADMINACCOUNT_NAME & vbcrlf & vbcrlf & _
             GET_NEWGUESTACCOUNT_NAME & vbcrlf & vbcrlf & _
             GET_LOCAL_AUDIT_POLICY

' 암호는 복잡성을 만족해야 함
Function GET_PASSWORDCOMPLEXITY()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim PASSWORD_COMPLEX_USE

    search_section = "System Access"
    search_key = "PasswordComplexity"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "암호는 복잡성을 만족해야함 " & vbcrlf
    
    if value < 1 then
        PASSWORD_COMPLEX_USE = CHECK_ITEM & "- 결과: 취약 [ 사용안함 ]"
    else
        PASSWORD_COMPLEX_USE = CHECK_ITEM & "- 결과: 양호 [ 사용 ] "
    end if

    'wscript.echo PASSWORD_COMPLEX_USE

    GET_PASSWORDCOMPLEXITY = PASSWORD_COMPLEX_USE
End Function

' 최근 암호 기억 
Function GET_PASSWORD_HISTORY()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim PASSWORD_HISTORY_USE

    search_section = "System Access"
    search_key = "PasswordHistorySize"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "최근 암호 기억 " & vbcrlf
    
    if value >= 6 then
        PASSWORD_HISTORY_USE = CHECK_ITEM & "- 결과: 양호 [ " & value & " ] 개 암호 기억됨"
    else
        PASSWORD_HISTORY_USE = CHECK_ITEM & "- 결과: 취약 [ " & value & " ] 개 암호 기억됨"
    end if

    'wscript.echo PASSWORD_HISTORY_USE

    GET_PASSWORD_HISTORY = PASSWORD_HISTORY_USE
End Function

' 최대 암호 사용 기간  
Function GET_MAX_PASSWORD_AGE()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim MAX_PASSWORD_AGE_USE

    search_section = "System Access"
    search_key = "MaximumPasswordAge"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "최대 암호 사용기간 " & vbcrlf
    
    if value >= 90 then
        MAX_PASSWORD_AGE_USE = CHECK_ITEM & "- 결과: 양호 [ " & value & " ] 일"
    else
        MAX_PASSWORD_AGE_USE = CHECK_ITEM & "- 결과: 취약 [ " & value & " ] 일"
    end if

    'wscript.echo MAX_PASSWORD_AGE_USE

    GET_MAX_PASSWORD_AGE = MAX_PASSWORD_AGE_USE
End Function

' 최소 암호 길이 
Function GET_MIN_PASSWORD_LENGTH()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim MIN_PASSWORD_LENGTH_USE

    search_section = "System Access"
    search_key = "MinimumPasswordLength"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "최소 암호 길이" & vbcrlf
    
    if value >= 8 then
        MIN_PASSWORD_LENGTH_USE = CHECK_ITEM & "- 결과: 양호 [ " & value & " ] 자리"
    else
        MIN_PASSWORD_LENGTH_USE = CHECK_ITEM & "- 결과: 취약 [ " & value & " ] 자리"
    end if

    'wscript.echo MIN_PASSWORD_LENGTH_USE

    GET_MIN_PASSWORD_LENGTH = MIN_PASSWORD_LENGTH_USE
End Function


' 최소 암호 사용기간 
Function GET_MIN_PASSWORD_AGE()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim MIN_PASSWORD_AGE_USE

    search_section = "System Access"
    search_key = "MinimumPasswordAge"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "최소 암호 사용기간 " & vbcrlf

    if value >= 1 then
        MIN_PASSWORD_AGE_USE = CHECK_ITEM & "- 결과: 양호 [ " & value & " ] 일"
    else
        MIN_PASSWORD_AGE_USE = CHECK_ITEM & "- 결과: 취약 [ " & value & " ] 일"
    end if

    'wscript.echo MIN_PASSWORD_AGE_USE

    GET_MIN_PASSWORD_AGE = MIN_PASSWORD_AGE_USE
End Function

' 해독 가능한 암호화를 사용하여 암호 저장 
Function GET_CLEAR_TEXT_PASSWORD()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim CLEAR_TEXT_PASSWORD_USE

    search_section = "System Access"
    search_key = "ClearTextPassword"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "해독 가능한 암호화를 사용하여 암호 저장 " & vbcrlf

    if value = 0 then
        CLEAR_TEXT_PASSWORD_USE = CHECK_ITEM & "- 결과: 양호 [ 사용안함 ]"
    else
        CLEAR_TEXT_PASSWORD_USE = CHECK_ITEM & "- 결과: 취약 [ 사용 ]"
    end if

    'wscript.echo CLEAR_TEXT_PASSWORD_USE

    GET_CLEAR_TEXT_PASSWORD = CLEAR_TEXT_PASSWORD_USE
End Function

' 계정 잠금 정책 확인
Function GET_ACCOUNT_LOCKOUT_POLICY()

    dim search_section, search_key, search_key_1, search_key_2
    dim RESULT, RESULT_1, RESULT_2 
    dim SPLIT_RESULT, SPLIT_RESULT_1, SPLIT_RESULT_2

    dim keyname, value, value1, value2
    dim CHECK_ITEM
    dim ACCOUNT_LOCKOUT_USE

    search_section = "System Access"
    search_key = "LockoutBadCount"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "계정 잠금 정책 " & vbcrlf

    if value = 0 then
        ACCOUNT_LOCKOUT_USE = CHECK_ITEM & "- 결과: 취약 [ 계정잠금 정책 사용안함 ]"
    else
        search_key_1 = "LockoutDuration"
        search_key_2 = "ResetLockoutCount"
     
        RESULT_1 = GET_LOCAL_POLICY(search_section, search_key_1)
        RESULT_2 = GET_LOCAL_POLICY(search_section, search_key_2)

        SPLIT_RESULT_1 = Split(RESULT_1,",")
        SPLIT_RESULT_2 = Split(RESULT_2,",")

        value1 = TRIM(SPLIT_RESULT_1(1))
        value2 = TRIM(SPLIT_RESULT_2(1))

        ACCOUNT_LOCKOUT_USE = CHECK_ITEM & "- 결과: 양호" & vbcrlf & _ 
                              "--> 계정 잠금 회수 [ " & value & " ] 번" & vbcrlf & _
                              "--> 계정 잠금 기간 [ " & value1 & " ] 분" & vbcrlf & _
                              "--> 계정 잠금 해제 [ " & value2 & " ] 분" 
    end if

    'wscript.echo ACCOUNT_LOCKOUT_USE
    
    GET_ACCOUNT_LOCKOUT_POLICY = ACCOUNT_LOCKOUT_USE
End Function

' 계정: Administrator 계정 상태
Function GET_ADMINACCOUT_STATUS()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim ADMINACCOUNT_USE

    search_section = "System Access"
    search_key = "EnableAdminAccount"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "계정: Administrator 계정 상태" & vbcrlf

    if value = 0 then
        ADMINACCOUNT_USE = CHECK_ITEM & "- 결과: 양호 [ 사용안함 ]"
    else
        ADMINACCOUNT_USE = CHECK_ITEM & "- 결과: 취약 [ 사용 ]"
    end if

    'wscript.echo ADMINACCOUNT_USE

    GET_ADMINACCOUT_STATUS = ADMINACCOUNT_USE
End Function


' 계정: Guest 계정 상태
Function GET_GUESTACCOUT_STATUS()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim GUESTACCOUNT_USE

    search_section = "System Access"
    search_key = "EnableGuestAccount"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    CHECK_ITEM = "계정: Guest 계정 상태" & vbcrlf

    if value = 0 then
        GUESTACCOUNT_USE = CHECK_ITEM & "- 결과: 양호 [ 사용안함 ]"
    else
        GUESTACCOUNT_USE = CHECK_ITEM & "- 결과: 취약 [ 사용 ]"
    end if

    'wscript.echo GUESTACCOUNT_USE

    GET_GUESTACCOUT_STATUS = GUESTACCOUNT_USE
End Function

' 계정: Administrator 계정 이름 바꾸기
Function GET_NEWADMINACCOUNT_NAME()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim COMPARE_RESULT
    dim DEFAULT_NAME
    dim NEW_ADMINACCOUNT_NAME

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

    CHECK_ITEM = "계정: Administrator 계정 이름 바꾸기" & vbcrlf

    if COMPARE_RESULT = 0 then
        NEW_ADMINACCOUNT_NAME = CHECK_ITEM & "- 결과: 취약" & vbcrlf & "- 계정이름: [ " & value & " ]"
    else
        NEW_ADMINACCOUNT_NAME = CHECK_ITEM & "- 결과: 양호" & vbcrlf & "- 계정이름: [ " & value & " ]"
    end if

    'wscript.echo NEW_ADMINACCOUNT_NAME

    GET_NEWADMINACCOUNT_NAME = NEW_ADMINACCOUNT_NAME
End Function

' 계정: Guest 계정 이름 바꾸기
Function GET_NEWGUESTACCOUNT_NAME()

    dim search_section, search_key
    dim RESULT, SPLIT_RESULT
    dim keyname, value
    dim CHECK_ITEM
    dim COMPARE_RESULT
    dim DEFAULT_NAME
    dim NEW_GUEST_NAME

    search_section = "System Access"
    search_key = "NewGuestName"

    RESULT = GET_LOCAL_POLICY(search_section, search_key)

    SPLIT_RESULT = Split(RESULT,",")

    keyname = TRIM(SPLIT_RESULT(0))
    value = TRIM(SPLIT_RESULT(1))

    DEFAULT_NAME="""" & "Guest" & """"
    DEFAULT_NAME = lcase(DEFAULT_NAME)
    value = lcase(value)

    COMPARE_RESULT = strcomp(value, DEFAULT_NAME, 1)

    CHECK_ITEM = "계정: Guest 계정 이름 바꾸기" & vbcrlf

    if COMPARE_RESULT = 0 then
        NEW_GUEST_NAME = CHECK_ITEM & "- 결과: 취약" & vbcrlf & "- 계정이름: [ " & value & " ]"
    else
        NEW_GUEST_NAME = CHECK_ITEM & "- 결과: 양호" & vbcrlf & "- 계정이름: [ " & value & " ]"
    end if

    'wscript.echo NEW_GUEST_NAME

    GET_NEWGUESTACCOUNT_NAME = NEW_GUEST_NAME
End Function

' 감사정책 체크하기 
Function GET_LOCAL_AUDIT_POLICY()

    dim search_section
    dim search_key, search_key_1, search_key_2, search_key_3, search_key_4, search_key_5, search_key_6, search_key_7, search_key_8
    dim keyname, value 
    dim RESULT, RESULT_1, RESULT_2, RESULT_3, RESULT_4, RESULT_5, RESULT_6, RESULT_7, RESULT_8, RESULT_9 
    dim varRESULT, SPLIT_RESULT
    dim CHECK_ITEM, LOCAL_AUDIT_POLICY

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

    CHECK_ITEM = "감사 정책 설정 확인" & vbcrlf

    For each varRESULT in RESULT
        SPLIT_RESULT = Split(varRESULT,",")

        keyname = TRIM(SPLIT_RESULT(0))
        value = TRIM(SPLIT_RESULT(1))

        select case value
            case 0
                value = "설정 안됨"
            case 1
                value = "성공"
            case 2
                value = "실패"
            case 3
                value = "성공,실패"
        End select

        select case keyname
            case "AuditObjectAccess"
                keyname = "* 개체 엑세스 감사"
            case "AuditAccountManage"
                keyname = "* 계정 관리 감사"
            case "AuditLogonEvents"
                keyname = "* 계정 로그온 이벤트 감사"
            case "AuditPrivilegeUse"
                keyname = "* 권한 사용 감사"
            case "AuditLogonEvents"
                keyname = "* 로그온 이벤트 감사"
            case "AuditSystemEvents"  
                keyname ="* 시스템 이벤트 감사"
            case "AuditProcessTracking"
                keyname ="* 프로세스 추적 감사"
            case "AuditDSAccess"
                keyname ="* 디렉토리 서비스 엑세스 감사"
            case "AuditPolicyChange"
                keyname ="* 정책변경 감사"
        End select        

        LOCAL_AUDIT_POLICY = LOCAL_AUDIT_POLICY & keyname & " : " & value & vbcrlf
        
    Next
        CHECK_ITEM = CHECK_ITEM & vbcrlf & LOCAL_AUDIT_POLICY
        'wscript.echo CHECK_ITEM

    GET_LOCAL_AUDIT_POLICY = CHECK_ITEM
End Function


Function GET_LOCAL_POLICY(search_section, search_key)

    dim objWSHShell, objFSO, env, objDUMP_SECPOLICY_FILE
    dim SCRIPTDIR, CMD, DUMP_SECPOLICY
    dim intEqualPos
    dim readini, strFilePath, strKey, strLeftString, strLine, strSection
    dim myFilePath

    Set objWSHShell = CreateObject("WScript.Shell")
    Set objFSO= CreateObject("Scripting.FileSystemObject")
    Set env = objWSHShell.environment("process")

    SCRIPTDIR = objFSO.GetParentFolderName(WScript.ScriptFullName)
    DUMP_SECPOLICY = SCRIPTDIR & "\" & env("COMPUTERNAME") & "_secedit_dump.txt"

    ' 로컬 보안 정책 파일이 존재하지 않을 경우 덤프 실행 
    If not objFSO.FileExists( DUMP_SECPOLICY ) Then 
        CMD = "%comspec% /c %SYSTEMROOT%\system32\"'cmd header
        objWSHShell.run CMD & "secedit /export /cfg " & """" & DUMP_SECPOLICY & """", 0
        'objWSHShell.run CMD & "secedit /export /cfg " & DUMP_SECPOLICY, 0
        wscript.sleep 1000
    End if

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    readIni     = ""
    'strFilePath = Trim( myFilePath )
    strSection  = Trim( search_section )
    strKey      = Trim( search_key )

    If objFSO.FileExists( DUMP_SECPOLICY ) Then
        Set objDUMP_SECPOLICY_FILE = objFSO.OpenTextFile( DUMP_SECPOLICY, ForReading, False, True )
        Do While objDUMP_SECPOLICY_FILE.AtEndOfStream = False
            strLine = objDUMP_SECPOLICY_FILE.ReadLine

            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objDUMP_SECPOLICY_FILE.ReadLine )

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

                    If objDUMP_SECPOLICY_FILE.AtEndOfStream Then Exit Do
                    strLine = Trim( objDUMP_SECPOLICY_FILE.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objDUMP_SECPOLICY_FILE.Close
    Else
        WScript.Echo DUMP_SECPOLICY & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If

GET_LOCAL_POLICY = strLeftString & "," & readini

End Function 

End if