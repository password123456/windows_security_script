
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

   '----------------------------------------------------
   ' Registry 가져오기
   '----------------------------------------------------
    
    Function GET_REGISTRY_VALUE(search_reg_key)

        dim objWSHShell
        dim value

        On Error Resume Next

        Set objWSHShell = CreateObject("WScript.Shell")
        Err.clear
        value = objWSHShell.RegRead(search_reg_key)

        If hex(Err.number) = "80070002" or hex(Err.number) = "80070005" Then
            value = "7749"
            GET_REGISTRY_VALUE = value
        else
            GET_REGISTRY_VALUE = value
        end if
       
        set objWSHShell = nothing

    End Function

    '----------------------------------------------------
    ' Registry 가져오기2
    '----------------------------------------------------

    Function GET_REGISTRY_ENUMVALUE(hive, key)

        dim reg, arrSubKeys, subkey

        Set reg = GetObject("winmgmts://./root/default:StdRegProv")
        'WScript.Echo key
        reg.Enumvalues hive, key, arrSubKeys
        If Not IsNull(arrSubKeys) Then
            For Each subkey In arrSubKeys
                GET_REGISTRY_ENUMVALUE hive, key & "\" & subkey
                GET_REGISTRY_ENUMVALUE = subkey
            Next
        else
            GET_REGISTRY_ENUMVALUE = "null"
        End If
    End function

    '-------------------------------------------
    ' SNMP 서비스 확인
    '-------------------------------------------

    dim CHECK_SNMP_SERVICE
    CHECK_SNMP_SERVICE = GET_SNMP_SERVICE_STATUS()

    Function GET_SNMP_SERVICE_STATUS()

        dim TITLE
        dim strKeyPath
        dim KEY
        dim strComputer
        dim objWMIService,colItems,objitem
        dim strSNMP, strWEAKNAME
        dim intWEAKNAME
        dim VAL_STATUS, VAL_COMMUNITY_CONFIG, VAL_COMMUNITYNAME
        dim RESULT_COMMUNITY_CONFIG
        dim RESULT

        Const HKEY_LOCAL_MACHINE = &H80000002

        TITLE = "3.4 SNMP 서비스 커뮤니티 스트링의 복잡성 설정"

        strComputer = "." 
        strSNMP = "snmp"
        strWEAKNAME = "public"

        strKeyPath = "SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities"

        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
        Set colItems = objWMIService.ExecQuery( _
                       "SELECT * FROM Win32_Service WHERE Name = 'SNMP'",,48) 

        For Each objItem in colItems 
            if StrComp(lcase(strSNMP), lcase(objItem.Name), 1) = 0 Then
                VAL_STATUS = "0"
            else
                VAL_STATUS = "1"
            end if
        Next

        if VAL_STATUS = "0" then
            
            VAL_COMMUNITYNAME = GET_REGISTRY_ENUMVALUE(HKEY_LOCAL_MACHINE, strKeyPath)

            if StrComp(lcase(strWEAKNAME), lcase(VAL_COMMUNITYNAME), 1) = "0" Then
                intWEAKNAME = "0"
            elseif StrComp(lcase("null"), lcase(VAL_COMMUNITYNAME), 1) = "0" Then
                intWEAKNAME = "0"
            else
                intWEAKNAME = "1" 
            end if

            if intWEAKNAME = "1" then
                RESULT = "양호" & vbcrlf & "- 내용: " & vbcrlf & _
                         "* SNMP 설정이름 : 양호 / " & VAL_COMMUNITYNAME 
            else
                RESULT = "취약" & vbcrlf & "- 내용: " & vbcrlf & _
                         "* SNMP 설정이름 : 취약 / " & VAL_COMMUNITYNAME
            end if

            KEY = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\" & _
                  "Parameters\ValidCommunities\" & VAL_COMMUNITYNAME

            VAL_COMMUNITY_CONFIG = GET_REGISTRY_VALUE(KEY)

            select case VAL_COMMUNITY_CONFIG
                Case "0"
                    RESULT_COMMUNITY_CONFIG = "취약 / 권한:없음(아무설정안됨)"
                Case "1"
                    RESULT_COMMUNITY_CONFIG = "취약 / 권한:알림"
                Case "4"
                    RESULT_COMMUNITY_CONFIG = "양호 / 권한:읽기전용"
                Case "8"
                    RESULT_COMMUNITY_CONFIG = "취약 / 권한:읽기/쓰기"
                Case "16"
                    RESULT_COMMUNITY_CONFIG = "취약 / 권한:읽기/만들기"
                Case "7749"
                    RESULT_COMMUNITY_CONFIG = "오류 / " & VAL_COMMUNITYNAME & " 키를 찾을수 없습니다."
            end select 
            
            RESULT = RESULT & vbcrlf & "* SNMP 권한 : " & RESULT_COMMUNITY_CONFIG
        else
            RESULT = "양호" & vbcrlf & "- 내용: [ " & strSNMP & " ] 서비스 설치되지 않음"

        end if
        call LOG_FILE(TITLE,RESULT)
    end function
end if
