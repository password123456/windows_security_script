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
    
'-------------------------------------------
' 보안업데이트 상태체크
' wsusscn2.cab 파일이 있으면 cab 파일 기준 체크
' 없으면 현재 로컬에서 윈도우 업데이트 상태 내역 확인
'-------------------------------------------

    dim KEY1
    dim objWSHShell, objFSO, i
    dim UpdateSession, UpdateServiceManager, UpdateService, UpdateSearcher, SearchResult, Updates, update
    dim STR_CURRENT_DITRECTORY, WSUS_DB, VAL_WINUPDATE_STATUS, VAL_MISSING_UPDATE, VAL_CHECKMETHOD
    dim VAL_UPDATE_OPTION, RESULT
    dim intNumSecurityUpdates
    Dim intNumNonSecurityUpdates

    intNumSecurityUpdates = "0"
    intNumNonSecurityUpdates = "0"

    KEY1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\AUOptions"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    STR_CURRENT_DITRECTORY = objFSO.GetParentFolderName(WScript.ScriptFullName)
    WSUS_DB = STR_CURRENT_DITRECTORY & "\wsusscn2.cab"
    
    VAL_WINUPDATE_STATUS = GET_REGISTRY_VALUE(KEY1)

    if VAL_WINUPDATE_STATUS = "1" then
        RESULT = "취약" & vbcrlf & "- 내용: " & vbcrlf & _ 
                 "윈도우 업데이트 사용 안함." & vbcrlf &_ 
                 "윈도우 업데이트를 활성화 후 다시 점검 해주세요."
    else
        select case VAL_WINUPDATE_STATUS 
            Case "2"
                VAL_UPDATE_OPTION = "업데이트를 확인하지만 다운로드 및 설치 여부는 직접 선택"
            Case "3"
                VAL_UPDATE_OPTION = "업데이트를 다운로드하지만 설치 여부는 직접 선택"
            Case "4"
                VAL_UPDATE_OPTION = "업데이트 자동 설치"
            Case "7749"
                VAL_UPDATE_OPTION = "오류 / AUOptions 키가 존재하지 않음"
        end select

        if (objFSO.fileexists(WSUS_DB)) then
            VAL_CHECKMETHOD = "wsusscn2.cab 파일로 보안업데이트 상태 확인함"

            wscript.echo "wsusscn2.cab 파일을 이용하여 보안업데이트를 체크합니다." & _
                         vbcrlf & "2~3분의 체크 시간이 소요 됩니다." & _ 
                         vbcrlf & "점검 완료시 점검 완료 팝업이 보여집니다. 점검 완료 팝업 확인 후 점검 파일을 확인 해주세요."

            Set UpdateSession = CreateObject("Microsoft.Update.Session")
            UpdateSession.ClientApplicationID = "Scripted Update Check"

            Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
            Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", WSUS_DB)
            Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
            UpdateSearcher.ServerSelection = 3
            UpdateSearcher.ServiceID = UpdateService.ServiceID

        else 
            VAL_CHECKMETHOD = "윈도우업데이트 서비스를 실행하여 업데이트 상태 확인함."

            wscript.echo "wsusscn2.cab 파일이 존재하지 않아 인터넷을 연결을 통한" & _
                         "윈도우 업데이트 상태를 체크합니다." & vbcrlf & vbcrlf & _ 
                         "인터넷 연결상태 / WSUS 연결상태에 따라 몇 분의 점검 시간이 소요 됩니다." & vbcrlf & vbcrlf & _
                         "점검 완료시 점검 완료 팝업이 보여집니다." & vbcrlf & _
                         "점검 완료 팝업 확인 후 점검 파일을 확인 해주세요."

            Set updateSession = CreateObject("Microsoft.Update.Session")
            UpdateSession.ClientApplicationID = "Scripted Update Check"
            Set updateSearcher = updateSession.CreateupdateSearcher()

        end If

        On Error Resume Next
        Err.clear

        Set SearchResult = UpdateSearcher.Search("IsInstalled=0 and Type='software' and IsHidden=0")

        If err.number <> 0 Then
            RESULT = "오류" & vbcrlf & "- 내용: " & vbcrlf & _
                     "* 업데이트 설치 옵션 : " & VAL_UPDATE_OPTION & vbcrlf & _ 
                     "* 업데이트 체크 결과 : " & vbcrlf & _
                     " - 윈도우 업데이트 확인을 실패하였습니다." & vbcrlf & _
                     " - 시스템이 인터넷이 가능한지, WSUS와 정상 연결 중인지 확인하세요." & vbcrlf & _
                     " - 오류 번호: [ " & hex(err.number) & " ]"
            Err.clear
        else

            if searchResult.Updates.Count <= 0 then
                RESULT = "양호" & vbcrlf & "- 내용: 최신업데이트 상태입니다." 
            else
                for i = 0 to searchResult.Updates.Count-1
                    Set update = searchResult.Updates.Item(i)
                    If update.MsrcSeverity = "Critical" Or update.MsrcSeverity = "Important" Or _
                       update.MsrcSeverity = "Moderate" Or update.MsrcSeverity = "Low" then
                        intNumSecurityUpdates = intNumSecurityUpdates + 1
                        VAL_MISSING_UPDATE = VAL_MISSING_UPDATE & vbcrlf & " [ " & update.MsrcSeverity & " ] " & update.Title    
                    else 
                        intNumNonSecurityUpdates = intNumNonSecurityUpdates + 1
                        VAL_MISSING_UPDATE = VAL_MISSING_UPDATE & vbcrlf & " [ 일반 ] " & update.Title                           
                        
                    end if
                next
                i = intNumSecurityUpdates + intNumNonSecurityUpdates
                RESULT = "취약" & vbcrlf & "- 체크: " & VAL_CHECKMETHOD & vbcrlf & _
                         "* 업데이트 설치 옵션: " & VAL_UPDATE_OPTION & vbcrlf & _
                         "* 미설치 업데이트 개수: [ " & i & " ] 개" & vbcrlf & _
                         "  - 중요업데이트: [ " & intNumSecurityUpdates & " ] " & vbcrlf & _
                         "  - 일반업데이트: [ " & intNumNonSecurityUpdates & " ] " & vbcrlf & _
                         "* 미설치 업데이트 내역: " & VAL_MISSING_UPDATE
            end if
        end if
    end if
    
    wscript.echo RESULT

end if
