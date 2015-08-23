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
' 보안업데이트 상태체크
' wsusscn2.cab 파일이 있으면 cab 파일 기준 체크
' 없으면 현재 로컬에서 윈도우 업데이트 상태 내역 확인
'-------------------------------------------

    dim wshshell, fso, i
    dim UpdateSession, UpdateServiceManager, UpdateService, UpdateSearcher, SearchResult, Updates, update
    dim STR_CURRENT_DITRECTORY, WSUS_DB, RESULT, MISSING_UPDATE

    Set wshshell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    STR_CURRENT_DITRECTORY = fso.GetParentFolderName(WScript.ScriptFullName)
    WSUS_DB = STR_CURRENT_DITRECTORY & "\wsusscn2.cab"
    wscript.echo wsus_db

    
    if (fso.fileexists(WSUS_DB)) then
        wscript.echo "wsusscn2.cab 파일을 이용하여 보안업데이트를 체크합니다." & _
                     vbCRLF & "체크 시간이 걸리므로 기다리세요."

        Set UpdateSession = CreateObject("Microsoft.Update.Session")
        Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
        Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", WSUS_DB)
        Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
        UpdateSearcher.ServerSelection = 3
        UpdateSearcher.ServiceID = UpdateService.ServiceID

    else 
        wscript.echo "wsusscn2.cab 파일이 존재하지 않아" & _
                     " 로컬의 마지막 윈도우 업데이트 상태를 체크합니다." & _ 
                     vbCRLF & "체크 시간이 걸리므로 기다리세요."

        Set updateSession = CreateObject("Microsoft.Update.Session")
        Set updateSearcher = updateSession.CreateupdateSearcher()

    end If

    Set SearchResult = UpdateSearcher.Search("IsInstalled=0 and Type='software'")
    Set Updates = SearchResult.Updates

    if searchResult.Updates.Count <= 0 then
        RESULT = "양호"
        wscript.echo "보안 업데이트 상태 : " & RESULT
        wscript.quit

    else
        RESULT = "취약"
        for i = 0 to searchResult.Updates.Count-1
            Set update = searchResult.Updates.Item(i)
            i = i+1
            MISSING_UPDATE = MISSING_UPDATE & vbcrlf & update.Title
            'WScript.Echo i + 1 & "> " & update.Title
        next

    wscript.echo "보안 업데이트 상태 : " & RESULT & vbcrlf & "미설치 업데이트 개수 : " & i & vbcrlf & "미설치 업데이트 내역 : " & MISSING_UPDATE

    end If

    Set WshShell = nothing

end if