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
' ���Ⱦ�����Ʈ ����üũ
' wsusscn2.cab ������ ������ cab ���� ���� üũ
' ������ ���� ���ÿ��� ������ ������Ʈ ���� ���� Ȯ��
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
        wscript.echo "wsusscn2.cab ������ �̿��Ͽ� ���Ⱦ�����Ʈ�� üũ�մϴ�." & _
                     vbCRLF & "üũ �ð��� �ɸ��Ƿ� ��ٸ�����."

        Set UpdateSession = CreateObject("Microsoft.Update.Session")
        Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
        Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", WSUS_DB)
        Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
        UpdateSearcher.ServerSelection = 3
        UpdateSearcher.ServiceID = UpdateService.ServiceID

    else 
        wscript.echo "wsusscn2.cab ������ �������� �ʾ�" & _
                     " ������ ������ ������ ������Ʈ ���¸� üũ�մϴ�." & _ 
                     vbCRLF & "üũ �ð��� �ɸ��Ƿ� ��ٸ�����."

        Set updateSession = CreateObject("Microsoft.Update.Session")
        Set updateSearcher = updateSession.CreateupdateSearcher()

    end If

    Set SearchResult = UpdateSearcher.Search("IsInstalled=0 and Type='software'")
    Set Updates = SearchResult.Updates

    if searchResult.Updates.Count <= 0 then
        RESULT = "��ȣ"
        wscript.echo "���� ������Ʈ ���� : " & RESULT
        wscript.quit

    else
        RESULT = "���"
        for i = 0 to searchResult.Updates.Count-1
            Set update = searchResult.Updates.Item(i)
            i = i+1
            MISSING_UPDATE = MISSING_UPDATE & vbcrlf & update.Title
            'WScript.Echo i + 1 & "> " & update.Title
        next

    wscript.echo "���� ������Ʈ ���� : " & RESULT & vbcrlf & "�̼�ġ ������Ʈ ���� : " & i & vbcrlf & "�̼�ġ ������Ʈ ���� : " & MISSING_UPDATE

    end If

    Set WshShell = nothing

end if