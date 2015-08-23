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
    ' 불필요한 서비스 확인
    '-------------------------------------------

    dim CHECK_WINDOWS_SERVICE
    CHECK_WINDOWS_SERVICE = GET_UNUSE_WINDDOWS_SERVICEY()

    Function GET_UNUSE_WINDDOWS_SERVICEY()

        dim strComputer,i
        dim objWMIService,colItems,objitem
        dim SERVICE_LIST,RESULT,COUNT,LIST

        strComputer = "." 
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
        Set colItems = objWMIService.ExecQuery( _
            "SELECT * FROM Win32_Service WHERE State = 'Running'",,48) 

        SERVICE_LIST=array("Application Layer Gateway Service","Certificate Propagation","DFSR","Diagnostic Policy Service","Diagnostic Service Host","Diagnostic System Host","Disk Defragmenter","Encrypting File System (EFS)","Extensible Authentication Protocol","IKE and AuthIP IPsec Keying Modules","Interactive Services Detection","Internet Connection Sharing (ICS)","IP Helper","IsmServ","KtmRm for Distributed Transaction Coordinator","Link-Layer Topology Discovery Mapper","Multimedia Class Scheduler","Network Access Protection Agent","PnP-X IP Bus Enumerator","Portable Device Enumerator Service","Print Spooler","Problem Reports and Solutions Control Panel Support","Protected Storage","Remote Access Auto Connection Manager","Remote Access Connection Manager","Routing and Remote Access","Secure Socket Tunneling Protocol Service","Shell Hardware Detection","Smart Card","Smart Card Removal Policy","SNMP Trap","Special Administration Console Helper","SSDP Discovery","Telephony","Thread Ordering Server","TPM Base Services","UPnP Device Host","Windows Audio","Windows Audio Endpoint Builder","Windows CardSpace","Windows Error Reporting Service","WinHTTP Web Proxy Auto-Discovery Service","Wired AutoConfig","Application Identity","Windows Remote Management (WS-Management)","Software Protection")

        COUNT = 0

        for each objItem in colItems 
            for i=Lbound(SERVICE_LIST) to Ubound(SERVICE_LIST)

            ' 0 - for binary comparison "Case sensitive
            ' 1 - for text compare not case sensitive

                if StrComp(lcase(SERVICE_LIST(i)), lcase(objItem.DisplayName), 1) = 0 Then
                    RESULT = TRUE
                        COUNT=COUNT+1
                    LIST= LIST & vbcrlf & objItem.DisplayName & "  | " & objItem.state
                    'wscript.echo "불필요서비스: "  &  LIST
                else
                    RESULT = FALSE
                end if
            next
        next

        wscript.echo "불필요 서비스 개수 : " & COUNT 
        wscript.echo "불필요 서비스 목록 : " &  LIST
        
    end function

end if