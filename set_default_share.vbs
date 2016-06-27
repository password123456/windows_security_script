	
    Function SET_DEFAULT_SHARE(share_name)
	
        dim objFSO
	dim objWSHShell, strCMD
	dim strSystemRoot, strNET
	dim ret
	
	Set objWSHShell = CreateObject("WScript.Shell")
        Set objFSO= CreateObject("Scripting.FileSystemObject")
            
	strSystemRoot = objWSHShell.ExpandEnvironmentStrings( "%SystemRoot%" )		
        strNET = strSystemRoot & "\system32\net.exe" '
		
        ' net.exe 존재할 경우만 실행
        If objFSO.FileExists( strNET ) Then		
		strCMD = "%comspec% /c %SYSTEMROOT%\system32\"'strCMD header
		ret = objWSHShell.run(strCMD & "net share " & share_name & " /delete" & """", 0,True)
		wscript.echo "Default SHARE: " & share_name & " / ret: "& ret
		ret = ret + ret
			
		if ret <> 0 then
                ret = 1
            else
		wscript.echo "Default SHARE: " & share_name & " delete."
                ret = 0
            end if		
            SET_DEFAULT_SHARE = ret			
		else
		    ret = 2
		    SET_DEFAULT_SHARE = ret
		end if
	end function
	
	dim retSetlog
	
	default_share = "IPC$,Admin$,C$,D$,E$,F$,G$,H$,I$"
	list=Split(default_share,",")
	
    for each share_list in list
	    retSetlog = SET_DEFAULT_SHARE(share_list)
	next
	if retSetlog = 2 then
	    wscript.echo "net.exe " & strNET & "not exits. "
	    
	elseif retSetlog = 0 then
	    wscript.echo "default_share: " & retSetlog & " / ok"

	else
	    wscript.echo "retSetlog: " & retSetlog & " / Fail"
	end if
