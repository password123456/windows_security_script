    Function SET_EVENTLOG_SIZE(eventlog, log_size)
	
        dim objFSO
		    dim objWSHShell, strCMD
		    dim strSystemRoot, strWEVUTIL
		    dim ret
	
	      Set objWSHShell = CreateObject("WScript.Shell")
        Set objFSO= CreateObject("Scripting.FileSystemObject")
	    
		    strSystemRoot = objWSHShell.ExpandEnvironmentStrings( "%SystemRoot%" )		
        strWEVUTIL = strSystemRoot & "\system32\wevtutil.exe" '

        ' if "wevtutil.exe" exits, run
        If objFSO.FileExists( strWEVUTIL ) Then		
			      wscript.echo log_list & " event log size set " & log_size & " bytes."
			      strCMD = "%comspec% /c %SYSTEMROOT%\system32\"'strCMD header
			      ret = objWSHShell.run(strCMD & "wevtutil sl " & log_list & " /ms:" & log_size & """", 0,True)
			      wscript.sleep 1000
			      wscript.echo ret
			
			      ret = ret + ret
			      if ret <> 0 then
                ret = 1
            else
                ret = 0
            end if		
            SET_EVENTLOG_SIZE = ret			
		    else
		        ret = 2
		        SET_EVENTLOG_SIZE = ret
		    end if
	end function
	
	dim retSetlog
	
	set_logsize = 314572800
	eventlog = "application,system,security"
	list=Split(eventlog,",")
	
  for each log_list in list
      If not len(log_list) = 0 Then
          retSetlog = SET_EVENTLOG_SIZE(eventlog,set_logsize)
      end if
	next
	
	if retSetlog = 2 then
	    wscript.echo "WEVUTIL: " & strWEVUTIL & "not exits. "
	    
	elseif retSetlog = 0 then
	    wscript.echo "SET_EVENTLOG_SIZE: " & retSetlog & " / ok"

	else
	    wscript.echo "SET_EVENTLOG_SIZE: " & retSetlog & " / Fail"
	end if
