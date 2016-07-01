
dim TITLE,KEY
dim KEY_TYPE, KEY_NAME_POS
dim ret_reg_key

TITLE = "save_file_name"

'registry-key to export
KEY = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit\ProcessCreationIncludeCmdLine_Enabled"

'registry-key type
KEY_TYPE = "REG_DWORD"

'position of registry_key
KEY_NAME_POS = 8

'value of registry_key
ret_reg_key = 1


dim test,test1,test2,test3,test4,test5

test = SAVE_REG(TITLE,KEY,KEY_TYPE,KEY_NAME_POS,ret_reg_key)
test1 = SAVE_REG(TITLE,KEY,KEY_TYPE,KEY_NAME_POS,2)
test2 = SAVE_REG(TITLE,KEY,KEY_TYPE,KEY_NAME_POS,3)
test3 = SAVE_REG(TITLE,KEY,"REG_SZ",KEY_NAME_POS,"2016-07-01")
test4 = SAVE_REG(TITLE,KEY,KEY_TYPE,KEY_NAME_POS,7749)
test5 = SAVE_REG(TITLE,KEY,KEY_TYPE,KEY_NAME_POS,398)

'-----------------------------
' create .reg file
'-----------------------------
Function SAVE_REG(title,reg_key,reg_keytype,reg_keyname_pos,reg_keyname_value)
    dim objFSO,objobjLOGFILE,objWSHShell
    dim objSCRIPTDIR, objLOGFILE
    dim checktime
	dim segments
	dim export_keyname, export_key
	dim hx_reg_keyname_value
	dim ret
	dim i,b
	
	checktime= year(now)& "." & right("0" & month(now),2) & "." & right("0" & day(now),2) & _
               "_" & right("0" & hour(now),2) & "." & right("0" & minute(now),2) & "." & right("0" & second(now),2)
	    		
    Set objWSHShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
			
	objSCRIPTDIR = objFSO.GetParentFolderName(WScript.ScriptFullName)
	objLOGFILE = objSCRIPTDIR & "\" & checktime& "_" & TITLE & "_backup.reg"
	   
	ret = StrComp(reg_keyname_value,"7749",1)
	wscript.echo "reg_keyname_value: " & reg_keyname_value & " ret:" & ret
	
	If ret = "-1" then
	    segments = Split(reg_key,"\")
		export_keyname = segments(reg_keyname_pos)
		wscript.echo "export_keyname: " & export_keyname
        b = UBound(segments)
        b = b - 1
            
        For i = 0 to b
            export_key = export_key & segments(i) & "\"
        next
		
	    export_key = Left(export_key, Len(export_key) - 1)
	    wscript.echo "reg_keyname_value: " & reg_keyname_value & " ret:" & ret
	    
	  	If reg_keytype = "REG_DWORD" then
		      hx_reg_keyname_value = LCase(hex(reg_keyname_value))
	        reg_keyname_value = "dword:" & String(8 - Len(hx_reg_keyname_value), "0") & hx_reg_keyname_value
		  elseif reg_keytype = "REG_SZ" then
		      reg_keyname_value = """" & reg_keyname_value & """"
	    end if
      
	    if objFSO.fileExists(objLOGFILE) then
	        set objobjLOGFILE = objFSO.OpenTextFile(objLOGFILE,8,true)
	    else
	        set objobjLOGFILE = objFSO.CreateTextFile(objLOGFILE,True)
	    end if
		  
	    objobjLOGFILE.writeline ("Windows Registry Editor Version 5.00")
		objobjLOGFILE.writeline ("")
	    objobjLOGFILE.writeline ("[" & export_key & "]")
	    objobjLOGFILE.writeline ("""" & export_keyname & """" & "=" & reg_keyname_value)
		objobjLOGFILE.writeline ("")

	    Set objFSO = Nothing
	    objobjLOGFILE.close
    end if
end Function
    
    
