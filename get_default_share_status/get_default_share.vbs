'--------------------------------------------------------------------
' Windows security check script
' supported on : windows 7 over
' created by password123456 / 2015.08
' Freeware, but You always need to cite your sources.
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
    ' �⺻ ���� Ȯ��
    '-------------------------------------------

    dim CHECK_DEFAULT_SHARE
    CHECK_DEFAULT_SHARE = GET_DEFAULT_SHARE()

    Function GET_DEFAULT_SHARE()

        dim strComputer,i
        dim objWMIService,colItems,objitem
        dim SHARE_LIST,RESULT,COUNT,LIST

        strComputer = "." 
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
        Set colItems = objWMIService.ExecQuery( _
             "SELECT * FROM Win32_Share",,48)  

        SHARE_LIST=array("Admin$","IPC$")
        COUNT = 0

        for each objItem in colItems 
            for i=Lbound(SHARE_LIST) to Ubound(SHARE_LIST)

            ' 0 - for binary comparison "Case sensitive
            ' 1 - for text compare not case sensitive

                if StrComp(lcase(SHARE_LIST(i)), lcase(objItem.Name), 1) = 0 Then
                    COUNT = COUNT + 1
                    'RESULT = "���/�⺻���� ��� ��"
                    LIST= LIST & vbcrlf & objItem.Name
                    'wscript.echo "�⺻���� ��� ����: "  &  LIST
                end if
            next
        next
        
        if COUNT <= 0 then
            RESULT = "��ȣ/�⺻���� ������"
            LIST = "����"
        else
            RESULT = "���/�⺻���� ��� ��"
        end if
            
        wscript.echo "�⺻ ���� ��� ���� : " & RESULT & vbcrlf & "�⺻ ���� ��� : " & LIST
        
    end function

end if