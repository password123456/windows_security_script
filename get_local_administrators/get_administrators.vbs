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

If WScript.Arguments.length =0 Then
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument with leading blank space, say [ uac ]
    objShell.ShellExecute "wscript.exe", Chr(34) & _
    WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    
Else

'-------------------------------------------
' 어드민 계정 목록 확인
'-------------------------------------------

    Dim admin_users 
    admin_users = GET_ADMINISTRATORS

Function ComputerName
    Dim WshNetwork
    Set WshNetwork = CreateObject("WScript.Network")
    ComputerName = WshNetwork.ComputerName
    Set WshNetwork = Nothing
End Function
 
Function GET_ADMINISTRATORS
    Dim objWMIService
    Dim Item, strObjectPath, Membername, MemberDomain
    Dim strgroup 
    Dim GrpMembers,GetAdminGroupUsers 
    
    GrpMembers = Empty
    strgroup = "Administrators"
    
    Dim sN, lN, sD, lD, COUNT
    Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2") _
                .ExecQuery("select * from Win32_GroupUser where " & "GroupComponent = " & chr(34) & "Win32_Group.Domain='" _
                    & Computername & "',Name='" & strGroup & "'" & Chr(34) )
    For Each Item In objWMIService
    strObjectPath = Item.PartComponent
    sN = inStrRev(strObjectPath, "Name=""",-1,1)
    lN = Len(strObjectPath)-(sN+6)
    sD = inStrRev(strObjectPath, "Domain=""",-1,1)
    lD = (sN-2)-(sD+8)
    Membername = mid(strObjectPath, sN+6, lN)
    MemberDomain = mid(strObjectPath, sD+8,lD)
    GrpMembers =  GrpMembers & "(" _
        & MemberDomain & "\" & Membername & ")" & chr(44) & vbcrlf
    Next
    
    COUNT = 1
    If Len(GrpMembers) = 0 then 
        GrpMembers = "<none>"
    else
        COUNT = COUNT +1
        GetAdminGroupUsers = GrpMembers
        wscript.echo "Administrators 계정개수 : " & COUNT & vbcrlf & "Administrator 계정 : " & vbcrlf & GetAdminGroupUsers
    end if
    Set objWMIService = Nothing
    
End Function

End If