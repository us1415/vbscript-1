
' LS_AQNGlobal.vbs
' Global Logon Script (VBScript)
' This logon script is executed by everybody (domain)
' -----------------------------------------------------------------' 
Option Explicit

Dim wShell, run_this
Dim objGroupList, objUser, strGroup
Dim strNetBIOSDomain
Dim objNetwork, strUserName 
Dim strServer_MAD

strServer_MAD = "\\DC-SERVER\" 
strUserName = "" 

' -----------------------------------------------------------------' 
' Get the current user information.

' Create a network object
Set objNetwork = WScript.CreateObject("WScript.Network") 
' Current Username
strUserName = objNetwork.UserName
' NetBIOS Domain name
strNetBIOSDomain = "LOCAL.DOMAIN"
' Bind to the user object in Active Directory with the WinNT provider.
'Set objUser = GetObject("WinNT://" & strNetBIOSDomain & "/" & strUserName & ",user")

' -----------------------------------------------------------------' 
' Login Script
	Set wShell = CreateObject("WScript.Shell")
	run_this = "powershell.exe -Nologo -ExecutionPolicy Bypass -File " + strServer_MAD + "sysvol\emea.local\scripts\LogonScript.ps1"
	wShell.Run run_this, 0, TRUE

' Quit
WScript.Quit

' -----------------------------------------------------------------' 
' -----------------------------------------------------------------' 
' Global Functions

Function IsMember(ByVal strGroup)
    ' Function to test for user group membership.
    ' strGroup is the AD name (sAMAccountName) of the group to test.
    ' objGroupList is a dictionary object, with global scope.
    ' Returns True if the user is a member of the group.

    If (IsEmpty(objGroupList) = True) Then
        Call LoadGroups
    End If
    IsMember = objGroupList.Exists(strGroup)
End Function

Sub LoadGroups
    ' Subroutine to populate dictionary object with group memberships.
    ' objUser is the user object, with global scope.
    ' objGroupList is a dictionary object, with global scope.

    Dim objGroup
    Set objGroupList = CreateObject("Scripting.Dictionary")
    objGroupList.CompareMode = vbTextCompare
    For Each objGroup In objUser.Groups
        objGroupList.Add objGroup.name, True
    Next
    Set objGroup = Nothing
End Sub


