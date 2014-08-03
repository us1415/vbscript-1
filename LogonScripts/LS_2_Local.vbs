
' LS_AQNLocal.vbs
' Local Logon Script (VBScript)
' This logon script executes the local logon script according to the user's office (domain)
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
Set objUser = GetObject("WinNT://" & strNetBIOSDomain & "/" & strUserName & ",user")

' -----------------------------------------------------------------' 
' Run the local logon script according to the user's office 

' - If the user belongs to the Technology office in Spain
If (IsMember("SG_TEC") = True) Then
	' Execute the logon script for Technology
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNTEC.vbs"
	wShell.Run run_this, 0, TRUE
	
' - If the user belongs to the Spanish office
ElseIf (IsMember("SG_SPA") = True) Then
	' Execute the logon script for Madrid
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNSPA.vbs"
	wShell.Run run_this, 0, TRUE	

' - If the user belongs to the Brazil office
ElseIf (IsMember("SG_BRA") = True) Then
	' Execute the logon script for Sao Paulo
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNBRA.vbs"
	wShell.Run run_this, 0, TRUE

' - If the user belongs to the USA office
ElseIf (IsMember("SG_USA") = True) Then
	' Execute the logon script for Miami
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNUSA.vbs"
	wShell.Run run_this, 0, TRUE

' - If the user belongs to the Mexico office
ElseIf (IsMember("SG_MEX") = True) Then
	' Execute the logon script for Ciudad de México
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNMEX.vbs"
	wShell.Run run_this, 0, TRUE

' - If the user belongs to the Chile office
ElseIf (IsMember("SG_CHI") = True) Then
	' Execute the logon script for Santiago
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNCHI.vbs"
	wShell.Run run_this, 0, TRUE

' - If the user belongs to the Argentina office
ElseIf (IsMember("SG_ARG") = True) Then
	' Execute the logon script for Buenos Aires
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNARG.vbs"
	wShell.Run run_this, 0, TRUE

' - If the user belongs to the Argentina office
ElseIf (IsMember("SG_URU") = True) Then
	' Execute the logon script for Buenos Aires
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNURU.vbs"
	wShell.Run run_this, 0, TRUE
	
' - If the user belongs to Global Office
ElseIf (IsMember("SG_GLO") = True) Then
	' Execute the logon script for Global
	Set wShell = CreateObject("WScript.Shell")
	run_this = strServer_MAD & "SYSVOL\local.domain\scripts\logonscripts\LS_AQNGLO.vbs"
	wShell.Run run_this, 0, TRUE	
	
' - If the user doesn't belong to any office
Else
	' There has been an error
	WScript.Echo "There has been an error. Please contact IT Support."

End If	

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


