' LS_USA.vbs
' USA Logon Script (VBScript)
' This logon script is executed only by the users that belong to the USA Office
' -----------------------------------------------------------------' 
Option Explicit

' -----------------------------------------------------------------' 
' Variable Declaration

Dim strRemotePath, strUserName 
strRemotePath = "" 
strUserName = "" 

' Drive letters for mapping
Dim strDriveLetter_RRHH, strDriveLetter_Informes, strDriveLetter_MIS, strDriveLetter_Public, strDriveLetter_BOS
Dim strDriveLetter_Home, strDriveLetter_Acct, strDriveLetter_Apps, strDriveLetter_Procurement, strDriveLetter_Quicken, strDriveLetter_Reports
Dim OsType, strEMLPath32, strEMLPath64
strDriveLetter_Apps = "F:"
strDriveLetter_Home = "H:" 
strDriveLetter_Informes = "I:"
strDriveLetter_MIS = "M:"
strDriveLetter_Public = "P:" 
strDriveLetter_Quicken = "Z:"
strDriveLetter_RRHH = "R:"
strDriveLetter_BOS = "S:"
strDriveLetter_Acct = "T:"
strDriveLetter_Reports = "U:"
strDriveLetter_Procurement = "W:"
 
' Company Servers
Dim strServer_USA, strServer_MAD
strServer_USA = "\\USA-DC\" 
strServer_MAD = "\\MAD-DC\" 

' VB Objects
Dim objShell, objFSO, objNetwork
Set objNetwork = WScript.CreateObject("WScript.Network") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.shell")

' Remote and local Paths
Dim strAppData, strSysRoot, strRobocopyPath, strWallpaper, strWallpaperPath, strLocalScripts, strCopy
Dim strOCSPath, strOCS
strAppData = objShell.ExpandEnvironmentStrings("%APPDATA%")
strSysRoot = objShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
strRobocopyPath = strServer_USA & "SYSVOL\local.domain\scripts\bin\robocopy.exe"
strOCSPath = strServer_USA & "SYSVOL\local.domain\scripts\bin\OcsLogon.exe"
strWallpaper = strServer_USA & "SYSVOL\local.domain\scripts\data\Company.jpg"
strWallpaperPath = strSysRoot & "\Company\"
strLocalScripts = "C:\Scripts\"
strCopy = "robocopy.exe"
strOCS = "OcsLogon.exe"

strEMLPath32 = strServer_USA & "SYSVOL\local.domain\scripts\bin\eml-Outlook2007-Win32.reg"
strEMLPath64 = strServer_USA & "SYSVOL\local.domain\scripts\bin\eml-Outlook2007-Win64.reg"

'Read Registry for EML
OsType = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")

' -----------------------------------------------------------------' 
' Copy Wallpaper image file
Dim objFolder, objFileCopy

' Create/Get Local Scripts Path
If objFSO.FolderExists(strLocalScripts) Then
   Set objFolder = objFSO.GetFolder(strLocalScripts)
Else
   Set objFolder = objFSO.CreateFolder(strLocalScripts)
End If

' Create/Get Windows Company Directory
If objFSO.FolderExists(strWallpaperPath) Then
   Set objFolder = objFSO.GetFolder(strWallpaperPath)
Else
   Set objFolder = objFSO.CreateFolder(strWallpaperPath)
End If

' Copy Robocopy to C:\Scripts Directory
Set objFileCopy = objFSO.GetFile(strRobocopyPath)
objFileCopy.Copy (strLocalScripts)

' Copy OCSLogon to C:\Scripts Directory
Set objFileCopy = objFSO.GetFile(strOCSPath)
objFileCopy.Copy (strLocalScripts)

' Copy Wallpaper to %SYSTEMROOT%\Company Directory
Set objFileCopy = objFSO.GetFile(strWallpaper)
objFileCopy.Copy (strWallpaperPath)

'------------------------------------------------------------------'
' Copy and Activate BGINFO and TeamViewer

Dim strBGInfoPath, strBGInfoLocalPath, strBGInfoBGI, strBG, strTVID, strTV, strDesktop, srtSysRoot, objExec

srtSysRoot = objShell.ExpandEnvironmentStrings("%SYSTEMROOT%")

strBGInfoPath = strServer_USA & "SYSVOL\local.domain\scripts\bin\Bginfo.exe"
strBGInfoLocalPath = srtSysRoot & "\SYSTEM32\"
strBGInfoBGI = strServer_USA & "SYSVOL\local.domain\scripts\data\corporate_USA.bgi"
strBG = "bginfo.exe"
strTVID = strServer_USA & "SYSVOL\local.domain\scripts\data\GetTeamviewerIDs.vbs"
strTV = strServer_USA & "SYSVOL\local.domain\scripts\bin\TeamViewerQS.exe"
strDesktop = objShell.SpecialFolders("Desktop")

' Check Desktop Directory
Set objFolder = objFSO.GetFolder(strDesktop)

' Copy BGInfo to %SYSTEMROOT%\System32 Directory
Set objFileCopy = objFSO.GetFile(strBGInfoPath)
objFileCopy.Copy (strBGInfoLocalPath)

' Copy .bgi file for BGInfo to Scripts Directory
Set objFileCopy = objFSO.GetFile(strBGInfoBGI)
objFileCopy.Copy (strLocalScripts)

' Copy .reg file for EML to Scripts Directory
Set objFileCopy = objFSO.GetFile(strEMLPath32)
objFileCopy.Copy (strLocalScripts)

Set objFileCopy = objFSO.GetFile(strEMLPath64)
objFileCopy.Copy (strLocalScripts)

' Copy .vbs file to get TeamViewer ID
'Set objFileCopy = objFSO.GetFile(strTVID)
'objFileCopy.Copy (strLocalScripts)

' Copy TeamViewerQS to Desktop
'Set objFileCopy = objFSO.GetFile(strTV)
'objFileCopy.Copy (strDesktop) & "\"

' Run BGInfo
Set objExec = objShell.Exec (strBGInfoLocalPath & strBG & " " & strLocalScripts & "corporate_USA.bgi" & " " & "/TIMER:0 /SILENT /NOLICPROMPT /LOG:c:\Scripts\BGInfoUSA.log")

' Run OCSLogin
Set objExec = objShell.Exec (strLocalScripts & strOCS & " /S /SERVER=http://inventory.local.domain/ocsinventory /PACKAGER /DEPLOY=2.0.5.0")

' Run EML Registry
If OsType = "x86" then
    Set objExec = objShell.Exec (srtSysRoot & "\" & "regedit.exe /s " & strLocalScripts & "eml-Outlook2007-Win32.reg")
else
    Set objExec = objShell.Exec (srtSysRoot & "\" & "regedit.exe /s " & strLocalScripts & "eml-Outlook2007-Win64.reg")
end if

' -----------------------------------------------------------------' 
' Map Generic Network drives

' Map User's network drive
strUserName = objNetwork.UserName
strRemotePath  = strServer_USA & strUserName & "$"
Call MapNetworkDrive(strDriveLetter_Home, strRemotePath)

' Map Public network drive
'strRemotePath  = strServer_USA & "Public_Local"
'Call MapNetworkDrive(strDriveLetter_Public, strRemotePath)

' -----------------------------------------------------------------' 
' Map Network drive from group - Starts Here

Dim objGroupList, objUser, strGroup
Dim strNetBIOSDomain

' Current Username
strUserName = objNetwork.UserName

' NetBIOS Domain name
strNetBIOSDomain = "local.domain"

' Bind to the user object in Active Directory with the WinNT provider.
Set objUser = GetObject("WinNT://" & strNetBIOSDomain & "/" & strUserName & ",user")

' Map Accounting (ACCT) network drive - If the user is a member of the group 
strGroup = "ACCT_Write"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_USA & "ACCT"
	Call MapNetworkDrive(strDriveLetter_Acct, strRemotePath)
	
End If

' Map Procurement_Unit network drive - If the user is a member of the group 
strGroup = "USA_PU_Write"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_USA & "PU_Info"
	Call MapNetworkDrive(strDriveLetter_Procurement, strRemotePath)
	
End If

' Map BOS network drive - If the user is a member of the group 
strGroup = "USA_BOS_Read"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_USA & "BOS_Info"
	Call MapNetworkDrive(strDriveLetter_BOS, strRemotePath)
	
End If

' Map Reports network drive - If the user is a member of the group 
strGroup = "USA_Reports_Write"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_USA & "USA_Reports"
	Call MapNetworkDrive(strDriveLetter_Reports, strRemotePath)
	
End If

' Map Accounting Quicken(ACCT\Quicken) network drive - If the user is a member of the group 
strGroup = "ACCT_Quicken"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_USA & "Quicken"
	Call MapNetworkDrive(strDriveLetter_Quicken, strRemotePath)
	
End If

' Map Applications (APPS) network drive - If the user is a member of the group 
strGroup = "APPS_Write"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_USA & "APPS"
	Call MapNetworkDrive(strDriveLetter_Apps, strRemotePath)
	
End If

' Map INFORMES network drive - If the user is a member of the group
strGroup = "Informes_US_Read"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_MAD & "Company\Informes\US"
	Call MapNetworkDrive(strDriveLetter_Informes, strRemotePath) 	
	
End If

' Map HHRR network drive - if the user is a member of the group
strGroup = "HHRR_USA_Write"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_MAD & "Company\RRHH"
	Call MapNetworkDrive(strDriveLetter_RRHH, strRemotePath) 	

End If

' Map MIS network drive - if the user is a member of the group
strGroup = "MIS_Write"
If (IsMember(strGroup) = True) Then

	strRemotePath  = strServer_MAD & "Company\MIS"
	Call MapNetworkDrive(strDriveLetter_MIS, strRemotePath) 	

End If

' Quit
WScript.Quit




' -----------------------------------------------------------------' 
' -----------------------------------------------------------------' 
' Global Functions

Sub MapNetworkDrive (strDriveLetter, strRemotePath) 
    ' Subroutine to Map a Network Drive.
    ' strDriveLetter is the drive letter to be used.
    ' strRemotePath is the path to be mapped.

    On Error Resume Next
    objNetwork.MapNetworkDrive strDriveLetter, strRemotePath
    If (Err.Number <> 0) Then
        On Error GoTo 0
		
		'Debug
		'WScript.Echo " Path: " & strRemotePath 

        objNetwork.RemoveNetworkDrive strDriveLetter, True, True
        objNetwork.MapNetworkDrive strDriveLetter, strRemotePath
    End If
    On Error GoTo 0

End Sub

Function IsMember(ByVal strGroup)
    ' Function to test for user group membership.
    ' strGroup is the NT name (sAMAccountName) of the group to test.
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
