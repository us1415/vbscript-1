On Error Resume Next
Dim allComp(500,4)
Dim arrStr, count2

' ForAppending = 8 ForReading = 1, ForWriting = 2
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Function EnvString(variable)
    set objShell = WScript.CreateObject( "WScript.Shell" )
    variable = "%" & variable & "%"
    EnvString = objShell.ExpandEnvironmentStrings(variable)    
    Set objShell = Nothing    
End Function

'Get Enviromental Vars
user=lcase(EnvString("username"))
comp=lcase(EnvString("ComputerName"))
domain=lcase(EnvString("UserDomain"))

'Edit this to point to a shared drive where all users have write access
strDirectory = "C:\"
'Edit this to point to a shared drive where all users have write access

strFile = "\" & domain & ".txt" 'Outputs to "Domain Name.txt" / Use 


strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Wow6432Node\TeamViewer\Version6"
strValueName = "ClientID"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
if len(dwValue) >= 9 then strText2 = domain & "," & comp & "," & dwValue & "," & user end if

strKeyPath = "SOFTWARE\TeamViewer\Version6"
strValueName = "ClientID"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
if len(dwValue) >= 9 then strText2 = domain & "," & comp & "," & dwValue & "," & user end if

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strDirectory) Then
   Set objFolder = objFSO.GetFolder(strDirectory)
   'wscript.echo("Folder Exists!")
Else
	Set objFolder = objFSO.CreateFolder(strDirectory)  
	If not objFSO.FolderExists(strDirectory) Then
		wscript.quit
	End If
	'wscript.echo("Folder Does Not Exist!")
End If

If objFSO.FileExists(strDirectory & strFile) Then
	'wscript.echo("File Exists!")
	Set objFolder = objFSO.GetFolder(strDirectory)
	Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile)
	count = 0
	Do while NOT objTextFile.AtEndOfStream
		arrStr = split(objTextFile.ReadLine,",")
		'wscript.echo(Count & "," & arrStr)
		if arrStr(1) <> comp Then
			count = count + 1
			allcomp(count,0)=arrStr(0) 'Domain / Company
			allcomp(count,1)=arrStr(1) 'Computer Name
			allcomp(count,2)=arrStr(2) 'Teamviewer ID
			allcomp(count,3)=arrStr(3) 'Username
		End If
	Loop
		
		'wscript.echo("EOF!")
		objTextFile.Close
		set objTextFile = nothing
		
		objFSO.DeleteFile(strDirectory & strFile)
		
		Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
		set objFile = nothing
		
		Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, ForAppending, True)
		count2 = Count
		'wscript.echo(count2)
		for count = 1 to count2
			strtext=allcomp(count,0)&","&allcomp(count,1)&","&allcomp(count,2)&","&allcomp(count,3)
			objTextFile.WriteLine(strText)
			'wscript.echo(strText)
		Next
		objTextFile.WriteLine(strText2)
		'wscript.echo(strText2)
		objTextFile.Close
		set objTextFile = nothing
Else
	Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
	set objFile = nothing
	Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, ForAppending, True)
	objTextFile.WriteLine(strText2)
	objTextFile.Close
	set objTextFile = nothing
End If

set objFolder = nothing
set objFSO = Nothing

wscript.quit