On Error Resume Next
 
Const ADS_SCOPE_SUBTREE = 2
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
 
strContainer = "DC=domain,DC=local"
intMaxPwAge = 106
 
strEmailFrom = "josea.munoz@gmail.com"
strEmailSubject = "Password Expiration / Caducidad de Contraseña / Vencimento de senha"
strSMTP = "dc-server.local"
 
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
 
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
 
objCommand.CommandText = _
    "SELECT AdsPath FROM 'LDAP://" & strContainer & "' WHERE objectCategory='user'"  
Set objRecordSet = objCommand.Execute
 
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    Set objUser = GetObject(objRecordSet.Fields("AdsPath").Value)
    
    whenPasswordExpires = DateDiff("d", objUser.PasswordLastChanged, Now)
		whenPasswordExpires = (120 - whenPasswordExpires)
    'WScript.Echo whenPasswordExpires
    
		strEmailBody = "Warning: your Company domain password will expire in " & whenPasswordExpires & " days." & vbcrlf
		strEmailBody = strEmailBody & "Please change it ASAP to avoid being locked out of the system." & vbcrlf
		strEmailBody = strEmailBody & "Kind Regards" & vbcrlf
		strEmailBody = strEmailBody & " " & vbcrlf
		strEmailBody = strEmailBody & "===============================================" & vbcrlf
		strEmailBody = strEmailBody & " " & vbcrlf
		strEmailBody = strEmailBody & "AtenciÛn: su contraseña del dominio Company caducará en " & whenPasswordExpires & " dÌas." & vbcrlf
		strEmailBody = strEmailBody & "Por favor, cámbiela para evitar problemas de acceso al sistema." & vbcrlf
		strEmailBody = strEmailBody & "Reciba un cordial saludo" & vbcrlf
		strEmailBody = strEmailBody & " " & vbcrlf
		strEmailBody = strEmailBody & "===============================================" & vbcrlf
		strEmailBody = strEmailBody & " " & vbcrlf
		strEmailBody = strEmailBody & "Atençao: a sua senha de acesso ao dominio Company vencera em " & whenPasswordExpires & " días." & vbcrlf
		strEmailBody = strEmailBody & "Por favor, alterar a mesma para evitar problemas de acesso ao sistema" & vbcrlf
		strEmailBody = strEmailBody & "Atencioamente"

    If (objUser.userAccountControl) <> "" Then
    
	    If (objUser.userAccountControl And ADS_UF_DONT_EXPIRE_PASSWD) <> 0 Then
						'WScript.Echo "Never Expires"
	    Else
	    	
				If DateDiff("d", objUser.PasswordLastChanged, Now) > intMaxPwAge Then
		        Emailer objUser.mail
		    End If
	    
	    End If
	    
	  End If
            
    objRecordSet.MoveNext
Loop
 
Sub Emailer(strEmailTo)
    Set objEmail = CreateObject("CDO.Message")
 
    objEmail.From = strEmailFrom
    objEmail.To = strEmailTo
    objEmail.Subject = strEmailSubject
    objEmail.Textbody = strEmailBody
    objEmail.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objEmail.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
    objEmail.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objEmail.Configuration.Fields.Update
 
    objEmail.Send
End Sub